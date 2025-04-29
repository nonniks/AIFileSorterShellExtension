using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Linq;
using System.IO;
using System.Threading.Tasks;
using System.Reflection;
using System.Configuration;
using SharpShell.Attributes;
using SharpShell.SharpContextMenu;

namespace AIFileSorterShellExtension
{
    [ComVisible(true)]
    [Guid("0702a7bb-5000-58ca-256a-7306f3614924")] // Explicitly define GUID
    [COMServerAssociation(AssociationType.Directory)]
    [COMServerAssociation(AssociationType.DirectoryBackground)] // Add support for right-clicking on empty space
    public class AISorterShellExtension : SharpContextMenu
    {
        // Write diagnostic info on load
        static AISorterShellExtension()
        {
            // Добавляем обработчик разрешения сборок
            AppDomain.CurrentDomain.AssemblyResolve += (sender, args) => {
                try {
                    // Логируем запрос на сборку
                    string logDir = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                        "AIFileSorter", "diagnostics"
                    );
                    
                    if (!Directory.Exists(logDir)) {
                        Directory.CreateDirectory(logDir);
                    }
                    
                    File.AppendAllText(
                        Path.Combine(logDir, "assembly_resolve.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Requested assembly: {args.Name}\n"
                    );
                    
                    // Проверяем запрашиваемую сборку
                    var requestedAssembly = new System.Reflection.AssemblyName(args.Name);
                    if (requestedAssembly.Name == "System.Runtime.CompilerServices.Unsafe" && 
                        requestedAssembly.Version.Major == 4 && 
                        requestedAssembly.Version.Minor == 0) 
                    {
                        // Ищем эту сборку в папке рядом с нашей DLL
                        string assemblyPath = Path.Combine(
                            Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location),
                            "System.Runtime.CompilerServices.Unsafe.dll"
                        );
                        
                        File.AppendAllText(
                            Path.Combine(logDir, "assembly_resolve.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Trying to load from: {assemblyPath}\n"
                        );
                        
                        if (File.Exists(assemblyPath)) {
                            return Assembly.LoadFrom(assemblyPath);
                        }
                    }
                    
                    return null;
                }
                catch {
                    return null;
                }
            };
            
            try
            {
                // Create a diagnostic file to confirm the extension is loaded
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                if (!Directory.Exists(diagDir))
                {
                    Directory.CreateDirectory(diagDir);
                }
                
                // Create a file that shows the extension was loaded
                File.WriteAllText(
                    Path.Combine(diagDir, $"extension_loaded_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                    $"Assembly version: {Assembly.GetExecutingAssembly().GetName().Version}\n" +
                    $"Loading time: {DateTime.Now}\n" +
                    $"Current user: {Environment.UserName}\n" +
                    $"OS version: {Environment.OSVersion}"
                );
            }
            catch
            {
                // Ignore errors in diagnostic code
            }
        }

        private string _apiKey;
        
        // Add static fields at the class level to track sort history
        private static DateTime _lastSortTime = DateTime.MinValue;
        private static List<Dictionary<string, string>> _lastSortOperations = new List<Dictionary<string, string>>();
        private static string _lastSortedFolder = null;
        private const int UndoTimeWindowMinutes = 2; // 2-minute window for undo

        protected override bool CanShowMenu()
        {
            try
            {
                // Log diagnostic information
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                if (!Directory.Exists(diagDir))
                {
                    Directory.CreateDirectory(diagDir);
                }
                
                // Log each time CanShowMenu is called
                File.AppendAllText(
                    Path.Combine(diagDir, "canshowmenu_calls.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - CanShowMenu called. SelectedItemPaths: {SelectedItemPaths?.Count() ?? 0}\n"
                );
                
                // Show the menu for all directories
                return true;
            }
            catch
            {
                // If logging fails, still try to show the menu
                return true;
            }
        }

        protected override ContextMenuStrip CreateMenu()
        {
            try
            {
                // Log menu creation
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                if (!Directory.Exists(diagDir))
                {
                    Directory.CreateDirectory(diagDir);
                }
                
                File.AppendAllText(
                    Path.Combine(diagDir, "menu_creation.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - CreateMenu called\n"
                );
                
                // Create the menu strip
                var menu = new ContextMenuStrip();

                // Create a simplified menu with just one main option and settings
                var sortFilesItem = new ToolStripMenuItem
                {
                    Text = "Sort Files",
                };
                
                // Add undo sort option
                var undoSortItem = new ToolStripMenuItem
                {
                    Text = "Undo Last Sort",
                    Enabled = CanUndoLastSort() // Only enable if undo is available
                };
                
                // Загрузка пользовательской иконки
                try
                {
                    // Если иконка добавлена как "Внедренный ресурс" 
                    string resourceName = "AIFileSorterShellExtension.generative.png";
                    using (var stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        if (stream != null)
                        {
                            sortFilesItem.Image = System.Drawing.Image.FromStream(stream);
                        }
                        else
                        {
                            // Пробуем загрузить как внешний файл
                            string iconPath = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "generative.png");
                            if (File.Exists(iconPath))
                            {
                                sortFilesItem.Image = System.Drawing.Image.FromFile(iconPath);
                            }
                            else
                            {
                                sortFilesItem.Image = System.Drawing.SystemIcons.Application.ToBitmap();
                            }
                        }
                    }
                }
                catch
                {
                    // В случае ошибки загрузки используем стандартную иконку
                    sortFilesItem.Image = System.Drawing.SystemIcons.Application.ToBitmap();
                }
                
                var settingsItem = new ToolStripMenuItem
                {
                    Text = "Settings...",
                };

                // Add event handlers with explicit delegate methods for debugging
                sortFilesItem.Click += OnSortFilesClicked;
                undoSortItem.Click += OnUndoSortClicked; // Add handler for undo option
                settingsItem.Click += OnSettingsClicked;

                // Add items to the menu
                menu.Items.Add(sortFilesItem);
                
                // Only add undo option if it's available
                if (CanUndoLastSort())
                {
                    menu.Items.Add(undoSortItem);
                }
                
                menu.Items.Add(new ToolStripSeparator());
                menu.Items.Add(settingsItem);

                return menu;
            }
            catch (Exception ex)
            {
                LogError(ex);
                // Fallback to minimal menu in case of error
                var menu = new ContextMenuStrip();
                var errorItem = new ToolStripMenuItem
                {
                    Text = "Error creating menu",
                };
                menu.Items.Add(errorItem);
                return menu;
            }
        }
        
        // Check if undo is available (within time window and has operations)
        private bool CanUndoLastSort()
        {
            // Check if within 2-minute window
            bool timeWindowValid = (DateTime.Now - _lastSortTime).TotalMinutes <= UndoTimeWindowMinutes;
            
            // Check if there are operations to undo
            bool hasOperations = _lastSortOperations != null && _lastSortOperations.Count > 0;
            
            // Check if folder path is valid
            bool hasFolderPath = !string.IsNullOrEmpty(_lastSortedFolder);
            
            return timeWindowValid && hasOperations && hasFolderPath;
        }

        // Handler for the undo button
        private void OnUndoSortClicked(object sender, EventArgs args)
        {
            try
            {
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                if (!Directory.Exists(diagDir))
                {
                    Directory.CreateDirectory(diagDir);
                }
                
                File.AppendAllText(
                    Path.Combine(diagDir, "click_events.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Undo Sort clicked\n"
                );
                
                if (!CanUndoLastSort())
                {
                    MessageBox.Show("Cannot undo the sort operation. Time window (2 minutes) has expired or no sort was performed.", 
                                  "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                
                // Use wait cursor for undo operation
                using (CursorManager.ShowWaitCursor())
                {
                    UndoLastSort();
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error undoing sort: {ex.Message}", "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CursorManager.RestoreDefaultCursor(); // Just in case
            }
        }

        // Method to undo the last sort
        private void UndoLastSort()
        {
            string diagDir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AIFileSorter", "diagnostics"
            );
            
            File.AppendAllText(
                Path.Combine(diagDir, "undo_operations.log"),
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Starting undo for {_lastSortOperations.Count} operations\n"
            );
            
            int successCount = 0;
            int failCount = 0;
            
            // Process operations in reverse order to handle nested folders correctly
            foreach (var op in _lastSortOperations.AsEnumerable().Reverse())
            {
                try
                {
                    string source = op["destination"]; // Current location (destination during the sort)
                    string destination = op["source"]; // Original location (source during the sort)
                    string itemType = op["type"];      // File or folder
                    
                    // Check if item exists at the current location
                    bool exists = itemType == "file" ? 
                        File.Exists(source) : 
                        Directory.Exists(source);
                    
                    if (!exists)
                    {
                        File.AppendAllText(
                            Path.Combine(diagDir, "undo_operations.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Item not found at current location: {source}\n"
                        );
                        failCount++;
                        continue;
                    }
                    
                    // Create destination directory if needed
                    string destDir = Path.GetDirectoryName(destination);
                    if (!Directory.Exists(destDir))
                    {
                        Directory.CreateDirectory(destDir);
                    }
                    
                    // Move the item back to its original location
                    if (itemType == "file")
                    {
                        File.Move(source, destination);
                    }
                    else
                    {
                        Directory.Move(source, destination);
                    }
                    
                    File.AppendAllText(
                        Path.Combine(diagDir, "undo_operations.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Successfully moved {itemType} back: {source} -> {destination}\n"
                    );
                    
                    successCount++;
                }
                catch (Exception ex)
                {
                    File.AppendAllText(
                        Path.Combine(diagDir, "undo_operations.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Error during undo: {ex.Message}\n"
                    );
                    failCount++;
                }
            }
            
            // Clear history after undo
            _lastSortOperations.Clear();
            _lastSortTime = DateTime.MinValue;
            _lastSortedFolder = null;
            
            File.AppendAllText(
                Path.Combine(diagDir, "undo_operations.log"),
                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Undo completed. Success: {successCount}, Fail: {failCount}\n"
            );
            
            MessageBox.Show($"Undo completed. {successCount} items were successfully restored.", 
                           "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        // Modify OnSortFilesClicked to update sorted time
        private void OnSortFilesClicked(object sender, EventArgs args)
        {
            try
            {
                // Log click event
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                File.AppendAllText(
                    Path.Combine(diagDir, "click_events.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Sort Files clicked\n"
                );
                
                // Используем настройку web search
                bool useWebSearch = LoadUseWebSearch();
                
                // Используем новый менеджер курсора, который заменяет системный курсор
                using (CursorManager.ShowWaitCursor())
                {
                    // Set the sort time at the start of the operation
                    _lastSortTime = DateTime.Now;
                    
                    // Вызываем синхронную обертку асинхронного метода
                    Task.Run(async () => 
                    {
                        var result = await SortFolderWithAIAsync(useWebSearch);
                        // If we have a sorted folder path and operations, store them for undo
                        if (result.Item1 && result.Item2 != null && result.Item2.Count > 0)
                        {
                            _lastSortOperations = result.Item2;
                        }
                        else
                        {
                            // Clear history if sort was unsuccessful
                            _lastSortTime = DateTime.MinValue;
                            _lastSortOperations.Clear();
                        }
                    }).GetAwaiter().GetResult();
                }
                
                // Курсор автоматически восстановится благодаря IDisposable
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error starting sort: {ex.Message}", "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CursorManager.RestoreDefaultCursor(); // На всякий случай
            }
        }

        private void OnSettingsClicked(object sender, EventArgs args)
        {
            try
            {
                // Log click event
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                File.AppendAllText(
                    Path.Combine(diagDir, "click_events.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Settings clicked\n"
                );
                
                ShowSettings();
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error opening settings: {ex.Message}", "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Новый асинхронный метод
        private async Task<Tuple<bool, List<Dictionary<string, string>>>> SortFolderWithAIAsync(bool useWebSearch)
        {
            List<Dictionary<string, string>> operations = new List<Dictionary<string, string>>();
            
            try
            {
                // Копируем всю логику из SortFolderWithAI
                string diagDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "diagnostics"
                );
                
                if (!Directory.Exists(diagDir))
                {
                    Directory.CreateDirectory(diagDir);
                }
                
                File.AppendAllText(
                    Path.Combine(diagDir, "sorting_process.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - SortFolderWithAIAsync started, useWebSearch={useWebSearch}\n"
                );

                // Определяем целевую папку
                string folderPath = null;
                
                // Проверяем если есть выбранная папка
                if (SelectedItemPaths != null && SelectedItemPaths.Any())
                {
                    folderPath = SelectedItemPaths.First();
                    File.AppendAllText(
                        Path.Combine(diagDir, "sorting_process.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Selected folder path: {folderPath}\n"
                    );
                }
                else
                {
                    // Если папка не выбрана (клик по фону папки), получаем текущий каталог проводника
                    // Пробуем получить каталог из FolderPath свойства, если оно есть
                    try
                    {
                        var folderPathProperty = this.GetType().GetProperty("FolderPath", 
                            System.Reflection.BindingFlags.NonPublic | 
                            System.Reflection.BindingFlags.Instance | 
                            System.Reflection.BindingFlags.Public);
                        
                        if (folderPathProperty != null)
                        {
                            folderPath = folderPathProperty.GetValue(this) as string;
                            
                            File.AppendAllText(
                                Path.Combine(diagDir, "sorting_process.log"),
                                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Got folder from FolderPath property: {folderPath}\n"
                            );
                        }
                    }
                    catch (Exception ex)
                    {
                        File.AppendAllText(
                            Path.Combine(diagDir, "sorting_process.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Error getting FolderPath: {ex.Message}\n"
                        );
                    }
                    
                    // Если всё еще нет пути, пробуем получить путь текущей директории из Shell
                    if (string.IsNullOrEmpty(folderPath))
                    {
                        try 
                        {
                            // Попытка получить путь текущей директории через Windows API
                            IntPtr shellWindow = Win32Helper.GetShellWindow();
                            if (shellWindow != IntPtr.Zero)
                            {
                                folderPath = Win32Helper.GetCurrentDirectoryFromExplorer(shellWindow);
                                
                                File.AppendAllText(
                                    Path.Combine(diagDir, "sorting_process.log"),
                                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Got folder from Explorer window: {folderPath}\n"
                                );
                            }
                        }
                        catch (Exception ex)
                        {
                            File.AppendAllText(
                                Path.Combine(diagDir, "sorting_process.log"),
                                $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Error getting current directory: {ex.Message}\n"
                            );
                        }
                    }
                }

                // Проверяем, получили ли мы путь к папке
                if (string.IsNullOrEmpty(folderPath))
                {
                    File.AppendAllText(
                        Path.Combine(diagDir, "sorting_process.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - No folder path found\n"
                    );
                    
                    MessageBox.Show("No folder selected. Please try again by right-clicking directly on a folder.", 
                                  "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
                }

                // Ensure the directory exists
                if (!Directory.Exists(folderPath))
                {
                    File.AppendAllText(
                        Path.Combine(diagDir, "sorting_process.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Directory does not exist: {folderPath}\n"
                    );
                    
                    MessageBox.Show($"Selected directory doesn't exist: {folderPath}", "AI File Sorter", 
                                  MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
                }

                // Load the API key from config or registry
                _apiKey = LoadApiKey();
                
                File.AppendAllText(
                    Path.Combine(diagDir, "sorting_process.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - API key loaded: {(!string.IsNullOrEmpty(_apiKey) ? "Yes" : "No")}\n"
                );
                
                if (string.IsNullOrEmpty(_apiKey))
                {
                    // Try to load from .env file
                    string envPath = Path.Combine(
                        Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                        "AI_Sort_Project", ".env"
                    );
                    
                    File.AppendAllText(
                        Path.Combine(diagDir, "sorting_process.log"),
                        $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Checking for .env file at: {envPath}\n"
                    );
                    
                    if (File.Exists(envPath))
                    {
                        string[] lines = File.ReadAllLines(envPath);
                        foreach (var line in lines)
                        {
                            if (line.StartsWith("OPENROUTER_API_KEY="))
                            {
                                _apiKey = line.Substring("OPENROUTER_API_KEY=".Length).Trim();
                                SaveApiKey(_apiKey); // Save this for future use
                                break;
                            }
                        }
                    }
                    
                    if (string.IsNullOrEmpty(_apiKey))
                    {
                        Application.UseWaitCursor = false;
                        Cursor.Current = Cursors.Default;
                        
                        File.AppendAllText(
                            Path.Combine(diagDir, "sorting_process.log"),
                            $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - No API key found, showing settings\n"
                        );
                        
                        MessageBox.Show("OpenRouter API key is not set. Please set your API key in the settings.",
                                    "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        ShowSettings();
                        return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
                    }
                }

                // Store the folder being sorted
                _lastSortedFolder = folderPath;
                
                // Initialize the sorter service
                File.AppendAllText(
                    Path.Combine(diagDir, "sorting_process.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Initializing AIFileSorterService\n"
                );
                
                var sorterService = new AIFileSorterService(_apiKey);
                
                // Sort the folder with modified return type
                var sortResult = await sorterService.SortFolderAsyncWithOperations(folderPath, useWebSearch);
                bool success = sortResult.Item1;
                operations = sortResult.Item2;

                File.AppendAllText(
                    Path.Combine(diagDir, "sorting_process.log"),
                    $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - Sorting completed, success={success}\n"
                );
                
                if (!success)
                {
                    // Only show errors, successful sorting should be silent
                    MessageBox.Show("An error occurred while sorting files. Please check the log for details.",
                                "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    MessageBox.Show("Files sorted successfully!", "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                
                return new Tuple<bool, List<Dictionary<string, string>>>(success, operations);
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error: {ex.Message}", "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
            }
        }
        
        // Helper method for logging errors consistently
        private void LogError(Exception ex)
        {
            try
            {
                string logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter"
                );
                
                if (!Directory.Exists(logDirectory))
                {
                    Directory.CreateDirectory(logDirectory);
                }
                
                File.WriteAllText(
                    Path.Combine(logDirectory, $"error_{DateTime.Now:yyyyMMdd_HHmmss}.log"),
                    ex.ToString()
                );
            }
            catch
            {
                // Suppress errors in the error logger to prevent cascading exceptions
            }
        }

        private void ShowSettings()
        {
            try
            {
                // Create settings form
                using (var settingsForm = new Form
                {
                    Text = "AI File Sorter Settings",
                    Size = new System.Drawing.Size(450, 250), // Увеличиваем высоту формы
                    FormBorderStyle = FormBorderStyle.FixedDialog,
                    StartPosition = FormStartPosition.CenterScreen,
                    MaximizeBox = false,
                    MinimizeBox = false
                })
                {
                    // Load current API key
                    string apiKey = LoadApiKey() ?? "";
                    bool useWebSearch = LoadUseWebSearch();
                    
                    // Create controls
                    var apiKeyLabel = new Label
                    {
                        Text = "OpenRouter API Key:",
                        Location = new System.Drawing.Point(20, 20),
                        Size = new System.Drawing.Size(150, 20)
                    };
                    
                    var apiKeyTextBox = new TextBox
                    {
                        Text = apiKey,
                        Location = new System.Drawing.Point(20, 50),
                        Size = new System.Drawing.Size(400, 20),
                        PasswordChar = '*'
                    };
                    
                    var showKeyCheckBox = new CheckBox
                    {
                        Text = "Show API Key",
                        Location = new System.Drawing.Point(20, 80),
                        Size = new System.Drawing.Size(150, 20)
                    };
                    
                    // Добавляем чекбокс для Web Search
                    var webSearchCheckBox = new CheckBox
                    {
                        Text = "Enable web search (more accurate categorization)",
                        Location = new System.Drawing.Point(20, 110),
                        Size = new System.Drawing.Size(350, 20),
                        Checked = useWebSearch
                    };
                    
                    var saveButton = new Button
                    {
                        Text = "Save",
                        Location = new System.Drawing.Point(250, 170),
                        Size = new System.Drawing.Size(80, 30),
                        DialogResult = DialogResult.OK
                    };
                    
                    var cancelButton = new Button
                    {
                        Text = "Cancel",
                        Location = new System.Drawing.Point(340, 170),
                        Size = new System.Drawing.Size(80, 30),
                        DialogResult = DialogResult.Cancel
                    };
                    
                    // Wire up events
                    showKeyCheckBox.CheckedChanged += (s, e) =>
                    {
                        apiKeyTextBox.PasswordChar = showKeyCheckBox.Checked ? '\0' : '*';
                    };
                    
                    saveButton.Click += (s, e) =>
                    {
                        SaveApiKey(apiKeyTextBox.Text);
                        SaveUseWebSearch(webSearchCheckBox.Checked);
                        settingsForm.Close();
                    };
                    
                    // Add controls to form
                    settingsForm.Controls.Add(apiKeyLabel);
                    settingsForm.Controls.Add(apiKeyTextBox);
                    settingsForm.Controls.Add(showKeyCheckBox);
                    settingsForm.Controls.Add(webSearchCheckBox);
                    settingsForm.Controls.Add(saveButton);
                    settingsForm.Controls.Add(cancelButton);
                    
                    // Show the form
                    settingsForm.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error showing settings: {ex.Message}", 
                                "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private string LoadApiKey()
        {
            try
            {
                // First try to get from environment variable
                string apiKey = Environment.GetEnvironmentVariable("OPENROUTER_API_KEY");
                if (!string.IsNullOrEmpty(apiKey))
                {
                    return apiKey;
                }
                
                // Try registry
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\AIFileSorter"))
                {
                    if (key != null)
                    {
                        apiKey = key.GetValue("ApiKey") as string;
                        if (!string.IsNullOrEmpty(apiKey))
                        {
                            return apiKey;
                        }
                    }
                }
                
                // Try .env file in Desktop\AI_Sort_Project folder
                string envPath = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                    "AI_Sort_Project", ".env"
                );
                
                if (File.Exists(envPath))
                {
                    string[] lines = File.ReadAllLines(envPath);
                    foreach (var line in lines)
                    {
                        if (line.StartsWith("OPENROUTER_API_KEY="))
                        {
                            return line.Substring("OPENROUTER_API_KEY=".Length).Trim();
                        }
                    }
                }
                
                return null;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return null;
            }
        }

        private void SaveApiKey(string apiKey)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\AIFileSorter"))
                {
                    if (key != null)
                    {
                        key.SetValue("ApiKey", apiKey);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
                MessageBox.Show($"Error saving API key: {ex.Message}", 
                                "AI File Sorter", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // Добавляем методы для хранения настройки web search
        private bool LoadUseWebSearch()
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"SOFTWARE\AIFileSorter"))
                {
                    if (key != null)
                    {
                        object value = key.GetValue("UseWebSearch");
                        if (value != null)
                        {
                            return Convert.ToBoolean(value);
                        }
                    }
                }
                // По умолчанию включаем web search
                return true;
            }
            catch (Exception ex)
            {
                LogError(ex);
                return true; // По умолчанию включено
            }
        }

        private void SaveUseWebSearch(bool useWebSearch)
        {
            try
            {
                using (var key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"SOFTWARE\AIFileSorter"))
                {
                    if (key != null)
                    {
                        key.SetValue("UseWebSearch", useWebSearch ? 1 : 0, Microsoft.Win32.RegistryValueKind.DWord);
                    }
                }
            }
            catch (Exception ex)
            {
                LogError(ex);
            }
        }
    }

    // Вспомогательный класс для работы с Windows API
    internal static class Win32Helper
    {
        [DllImport("user32.dll")]
        public static extern IntPtr GetShellWindow();

        [DllImport("user32.dll", SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        public static extern IntPtr SendMessage(IntPtr hWnd, UInt32 Msg, IntPtr wParam, IntPtr lParam);
        
        // Cursor-related Win32 API functions
        [DllImport("user32.dll")]
        public static extern IntPtr LoadCursor(IntPtr hInstance, int lpCursorName);
        
        [DllImport("user32.dll")]
        public static extern IntPtr SetCursor(IntPtr hCursor);
        
        // SystemParametersInfo to force UI updates
        [DllImport("user32.dll")]
        public static extern bool SystemParametersInfo(uint uiAction, uint uiParam, IntPtr pvParam, uint fWinIni);
        
        // Constants for cursor management
        public const int IDC_ARROW = 32512;   // Standard arrow
        public const int IDC_WAIT = 32514;    // Hourglass/wait
        
        // Constants for SystemParametersInfo
        public const uint SPI_SETCURSORS = 0x0057;
        
        // Получение текущей директории из активного окна проводника
        public static string GetCurrentDirectoryFromExplorer(IntPtr shellWindow)
        {
            try
            {
                GetWindowThreadProcessId(shellWindow, out uint processId);
                
                // Здесь должна быть более сложная логика для получения пути через Explorer COM интерфейсы
                // Это упрощенная версия, которая может не работать во всех случаях
                
                // Как простое решение, используем текущий каталог процесса
                return Environment.CurrentDirectory;
            }
            catch
            {
                return null;
            }
        }
    }
}

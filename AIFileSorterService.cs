using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Linq;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace AIFileSorterShellExtension
{
    public class AIFileSorterService
    {
        private const string OpenRouterApiUrl = "https://openrouter.ai/api/v1/chat/completions";
        private readonly string _apiKey;
        private readonly string _defaultModel = "meta-llama/llama-4-maverick:free";
        private readonly HttpClient _httpClient;
        private readonly string _logFilePath;

        // Create a logger that writes to a file
        private void LogInfo(string message)
        {
            try
            {
                File.AppendAllText(_logFilePath, $"[INFO] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
            }
            catch { /* Ignore logging errors */ }
        }

        private void LogError(string message, Exception ex = null)
        {
            try
            {
                File.AppendAllText(_logFilePath, $"[ERROR] {DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}\n");
                if (ex != null)
                {
                    File.AppendAllText(_logFilePath, $"Exception: {ex}\n");
                }
            }
            catch { /* Ignore logging errors */ }
        }

        public AIFileSorterService(string apiKey)
        {
            _apiKey = apiKey;
            _httpClient = new HttpClient();
            _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", _apiKey);
            _httpClient.DefaultRequestHeaders.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));

            // Set up logging
            string logDirectory = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                "AIFileSorter"
            );

            if (!Directory.Exists(logDirectory))
            {
                Directory.CreateDirectory(logDirectory);
            }

            _logFilePath = Path.Combine(logDirectory, "AIFileSorter.log");
        }

        public async Task<bool> SortFolderAsync(string folderPath, bool useWebSearch = true)
        {
            try
            {
                LogInfo($"Starting folder sort for {folderPath}");

                // Get files and folders to analyze
                var files = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .ToList();

                var folders = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .Where(f => !f.StartsWith(".") && f != "__pycache__")
                    .ToList();

                LogInfo($"Found {files.Count} files and {folders.Count} folders");

                if (files.Count == 0 && folders.Count == 0)
                {
                    LogInfo("No files or folders to sort");
                    return false;
                }

                // Analyze which folders are likely already categorized
                var (categorizedFolders, foldersToSort) = AnalyzeFolderStructure(folders);
                LogInfo($"Found {categorizedFolders.Count} already categorized folders and {foldersToSort.Count} folders to sort");

                // Prepare and send the request to OpenRouter API
                var categorization = await GetCategorizationFromLLMAsync(files, foldersToSort, useWebSearch);
                if (categorization == null)
                {
                    LogError("Failed to get valid categorization from LLM");
                    return false;
                }

                // Move files and folders based on categorization
                MoveFilesAndFolders(categorization, folderPath);

                // Clean up empty folders
                RemoveEmptyFolders(folderPath);

                LogInfo("Folder sorting completed successfully");
                return true;
            }
            catch (Exception ex)
            {
                LogError("Error during folder sorting", ex);
                return false;
            }
        }

        // Add new method that returns operations
        public async Task<Tuple<bool, List<Dictionary<string, string>>>> SortFolderAsyncWithOperations(string folderPath, bool useWebSearch = true)
        {
            List<Dictionary<string, string>> operations = new List<Dictionary<string, string>>();
            
            try
            {
                LogInfo($"Starting folder sort for {folderPath}");

                // Get files and folders to analyze
                var files = Directory.GetFiles(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .ToList();

                var folders = Directory.GetDirectories(folderPath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .Where(f => !f.StartsWith(".") && f != "__pycache__")
                    .ToList();

                LogInfo($"Found {files.Count} files and {folders.Count} folders");

                if (files.Count == 0 && folders.Count == 0)
                {
                    LogInfo("No files or folders to sort");
                    return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
                }

                // Analyze which folders are likely already categorized
                var (categorizedFolders, foldersToSort) = AnalyzeFolderStructure(folders);
                LogInfo($"Found {categorizedFolders.Count} already categorized folders and {foldersToSort.Count} folders to sort");

                // Prepare and send the request to OpenRouter API
                var categorization = await GetCategorizationFromLLMAsync(files, foldersToSort, useWebSearch);
                if (categorization == null)
                {
                    LogError("Failed to get valid categorization from LLM");
                    return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
                }

                // Move files and folders based on categorization
                operations = MoveFilesAndFoldersWithOperations(categorization, folderPath);

                // Clean up empty folders
                RemoveEmptyFolders(folderPath);

                LogInfo("Folder sorting completed successfully");
                return new Tuple<bool, List<Dictionary<string, string>>>(true, operations);
            }
            catch (Exception ex)
            {
                LogError("Error during folder sorting", ex);
                return new Tuple<bool, List<Dictionary<string, string>>>(false, operations);
            }
        }

        private async Task<Dictionary<string, object>> GetCategorizationFromLLMAsync(List<string> files, List<string> folders, bool useWebSearch)
        {
            try
            {
                LogInfo("Preparing API request to OpenRouter");

                // Get folder structure
                var folderStructure = GetFolderStructure(Path.GetDirectoryName(folders.FirstOrDefault() ?? ""));

                // Format folder structure for the prompt
                string folderStructureText = "";
                if (folderStructure.Count > 0)
                {
                    folderStructureText = "Existing folder structure:\n";
                    // Fix: Replace deconstruction with explicit KeyValuePair usage
                    foreach (var pair in folderStructure)
                    {
                        string folderPath = pair.Key;
                        List<string> subfolders = pair.Value;

                        // Fix: Use Count property instead of Count() method
                        string subfoldersText = subfolders.Count > 0 ? string.Join(", ", subfolders) : "No subfolders";
                        folderStructureText += $"- {folderPath} ({subfoldersText})\n";
                    }
                }

                // Prepare the prompt
                string filesStr = files.Count > 0 ? string.Join(", ", files) : "No files to sort";
                string foldersStr = folders.Count > 0 ? string.Join(", ", folders) : "No folders to sort";

                string prompt = @"
                You are a file organization expert. Your task is to create a logical folder hierarchy for the following files and folders:
                
                # INPUT DATA
                Files to categorize:
                " + filesStr + @"
                
                Folders to categorize:
                " + foldersStr + @"
                
                Existing structure:
                " + folderStructureText + @"
                
                # SORTING RULES AND PRIORITIES
                
                ## Core Rules
                1. Focus primarily on file and folder NAMES rather than just extensions - analyze what content they likely contain
                2. Create a logical, intuitive hierarchy that would make sense to a human user
                3. Group related content together (e.g., everything related to one game should go together)
                4. Create a balanced hierarchy - not too deep, not too flat (max 3 levels depth)
                5. ALL files and folders MUST be included in your categorization
                6. CRITICAL: Each file must be assigned to ONLY ONE category - DO NOT duplicate files across categories
                7. Leave no files uncategorized - every file and folder must have a destination
                
                ## Game Files & Mods
                - Game folders should go into ""Games/[Game Name]"" (e.g., ""Minecraft"", ""The Witcher 3"")
                - Game mods must be sorted by specific game they belong to (e.g., ""Games/Skyrim/Mods"")
                - For mods with unclear game association, search Nexus Mods site (nexusmods.com) to identify the correct game
                - Common mod folders contain words like ""mod"", ""addon"", ""patch"", ""texture"", ""skin"", ""dlc""
                - Save files should go into ""Games/[Game Name]/Saves"" directory
                
                ## Archive Analysis
                - Look INSIDE the filename of archives (.zip, .rar, .7z) to determine their category:
                  * Archives containing ""mod"", ""patch"" → likely game mods
                  * Archives containing ""setup"", ""install"" → likely software
                  * Archives with game names → likely game-related content
                  * Archives with version numbers (v1.2, etc.) → often software or updates
                
                ## Torrent Files
                - All .torrent files must go to ""Torrents"" folder with proper subcategories
                - Categorize torrents based on their content type:
                  * Game torrents → ""Torrents/Games""
                  * Movie/TV torrents → ""Torrents/Movies"" or ""Torrents/TV Shows""
                  * Software torrents → ""Torrents/Software""
                  * Music torrents → ""Torrents/Music""
                
                ## Special Folder Handling
                - Development projects go into ""Development"" with subcategories by programming language/framework
                - Existing category folders (like ""Documents"", ""Pictures"", ""Games"") should NOT be moved
                - If you're unsure about a folder's purpose, examine its name carefully for clues
                - Unknown or mixed-content folders should be sorted by the majority of their likely content
                

                # CRITICAL CHECK BEFORE RESPONDING
                Before you submit your answer, verify:
                1. Every file from the input is present EXACTLY ONCE in your output
                2. Every folder from the input is present EXACTLY ONCE in your output
                3. No file or folder is duplicated across multiple categories
                4. All paths use forward slashes (/) not backslashes
                5. All your JSON is valid with proper syntax (double quotes, no trailing commas)

                # RESPONSE FORMAT
                Respond ONLY with a JSON object with these two sections:
                1. ""files"": Maps destination paths to lists of files
                2. ""folders"": Maps destination paths to lists of folders

                Example format:
                {
                  ""files"": {
                    ""Games/Skyrim/Mods"": [""skyrim_texture_pack.zip"", ""better_weapons.rar""],
                    ""Documents/Work"": [""report.docx""],
                    ""Torrents/Games"": [""skyrim.torrent"", ""doom.torrent""]
                  },
                  ""folders"": {
                    ""Games/Minecraft"": [""Minecraft_Server"", ""Minecraft_Mods""],
                    ""Development/Python"": [""python_project""],
                    ""Torrents/Movies"": [""HDmovies_folder""]
                  }
                }

                The JSON must be valid with double quotes around keys and string values.
                Include ALL files and folders exactly once.
                ";

                // Set up the request model
                string model = _defaultModel;
                if (useWebSearch)
                {
                    model += ":online";
                    LogInfo($"Using web search capability with model: {model}");
                }

                // Create request payload
                var requestData = new
                {
                    model = model,
                    messages = new[]
                    {
                        new { role = "user", content = prompt }
                    }
                };

                // Add web search options if needed
                if (useWebSearch)
                {
                    // OpenRouter doesn't directly support adding this in the C# payload structure as-is
                    // We need to manually create the JSON with these options
                    var requestDataWithSearch = new
                    {
                        model = model,
                        messages = new[]
                        {
                            new { role = "user", content = prompt }
                        },
                        web_search_options = new
                        {
                            search_context_size = "medium",
                            max_results = 5
                        }
                    };

                    // Replace System.Text.Json serialization with Newtonsoft.Json
                    string jsonContent = JsonConvert.SerializeObject(requestDataWithSearch);
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    // Send the request
                    LogInfo("Sending API request with web search");
                    var response = await _httpClient.PostAsync(OpenRouterApiUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        LogError($"API request failed: {response.StatusCode}, {errorContent}");
                        return null;
                    }

                    // Process response using Newtonsoft.Json
                    string responseJson = await response.Content.ReadAsStringAsync();
                    return ExtractCategorization(responseJson);
                }
                else
                {
                    // No web search, simpler request
                    var jsonContent = JsonConvert.SerializeObject(requestData);
                    var content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    // Send the request
                    LogInfo("Sending API request without web search");
                    var response = await _httpClient.PostAsync(OpenRouterApiUrl, content);

                    if (!response.IsSuccessStatusCode)
                    {
                        string errorContent = await response.Content.ReadAsStringAsync();
                        LogError($"API request failed: {response.StatusCode}, {errorContent}");
                        return null;
                    }

                    // Process response using Newtonsoft.Json
                    string responseJson = await response.Content.ReadAsStringAsync();
                    return ExtractCategorization(responseJson);
                }
            }
            catch (Exception ex)
            {
                LogError("Error getting categorization from LLM", ex);
                return null;
            }
        }

        private Dictionary<string, object> ExtractCategorization(string responseJson)
        {
            try
            {
                // Use JObject instead of JsonDocument
                JObject jsonResponse = JObject.Parse(responseJson);
                
                // Extract content from response
                string content = jsonResponse["choices"][0]["message"]["content"].ToString();
                
                LogInfo("Received response from API, extracting JSON content");

                // Save the raw response for debugging
                string debugDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "debug"
                );

                if (!Directory.Exists(debugDir))
                {
                    Directory.CreateDirectory(debugDir);
                }

                File.WriteAllText(
                    Path.Combine(debugDir, $"llm_response_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                    content
                );

                // Extract JSON data from the content
                return ExtractJsonFromText(content);
            }
            catch (Exception ex)
            {
                LogError("Error extracting categorization from response", ex);
                return null;
            }
        }

        private Dictionary<string, object> ExtractJsonFromText(string text)
        {
            try
            {
                // Save raw response for debugging (without changing existing logs)
                string debugDir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "debug"
                );

                if (!Directory.Exists(debugDir))
                {
                    Directory.CreateDirectory(debugDir);
                }

                File.WriteAllText(
                    Path.Combine(debugDir, $"raw_llm_response_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                    text
                );

                // Try to extract JSON object using improved regex patterns
                
                // First search for JSON with 'files' and 'folders' keys - most accurate search
                Regex jsonWithKeysRegex = new Regex(@"\{(?:\s*""[^""]*""\s*:(?:[^{}]|(?<o>\{)|(?<-o>\}))*(?(o)(?!))){2,}\s*\}", 
                    RegexOptions.Singleline);
                Match matchWithKeys = jsonWithKeysRegex.Match(text);

                string jsonStr;
                if (matchWithKeys.Success)
                {
                    jsonStr = matchWithKeys.Value;
                }
                else 
                {
                    // More general search - any JSON object
                    Regex jsonRegex = new Regex(@"\{[\s\S]*\}");
                    Match match = jsonRegex.Match(text);

                    if (match.Success) 
                    {
                        jsonStr = match.Value;
                    }
                    else 
                    {
                        // Write original text without changing logs
                        return null;
                    }
                }

                // Save extracted JSON for debugging
                File.WriteAllText(
                    Path.Combine(debugDir, $"extracted_json_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                    jsonStr
                );

                try
                {
                    // Use JsonConvert instead of JsonSerializer
                    var result = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonStr);
                    
                    // Check for presence of required keys
                    if (result != null && result.ContainsKey("files") && result.ContainsKey("folders"))
                    {
                        return result;
                    }
                    
                    // If required keys are missing, try to fix JSON
                }
                catch
                {
                    // If not able to parse immediately, try to fix common issues
                }

                // Multi-stage JSON correction
                
                // Fix single quotes to double quotes for keys
                jsonStr = Regex.Replace(jsonStr, @"'([^']*)':", @"""$1"":");
                
                // Fix single quotes to double quotes for values
                jsonStr = Regex.Replace(jsonStr, @":\s*'([^']*)'", @": ""$1""");
                
                // Fix quote escaping problems
                jsonStr = Regex.Replace(jsonStr, @"\\""", @"""");
                jsonStr = Regex.Replace(jsonStr, @"\\\\", @"\");
                
                // Remove commas before closing brackets
                jsonStr = Regex.Replace(jsonStr, @",\s*([\]}])", "$1");
                
                // Add quotes to keys without quotes
                jsonStr = Regex.Replace(jsonStr, @"([{,])\s*([a-zA-Z0-9_]+)\s*:", "$1\"$2\":");

                // Save fixed JSON
                File.WriteAllText(
                    Path.Combine(debugDir, $"fixed_json_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                    jsonStr
                );

                try
                {
                    var result = JsonConvert.DeserializeObject<Dictionary<string, object>>(jsonStr);
                    if (result != null && (result.ContainsKey("files") || result.ContainsKey("folders")))
                    {
                        // Ensure both keys are present
                        if (!result.ContainsKey("files")) result["files"] = new JObject();
                        if (!result.ContainsKey("folders")) result["folders"] = new JObject();
                        
                        return result;
                    }
                }
                catch
                {
                    // If still not working after all fixes, try to assemble JSON from fragments
                    
                    // Search for parts of JSON objects
                    Regex filesKeyRegex = new Regex(@"""files""\s*:\s*\{[^{}]*(((?'Open'\{)[^{}]*)+((?'Close-Open'\})[^{}]*)+)*(?(Open)(?!))\}", RegexOptions.Singleline);
                    Regex foldersKeyRegex = new Regex(@"""folders""\s*:\s*\{[^{}]*(((?'Open'\{)[^{}]*)+((?'Close-Open'\})[^{}]*)+)*(?(Open)(?!))\}", RegexOptions.Singleline);
                    
                    Match filesMatch = filesKeyRegex.Match(text);
                    Match foldersMatch = foldersKeyRegex.Match(text);
                    
                    if (filesMatch.Success || foldersMatch.Success)
                    {
                        var manualJson = "{\n";
                        if (filesMatch.Success) 
                        {
                            manualJson += filesMatch.Value + ",\n";
                        }
                        else
                        {
                            manualJson += "\"files\": {},\n";
                        }
                        
                        if (foldersMatch.Success)
                        {
                            manualJson += foldersMatch.Value + "\n";
                        }
                        else
                        {
                            manualJson += "\"folders\": {}\n";
                        }
                        manualJson += "}";
                        
                        // Save manually assembled JSON
                        File.WriteAllText(
                            Path.Combine(debugDir, $"manual_json_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                            manualJson
                        );
                        
                        try 
                        {
                            var result = JsonConvert.DeserializeObject<Dictionary<string, object>>(manualJson);
                            if (result != null)
                            {
                                return result;
                            }
                        }
                        catch 
                        {
                            // Failed even with manual assembly
                        }
                    }
                    
                    // Last resort - create minimal working structure
                    var emptyStructure = new Dictionary<string, object>
                    {
                        ["files"] = new JObject(),
                        ["folders"] = new JObject()
                    };
                    
                    // Save empty JSON structure for debugging
                    File.WriteAllText(
                        Path.Combine(debugDir, $"empty_json_{DateTime.Now:yyyyMMdd_HHmmss}.txt"),
                        JsonConvert.SerializeObject(emptyStructure)
                    );
                    
                    return emptyStructure;
                }

                return null;
            }
            catch (Exception ex)
            {
                // This log was already in the original code, don't change it
                LogError("Error extracting JSON from text", ex);
                return null;
            }
        }

        private (List<string> Categorized, List<string> Uncategorized) AnalyzeFolderStructure(List<string> folders)
        {
            // Common category folder names
            HashSet<string> categoryFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Documents", "Music", "Pictures", "Videos", "Games", "Applications", "Software",
                "Development", "Projects", "Work", "Personal", "Archives", "Torrents", "Torrent",
                "Books", "PDFs", "Install Files", "Programs", "Installers", "Temp", "Images"
            };

            List<string> alreadyCategorized = new List<string>();
            List<string> needSorting = new List<string>();

            foreach (string folder in folders)
            {
                if (categoryFolders.Contains(folder) ||
                    categoryFolders.Any(cat => folder.StartsWith(cat, StringComparison.OrdinalIgnoreCase)))
                {
                    alreadyCategorized.Add(folder);
                }
                else
                {
                    needSorting.Add(folder);
                }
            }

            return (alreadyCategorized, needSorting);
        }

        private Dictionary<string, List<string>> GetFolderStructure(string directory, int maxDepth = 3)
        {
            Dictionary<string, List<string>> structure = new Dictionary<string, List<string>>();

            void ExploreFolder(string folderPath, int currentDepth = 0, string relativePath = "")
            {
                if (currentDepth > maxDepth) return;

                try
                {
                    // Path validation and exception protection
                    if (string.IsNullOrEmpty(folderPath) || !Directory.Exists(folderPath))
                    {
                        LogInfo($"Skipping invalid or non-existent folder: {folderPath}");
                        return;
                    }
                    
                    // Safe subdirectory retrieval with error handling
                    string[] subdirectories;
                    try
                    {
                        subdirectories = Directory.GetDirectories(folderPath);
                    }
                    catch (Exception ex)
                    {
                        LogInfo($"Error accessing subdirectories in {folderPath}: {ex.Message}");
                        subdirectories = new string[0];
                    }

                    foreach (var subdir in subdirectories)
                    {
                        string dirName = Path.GetFileName(subdir);
                        string newRelativePath = string.IsNullOrEmpty(relativePath) ? dirName : Path.Combine(relativePath, dirName);

                        // Protection from cyclic references and other path issues
                        try
                        {
                            if (FolderNameMatchesCategory(dirName))
                            {
                                // Add path to result...
                            }
                            else
                            {
                                ExploreFolder(subdir, currentDepth + 1, newRelativePath);
                            }
                        }
                        catch (Exception ex)
                        {
                            LogInfo($"Error processing subdirectory {subdir}: {ex.Message}");
                        }
                    }

                    // Process files...
                }
                catch (Exception ex)
                {
                    LogError($"Error exploring folder {folderPath}", ex);
                }
            }

            ExploreFolder(directory);
            return structure;
        }

        // Добавьте недостающий метод для проверки соответствия имени папки категории
        private bool FolderNameMatchesCategory(string folderName)
        {
            // Общие названия категорий папок
            HashSet<string> categoryFolders = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                "Documents", "Music", "Pictures", "Videos", "Games", "Applications", "Software",
                "Development", "Projects", "Work", "Personal", "Archives", "Torrents", "Torrent",
                "Books", "PDFs", "Install Files", "Programs", "Installers", "Temp", "Images"
            };
            
            return categoryFolders.Contains(folderName);
        }

        private void MoveFilesAndFolders(Dictionary<string, object> categorization, string basePath)
        {
            try
            {
                // Create history folder for undo operations
                string historyFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "history"
                );

                if (!Directory.Exists(historyFolder))
                {
                    Directory.CreateDirectory(historyFolder);
                }

                List<Dictionary<string, string>> moveRecord = new List<Dictionary<string, string>>();
                
                // Track moved files and folders to avoid duplication
                HashSet<string> movedFiles = new HashSet<string>();
                HashSet<string> movedFolders = new HashSet<string>();
                
                // Check for duplicate files in model response
                HashSet<string> allFilesToMove = new HashSet<string>();
                bool hasDuplicates = false;
                
                // Check for duplicate files in model response
                if (categorization.TryGetValue("files", out object filesObj) && filesObj is JObject filesObject)
                {
                    foreach (var prop in filesObject.Properties())
                    {
                        JArray filesArray = prop.Value as JArray;
                        if (filesArray != null)
                        {
                            foreach (var fileToken in filesArray)
                            {
                                string fileName = fileToken.ToString();
                                if (!allFilesToMove.Add(fileName))
                                {
                                    LogError($"Duplicate file detected in model response: {fileName} in category {prop.Name}");
                                    hasDuplicates = true;
                                }
                            }
                        }
                    }
                }
                
                if (hasDuplicates)
                {
                    LogError("Model returned duplicate files across categories. Some files may not be moved correctly.");
                }

                // Process files
                if (categorization.TryGetValue("files", out filesObj) && filesObj is JObject filesObject2)
                {
                    foreach (var prop in filesObject2.Properties())
                    {
                        string folderPath = prop.Name;

                        // Create target folder if it doesn't exist
                        string fullFolderPath = Path.Combine(basePath, folderPath);
                        if (!Directory.Exists(fullFolderPath))
                        {
                            Directory.CreateDirectory(fullFolderPath);
                            LogInfo($"Created folder: {folderPath}");
                        }

                        // Move each file to folder
                        JArray filesArray = prop.Value as JArray;
                        if (filesArray != null)
                        {
                            foreach (var fileToken in filesArray)
                            {
                                string fileName = fileToken.ToString();
                                string source = Path.Combine(basePath, fileName);
                                string destination = Path.Combine(fullFolderPath, fileName);

                                // Check if file has already been moved
                                if (movedFiles.Contains(fileName))
                                {
                                    LogInfo($"Skipping already moved file: {fileName}");
                                    continue;
                                }

                                if (File.Exists(source))
                                {
                                    try
                                    {
                                        File.Move(source, destination);
                                        movedFiles.Add(fileName);
                                        LogInfo($"Moved file: {fileName} to {folderPath}");

                                        // Record the move
                                        moveRecord.Add(new Dictionary<string, string>
                                        {
                                            ["source"] = source,
                                            ["destination"] = destination,
                                            ["item"] = fileName,
                                            ["target_folder"] = folderPath,
                                            ["type"] = "file"
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError($"Error moving file {fileName} to {folderPath}", ex);
                                    }
                                }
                                else
                                {
                                    LogInfo($"File not found: {source}");
                                }
                            }
                        }
                    }
                }

                // Process folders similarly
                if (categorization.TryGetValue("folders", out object foldersObj))
                {
                    if (foldersObj is JObject foldersObject)
                    {
                        foreach (var prop in foldersObject.Properties())
                        {
                            string folderPath = prop.Name;

                            // Create target folder if it doesn't exist
                            string fullFolderPath = Path.Combine(basePath, folderPath);
                            if (!Directory.Exists(fullFolderPath))
                            {
                                Directory.CreateDirectory(fullFolderPath);
                                LogInfo($"Created folder: {folderPath}");
                            }

                            // Move each folder to destination
                            JArray foldersArray = prop.Value as JArray;
                            if (foldersArray != null)
                            {
                                foreach (var folderToken in foldersArray)
                                {
                                    string folderName = folderToken.ToString();
                                    string source = Path.Combine(basePath, folderName);
                                    string destination = Path.Combine(fullFolderPath, folderName);

                                    // Check if folder has already been moved
                                    if (movedFolders.Contains(folderName))
                                    {
                                        LogInfo($"Skipping already moved folder: {folderName}");
                                        continue;
                                    }

                                    if (Directory.Exists(source))
                                    {
                                        try
                                        {
                                            // Check if destination exists
                                            if (Directory.Exists(destination))
                                            {
                                                string folderNameBase = Path.GetFileName(destination);
                                                string baseDestination = Path.GetDirectoryName(destination);
                                                int suffix = 1;

                                                while (Directory.Exists(destination))
                                                {
                                                    string newFolderName = $"{folderNameBase}_{suffix}";
                                                    destination = Path.Combine(baseDestination, newFolderName);
                                                    suffix++;
                                                }

                                                LogInfo($"Renamed folder to avoid conflict: {folderName} → {Path.GetFileName(destination)}");
                                            }

                                            Directory.Move(source, destination);
                                            movedFolders.Add(folderName);
                                            LogInfo($"Moved folder: {folderName} to {folderPath}");

                                            // Record the move
                                            moveRecord.Add(new Dictionary<string, string>
                                            {
                                                ["source"] = source,
                                                ["destination"] = destination,
                                                ["item"] = folderName,
                                                ["target_folder"] = folderPath,
                                                ["type"] = "folder"
                                            });
                                        }
                                        catch (Exception ex)
                                        {
                                            LogError($"Error moving folder {folderName} to {folderPath}", ex);
                                        }
                                    }
                                    else
                                    {
                                        LogInfo($"Folder not found: {source}");
                                    }
                                }
                            }
                        }
                    }
                }

                // Save the move record for potential undo
                if (moveRecord.Count > 0)
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string historyFile = Path.Combine(historyFolder, $"sort_history_{timestamp}.json");

                    var historyData = new Dictionary<string, object>
                    {
                        ["timestamp"] = timestamp,
                        ["downloads_folder"] = basePath,
                        ["moves"] = moveRecord
                    };

                    // Используем JsonConvert вместо JsonSerializer
                    string historyJson = JsonConvert.SerializeObject(historyData, Formatting.Indented);
                    File.WriteAllText(historyFile, historyJson);

                    LogInfo($"Saved move history to {historyFile}");
                }

                // После завершения сортировки проверяем, остались ли несортированные файлы
                string[] remainingFiles = Directory.GetFiles(basePath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .Where(f => !movedFiles.Contains(f))
                    .ToArray();
                
                if (remainingFiles.Length > 0)
                {
                    LogInfo($"Found {remainingFiles.Length} unsorted files. Creating 'Unsorted' folder for them.");
                    
                    // Создаем папку для несортированных файлов
                    string unsortedFolder = Path.Combine(basePath, "Unsorted");
                    if (!Directory.Exists(unsortedFolder))
                    {
                        Directory.CreateDirectory(unsortedFolder);
                    }
                    
                    // Перемещаем несортированные файлы
                    foreach (string fileName in remainingFiles)
                    {
                        string source = Path.Combine(basePath, fileName);
                        string destination = Path.Combine(unsortedFolder, fileName);
                        
                        try
                        {
                            File.Move(source, destination);
                            LogInfo($"Moved unsorted file: {fileName} to Unsorted folder");
                            
                            moveRecord.Add(new Dictionary<string, string>
                            {
                                ["source"] = source,
                                ["destination"] = destination,
                                ["item"] = fileName,
                                ["target_folder"] = "Unsorted",
                                ["type"] = "file"
                            });
                        }
                        catch (Exception ex)
                        {
                            LogError($"Error moving unsorted file {fileName}", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error moving files and folders", ex);
            }
        }

        private List<Dictionary<string, string>> MoveFilesAndFoldersWithOperations(Dictionary<string, object> categorization, string basePath)
        {
            List<Dictionary<string, string>> moveRecord = new List<Dictionary<string, string>>();
            
            try
            {
                // Create history folder for undo operations
                string historyFolder = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                    "AIFileSorter", "history"
                );

                if (!Directory.Exists(historyFolder))
                {
                    Directory.CreateDirectory(historyFolder);
                }

                // Отслеживаем перемещенные файлы и папки, чтобы избежать дублирования
                HashSet<string> movedFiles = new HashSet<string>();
                HashSet<string> movedFolders = new HashSet<string>();
                
                // Process files
                if (categorization.TryGetValue("files", out object filesObj) && filesObj is JObject filesObject)
                {
                    foreach (var prop in filesObject.Properties())
                    {
                        string folderPath = prop.Name;

                        // Create target folder if it doesn't exist
                        string fullFolderPath = Path.Combine(basePath, folderPath);
                        if (!Directory.Exists(fullFolderPath))
                        {
                            Directory.CreateDirectory(fullFolderPath);
                            LogInfo($"Created folder: {folderPath}");
                        }

                        // Move each file to the folder
                        JArray filesArray = prop.Value as JArray;
                        if (filesArray != null)
                        {
                            foreach (var fileToken in filesArray)
                            {
                                string fileName = fileToken.ToString();
                                
                                // Skip if already moved
                                if (movedFiles.Contains(fileName))
                                {
                                    LogInfo($"Skipping already moved file: {fileName}");
                                    continue;
                                }
                                
                                string source = Path.Combine(basePath, fileName);
                                string destination = Path.Combine(fullFolderPath, fileName);

                                if (File.Exists(source))
                                {
                                    try
                                    {
                                        File.Move(source, destination);
                                        movedFiles.Add(fileName);
                                        LogInfo($"Moved file: {fileName} to {folderPath}");

                                        // Record the move for undo
                                        moveRecord.Add(new Dictionary<string, string>
                                        {
                                            ["source"] = source,
                                            ["destination"] = destination,
                                            ["item"] = fileName,
                                            ["target_folder"] = folderPath,
                                            ["type"] = "file"
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError($"Error moving file {fileName} to {folderPath}", ex);
                                    }
                                }
                                else
                                {
                                    LogInfo($"File not found: {source}");
                                }
                            }
                        }
                    }
                }

                // Process folders similarly
                if (categorization.TryGetValue("folders", out object foldersObj) && foldersObj is JObject foldersObject)
                {
                    foreach (var prop in foldersObject.Properties())
                    {
                        string folderPath = prop.Name;

                        // Create target folder if it doesn't exist
                        string fullFolderPath = Path.Combine(basePath, folderPath);
                        if (!Directory.Exists(fullFolderPath))
                        {
                            Directory.CreateDirectory(fullFolderPath);
                            LogInfo($"Created folder: {folderPath}");
                        }

                        // Move each folder to destination
                        JArray foldersArray = prop.Value as JArray;
                        if (foldersArray != null)
                        {
                            foreach (var folderToken in foldersArray)
                            {
                                string folderName = folderToken.ToString();
                                string source = Path.Combine(basePath, folderName);
                                string destination = Path.Combine(fullFolderPath, folderName);

                                // Check if folder has already been moved
                                if (movedFolders.Contains(folderName))
                                {
                                    LogInfo($"Skipping already moved folder: {folderName}");
                                    continue;
                                }

                                if (Directory.Exists(source))
                                {
                                    try
                                    {
                                        // Check if destination exists
                                        if (Directory.Exists(destination))
                                        {
                                            string folderNameBase = Path.GetFileName(destination);
                                            string baseDestination = Path.GetDirectoryName(destination);
                                            int suffix = 1;

                                            while (Directory.Exists(destination))
                                            {
                                                string newFolderName = $"{folderNameBase}_{suffix}";
                                                destination = Path.Combine(baseDestination, newFolderName);
                                                suffix++;
                                            }

                                            LogInfo($"Renamed folder to avoid conflict: {folderName} → {Path.GetFileName(destination)}");
                                        }

                                        Directory.Move(source, destination);
                                        movedFolders.Add(folderName);
                                        LogInfo($"Moved folder: {folderName} to {folderPath}");

                                        // Record the move
                                        moveRecord.Add(new Dictionary<string, string>
                                        {
                                            ["source"] = source,
                                            ["destination"] = destination,
                                            ["item"] = folderName,
                                            ["target_folder"] = folderPath,
                                            ["type"] = "folder"
                                        });
                                    }
                                    catch (Exception ex)
                                    {
                                        LogError($"Error moving folder {folderName} to {folderPath}", ex);
                                    }
                                }
                                else
                                {
                                    LogInfo($"Folder not found: {source}");
                                }
                            }
                        }
                    }
                }

                // Save the move record for potential manual undo
                if (moveRecord.Count > 0)
                {
                    string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmmss");
                    string historyFile = Path.Combine(historyFolder, $"sort_history_{timestamp}.json");

                    var historyData = new Dictionary<string, object>
                    {
                        ["timestamp"] = timestamp,
                        ["downloads_folder"] = basePath,
                        ["moves"] = moveRecord
                    };

                    string historyJson = JsonConvert.SerializeObject(historyData, Formatting.Indented);
                    File.WriteAllText(historyFile, historyJson);

                    LogInfo($"Saved move history to {historyFile}");
                }

                // После завершения сортировки проверяем, остались ли несортированные файлы
                string[] remainingFiles = Directory.GetFiles(basePath, "*", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .Where(f => !movedFiles.Contains(f))
                    .ToArray();
                
                if (remainingFiles.Length > 0)
                {
                    LogInfo($"Found {remainingFiles.Length} unsorted files. Creating 'Unsorted' folder for them.");
                    
                    // Создаем папку для несортированных файлов
                    string unsortedFolder = Path.Combine(basePath, "Unsorted");
                    if (!Directory.Exists(unsortedFolder))
                    {
                        Directory.CreateDirectory(unsortedFolder);
                    }
                    
                    // Перемещаем несортированные файлы
                    foreach (string fileName in remainingFiles)
                    {
                        string source = Path.Combine(basePath, fileName);
                        string destination = Path.Combine(unsortedFolder, fileName);
                        
                        try
                        {
                            File.Move(source, destination);
                            LogInfo($"Moved unsorted file: {fileName} to Unsorted folder");
                            
                            moveRecord.Add(new Dictionary<string, string>
                            {
                                ["source"] = source,
                                ["destination"] = destination,
                                ["item"] = fileName,
                                ["target_folder"] = "Unsorted",
                                ["type"] = "file"
                            });
                        }
                        catch (Exception ex)
                        {
                            LogError($"Error moving unsorted file {fileName}", ex);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError("Error moving files and folders", ex);
            }

            return moveRecord;
        }

        private void RemoveEmptyFolders(string path)
        {
            try
            {
                LogInfo($"Cleaning up empty folders in {path}");
                int foldersRemoved = 0;

                // Get all directories in the path
                string[] directories = Directory.GetDirectories(path, "*", SearchOption.AllDirectories);

                // Sort by directory depth (deepest first)
                Array.Sort(directories, (a, b) => b.Split(Path.DirectorySeparatorChar).Length.CompareTo(a.Split(Path.DirectorySeparatorChar).Length));

                foreach (string dir in directories)
                {
                    // Skip the root folder
                    if (dir == path) continue;

                    try
                    {
                        // If directory is empty (no files and no subdirectories), delete it
                        if (Directory.GetFiles(dir).Length == 0 && Directory.GetDirectories(dir).Length == 0)
                        {
                            Directory.Delete(dir);
                            foldersRemoved++;
                            LogInfo($"Removed empty folder: {dir}");
                        }
                    }
                    catch (Exception ex)
                    {
                        LogError($"Error removing empty folder {dir}", ex);
                    }
                }

                LogInfo($"Removed {foldersRemoved} empty folders");
            }
            catch (Exception ex)
            {
                LogError("Error removing empty folders", ex);
            }
        }
    }
}

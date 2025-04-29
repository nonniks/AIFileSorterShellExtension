# AI File Sorter

AI-powered automatic file organization tool for Windows that intelligently categorizes and sorts files directly from Explorer's context menu.

## Overview

AI File Sorter is a Windows shell extension that integrates with Explorer's context menu, allowing you to organize files and folders with a single click. The tool uses OpenRouter's language models to analyze file names and create a logical folder structure automatically.

## Features

- **AI-Powered Organization**: Intelligently categorizes files based on names and types
- **Windows Explorer Integration**: Simple right-click access in any folder
- **Game & Mod Recognition**: Special handling for game files and mods
- **Web Search Option**: Optional online search for more accurate categorization
- **Undo Functionality**: Revert sorting operations within 2 minutes
- **Simple Installation**: User-friendly installer with minimal setup

## Installation

### Using the Installer Package

1. Download the latest release from the Releases section
2. Extract the ZIP file
3. Right-click on `Install.ps1` and select "Run with PowerShell"
4. If prompted about execution policy, choose "Yes" to proceed
5. Follow the on-screen instructions

### Building from Source

1. Clone this repository
2. Open the solution in Visual Studio
3. Build in Release mode
4. Run `Install-Extension.ps1` as administrator

## Usage

1. **First-time Setup**:
   - You'll need an OpenRouter API key (get one at https://openrouter.ai/)
   - Enter the key in the Settings dialog on first use

2. **Sorting Files**:
   - Right-click on any folder (or inside a folder)
   - Select "Sort Files" from the context menu
   - Wait for the AI to analyze and organize your files

3. **Configuration Options**:
   - Right-click and select "Settings..." to access configuration
   - Enable/disable web search for more accurate results
   - Update your API key

## System Requirements

- Windows 10/11
- .NET Framework 4.7.2 or higher
- Administrator privileges (for installation only)

## Technical Details

- Built with C# and .NET Framework
- Uses SharpShell for Windows Explorer integration
- Communicates with OpenRouter API for LLM processing
- No data is stored or shared beyond what's needed for sorting

## License

This project is licensed under the Apache License 2.0

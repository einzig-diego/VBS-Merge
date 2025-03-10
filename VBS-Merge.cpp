#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <stdexcept>
#include <windows.h>

std::string readFile(const std::string& filePath) {
    std::ifstream file(filePath.c_str());
    if (!file) {
        throw std::runtime_error("Could not open file: " + filePath);
    }
    std::stringstream buffer;
    buffer << file.rdbuf();
    return buffer.str();
}

void writeFile(const std::string& filePath, const std::string& content) {
    std::ofstream file(filePath.c_str());
    if (!file) {
        throw std::runtime_error("Could not write to file: " + filePath);
    }
    file << content;
}

std::string processIncludes(const std::string& content, const std::string& parentPath) {
    std::stringstream processedContent;
    std::istringstream stream(content);
    std::string line;

    while (std::getline(stream, line)) {
        size_t commentPos = line.find("'");
        size_t mergePos = line.find("#merge");
        if (commentPos != std::string::npos && mergePos != std::string::npos && mergePos > commentPos) {
            size_t startPos = line.find("#merge") + 6;
            std::string includeFileName = line.substr(startPos);
            includeFileName.erase(0, includeFileName.find_first_not_of(" \t"));  // Trim leading whitespace
            std::string includeFilePath = parentPath + "\\" + includeFileName + ".vbsm.vbs";
            
            std::string fileContent = readFile(includeFilePath);
            processedContent << fileContent << "\n";
        } else {
            processedContent << line << "\n";
        }
    }
    return processedContent.str();
}

void addContextMenuOption() {
    HKEY hKey;
    const char* keyPath = "Software\\Classes\\VBSFile\\shell\\Merge includes with VBS-Merge\\command";
    
    // Get the path of the current executable
    char exePath[MAX_PATH];
    GetModuleFileName(NULL, exePath, MAX_PATH);
    std::string command = "\"" + std::string(exePath) + "\" \"%1\"";

    // Check if the registry key already exists
    if (RegOpenKeyEx(HKEY_CURRENT_USER, keyPath, 0, KEY_READ, &hKey) == ERROR_SUCCESS) {
        RegCloseKey(hKey);
    } else {
        // Create the registry key
        if (RegCreateKeyEx(HKEY_CURRENT_USER, keyPath, 0, NULL, REG_OPTION_NON_VOLATILE, KEY_WRITE, NULL, &hKey, NULL) == ERROR_SUCCESS) {
            // Set the default value of the key
            RegSetValueEx(hKey, NULL, 0, REG_SZ, (const BYTE*)command.c_str(), command.length() + 1);
            RegCloseKey(hKey);
            std::cout << "Context menu option added." << std::endl;
        } else {
            std::cerr << "Failed to add context menu option." << std::endl;
        }
    }
}

int main(int argc, char* argv[]) {
    // Add context menu option when the program runs
    addContextMenuOption();
    std::cout << "Copyright 2025 Diego Manuel Niess Zafra" << std::endl;
    if (argc != 2) {
        std::cerr << "Usage: vbs-merge <filename.main.vbs>" << std::endl;
        return 1;
    }

    std::string filePath = argv[1];
    std::string parentPath = filePath.substr(0, filePath.find_last_of("/\\"));

    // Check if the file name follows the *.main.vbs scheme
    if (filePath.substr(filePath.find_last_of(".") - 4) != "main.vbs") {
        std::cerr << "Error: File name does not follow the *.main.vbs scheme." << std::endl;
        system("pause");
        return 1;
    }

    try {
        std::string content = readFile(filePath);
        std::string processedContent = processIncludes(content, parentPath);
        std::string outputPath = filePath.substr(0, filePath.find_last_of(".main.vbs") - sizeof(".main.vbs") + 2) + ".vbs";
        writeFile(outputPath, processedContent);
        std::cout << "Processed file saved as: " << outputPath << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        system("pause");
        return 1;
    }

    return 0;
}

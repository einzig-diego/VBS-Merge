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
        if (line.find("#merge") != std::string::npos) {
            size_t startPos = line.find("\"") + 1;
            size_t endPos = line.find_last_of("\"");
            std::string includeFilePath = parentPath + "\\" + line.substr(startPos, endPos - startPos);
            processedContent << readFile(includeFilePath);
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
        std::cout << "Context menu option already exists." << std::endl;
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
        std::cerr << "Usage: vbsmerge <filename.vbs>" << std::endl;
        return 1;
    }

    std::string filePath = argv[1];
    std::string parentPath = filePath.substr(0, filePath.find_last_of("/\\"));

    try {
        std::string content = readFile(filePath);
        std::string processedContent = processIncludes(content, parentPath);
        std::string outputPath = filePath.substr(0, filePath.find_last_of(".")) + ".merged.vbs";
        writeFile(outputPath, processedContent);
        std::cout << "Processed file saved as: " << outputPath << std::endl;
    } catch (const std::exception& e) {
        std::cerr << "Error: " << e.what() << std::endl;
        return 1;
    }

    return 0;
}

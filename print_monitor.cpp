/*
 * Windows Print Job Monitoring and Analytics System
 * 
 * This application monitors print jobs on Windows systems, collecting metadata
 * and exporting the data to CSV format for analytics.
 * 
 * Features:
 * - Real-time monitoring of all connected printers
 * - Collection of comprehensive print job metadata
 * - Automatic periodic data saving
 * - Interactive command interface
 * - RFC-4180 compliant CSV export
 * - Comprehensive error logging
 * 
 * Compilation:
 * To compile this application, use g++ with the following command:
 * g++ -o print_monitor.exe print_monitor.cpp -lwinmm -lwinspool -std=c++17
 * 
 * Usage:
 * - Run the executable to start the monitoring system
 * - Use the command interface to control monitoring and export data
 * - Check print_monitor.log for detailed logs
 * - CSV files are saved in the same directory as the executable
 */

#include <windows.h>
#include <lmcons.h>
#include <winspool.h>
#include <iostream>
#include <fstream>
#include <sstream>
#include <string>
#include <vector>
#include <map>
#include <thread>
#include <chrono>
#include <iomanip>
#include <ctime>
#include <mutex>
#include <algorithm>
#include <cctype>
#include <locale>
#include <io.h>
#include <fcntl.h>

// Function declarations
std::string getCurrentTimestamp();
std::string wideStringToUtf8(const wchar_t* wideStr);
std::string ansiStringToUtf8(const char* ansiStr);

// Print job data structure to store collected metadata
struct PrintJob {
    std::string printerName;      // Name of the printer
    std::string timestamp;        // When the job was detected
    std::string status;           // Current status of the job
    int pages = 0;               // Number of pages in the job
    int documentSize = 0;        // Size of the document in bytes
    std::string colorMode;       // Color or monochrome printing
    std::string duplexSetting;   // Duplex mode (simplex/duplex)
    std::string paperSize;       // Paper size used
    std::string userAccount;     // User who initiated the job
    std::string jobId;           // System-assigned job identifier
};

// Global variables for monitoring
bool monitoringActive = false;
std::vector<PrintJob> printJobs;
std::mutex jobsMutex;
std::thread monitorThread;
std::mutex logMutex; // For logging synchronization

// Log message to file
void logMessage(const std::string& level, const std::string& message) {
    std::lock_guard<std::mutex> lock(logMutex);
    
    std::string timestamp = getCurrentTimestamp();
    std::string logEntry = "[" + timestamp + "] [" + level + "] " + message + "\n";
    
    // Write to log file
    std::ofstream logFile("print_monitor.log", std::ios::app);
    if (logFile.is_open()) {
        logFile << logEntry;
        logFile.close();
    }
    
    // Also output to console
    if (level == "ERROR") {
        std::cerr << logEntry;
    } else {
        std::cout << logEntry;
    }
}

// Function to get current timestamp in ISO 8601 format
std::string getCurrentTimestamp() {
    auto now = std::chrono::system_clock::now();
    auto time_t = std::chrono::system_clock::to_time_t(now);
    
    std::stringstream ss;
    ss << std::put_time(std::localtime(&time_t), "%Y-%m-%dT%H:%M:%S");
    
    // Add milliseconds
    auto ms = std::chrono::duration_cast<std::chrono::milliseconds>(
        now.time_since_epoch()) % 1000;
    ss << "." << std::setfill('0') << std::setw(3) << ms.count();
    
    // Add timezone info (simple way - could be improved to detect automatically)
    ss << "+00:00"; // Assuming UTC for simplicity
    
    return ss.str();
}

// Function to convert ANSI string to UTF-8 string
std::string ansiStringToUtf8(const char* ansiStr) {
    if (!ansiStr) return "";
    return std::string(ansiStr);
}

// Function to convert wide string to UTF-8 string
std::string wideStringToUtf8(const wchar_t* wideStr) {
    if (!wideStr) return "";

    int size_needed = WideCharToMultiByte(CP_UTF8, 0, wideStr, -1, NULL, 0, NULL, NULL);
    if (size_needed == 0) return "";

    std::string result(size_needed - 1, 0);
    WideCharToMultiByte(CP_UTF8, 0, wideStr, -1, &result[0], size_needed, NULL, NULL);

    return result;
}

// Function to get color mode from device mode
std::string getColorMode(DEVMODE* pDevMode) {
    if (!pDevMode) return "Unknown";
    
    if (pDevMode->dmFields & DM_COLOR) {
        return (pDevMode->dmColor == DMCOLOR_COLOR) ? "Color" : "Monochrome";
    }
    return "Unknown";
}

// Function to get duplex setting from device mode
std::string getDuplexSetting(DEVMODE* pDevMode) {
    if (!pDevMode) return "Unknown";
    
    if (pDevMode->dmFields & DM_DUPLEX) {
        switch(pDevMode->dmDuplex) {
            case DMDUP_SIMPLEX: return "Simplex";
            case DMDUP_VERTICAL: return "Duplex Vertical";
            case DMDUP_HORIZONTAL: return "Duplex Horizontal";
            default: return "Unknown";
        }
    }
    return "Unknown";
}

// Function to get paper size from device mode
std::string getPaperSize(DEVMODE* pDevMode) {
    if (!pDevMode) return "Unknown";
    
    if (pDevMode->dmFields & DM_PAPERSIZE) {
        switch(pDevMode->dmPaperSize) {
            case DMPAPER_LETTER: return "Letter";
            case DMPAPER_LEGAL: return "Legal";
            case DMPAPER_A4: return "A4";
            case DMPAPER_A3: return "A3";
            case DMPAPER_A5: return "A5";
            default: return "Custom";
        }
    }
    return "Unknown";
}

// Function to get current user
std::string getCurrentUser() {
    char username[UNLEN + 1];
    DWORD username_len = UNLEN + 1;
    if (GetUserNameA(username, &username_len)) {
        return std::string(username);
    }
    return "Unknown";
}

// Helper function to get extended print job information
bool getExtendedJobInfo(HANDLE hPrinter, DWORD jobId, PrintJob& job) {
    // Request detailed job information
    DWORD bytesNeeded = 0;
    DWORD numJobs = 0;
    
    // Enumerate the specific job to get more details
    if (!EnumJobs(hPrinter, jobId, 1, 2, NULL, 0, &bytesNeeded, &numJobs)) {
        // If job 2 fails, the printer might not support detailed info
        return false;
    }
    
    if (bytesNeeded == 0) {
        return false;
    }
    
    std::vector<BYTE> jobBuffer(bytesNeeded);
    JOB_INFO_2* pJobInfo2 = reinterpret_cast<JOB_INFO_2*>(jobBuffer.data());

    if (EnumJobs(hPrinter, jobId, 1, 2, reinterpret_cast<LPBYTE>(pJobInfo2), bytesNeeded, &bytesNeeded, &numJobs)) {
        if (numJobs > 0) {
            // Try to get device mode information for color/duplex/paper settings
            if (pJobInfo2->pDevMode) {
                DEVMODE* pDevMode = pJobInfo2->pDevMode;
                job.colorMode = getColorMode(pDevMode);
                job.duplexSetting = getDuplexSetting(pDevMode);
                job.paperSize = getPaperSize(pDevMode);
            }
            return true;
        }
    }
    
    return false;
}

// Main monitoring function that uses Windows Print Spooler APIs
void monitorPrintJobs() {
    while (monitoringActive) {
        DWORD flags = PRINTER_ENUM_LOCAL | PRINTER_ENUM_CONNECTIONS;
        DWORD bytesNeeded = 0;
        DWORD numPrinters = 0;

        // First call to get required buffer size
        EnumPrinters(flags, NULL, 2, NULL, 0, &bytesNeeded, &numPrinters);

        if (bytesNeeded == 0) {
            logMessage("WARN", "No printers found during monitoring cycle");
            std::this_thread::sleep_for(std::chrono::seconds(5)); // Wait before retrying
            continue;
        }

        // Allocate buffer for printer info
        std::vector<BYTE> buffer(bytesNeeded);
        PRINTER_INFO_2A* pPrinterInfo2 = reinterpret_cast<PRINTER_INFO_2A*>(buffer.data());

        // Get printer information
        if (!EnumPrinters(flags, NULL, 2, reinterpret_cast<LPBYTE>(pPrinterInfo2), bytesNeeded, &bytesNeeded, &numPrinters)) {
            logMessage("ERROR", "Failed to enumerate printers. Error: " + std::to_string(GetLastError()));
            std::this_thread::sleep_for(std::chrono::seconds(5)); // Wait before retrying
            continue;
        }

        // Monitor print jobs for each printer
        for (DWORD i = 0; i < numPrinters && monitoringActive; ++i) {
            // Open the printer
            HANDLE hPrinter = NULL;
            const char* printerNameAnsi = pPrinterInfo2[i].pPrinterName;

            PRINTER_DEFAULTS pd = { NULL, NULL, PRINTER_ACCESS_USE };
            if (OpenPrinterA(const_cast<LPSTR>(printerNameAnsi), &hPrinter, &pd)) {
                // Enumerate jobs on this printer
                DWORD jobBytesNeeded = 0;
                DWORD numJobs = 0;
                
                // First call to get required buffer size
                EnumJobs(hPrinter, 0, 1000, 2, NULL, 0, &jobBytesNeeded, &numJobs);

                if (jobBytesNeeded > 0) {
                    std::vector<BYTE> jobBuffer(jobBytesNeeded);
                    JOB_INFO_2A* pJobInfo = reinterpret_cast<JOB_INFO_2A*>(jobBuffer.data());

                    if (EnumJobs(hPrinter, 0, 1000, 2, reinterpret_cast<LPBYTE>(pJobInfo), jobBytesNeeded, &jobBytesNeeded, &numJobs)) {
                        for (DWORD j = 0; j < numJobs && monitoringActive; ++j) {
                            PrintJob job;
                            job.printerName = ansiStringToUtf8(pPrinterInfo2[i].pPrinterName);
                            job.timestamp = getCurrentTimestamp();

                            // Map job status
                            if (pJobInfo[j].Status & JOB_STATUS_PAUSED)
                                job.status = "Paused";
                            else if (pJobInfo[j].Status & JOB_STATUS_ERROR)
                                job.status = "Error";
                            else if (pJobInfo[j].Status & JOB_STATUS_DELETING)
                                job.status = "Deleting";
                            else if (pJobInfo[j].Status & JOB_STATUS_SPOOLING)
                                job.status = "Spooling";
                            else if (pJobInfo[j].Status & JOB_STATUS_PRINTING)
                                job.status = "Printing";
                            else if (pJobInfo[j].Status & JOB_STATUS_OFFLINE)
                                job.status = "Offline";
                            else if (pJobInfo[j].Status & JOB_STATUS_PAPEROUT)
                                job.status = "Paper Out";
                            else if (pJobInfo[j].Status & JOB_STATUS_DELETED)
                                job.status = "Deleted";
                            else if (pJobInfo[j].Status & JOB_STATUS_BLOCKED_DEVQ)
                                job.status = "Blocked";
                            else if (pJobInfo[j].Status & JOB_STATUS_USER_INTERVENTION)
                                job.status = "User Intervention Required";
                            else
                                job.status = "Queued";

                            job.pages = pJobInfo[j].TotalPages > 0 ? pJobInfo[j].TotalPages : pJobInfo[j].PagesPrinted;
                            job.documentSize = static_cast<int>(pJobInfo[j].Size);
                            job.userAccount = ansiStringToUtf8(pJobInfo[j].pUserName);
                            job.jobId = std::to_string(pJobInfo[j].JobId);

                            // Try to get extended information from the printer
                            // The getExtendedJobInfo function might need adjustment since we're already using level 2
                            if (pJobInfo[j].pDevMode) {
                                DEVMODEA* pDevMode = pJobInfo[j].pDevMode;
                                job.colorMode = getColorMode(reinterpret_cast<DEVMODE*>(pDevMode));
                                job.duplexSetting = getDuplexSetting(reinterpret_cast<DEVMODE*>(pDevMode));
                                job.paperSize = getPaperSize(reinterpret_cast<DEVMODE*>(pDevMode));
                            }
                            
                            {
                                std::lock_guard<std::mutex> lock(jobsMutex);
                                
                                // Check if this job is already recorded to avoid duplicates
                                bool exists = false;
                                for (const auto& existingJob : printJobs) {
                                    if (existingJob.jobId == job.jobId && existingJob.printerName == job.printerName) {
                                        exists = true;
                                        break;
                                    }
                                }
                                
                                if (!exists) {
                                    printJobs.push_back(job);
                                    
                                    // Keep only the last 1000 jobs to prevent memory issues
                                    if (printJobs.size() > 1000) {
                                        printJobs.erase(printJobs.begin(), printJobs.begin() + 100); // Remove oldest 100
                                    }
                                }
                            }
                            
                            if (monitoringActive) {
                                logMessage("INFO", "Detected print job: " + job.jobId 
                                          + " on " + job.printerName 
                                          + " - Status: " + job.status);
                            }
                        }
                    } else {
                        logMessage("ERROR", "Failed to enumerate jobs. Error: " + std::to_string(GetLastError()));
                    }
                }
                
                ClosePrinter(hPrinter);
            } else {
                logMessage("ERROR", "Could not open printer: " + ansiStringToUtf8(pPrinterInfo2[i].pPrinterName)
                          + ". Error: " + std::to_string(GetLastError()));
            }
        }
        
        // Wait before checking again to reduce CPU usage, but check if monitoring is still active
        for (int i = 0; i < 10 && monitoringActive; ++i) {
            std::this_thread::sleep_for(std::chrono::seconds(1));
        }
    }
}

// Start monitoring print jobs
void startMonitoring() {
    if (monitoringActive) {
        logMessage("INFO", "Monitoring is already active.");
        return;
    }
    
    try {
        monitoringActive = true;
        monitorThread = std::thread(monitorPrintJobs);
        logMessage("INFO", "Print job monitoring started.");
    } catch (const std::exception& e) {
        logMessage("ERROR", std::string("Failed to start monitoring thread: ") + e.what());
    }
}

// Stop monitoring print jobs
void stopMonitoring() {
    if (!monitoringActive) {
        logMessage("INFO", "Monitoring is not active.");
        return;
    }
    
    try {
        monitoringActive = false;
        if (monitorThread.joinable()) {
            monitorThread.join();
        }
        logMessage("INFO", "Print job monitoring stopped.");
    } catch (const std::exception& e) {
        logMessage("ERROR", std::string("Failed to stop monitoring thread: ") + e.what());
    }
}

// Export print jobs to CSV file
bool exportToCSV(const std::string& filename) {
    try {
        std::lock_guard<std::mutex> lock(jobsMutex);
        
        std::ofstream file(filename);
        if (!file.is_open()) {
            logMessage("ERROR", "Could not open file for writing: " + filename);
            return false;
        }
        
        // Write CSV header following RFC-4180
        file << "\"Printer Name\",\"Timestamp\",\"Status\",\"Pages\",\"Document Size\",\"Color Mode\",\"Duplex Setting\",\"Paper Size\",\"User Account\",\"Job ID\"\n";
        
        // Write each print job as a CSV row, properly escaping values per RFC-4180
        for (const auto& job : printJobs) {
            // Properly escape and quote each field per RFC-4180
            std::string printerName = job.printerName;
            std::string status = job.status;
            std::string colorMode = job.colorMode;
            std::string duplexSetting = job.duplexSetting;
            std::string paperSize = job.paperSize;
            std::string userAccount = job.userAccount;
            std::string jobId = job.jobId;
            
            // Replace double quotes with two double quotes (RFC-4180 section 2.4)
            size_t pos;
            pos = 0;
            while ((pos = printerName.find('"', pos)) != std::string::npos) {
                printerName.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = status.find('"', pos)) != std::string::npos) {
                status.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = colorMode.find('"', pos)) != std::string::npos) {
                colorMode.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = duplexSetting.find('"', pos)) != std::string::npos) {
                duplexSetting.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = paperSize.find('"', pos)) != std::string::npos) {
                paperSize.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = userAccount.find('"', pos)) != std::string::npos) {
                userAccount.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            pos = 0;
            while ((pos = jobId.find('"', pos)) != std::string::npos) {
                jobId.replace(pos, 1, "\"\"");
                pos += 2;
            }
            
            file << "\"" << printerName << "\","
                 << "\"" << job.timestamp << "\","
                 << "\"" << status << "\","
                 << job.pages << ","
                 << job.documentSize << ","
                 << "\"" << colorMode << "\","
                 << "\"" << duplexSetting << "\","
                 << "\"" << paperSize << "\","
                 << "\"" << userAccount << "\","
                 << "\"" << jobId << "\"\n";
        }
        
        file.close();
        logMessage("INFO", "Data exported to: " + filename + " (" + std::to_string(printJobs.size()) + " records)");
        return true;
    } catch (const std::exception& e) {
        logMessage("ERROR", std::string("Exception during CSV export: ") + e.what());
        return false;
    }
}

// Function to perform periodic saves
void periodicSave() {
    while (monitoringActive) {
        // Sleep for 30 minutes (1800 seconds)
        for (int i = 0; i < 1800 && monitoringActive; ++i) {
            std::this_thread::sleep_for(std::chrono::seconds(1));
        }
        
        if (monitoringActive) {
            std::string filename = "print_jobs_auto_save_" + getCurrentTimestamp().substr(0, 19) + ".csv";
            // Replace colons in timestamp with hyphens for valid filename
            std::replace(filename.begin(), filename.end(), ':', '-');
            exportToCSV(filename);
        }
    }
}

// Force save data to default file
void forceSave() {
    std::string filename = "print_jobs_" + getCurrentTimestamp().substr(0, 19) + ".csv";
    // Replace colons in timestamp with hyphens for valid filename
    std::replace(filename.begin(), filename.end(), ':', '-');
    exportToCSV(filename);
}

// Show current statistics
void showStatistics() {
    std::lock_guard<std::mutex> lock(jobsMutex);
    
    std::cout << "\n=== Print Job Statistics ===" << std::endl;
    std::cout << "Total print jobs recorded: " << printJobs.size() << std::endl;
    
    if (!printJobs.empty()) {
        // Count jobs by status
        std::map<std::string, int> statusCount;
        int totalPages = 0;
        int totalSize = 0;
        
        for (const auto& job : printJobs) {
            statusCount[job.status]++;
            totalPages += job.pages;
            totalSize += job.documentSize;
        }
        
        std::cout << "Jobs by status:" << std::endl;
        for (const auto& pair : statusCount) {
            std::cout << "  " << pair.first << ": " << pair.second << std::endl;
        }
        
        std::cout << "Total pages printed: " << totalPages << std::endl;
        std::cout << "Total document size: " << totalSize << " bytes" << std::endl;
        std::cout << "Average pages per job: " << (double)totalPages / printJobs.size() << std::endl;
    }
    
    std::cout << "Monitoring status: " << (monitoringActive ? "ACTIVE" : "STOPPED") << std::endl;
    std::cout << "============================\n" << std::endl;
}

// Show help information
void showHelp() {
    std::cout << "\n=== Print Job Monitor Help ===" << std::endl;
    std::cout << "Commands:" << std::endl;
    std::cout << "  start         - Start monitoring print jobs" << std::endl;
    std::cout << "  stop          - Stop monitoring print jobs" << std::endl;
    std::cout << "  save          - Force save current data to CSV" << std::endl;
    std::cout << "  export [file] - Export to specified CSV file" << std::endl;
    std::cout << "  stats         - Show current statistics" << std::endl;
    std::cout << "  help          - Show this help message" << std::endl;
    std::cout << "  quit/exit     - Quit the application" << std::endl;
    std::cout << "==============================\n" << std::endl;
}

// Main command loop
void commandLoop() {
    std::cout << "Windows Print Job Monitoring System" << std::endl;
    std::cout << "Type 'help' for available commands or 'quit' to exit.\n" << std::endl;
    
    std::string input;
    while (true) {
        std::cout << "> ";
        std::getline(std::cin, input);
        
        // Convert to lowercase for easier command matching
        std::transform(input.begin(), input.end(), input.begin(), ::tolower);
        
        if (input == "start") {
            startMonitoring();
        }
        else if (input == "stop") {
            stopMonitoring();
        }
        else if (input == "save") {
            forceSave();
        }
        else if (input.substr(0, 6) == "export") {
            std::string filename = "print_jobs_export.csv";
            size_t pos = input.find(' ', 6);
            if (pos != std::string::npos) {
                filename = input.substr(pos + 1);
            }
            
            if (filename.length() > 0) {
                exportToCSV(filename);
            } else {
                std::cout << "Please specify a filename for export." << std::endl;
            }
        }
        else if (input == "stats") {
            showStatistics();
        }
        else if (input == "help") {
            showHelp();
        }
        else if (input == "quit" || input == "exit") {
            if (monitoringActive) {
                stopMonitoring();
            }
            std::cout << "Exiting..." << std::endl;
            break;
        }
        else if (input.empty()) {
            continue;
        }
        else {
            std::cout << "Unknown command. Type 'help' for available commands." << std::endl;
        }
    }
}

std::thread periodicSaveThread;

int main() {
    try {
        logMessage("INFO", "Initializing Windows Print Job Monitoring System...");
        
        // Initialize random seed for any simulated jobs
        srand(static_cast<unsigned>(time(nullptr)));
        
        // Start the periodic save thread
        periodicSaveThread = std::thread(periodicSave);
        
        // Start the command loop
        commandLoop();
        
        // Stop monitoring if still active
        if (monitoringActive) {
            stopMonitoring();
        }
        
        // Wait for periodic save thread to finish
        monitoringActive = false;
        if (periodicSaveThread.joinable()) {
            periodicSaveThread.join();
        }
        
        logMessage("INFO", "Windows Print Job Monitoring System exited normally.");
    } catch (const std::exception& e) {
        logMessage("ERROR", std::string("Uncaught exception in main: ") + e.what());
        return 1;
    } catch (...) {
        logMessage("ERROR", "Unknown exception in main.");
        return 1;
    }
    
    return 0;
}
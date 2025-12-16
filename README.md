# PrintingMonitor
A simple printing monitor written in C++

1. Compile the application as described above
2. Run the executable to start the monitoring system
3. Use the interactive command interface to control the application:
   - `start` - Start monitoring print jobs
   - `stop` - Stop monitoring print jobs
   - `save` - Force save current data to CSV
   - `export [filename]` - Export to specified CSV file
   - `stats` - Show current statistics
   - `help` - Show help information
   - `quit` or `exit` - Quit the application

## Data Collection Fields
The system captures the following print job attributes:
- Printer name/identifier
- Timestamp (date and time with timezone in ISO 8601 format)
- Job status (queued, printing, completed, error, etc.)
- Number of pages
- Document size (bytes)
- Color mode (monochrome/color)
- Duplex setting (simplex/duplex)
- Paper size/type
- User account/initiator
- Job ID (system-assigned)

## CSV Export Format
The exported CSV files follow RFC-4180 standards with proper field escaping and quoting:
```
"Printer Name","Timestamp","Status","Pages","Document Size","Color Mode","Duplex Setting","Paper Size","User Account","Job ID"
"HP_LaserJet_123","2023-12-16T10:30:45.123+00:00","Completed",5,25000,"Color","Duplex","A4","john_doe","101"
```

## Error Handling and Logging
The application maintains detailed logs in `print_monitor.log` with timestamps, log levels, and error messages. The log format follows:
```
[2023-12-16T10:30:45.123+00:00] [INFO] Print job monitoring started.
[2023-12-16T10:31:20.456+00:00] [ERROR] Failed to enumerate printers. Error: 5
```

## Architecture
The application uses an event-driven architecture with multiple threads:
- **Main Thread**: Handles the command interface
- **Monitoring Thread**: Continuously polls the print spooler for new jobs
- **Periodic Save Thread**: Automatically saves data every 30 minutes

## Performance Considerations
- The monitoring process sleeps between polling cycles to minimize CPU impact
- Memory usage is controlled by keeping only the last 1000 print jobs in memory
- Duplicate detection prevents redundant entries in the dataset

## Security Considerations
- The application requires appropriate permissions to access the print spooler service
- User account information is collected only from jobs the current user can access
- All file operations are performed in the application directory

## Troubleshooting
- If monitoring fails, check the `print_monitor.log` file for error details
- Ensure the application has sufficient permissions to access printer information
- The application may need to be run as administrator to monitor all print jobs on the system

## Limitations
- Requires Windows operating system
- Some printer-specific metadata may not be available for all printer types
- Historical jobs (created before monitoring started) are not captured

## License
This code is provided as-is for educational and development purposes.

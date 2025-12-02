# SAP GUI Connector

Python script for connecting to SAP GUI using COM interface.

## Description

This project provides a simple Python script that establishes a connection to SAP GUI using the Windows COM interface. It can be used as a foundation for automating SAP GUI interactions.

## Features

- Connect to running SAP GUI instance
- Retrieve active session information
- Display system details (system name, client, user, transaction, etc.)
- Error handling and user-friendly messages

## Requirements

- Windows OS (required for COM interface)
- SAP GUI installed and running
- Python 3.6 or higher
- pywin32 library

## Installation

1. Clone the repository:
```bash
git clone https://github.com/YOUR_USERNAME/sap-gui-connector.git
cd sap-gui-connector
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

## Prerequisites

Before running the script, ensure:

1. SAP GUI is installed on your system
2. SAP GUI Scripting is enabled:
   - Open SAP GUI
   - Go to Options → Accessibility & Scripting → Scripting
   - Enable "Enable scripting"
3. You have an active SAP GUI connection and session open

## Usage

Run the script:
```bash
python sap_connector.py
```

The script will:
1. Check if pywin32 is installed
2. Connect to the running SAP GUI instance
3. Retrieve and display information about the active session

## Expected Output

```
SAP GUI Connector
==================================================
Successfully connected to SAP GUI!
Connection: SAP System Description
Session ID: ses[0]

==================================================
SAP Session Information:
==================================================
System Name: XXX
Client: 100
User: USERNAME
Language: EN
Transaction: SE38
Program: SAPMSSY0
Screen Number: 1000
==================================================

Connection test successful!
```

## Enabling SAP GUI Scripting

If scripting is not enabled, you'll need to:

1. **On the client side (SAP GUI):**
   - Open SAP Logon
   - Click on the settings icon
   - Go to Accessibility & Scripting → Scripting
   - Check "Enable scripting"

2. **On the server side (if you have admin rights):**
   - Execute transaction RZ11
   - Display parameter: sapgui/user_scripting
   - Set value to TRUE

## Troubleshooting

**Error: "SAP GUI not found"**
- Make sure SAP GUI is installed and running

**Error: "No SAP GUI connections found"**
- Open SAP GUI and connect to a system before running the script

**Error: "SAP GUI Scripting engine not available"**
- Enable SAP GUI scripting in SAP GUI options

**ImportError: No module named 'win32com'**
- Install pywin32: `pip install pywin32`

## Future Enhancements

- Execute SAP transactions
- Read and write data from SAP screens
- Export data to various formats
- Command-line interface for automation

## License

MIT License

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Author

Created with Claude Code

## Disclaimer

This tool is for educational and authorized automation purposes only. Always ensure you have proper authorization before automating SAP systems.

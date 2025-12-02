"""
SAP GUI Connector Script
Connects to SAP GUI using COM interface
"""

import sys
import subprocess


def check_win32com():
    """Check if pywin32 is installed"""
    try:
        import win32com.client
        return True
    except ImportError:
        return False


def get_sap_gui_connection():
    """
    Establish connection to SAP GUI
    Returns: tuple (session, connection) or (None, None) if failed
    """
    try:
        import win32com.client

        # Get SAP GUI Scripting object
        sap_gui_auto = win32com.client.GetObject("SAPGUI")
        if not sap_gui_auto:
            print("SAP GUI not found. Please ensure SAP GUI is installed and running.")
            return None, None

        # Get the scripting engine
        application = sap_gui_auto.GetScriptingEngine
        if not application:
            print("SAP GUI Scripting engine not available.")
            return None, None

        # Check if there are any open connections
        if application.Children.Count == 0:
            print("No SAP GUI connections found. Please open SAP GUI and connect to a system.")
            return None, None

        # Get the first connection
        connection = application.Children(0)

        # Check if there are any sessions
        if connection.Children.Count == 0:
            print("No active sessions found in the connection.")
            return None, None

        # Get the first session
        session = connection.Children(0)

        print("Successfully connected to SAP GUI!")
        print(f"Connection: {connection.Description}")
        print(f"Session ID: {session.Id}")

        return session, connection

    except Exception as e:
        print(f"Error connecting to SAP GUI: {str(e)}")
        print("\nPossible reasons:")
        print("1. SAP GUI is not running")
        print("2. SAP GUI Scripting is not enabled")
        print("3. No active SAP session is open")
        return None, None


def get_session_info(session):
    """
    Get information about the current SAP session
    """
    if not session:
        return

    try:
        info = session.Info
        print("\n" + "="*50)
        print("SAP Session Information:")
        print("="*50)
        print(f"System Name: {info.SystemName}")
        print(f"Client: {info.Client}")
        print(f"User: {info.User}")
        print(f"Language: {info.Language}")
        print(f"Transaction: {info.Transaction}")
        print(f"Program: {info.Program}")
        print(f"Screen Number: {info.ScreenNumber}")
        print("="*50)
    except Exception as e:
        print(f"Error getting session info: {str(e)}")


def main():
    """
    Main function to test SAP GUI connection
    """
    print("SAP GUI Connector")
    print("="*50)

    # Check if running on Windows
    if sys.platform != "win32":
        print("Error: This script only works on Windows as it uses COM interface.")
        return 1

    # Check if pywin32 is installed
    if not check_win32com():
        print("Error: pywin32 is not installed.")
        print("Please install it using: pip install pywin32")
        return 1

    # Try to connect to SAP GUI
    session, connection = get_sap_gui_connection()

    if session and connection:
        # Get and display session information
        get_session_info(session)
        print("\nConnection test successful!")
        return 0
    else:
        print("\nConnection test failed!")
        return 1


if __name__ == "__main__":
    exit_code = main()
    sys.exit(exit_code)

# SAP GUI Connection Manager
# ===========================
# Provides the SAPManager class to automate SAP GUI connections via COM scripting.
#
# Usage:
#   sap = SAPManager()
#   session = sap.get_session()    # Opens SAP GUI, connects to PR1, returns a session
#   ... (run your SAP script) ...
#   sap.close_connection(session)  # Logs off and closes the SAP window
#
# What get_session() does:
#   1. Connects to an existing SAP GUI instance, or launches a new one if none is running
#   2. Opens a NEW connection to the PR1 server (never reuses existing sessions)
#   3. Handles the "Multiple Logon" popup automatically
#   4. Handles the session manager screen if it appears instead of Easy Access
#   5. Returns the session object used to script SAP GUI actions (findById, press, etc.)
#
# Other methods:
#   - return_to_main_menu(session)  : sends /n and waits until Easy Access screen appears
#   - close_connection(session)     : sends /nex to log off (only closes this script's connection)

import subprocess
import os
import platform
import time
import win32com.client

class SAPManager:
    def __init__(self, connection_name="-PR1 [PWC SAP ECC 6.0]", xml_name="PWC_SAPUILandscape.xml"):
        self.connection_name = connection_name
        self.xml_name = xml_name
        self.session = None
        self.connection = None
        self.application = None
    
    def _handle_session_manager_screen(self, session):
        """Navigate past the session manager screen if it appears.

        Sometimes SAP opens to a screen titled just 'SAP' with a
        'Start SAP Easy Access' button instead of going directly to the
        main menu. Tries multiple approaches to click through.
        """
        try:
            title = session.findById("wnd[0]").text
            if "Easy Access" in title:
                return  # Already on the main menu
            print(f"Session manager screen detected (title: '{title}'). Clicking through...")

            # Try to find and click the 'Start SAP Easy Access' button
            # by scanning all children in the user area
            try:
                usr = session.findById("wnd[0]/usr")
                for i in range(usr.Children.Count):
                    child = usr.Children(i)
                    try:
                        child_text = child.text if hasattr(child, 'text') else ""
                        if "Easy Access" in child_text or "Start" in child_text:
                            child.press()
                            time.sleep(2)
                            print(f"Clicked '{child_text}' button (id: {child.id}).")
                            return
                    except Exception:
                        continue
            except Exception:
                pass

            # Fallback: press Enter which may activate the default button
            try:
                session.findById("wnd[0]").sendVKey(0)
                time.sleep(2)
                print("Pressed Enter to pass session manager screen.")
            except Exception:
                pass
        except Exception as e:
            print(f"Warning: Could not handle session manager screen: {e}")

    def return_to_main_menu(self, session, timeout=5):
        """Navigate back to main menu and wait until it's confirmed.

        Sends /n and polls until the SAP Easy Access main menu is detected
        (title contains 'Easy Access'). This serves as a reliable signal
        that the previous transaction has fully completed.

        Returns True if main menu reached, False on timeout.
        """
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/n"
            session.findById("wnd[0]").sendVKey(0)
        except Exception as e:
            print(f"Warning: Could not send /n command: {e}")
            return False

        start = time.time()
        while time.time() - start < timeout:
            try:
                title = session.findById("wnd[0]").text
                if "Easy Access" in title:
                    print("Returned to SAP main menu.")
                    return True
            except Exception:
                pass
            time.sleep(0.5)

        print(f"Warning: Timed out waiting for main menu after {timeout}s.")
        return False
    
    def close_connection(self, session):
        """Close the connection this script opened.

        Sends /nex to log off and close the SAP window, then clears
        internal references. Only affects this script's connection —
        other connections remain untouched.

        Returns True if closed successfully, False on error.
        """
        try:
            session.findById("wnd[0]/tbar[0]/okcd").text = "/nex"
            session.findById("wnd[0]").sendVKey(0)
            print(f"[{self.connection_name}] Connection closed.")
            self.session = None
            self.connection = None
            return True
        except Exception as e:
            print(f"Warning: Could not close connection: {e}")
            return False

    def _handle_multiple_logon_popup(self, session=None):
        """Internal helper to allow the multiple logon (continue with current session)."""
        session = session or self.session
        try:
            popup = session.findById("wnd[1]", False)
            if popup is not None:
                if "License Information" in popup.text or "Multiple Logon" in popup.text:
                    session.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").Select()
                    session.findById("wnd[1]/tbar[0]/btn[0]").press()
                    print(f"[{self.connection_name}] Allowed Multiple Logon - new session created.")
        except Exception:
            pass
    
    def get_session(self):
        if platform.system() != "Windows":
            raise RuntimeError("Windows OS is required for SAP GUI Scripting.")
        
        # 1. Connect to existing SAP GUI or start new one
        try:
            try:
                sap_gui_auto = win32com.client.GetObject("SAPGUI")
                self.application = sap_gui_auto.GetScriptingEngine
                print("Connected to existing SAP GUI instance.")
            except Exception:
                # No SAP GUI running — start it
                exe_path = r"C:\Program Files\SAP\FrontEnd\SAPGUI\saplgpad.exe"
                xml_full_path = os.path.expandvars(rf"%ProgramFiles%\SAP\PWCConfig\{self.xml_name}")

                if not os.path.exists(exe_path):
                    raise FileNotFoundError(f"SAP Executable not found at {exe_path}")

                subprocess.Popen(
                    f'"{exe_path}" /LSXML_FILE="{xml_full_path}"',
                    shell=True,
                    stdin=subprocess.DEVNULL,
                    stdout=subprocess.DEVNULL,
                    stderr=subprocess.DEVNULL,
                    creationflags=subprocess.DETACHED_PROCESS,
                )

                # Poll until SAP GUI is available instead of fixed sleep
                for _ in range(15):
                    time.sleep(1)
                    try:
                        sap_gui_auto = win32com.client.GetObject("SAPGUI")
                        self.application = sap_gui_auto.GetScriptingEngine
                        break
                    except Exception:
                        continue
                else:
                    raise RuntimeError("SAP GUI did not become available within 15 seconds")
                print("Started new SAP GUI instance.")
        except Exception as e:
            raise RuntimeError(f"Failed to connect to SAP GUI: {e}")

        # 2. Always open a new connection to avoid interfering with
        #    other scripts that may be using existing sessions
        for attempt in range(2):
            try:
                print(f"[{self.connection_name}] Opening new connection...")
                self.connection = self.application.OpenConnection(self.connection_name, True)
                time.sleep(2)
                self.session = self.connection.Children(0)
                self._handle_multiple_logon_popup()
                print(f"New connection opened with 1 session")
                break
            except Exception as e:
                if attempt == 0:
                    print(f"Connection attempt failed, retrying: {e}")
                    time.sleep(3)
                else:
                    raise ConnectionError(f"Could not open new connection to {self.connection_name}: {str(e)}")

        # Handle session manager screen if it appears instead of Easy Access
        self._handle_session_manager_screen(self.session)

        return self.session
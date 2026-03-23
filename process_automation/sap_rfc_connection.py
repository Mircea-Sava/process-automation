# SAP RFC Connection
# ==================
# Low-level SAP connection via RFC (COM) with secure credential storage.
#
# Classes:
#   - WinCredentialStore : DPAPI-encrypted credential storage in Windows Credential Manager
#   - RFCConnection      : SAP.Functions COM wrapper for BBP_RFC_READ_TABLE and RFC_READ_TEXT
#
# Usage:
#   from process_automation import RFCConnection, WinCredentialStore
#
#   WinCredentialStore.ensure_credentials()
#   conn = RFCConnection()
#   conn.login()
#   data = conn.bbp_rfc_read_table("KNA1", "KUNNR,NAME1", ["MANDT = '400'"])
#   conn.log_off()

import getpass
import base64
import win32cred
import win32crypt
from win32com.client import Dispatch


class WinCredentialStore:
    """DPAPI-encrypted credential storage in Windows Credential Manager."""

    TARGET_NAME = "SAP_CREDENTIALS"

    @staticmethod
    def save_credentials():
        """Prompt for user/password, encrypt with DPAPI, store in Credential Manager."""
        user = input("User: ")
        pwd = getpass.getpass(prompt="Password: ")
        combined = f"{user}:{pwd}".encode("utf-8")

        protected_blob = win32crypt.CryptProtectData(combined, None, None, None, None, 0)
        protected_blob_b64 = base64.b64encode(protected_blob).decode("utf-8")

        credential = {
            "Type": win32cred.CRED_TYPE_GENERIC,
            "TargetName": WinCredentialStore.TARGET_NAME,
            "UserName": "",
            "CredentialBlob": protected_blob_b64,
            "Persist": win32cred.CRED_PERSIST_LOCAL_MACHINE,
        }
        win32cred.CredWrite(credential, 0)
        print("Credentials saved securely to Windows Credential Manager.")

    @staticmethod
    def fetch_credentials():
        """Retrieve and decrypt credentials. Returns (user, pwd) or (None, None)."""
        try:
            cred = win32cred.CredRead(WinCredentialStore.TARGET_NAME, win32cred.CRED_TYPE_GENERIC)
            protected_blob = base64.b64decode(cred["CredentialBlob"])
            _, decrypted_bytes = win32crypt.CryptUnprotectData(protected_blob, None, None, None, 0)
            combined = decrypted_bytes.decode("utf-8")
            user, pwd = combined.split(":", 1)
            return user, pwd
        except Exception as e:
            print(f"Failed to retrieve credentials: {e}")
            return None, None

    @staticmethod
    def delete_credentials():
        """Remove credentials from Windows Credential Manager."""
        try:
            win32cred.CredDelete(WinCredentialStore.TARGET_NAME, win32cred.CRED_TYPE_GENERIC, 0)
            print("Credentials deleted from Windows Credential Manager.")
        except Exception as e:
            print(f"Failed to delete credentials: {e}")

    @staticmethod
    def ensure_credentials():
        """Check if credentials exist; prompt to save if not."""
        user, pwd = WinCredentialStore.fetch_credentials()
        if not user or not pwd:
            print("No credentials found. Please enter your SAP credentials.")
            WinCredentialStore.save_credentials()
        else:
            print("Credentials already exist in Windows Credential Manager.")


class RFCConnection:
    """SAP.Functions COM wrapper for RFC calls (BBP_RFC_READ_TABLE, RFC_READ_TEXT)."""

    def __init__(self):
        self.R3 = Dispatch("SAP.Functions")
        self.system = "-PR1 [PWC SAP ECC 6.0]"

    def login(self):
        """Set connection parameters, log on, return connection object (or None on failure)."""
        self.R3.Connection.System = self.system
        self.R3.Connection.Client = "400"
        user, pwd = WinCredentialStore.fetch_credentials()
        if not user or not pwd:
            print("Credential retrieval failed. Please save credentials.")
            return None
        self.R3.Connection.User = user
        self.R3.Connection.Password = pwd
        self.R3.Connection.Language = "EN"
        self.R3.Connection.Logon(0, True)
        if self.R3.Connection.IsConnected != 1:
            return None
        return self.R3

    def log_off(self):
        """Log off SAP."""
        self.R3.Connection.Logoff()

    def bbp_rfc_read_table(self, tbl, cols, filter, rowcount=0, rowskips=0):
        """
        Fetch data from a SAP table via BBP_RFC_READ_TABLE.

        Parameters
        ----------
        tbl      : SAP table name
        cols     : comma-separated column names
        filter   : list of filter condition strings (OPTIONS rows)
        rowcount : max rows to return (0 = no limit)
        rowskips : number of rows to skip from the top (for pagination)
        """
        if self.R3.Connection.IsConnected != 1:
            print("Error - logon to SAP Failed")
            self.login()

        if self.R3 is None:
            return None

        MyFunc = self.R3.Add("BBP_RFC_READ_TABLE")

        QUERY_TABLE = MyFunc.Exports("QUERY_TABLE")
        ROWCOUNT_param = MyFunc.Exports("ROWCOUNT")
        ROWSKIPS_param = MyFunc.Exports("ROWSKIPS")
        OPTIONS = MyFunc.Tables("OPTIONS")
        FIELDS = MyFunc.Tables("FIELDS")

        QUERY_TABLE.Value = tbl
        ROWCOUNT_param.Value = rowcount
        ROWSKIPS_param.Value = rowskips

        OPTIONS.Data = [filter]
        FIELDS.Data = cols.split(",")

        DATA = MyFunc.Tables("DATA")
        DATA.FreeTable()

        if MyFunc.call != True:
            return None

        if DATA.Rowcount > 0:
            unique = tuple(set(DATA.Data))
            dic = []
            for i in unique:
                d = {}
                for j in FIELDS.Data:
                    istart = int(j[1])
                    iEnd = int(j[2])
                    d[j[4]] = (i[0][istart:iEnd + istart]).strip()
                dic.append(d)
        else:
            dic = None

        return dic

    def read_text(self, td_name, td_line):
        """Read SAP text objects via RFC_READ_TEXT."""
        if self.R3.Connection.IsConnected != 1:
            print("Error - logon to SAP Failed")
            self.login()

        if self.R3 is None:
            return None

        MyFunc = self.R3.Add("RFC_READ_TEXT")
        DATA = MyFunc.Tables("TEXT_LINES")
        DATA.FreeTable()
        DATA.DATA = (("", f"{td_line}", f"{td_name}", "LTXT", "", "000", "", ""),)

        if MyFunc.call != True:
            return None

        text_lines = ""
        if DATA.Rowcount > 0:
            for row in DATA.DATA:
                text_lines += str(row[7])
            return text_lines

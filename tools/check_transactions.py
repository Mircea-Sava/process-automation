"""
SAP Transaction Checker
Scans all .py files in the repo for SAP transaction codes, then opens SAP
and tries each transaction to verify it opens successfully.
"""

import os
import re
import sys
import time

REPO_ROOT = os.path.join(os.environ["USERPROFILE"], "Documents", "ADO", "Plant5")

# Add functions folder so we can import SAPManager
from pathlib import Path
functions_path = str(Path(__file__).resolve().parent.parent / "functions")
if functions_path not in sys.path:
    sys.path.append(functions_path)

from sap_connection import SAPManager


# ── 1. Scan for transaction codes ────────────────────────────────────────────

# Pattern 1: okcd").text = "TRANSACTION"
TCODE_PATTERN_OKCD = re.compile(r'okcd["\'\)]*\.text\s*=\s*["\']([a-zA-Z0-9_/]+)["\']')

# Pattern 2: transaction="TCODE" (used by run_extract interface)
TCODE_PATTERN_PARAM = re.compile(r'transaction\s*=\s*["\']([a-zA-Z0-9_]+)["\']', re.IGNORECASE)

# Navigation commands to skip (not real transactions)
SKIP_TCODES = {"/n", "/nex", "/i"}

# Directories to skip when scanning
SKIP_DIRS = {"functions", "templates", "tools"}


def find_all_transactions(root_dir):
    """Walk all .py files and return {tcode: [list of files using it]}."""
    tcode_map = {}

    for dirpath, dirnames, filenames in os.walk(root_dir):
        dirnames[:] = [d for d in dirnames if d not in SKIP_DIRS]
        for fname in filenames:
            if not fname.endswith(".py"):
                continue
            filepath = os.path.join(dirpath, fname)
            try:
                with open(filepath, "r", encoding="utf-8") as f:
                    content = f.read()
            except Exception:
                continue

            # Skip commented-out lines
            for line in content.splitlines():
                stripped = line.strip()
                if stripped.startswith("#"):
                    continue
                matches = TCODE_PATTERN_OKCD.findall(stripped) + TCODE_PATTERN_PARAM.findall(stripped)
                for tcode in matches:
                    tcode = tcode.upper()
                    if tcode in SKIP_TCODES:
                        continue
                    rel_path = os.path.relpath(filepath, root_dir)
                    tcode_map.setdefault(tcode, [])
                    if rel_path not in tcode_map[tcode]:
                        tcode_map[tcode].append(rel_path)

    return tcode_map


# ── 2. Test each transaction ─────────────────────────────────────────────────

def test_transaction(session, sap, tcode):
    """Try to open a transaction and check if it errors or opens successfully.
    Returns (success: bool, message: str).
    """
    try:
        session.findById("wnd[0]/tbar[0]/okcd").text = tcode
        session.findById("wnd[0]").sendVKey(0)
        time.sleep(1)

        # Check the status bar for errors
        try:
            statusbar = session.findById("wnd[0]/sbar")
            msg_type = statusbar.MessageType  # E=Error, S=Success, W=Warning, I=Info
            msg_text = statusbar.Text
        except Exception:
            msg_type = ""
            msg_text = ""

        # Check window title
        title = ""
        try:
            title = session.findById("wnd[0]").text
        except Exception:
            pass

        # If status bar shows error, the transaction failed
        if msg_type == "E":
            sap.return_to_main_menu(session)
            return False, f"Error: {msg_text}"

        # If we're still on Easy Access with no screen change, likely failed
        if "Easy Access" in title and msg_type != "S":
            return False, f"Transaction did not open (stayed on main menu). {msg_text}"

        # Transaction opened successfully - go back to main menu
        sap.return_to_main_menu(session)
        return True, f"OK - opened screen: {title}"

    except Exception as e:
        # Try to recover to main menu
        try:
            sap.return_to_main_menu(session)
        except Exception:
            pass
        return False, f"Exception: {e}"


# ── 3. Main ──────────────────────────────────────────────────────────────────

def main():
    print("=" * 65)
    print("         SAP Transaction Checker")
    print("=" * 65)
    print()

    # Step 1: Find all transactions
    print("Scanning repository for SAP transaction codes...")
    tcode_map = find_all_transactions(REPO_ROOT)

    if not tcode_map:
        print("No SAP transaction codes found in the repository.")
        return

    print(f"\nFound {len(tcode_map)} unique transaction(s):\n")
    for tcode, files in sorted(tcode_map.items()):
        print(f"  {tcode:<20} used in: {', '.join(files)}")

    print()

    # Step 2: Ask which to test
    print("Options:")
    print("  1) Test ALL transactions")
    print("  2) Pick specific ones to test")
    print("  3) Exit")
    print()

    choice = input("Select (1/2/3): ").strip()
    if choice == "3":
        return

    tcodes_to_test = sorted(tcode_map.keys())

    if choice == "2":
        print()
        for i, tcode in enumerate(tcodes_to_test, 1):
            print(f"  {i}) {tcode}")
        print()
        picks = input("Enter numbers separated by commas (e.g. 1,3): ").strip()
        try:
            indices = [int(x.strip()) - 1 for x in picks.split(",")]
            tcodes_to_test = [tcodes_to_test[i] for i in indices]
        except (ValueError, IndexError):
            print("Invalid selection. Exiting.")
            return

    # Step 3: Connect to SAP
    print()
    print("Connecting to SAP...")
    try:
        sap = SAPManager()
        session = sap.get_session()
        print("SAP session ready.\n")
    except Exception as e:
        print(f"Failed to connect to SAP: {e}")
        return

    # Step 4: Test each transaction
    print("-" * 65)
    results = []
    for tcode in tcodes_to_test:
        print(f"  Testing: {tcode:<20} ... ", end="", flush=True)
        success, message = test_transaction(session, sap, tcode)
        status = "PASS" if success else "FAIL"
        print(f"[{status}] {message}")
        results.append((tcode, success, message))
        time.sleep(0.5)

    # Step 5: Close SAP connection
    print("-" * 65)
    sap.close_connection(session)

    # Step 6: Summary
    passed = sum(1 for _, s, _ in results if s)
    failed = sum(1 for _, s, _ in results if not s)

    print(f"\nResults: {passed} passed, {failed} failed out of {len(results)} transaction(s)\n")

    if failed:
        print("Failed transactions:")
        for tcode, success, message in results:
            if not success:
                files = ", ".join(tcode_map.get(tcode, []))
                print(f"  [{tcode}] {message}")
                print(f"    Used in: {files}")
        print()


if __name__ == "__main__":
    main()

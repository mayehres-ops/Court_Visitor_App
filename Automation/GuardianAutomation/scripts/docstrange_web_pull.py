# docstrange_web_pull.py — attach to your real Edge profile via CDP (no persistent context)
# Steps: close Edge → run this script → Edge opens with your real profile → sign in / proceed.

import os
import sys
import time
import socket
import subprocess
from pathlib import Path
from playwright.sync_api import sync_playwright

EDGE_PORT = int(os.environ.get("EDGE_REMOTE_PORT", "9222"))
EDGE_PROFILE_NAME = os.environ.get("EDGE_PROFILE_NAME", "Default")  # change to "Profile 1" etc if needed

def _find_edge_exe():
    candidates = [
        r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        r"C:\Program Files\Microsoft\Edge\Application\msedge.exe",
    ]
    for p in candidates:
        if Path(p).exists():
            return p
    raise FileNotFoundError("Could not find msedge.exe in the usual locations.")

def _port_is_open(host: str, port: int) -> bool:
    try:
        with socket.create_connection((host, port), timeout=0.4):
            return True
    except OSError:
        return False

def _launch_edge_detached(edge_exe: str, port: int, profile_name: str):
    # Use your real profile; do NOT pass an "empty" user-data-dir.
    # We only add remote debugging and the profile selector.
    args = [
        edge_exe,
        f"--remote-debugging-port={port}",
        f"--profile-directory={profile_name}",
        "about:blank",
    ]
    # Spawn detached so the console isn’t tied to Edge lifetime.
    DETACHED_PROCESS = 0x00000008
    subprocess.Popen(args, creationflags=DETACHED_PROCESS)

def main():
    print("✅ Starting Edge with your REAL profile (CDP attach).")
    print("   Make sure ALL Edge windows are CLOSED before running this.\n")

    edge_exe = _find_edge_exe()

    # If something else already listens on the port, we’ll just try to attach.
    if not _port_is_open("127.0.0.1", EDGE_PORT):
        print(f"➡️  Launching Edge: {edge_exe}")
        print(f"    Profile: {EDGE_PROFILE_NAME} | Debug port: {EDGE_PORT}")
        _launch_edge_detached(edge_exe, EDGE_PORT, EDGE_PROFILE_NAME)

    # Wait for Edge to open the debugging port
    print("⏳ Waiting for Edge remote debugging port…")
    for _ in range(50):  # ~20 seconds max
        if _port_is_open("127.0.0.1", EDGE_PORT):
            break
        time.sleep(0.4)
    else:
        print("❌ Edge did not open the debugging port. Close Edge and try again.")
        sys.exit(1)

    with sync_playwright() as pw:
        url = f"http://127.0.0.1:{EDGE_PORT}"
        print(f"🔌 Connecting over CDP to {url} …")
        browser = pw.chromium.connect_over_cdp(url)

        # Use the first existing context (persistent profile). If none, create one.
        context = browser.contexts[0] if browser.contexts else browser.new_context()
        page = context.pages[0] if context.pages else context.new_page()

        # Go to DocStrange landing. This should be a normal Edge window with your profile.
        print("🌐 Navigating to DocStrange…")
        page.goto("https://docstrange.nanonets.com/", wait_until="domcontentloaded")

        print("\n🎉 Edge is open on DocStrange using your real profile.")
        print("   If you see a Google sign-in prompt, complete it in the Edge window.")
        print("   When you’re on the tool page (Drag & drop / Select Output Format), you’re set.\n")

        # Keep the script alive briefly for convenience; Edge stays open regardless.
        for i in range(5, 0, -1):
            print(f"Closing Playwright connection in {i}s (Edge stays open)…", end="\r")
            time.sleep(1)
        print()

        # Just detach; DO NOT close the Edge browser we started.
        browser.close()

if __name__ == "__main__":
    main()


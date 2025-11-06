#!/usr/bin/env python3
"""
Court Visitor App - Setup Wizard (GUI)
Checks and installs all dependencies with a user-friendly interface
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import subprocess
import sys
import os
import webbrowser
from pathlib import Path
import threading

class SetupWizard:
    def __init__(self, root):
        self.root = root
        self.root.title("Court Visitor App - Setup Wizard")
        self.root.geometry("700x600")
        self.root.resizable(False, False)

        # Dependency status
        self.dependencies = {
            'python': {'name': 'Python 3.10+', 'status': 'checking', 'required': True},
            'pip': {'name': 'Pip (Python Package Manager)', 'status': 'checking', 'required': True},
            'packages': {'name': 'Python Packages', 'status': 'checking', 'required': True},
            'tesseract': {'name': 'Tesseract OCR', 'status': 'checking', 'required': True},
            'poppler': {'name': 'Poppler PDF Tools', 'status': 'checking', 'required': True},
            'word': {'name': 'Microsoft Word', 'status': 'checking', 'required': False},
        }

        self.setup_ui()
        self.check_dependencies()

    def setup_ui(self):
        # Header
        header_frame = tk.Frame(self.root, bg='#667eea', height=100)
        header_frame.pack(fill='x')
        header_frame.pack_propagate(False)

        title = tk.Label(header_frame, text="Court Visitor App",
                        font=('Segoe UI', 24, 'bold'),
                        bg='#667eea', fg='white')
        title.pack(pady=15)

        subtitle = tk.Label(header_frame, text="Setup Wizard",
                           font=('Segoe UI', 12),
                           bg='#667eea', fg='white')
        subtitle.pack()

        # Main content
        content_frame = tk.Frame(self.root, padx=30, pady=20)
        content_frame.pack(fill='both', expand=True)

        # Instructions
        instructions = tk.Label(content_frame,
                               text="Checking system dependencies...",
                               font=('Segoe UI', 11))
        instructions.pack(pady=(0, 15))

        # Dependencies list
        self.dep_frame = tk.Frame(content_frame)
        self.dep_frame.pack(fill='both', expand=True)

        self.dep_labels = {}
        for key, dep in self.dependencies.items():
            frame = tk.Frame(self.dep_frame)
            frame.pack(fill='x', pady=5)

            # Status icon
            icon_label = tk.Label(frame, text="⏳", font=('Segoe UI', 14), width=3)
            icon_label.pack(side='left')

            # Dependency name
            name_label = tk.Label(frame, text=dep['name'],
                                 font=('Segoe UI', 10), anchor='w')
            name_label.pack(side='left', fill='x', expand=True)

            # Action button (hidden initially)
            action_btn = tk.Button(frame, text="Install",
                                  command=lambda k=key: self.install_dependency(k))
            action_btn.pack(side='right')
            action_btn.pack_forget()

            self.dep_labels[key] = {
                'icon': icon_label,
                'name': name_label,
                'button': action_btn,
                'frame': frame
            }

        # Log output
        log_label = tk.Label(content_frame, text="Installation Log:",
                            font=('Segoe UI', 10, 'bold'))
        log_label.pack(pady=(15, 5), anchor='w')

        self.log_text = scrolledtext.ScrolledText(content_frame,
                                                  height=8,
                                                  font=('Consolas', 9),
                                                  bg='#f5f5f5')
        self.log_text.pack(fill='both')

        # Bottom buttons
        button_frame = tk.Frame(self.root, padx=30, pady=15)
        button_frame.pack(fill='x', side='bottom')

        self.finish_btn = tk.Button(button_frame, text="Finish",
                                    command=self.finish_setup,
                                    state='disabled',
                                    font=('Segoe UI', 10, 'bold'),
                                    bg='#667eea', fg='white',
                                    padx=20, pady=5)
        self.finish_btn.pack(side='right')

        self.recheck_btn = tk.Button(button_frame, text="Re-check",
                                     command=self.check_dependencies,
                                     font=('Segoe UI', 10),
                                     padx=15, pady=5)
        self.recheck_btn.pack(side='right', padx=10)

    def log(self, message):
        """Add message to log"""
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def update_status(self, key, status):
        """Update dependency status"""
        self.dependencies[key]['status'] = status
        label = self.dep_labels[key]

        if status == 'ok':
            label['icon'].config(text="✅")
            label['button'].pack_forget()
        elif status == 'missing':
            label['icon'].config(text="❌")
            label['button'].pack(side='right')
        elif status == 'warning':
            label['icon'].config(text="⚠️")
            label['button'].pack(side='right')
        else:
            label['icon'].config(text="⏳")

        self.check_if_complete()

    def check_dependencies(self):
        """Check all dependencies in background thread"""
        self.log("Starting dependency check...")
        self.recheck_btn.config(state='disabled')

        thread = threading.Thread(target=self._check_dependencies_thread, daemon=True)
        thread.start()

    def _check_dependencies_thread(self):
        """Background thread for checking dependencies"""

        # 1. Check Python
        self.log("\n[1/6] Checking Python...")
        try:
            result = subprocess.run(['python', '--version'],
                                   capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version = result.stdout.strip()
                self.log(f"  ✓ {version}")

                # Check version >= 3.10
                version_parts = version.split()[1].split('.')
                major, minor = int(version_parts[0]), int(version_parts[1])

                if major >= 3 and minor >= 10:
                    self.update_status('python', 'ok')
                else:
                    self.log(f"  ✗ Python version too old (need 3.10+)")
                    self.update_status('python', 'missing')
            else:
                self.log("  ✗ Python not found")
                self.update_status('python', 'missing')
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.update_status('python', 'missing')

        # 2. Check Pip
        self.log("\n[2/6] Checking pip...")
        try:
            result = subprocess.run(['python', '-m', 'pip', '--version'],
                                   capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                self.log(f"  ✓ {result.stdout.strip()}")
                self.update_status('pip', 'ok')
            else:
                self.log("  ✗ Pip not found")
                self.update_status('pip', 'missing')
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.update_status('pip', 'missing')

        # 3. Check Python packages
        self.log("\n[3/6] Checking Python packages...")
        required_packages = ['openpyxl', 'pandas', 'pytesseract', 'pdf2image',
                            'pdfplumber', 'PIL', 'google-auth', 'pywin32']

        missing_packages = []
        for package in required_packages:
            pkg_name = 'pillow' if package == 'PIL' else package
            result = subprocess.run(['python', '-m', 'pip', 'show', pkg_name],
                                   capture_output=True, text=True)
            if result.returncode != 0:
                missing_packages.append(package)
                self.log(f"  ✗ {package} not installed")
            else:
                self.log(f"  ✓ {package} installed")

        if missing_packages:
            self.log(f"\n  Missing {len(missing_packages)} package(s)")
            self.update_status('packages', 'missing')
        else:
            self.log("  ✓ All packages installed")
            self.update_status('packages', 'ok')

        # 4. Check Tesseract
        self.log("\n[4/6] Checking Tesseract OCR...")
        try:
            result = subprocess.run(['tesseract', '--version'],
                                   capture_output=True, text=True, timeout=5)
            if result.returncode == 0:
                version_line = result.stdout.split('\n')[0]
                self.log(f"  ✓ {version_line}")
                self.update_status('tesseract', 'ok')
            else:
                self.log("  ✗ Tesseract not found in PATH")
                self.update_status('tesseract', 'missing')
        except FileNotFoundError:
            self.log("  ✗ Tesseract not installed")
            self.update_status('tesseract', 'missing')
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.update_status('tesseract', 'missing')

        # 5. Check Poppler
        self.log("\n[5/6] Checking Poppler...")
        try:
            result = subprocess.run(['pdftoppm', '-v'],
                                   capture_output=True, text=True, timeout=5)
            # pdftoppm returns version info on stderr
            if 'poppler' in result.stderr.lower():
                self.log("  ✓ Poppler installed")
                self.update_status('poppler', 'ok')
            else:
                self.log("  ✗ Poppler not found")
                self.update_status('poppler', 'missing')
        except FileNotFoundError:
            self.log("  ✗ Poppler not installed")
            self.update_status('poppler', 'missing')
        except Exception as e:
            self.log(f"  ✗ Error: {e}")
            self.update_status('poppler', 'missing')

        # 6. Check Microsoft Word
        self.log("\n[6/6] Checking Microsoft Word...")
        try:
            result = subprocess.run([sys.executable, '-c',
                                    'import win32com.client; word = win32com.client.Dispatch("Word.Application"); word.Quit()'],
                                   capture_output=True, text=True, timeout=10)
            if result.returncode == 0:
                self.log("  ✓ Microsoft Word installed")
                self.update_status('word', 'ok')
            else:
                self.log("  ⚠ Microsoft Word not found (optional)")
                self.update_status('word', 'warning')
        except Exception as e:
            self.log(f"  ⚠ Microsoft Word not found (optional)")
            self.update_status('word', 'warning')

        self.log("\n✓ Dependency check complete!")
        self.root.after(0, lambda: self.recheck_btn.config(state='normal'))

    def install_dependency(self, key):
        """Install or provide instructions for a dependency"""

        if key == 'python':
            messagebox.showinfo("Install Python",
                               "Please download and install Python 3.10 or higher from:\n\n"
                               "https://www.python.org/downloads/windows/\n\n"
                               "IMPORTANT: Check 'Add Python to PATH' during installation!")
            webbrowser.open("https://www.python.org/downloads/windows/")

        elif key == 'pip':
            self.log("\nUpgrading pip...")
            self.run_command([sys.executable, '-m', 'ensurepip', '--upgrade'])
            self.check_dependencies()

        elif key == 'packages':
            self.log("\nInstalling Python packages...")
            self.log("This may take 2-5 minutes...\n")
            requirements_file = Path(__file__).parent / "requirements.txt"
            if requirements_file.exists():
                self.run_command([sys.executable, '-m', 'pip', 'install', '-r', str(requirements_file)])
            else:
                self.log("Error: requirements.txt not found!")
            self.check_dependencies()

        elif key == 'tesseract':
            messagebox.showinfo("Install Tesseract OCR",
                               "Tesseract OCR is required for PDF text extraction.\n\n"
                               "Instructions:\n"
                               "1. Click OK to open the download page\n"
                               "2. Download 'tesseract-ocr-w64-setup-5.x.x.exe'\n"
                               "3. Run the installer\n"
                               "4. Use the default installation path\n"
                               "5. Click 'Re-check' in this wizard after installation")
            webbrowser.open("https://github.com/UB-Mannheim/tesseract/wiki")

        elif key == 'poppler':
            messagebox.showinfo("Install Poppler",
                               "Poppler is required for PDF to image conversion.\n\n"
                               "Instructions:\n"
                               "1. Click OK to open the download page\n"
                               "2. Download the latest Release ZIP\n"
                               "3. Extract to C:\\poppler\\\n"
                               "4. Add C:\\poppler\\Library\\bin to PATH\n"
                               "5. Click 'Re-check' in this wizard after installation")
            webbrowser.open("https://github.com/oschwartz10612/poppler-windows/releases")

        elif key == 'word':
            messagebox.showinfo("Install Microsoft Word",
                               "Microsoft Word is required for document generation.\n\n"
                               "Please install Microsoft Office if you haven't already.")

    def run_command(self, cmd):
        """Run command and log output"""
        try:
            process = subprocess.Popen(cmd, stdout=subprocess.PIPE,
                                      stderr=subprocess.STDOUT, text=True)
            for line in process.stdout:
                self.log(line.rstrip())
            process.wait()
            if process.returncode == 0:
                self.log("\n✓ Command completed successfully")
            else:
                self.log(f"\n✗ Command failed with exit code {process.returncode}")
        except Exception as e:
            self.log(f"\n✗ Error: {e}")

    def check_if_complete(self):
        """Check if all required dependencies are installed"""
        all_ok = True
        for key, dep in self.dependencies.items():
            if dep['required'] and dep['status'] != 'ok':
                all_ok = False
                break

        if all_ok:
            self.finish_btn.config(state='normal')
        else:
            self.finish_btn.config(state='disabled')

    def finish_setup(self):
        """Complete setup"""
        messagebox.showinfo("Setup Complete!",
                           "All required dependencies are installed!\n\n"
                           "Next steps:\n"
                           "1. Read INSTALLATION_GUIDE.md\n"
                           "2. Setup Google API credentials\n"
                           "3. Add credentials to Config\\API\\\n"
                           "4. Launch the app\n\n"
                           "The wizard will now close.")
        self.root.destroy()


def main():
    root = tk.Tk()
    app = SetupWizard(root)
    root.mainloop()


if __name__ == "__main__":
    main()

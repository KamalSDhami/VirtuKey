#!/usr/bin/env python3
"""
VirtuKey Installer - Simple and Clean Multi-Step GUI Installer
Author: KamalSDhami
Version: 1.0
"""

import tkinter as tk
from tkinter import messagebox, filedialog
import os
import shutil
from pathlib import Path
import sys
import winreg
import subprocess
import time
import ctypes

# Try to import psutil, fall back to subprocess if not available
try:
    import psutil
    PSUTIL_AVAILABLE = True
except ImportError:
    PSUTIL_AVAILABLE = False

class VirtuKeyInstaller:
    def __init__(self):
        # Without this, Windows bitmap-stretches the window and text is blurry
        try:
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass
        self.root = tk.Tk()
        self.scale = max(1.0, self.root.winfo_fpixels('1i') / 96)
        self.root.resizable(False, False)
        self.root.configure(bg='#f8fafc')  # Light gray background
        
        # Modern color scheme
        self.colors = {
            'primary': '#2563eb',      # Modern blue
            'primary_hover': '#1d4ed8',
            'secondary': '#64748b',    # Slate gray
            'success': '#059669',      # Emerald green
            'danger': '#dc2626',       # Red
            'warning': '#d97706',      # Amber
            'background': '#f8fafc',   # Very light gray
            'surface': '#ffffff',      # White
            'card': '#ffffff',         # White for cards
            'border': '#e2e8f0',       # Light border
            'text_primary': '#1e293b', # Dark text
            'text_secondary': '#64748b', # Gray text
            'text_muted': '#94a3b8'    # Light gray text
        }
        
        # Installation variables
        default_path = os.path.join(os.path.expanduser("~"), "AppData", "Local", "VirtuKey")
        self.install_path = tk.StringVar(value=default_path)
        self.create_desktop_shortcut = tk.BooleanVar(value=True)
        self.create_startmenu_shortcut = tk.BooleanVar(value=True)
        self.auto_start = tk.BooleanVar(value=False)
        
        # Check if already installed
        self.is_installed = self.check_installation()
        self.mode = "uninstall" if self.is_installed else "install"  # install or uninstall
        
        # Set window title based on mode
        title = "VirtuKey Uninstaller" if self.is_installed else "VirtuKey Setup"
        self.root.title(title)
        
        # Current step (0-4)
        self.current_step = 0
        self.total_steps = 5
        
        # Setup the UI
        self.create_ui()
        self.show_step(0)
        
    def px(self, value):
        """Scale a pixel size for the current display DPI"""
        return int(value * self.scale)
        
    def create_ui(self):
        # Modern header with gradient-like appearance - reduced height
        header_frame = tk.Frame(self.root, bg=self.colors['primary'], height=self.px(80))
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        # Header content container - reduced padding
        header_content = tk.Frame(header_frame, bg=self.colors['primary'])
        header_content.pack(expand=True, fill=tk.BOTH, padx=30, pady=15)
        
        # Logo and title with modern typography - smaller fonts
        title_label = tk.Label(header_content, text="VirtuKey", 
                              bg=self.colors['primary'], fg='white', 
                              font=('Segoe UI', 20, 'bold'))
        title_label.pack(side=tk.LEFT, pady=5)
        
        subtitle_label = tk.Label(header_content, text="Setup", 
                                 bg=self.colors['primary'], fg='#e2e8f0', 
                                 font=('Segoe UI', 12))
        subtitle_label.pack(side=tk.LEFT, padx=(8, 0), pady=8)
        
        # Progress indicator (modern step dots)
        self.progress_frame = tk.Frame(header_content, bg=self.colors['primary'])
        self.progress_frame.pack(side=tk.RIGHT, pady=8)
        
        self.progress_dots = []
        for i in range(self.total_steps):
            dot_frame = tk.Frame(self.progress_frame, bg=self.colors['primary'], width=12, height=12)
            dot_frame.pack(side=tk.LEFT, padx=3)
            dot_frame.pack_propagate(False)
            
            dot = tk.Label(dot_frame, text="●", bg=self.colors['primary'], 
                          fg='#94a3b8', font=('Segoe UI', 8))
            dot.pack(expand=True)
            self.progress_dots.append(dot)
        
        # Main content area with card-like appearance - further reduced padding
        content_container = tk.Frame(self.root, bg=self.colors['background'])
        content_container.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        # Content card with shadow effect (simulated with borders)
        content_shadow = tk.Frame(content_container, bg='#d1d5db', height=2)
        content_shadow.pack(fill=tk.X, pady=(2, 0))
        
        self.content_frame = tk.Frame(content_container, bg=self.colors['surface'], 
                                     relief='flat', bd=0)
        self.content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 2))
        
        # Button area with modern styling - further reduced padding
        button_container = tk.Frame(self.root, bg=self.colors['background'])
        button_container.pack(fill=tk.X, padx=15, pady=(0, 10))
        
        # Button card - reduced height
        button_shadow = tk.Frame(button_container, bg='#d1d5db', height=2)
        button_shadow.pack(fill=tk.X, pady=(2, 0))
        
        button_frame = tk.Frame(button_container, bg=self.colors['surface'], height=self.px(64))
        button_frame.pack(fill=tk.X, pady=(0, 2))
        button_frame.pack_propagate(False)
        
        # Button content with proper spacing - further reduced padding
        button_inner = tk.Frame(button_frame, bg=self.colors['surface'])
        button_inner.pack(expand=True, fill=tk.BOTH, padx=20, pady=12)
        
        # Modern styled buttons
        self.back_button = tk.Button(button_inner, text="← Back", 
                                    command=self.go_back,
                                    bg='#f1f5f9', fg=self.colors['text_secondary'],
                                    font=('Segoe UI', 10, 'bold'),
                                    padx=30, pady=12, relief='flat', bd=0,
                                    state=tk.DISABLED, cursor='hand2')
        # Don't pack initially - will be managed by show_step
        
        # Style disabled button
        self.back_button.configure(bg='#f8fafc', fg='#cbd5e1')
        
        self.cancel_button = tk.Button(button_inner, text="Cancel", 
                                      command=self.cancel_installation,
                                      bg='#fef2f2', fg=self.colors['danger'],
                                      font=('Segoe UI', 10, 'bold'),
                                      padx=30, pady=12, relief='flat', bd=0,
                                      cursor='hand2')
        # Don't pack initially - will be managed by show_step
        
        self.next_button = tk.Button(button_inner, text="Next →", 
                                    command=self.go_next,
                                    bg=self.colors['primary'], fg='white',
                                    font=('Segoe UI', 10, 'bold'),
                                    padx=30, pady=12, relief='flat', bd=0,
                                    cursor='hand2')
        self.next_button.pack(side=tk.RIGHT)  # Next button always on the right
        
        # Add hover effects
        self.add_button_hover_effects()
        
    def add_button_hover_effects(self):
        """Add modern hover effects to buttons"""
        def on_enter_next(e):
            if self.next_button['state'] != 'disabled':
                self.next_button.configure(bg=self.colors['primary_hover'])
        
        def on_leave_next(e):
            if self.next_button['state'] != 'disabled':
                self.next_button.configure(bg=self.colors['primary'])
        
        def on_enter_back(e):
            if self.back_button['state'] != 'disabled':
                self.back_button.configure(bg='#e2e8f0')
        
        def on_leave_back(e):
            if self.back_button['state'] != 'disabled':
                self.back_button.configure(bg='#f1f5f9')
        
        def on_enter_cancel(e):
            self.cancel_button.configure(bg='#fee2e2')
        
        def on_leave_cancel(e):
            self.cancel_button.configure(bg='#fef2f2')
        
        self.next_button.bind("<Enter>", on_enter_next)
        self.next_button.bind("<Leave>", on_leave_next)
        self.back_button.bind("<Enter>", on_enter_back)
        self.back_button.bind("<Leave>", on_leave_back)
        self.cancel_button.bind("<Enter>", on_enter_cancel)
        self.cancel_button.bind("<Leave>", on_leave_cancel)
        
    def update_progress_dots(self):
        """Update progress indicator dots"""
        for i, dot in enumerate(self.progress_dots):
            if i < self.current_step:
                dot.configure(fg='#10b981')  # Completed - green
            elif i == self.current_step:
                dot.configure(fg='white')     # Current - white
            else:
                dot.configure(fg='#94a3b8')  # Future - gray
        
    def check_installation(self):
        """Check if VirtuKey is already installed"""
        install_dir = Path(self.install_path.get())
        exe_file = install_dir / "VirtuKey.exe"
        dll_file = install_dir / "VirtualDesktopAccessor.dll"
        
        # Check if installation directory exists and contains main files
        if install_dir.exists() and (exe_file.exists() or dll_file.exists()):
            return True
        return False
        
    def clear_content(self):
        """Clear the content frame"""
        for widget in self.content_frame.winfo_children():
            widget.destroy()
            
    def show_step(self, step):
        """Show the specified installation step"""
        self.current_step = step
        self.clear_content()
        
        # Update progress dots
        self.update_progress_dots()
        
        # Update mode based on user selection if on welcome page
        if step == 0 and self.is_installed and hasattr(self, 'action_mode'):
            if self.action_mode.get() == "reinstall":
                self.mode = "reinstall"
            else:
                self.mode = "uninstall"
        
        # Update buttons based on mode with modern styling
        if step > 0:
            # Pages 2+ - Show Back and Next, hide Cancel
            self.back_button.config(state=tk.NORMAL, bg='#f1f5f9', fg=self.colors['text_secondary'])
            self.back_button.pack(side=tk.LEFT)  # Show Back button on left
            self.cancel_button.pack_forget()  # Hide Cancel button
        else:
            # Page 1 - Show Cancel and Next, hide Back
            self.back_button.pack_forget()  # Hide Back button
            self.cancel_button.pack(side=tk.LEFT)  # Show Cancel button on left
        
        # Reset Next button state first
        self.next_button.config(state=tk.NORMAL)
        
        if step == self.total_steps - 1:  # Last step
            self.next_button.config(text="Finish", command=self.finish_installation,
                                   bg=self.colors['success'], fg='white')
        elif step == self.total_steps - 2:  # Installation step
            if self.mode == "uninstall":
                self.next_button.config(text="Uninstall", command=self.start_uninstallation,
                                       bg=self.colors['danger'], fg='white')
            elif self.mode == "reinstall":
                self.next_button.config(text="Reinstall", command=self.start_reinstallation,
                                       bg=self.colors['warning'], fg='white')
            else:
                self.next_button.config(text="Install", command=self.start_installation,
                                       bg=self.colors['success'], fg='white')
        else:
            self.next_button.config(text="Next →", command=self.go_next,
                                   bg=self.colors['primary'], fg='white')
        
        # Show the appropriate content based on mode
        if step == 0:
            self.show_welcome()
        elif step == 1:
            if self.mode == "uninstall":
                self.show_uninstall_options()
            else:
                self.show_license()
                # Disable Next button initially for license page
                self.next_button.config(state=tk.DISABLED, bg='#cbd5e1', fg='#9ca3af')
        elif step == 2:
            if self.mode == "uninstall":
                self.show_uninstall_summary()
            else:
                self.show_options()
        elif step == 3:
            self.show_installation()
        elif step == 4:
            self.show_complete()
    
    def show_welcome(self):
        """Welcome page with modern styling"""
        # Main container with further reduced padding
        container = tk.Frame(self.content_frame, bg=self.colors['surface'])
        container.pack(expand=True, fill=tk.BOTH, padx=20, pady=15)
        
        if self.mode == "uninstall":
            # Modern uninstall mode
            # Icon/Header section
            header_section = tk.Frame(container, bg=self.colors['surface'])
            header_section.pack(fill=tk.X, pady=(0, 15))
            
            # Large icon
            icon_label = tk.Label(header_section, text="⚠️", 
                                 bg=self.colors['surface'], font=('Segoe UI', 30))
            icon_label.pack(pady=(0, 8))
            
            title = tk.Label(header_section, text="VirtuKey is already installed", 
                            bg=self.colors['surface'], fg=self.colors['text_primary'], 
                            font=('Segoe UI', 18, 'bold'))
            title.pack(pady=(0, 8))
            
            subtitle = tk.Label(header_section, text="Choose what you'd like to do:", 
                               bg=self.colors['surface'], fg=self.colors['text_secondary'], 
                               font=('Segoe UI', 11))
            subtitle.pack()
            
            # Mode selection cards
            options_frame = tk.Frame(container, bg=self.colors['surface'])
            options_frame.pack(fill=tk.X, pady=(12, 15))
            
            self.action_mode = tk.StringVar(value="uninstall")
            
            # Uninstall option card
            uninstall_card = tk.Frame(options_frame, bg='#fef2f2', relief='flat', bd=1)
            uninstall_card.pack(fill=tk.X, pady=(0, 8), padx=12)
            
            uninstall_rb = tk.Radiobutton(uninstall_card, text="🗑️  Uninstall VirtuKey", 
                                         variable=self.action_mode, value="uninstall",
                                         bg='#fef2f2', fg=self.colors['text_primary'],
                                         font=('Segoe UI', 10, 'bold'), relief='flat')
            uninstall_rb.pack(anchor=tk.W, pady=8, padx=12)
            
            uninstall_desc = tk.Label(uninstall_card, 
                                     text="Remove VirtuKey from your computer completely",
                                     bg='#fef2f2', fg=self.colors['text_secondary'],
                                     font=('Segoe UI', 8))
            uninstall_desc.pack(anchor=tk.W, padx=12, pady=(0, 8))
            
            # Reinstall option card
            reinstall_card = tk.Frame(options_frame, bg='#eff6ff', relief='flat', bd=1)
            reinstall_card.pack(fill=tk.X, padx=12)
            
            reinstall_rb = tk.Radiobutton(reinstall_card, text="🔄  Reinstall VirtuKey", 
                                         variable=self.action_mode, value="reinstall",
                                         bg='#eff6ff', fg=self.colors['text_primary'],
                                         font=('Segoe UI', 10, 'bold'), relief='flat')
            reinstall_rb.pack(anchor=tk.W, pady=8, padx=12)
            
            reinstall_desc = tk.Label(reinstall_card, 
                                     text="Remove current installation and install fresh copy",
                                     bg='#eff6ff', fg=self.colors['text_secondary'],
                                     font=('Segoe UI', 8))
            reinstall_desc.pack(anchor=tk.W, padx=12, pady=(0, 8))
            
            # Installation info card
            info_card = tk.Frame(container, bg='#f8fafc', relief='flat', bd=1)
            info_card.pack(fill=tk.X, pady=(12, 0), padx=12)
            
            info_title = tk.Label(info_card, text="Current Installation", 
                                 bg='#f8fafc', fg=self.colors['text_primary'],
                                 font=('Segoe UI', 9, 'bold'))
            info_title.pack(anchor=tk.W, padx=12, pady=(8, 4))
            
            install_info = tk.Label(info_card, 
                                   text=self.install_path.get(),
                                   bg='#f8fafc', fg=self.colors['text_secondary'],
                                   font=('Segoe UI', 8))
            install_info.pack(anchor=tk.W, padx=12, pady=(0, 8))
            
        else:
            # Modern install mode
            # Hero section
            hero_section = tk.Frame(container, bg=self.colors['surface'])
            hero_section.pack(fill=tk.X, pady=(0, 15))
            
            # Large welcome icon
            icon_label = tk.Label(hero_section, text="🚀", 
                                 bg=self.colors['surface'], font=('Segoe UI', 30))
            icon_label.pack(pady=(0, 10))
            
            title = tk.Label(hero_section, text="Welcome to VirtuKey", 
                            bg=self.colors['surface'], fg=self.colors['text_primary'], 
                            font=('Segoe UI', 18, 'bold'))
            title.pack(pady=(0, 5))
            
            subtitle = tk.Label(hero_section, text="Virtual Desktop Manager", 
                               bg=self.colors['surface'], fg=self.colors['text_secondary'], 
                               font=('Segoe UI', 11))
            subtitle.pack(pady=(0, 15))
            
            # Features section
            features_frame = tk.Frame(container, bg=self.colors['surface'])
            features_frame.pack(fill=tk.X, pady=(0, 15))
            
            features_title = tk.Label(features_frame, text="What you'll get:", 
                                     bg=self.colors['surface'], fg=self.colors['text_primary'],
                                     font=('Segoe UI', 13, 'bold'))
            features_title.pack(pady=(0, 8))  # Centered by default
            
            features = [
                ("🖥️", "Multiple virtual desktops"),
                ("⌨️", "Convenient hotkey shortcuts"),
                ("⚡", "Lightning-fast desktop switching"),
                ("🎯", "Better workflow organization")
            ]
            
            for icon, text in features:
                feature_frame = tk.Frame(features_frame, bg=self.colors['surface'])
                feature_frame.pack(pady=3)  # Center the entire feature frame
                
                # Create inner frame to hold icon and text together
                feature_content = tk.Frame(feature_frame, bg=self.colors['surface'])
                feature_content.pack()  # Center the content within the frame
                
                icon_label = tk.Label(feature_content, text=icon, 
                                     bg=self.colors['surface'], font=('Segoe UI', 11))
                icon_label.pack(side=tk.LEFT, padx=(0, 6))
                
                text_label = tk.Label(feature_content, text=text, 
                                     bg=self.colors['surface'], fg=self.colors['text_secondary'],
                                     font=('Segoe UI', 10))
                text_label.pack(side=tk.LEFT)
            
            # Call to action
            cta_frame = tk.Frame(container, bg=self.colors['surface'])
            cta_frame.pack(fill=tk.X, pady=(15, 0))
            
            cta_text = tk.Label(cta_frame, 
                               text="Click Next to begin the installation process.",
                               bg=self.colors['surface'], fg=self.colors['text_secondary'],
                               font=('Segoe UI', 10))
            cta_text.pack()
        
    def show_license(self):
        """License agreement page"""
        license_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        license_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=10)
        
        title = tk.Label(license_frame, text="License Agreement", 
                        bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 14, 'bold'))
        title.pack(pady=(0, 10))
        
        # Create a frame for the text area with fixed height
        text_container = tk.Frame(license_frame, bg=self.colors['surface'])
        text_container.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # License text with scrollbar - reduced height to fit page
        license_text = tk.Text(text_container, wrap=tk.WORD, height=12, 
                              font=('Consolas', 9), bg=self.colors['background'], 
                              yscrollcommand=lambda *args: scrollbar.set(*args))
        
        scrollbar = tk.Scrollbar(text_container, orient=tk.VERTICAL, 
                                command=license_text.yview)
        
        license_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        license_content = """MIT License

Copyright (c) 2024 KamalSDhami

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.

Additional Terms:
By installing VirtuKey, you acknowledge that this software is provided as-is 
and the developer is not responsible for any issues that may arise from its use.

This software is designed to enhance productivity through virtual desktop 
management and does not collect any personal data or transmit information 
to external servers."""
        
        license_text.insert(tk.END, license_content)
        license_text.config(state=tk.DISABLED, yscrollcommand=scrollbar.set)
        
        # Accept checkbox - fixed at bottom with padding
        checkbox_frame = tk.Frame(license_frame, bg=self.colors['surface'], height=40)
        checkbox_frame.pack(fill=tk.X, side=tk.BOTTOM)
        checkbox_frame.pack_propagate(False)
        
        self.accept_license = tk.BooleanVar(value=False)
        accept_cb = tk.Checkbutton(checkbox_frame, text="I accept the license agreement",
                                  variable=self.accept_license,
                                  bg=self.colors['surface'], font=('Segoe UI', 10, 'bold'),
                                  command=self.on_license_accept)
        accept_cb.pack(pady=8, anchor=tk.W)
        
        # Disable Next initially
        self.next_button.config(state=tk.DISABLED)
        
    def on_license_accept(self):
        """Enable/disable Next button based on license acceptance"""
        if self.accept_license.get():
            self.next_button.config(state=tk.NORMAL, bg=self.colors['primary'], fg='white')
        else:
            self.next_button.config(state=tk.DISABLED, bg='#cbd5e1', fg='#9ca3af')
            
    def show_options(self):
        """Installation options page"""
        options_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        options_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=8)
        
        title = tk.Label(options_frame, text="Installation Options", 
                        bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 13, 'bold'))
        title.pack(pady=(0, 8))
        
        # Installation path section
        path_section = tk.Frame(options_frame, bg=self.colors['surface'])
        path_section.pack(fill=tk.X, pady=(0, 8))
        
        path_label = tk.Label(path_section, text="Installation Directory:", 
                             bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 9, 'bold'))
        path_label.pack(anchor=tk.W, pady=(0, 2))
        
        path_frame = tk.Frame(path_section, bg=self.colors['surface'])
        path_frame.pack(fill=tk.X, pady=(0, 2))
        
        path_entry = tk.Entry(path_frame, textvariable=self.install_path, 
                             font=('Segoe UI', 9), width=40)
        path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 6))
        
        browse_btn = tk.Button(path_frame, text="Browse...", 
                              command=self.browse_folder,
                              bg='#95a5a6', fg='white', font=('Segoe UI', 8),
                              padx=10, pady=3)
        browse_btn.pack(side=tk.RIGHT)
        
        # Space info
        space_info = tk.Label(path_section, text="Space required: ~2.5 MB", 
                             bg=self.colors['surface'], fg=self.colors['text_muted'], font=('Segoe UI', 8))
        space_info.pack(anchor=tk.W, pady=(2, 0))
        
        # Location info
        location_info = tk.Label(path_section, text="Default location is in user folder (no admin rights required)", 
                                bg=self.colors['surface'], fg=self.colors['success'], font=('Segoe UI', 7))
        location_info.pack(anchor=tk.W, pady=(1, 0))
        
        # Separator
        separator = tk.Frame(options_frame, bg=self.colors['border'], height=1)
        separator.pack(fill=tk.X, pady=8)
        
        # Additional options
        options_label = tk.Label(options_frame, text="Additional Options:", 
                                bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 9, 'bold'))
        options_label.pack(anchor=tk.W, pady=(0, 6))
        
        # Options checkboxes - more compact spacing
        options_container = tk.Frame(options_frame, bg=self.colors['surface'])
        options_container.pack(fill=tk.X)
        
        desktop_cb = tk.Checkbutton(options_container, text="Create desktop shortcut",
                                   variable=self.create_desktop_shortcut,
                                   bg=self.colors['surface'], font=('Segoe UI', 9))
        desktop_cb.pack(anchor=tk.W, pady=1)
        
        startmenu_cb = tk.Checkbutton(options_container, text="Create Start Menu shortcuts",
                                     variable=self.create_startmenu_shortcut,
                                     bg=self.colors['surface'], font=('Segoe UI', 9))
        startmenu_cb.pack(anchor=tk.W, pady=1)
        
        autostart_cb = tk.Checkbutton(options_container, text="Start VirtuKey automatically with Windows",
                                     variable=self.auto_start,
                                     bg=self.colors['surface'], font=('Segoe UI', 9))
        autostart_cb.pack(anchor=tk.W, pady=1)
        
    def show_uninstall_options(self):
        """Uninstall options page"""
        uninstall_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        uninstall_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=8)
        
        title = tk.Label(uninstall_frame, text="Uninstall Options", 
                        bg=self.colors['surface'], fg=self.colors['danger'], font=('Segoe UI', 13, 'bold'))
        title.pack(pady=(0, 12))
        
        # Current installation info
        info_section = tk.Frame(uninstall_frame, bg=self.colors['surface'])
        info_section.pack(fill=tk.X, pady=(0, 12))
        
        info_label = tk.Label(info_section, text="Current Installation:", 
                             bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 9, 'bold'))
        info_label.pack(anchor=tk.W, pady=(0, 4))
        
        path_info = tk.Label(info_section, text=f"Location: {self.install_path.get()}", 
                            bg=self.colors['surface'], fg=self.colors['text_secondary'], font=('Segoe UI', 9))
        path_info.pack(anchor=tk.W, pady=1)
        
        # Separator
        separator = tk.Frame(uninstall_frame, bg=self.colors['border'], height=1)
        separator.pack(fill=tk.X, pady=12)
        
        # Removal options
        options_label = tk.Label(uninstall_frame, text="What to Remove:", 
                                bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 9, 'bold'))
        options_label.pack(anchor=tk.W, pady=(0, 8))
        
        # Initialize uninstall options
        self.remove_shortcuts = tk.BooleanVar(value=True)
        self.remove_settings = tk.BooleanVar(value=False)
        
        # Options checkboxes
        options_container = tk.Frame(uninstall_frame, bg=self.colors['surface'])
        options_container.pack(fill=tk.X)
        
        shortcuts_cb = tk.Checkbutton(options_container, text="Remove desktop and Start Menu shortcuts",
                                     variable=self.remove_shortcuts,
                                     bg=self.colors['surface'], font=('Segoe UI', 9))
        shortcuts_cb.pack(anchor=tk.W, pady=2)
        
        settings_cb = tk.Checkbutton(options_container, text="Remove user settings and configuration",
                                    variable=self.remove_settings,
                                    bg=self.colors['surface'], font=('Segoe UI', 9))
        settings_cb.pack(anchor=tk.W, pady=2)
        
        # Warning
        warning = tk.Label(uninstall_frame, 
                          text="⚠️ This action cannot be undone. VirtuKey will be completely removed from your system.",
                          bg=self.colors['surface'], fg=self.colors['danger'], font=('Segoe UI', 8, 'italic'))
        warning.pack(pady=12)
        
    def show_uninstall_summary(self):
        """Uninstall summary page"""
        summary_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        summary_frame.pack(expand=True, fill=tk.BOTH, padx=15, pady=15)
        
        title = tk.Label(summary_frame, text="Ready to Uninstall", 
                        bg=self.colors['surface'], fg=self.colors['danger'], font=('Segoe UI', 13, 'bold'))
        title.pack(pady=(0, 15))
        
        # Summary text
        summary_text = f"""VirtuKey will be uninstalled with the following settings:

Installation Directory: {self.install_path.get()}
Remove Shortcuts: {'Yes' if self.remove_shortcuts.get() else 'No'}
Remove Settings: {'Yes' if self.remove_settings.get() else 'No'}

Note: If VirtuKey is running, you will be asked to close it automatically.

Click Uninstall to begin the removal process."""
        
        summary_label = tk.Label(summary_frame, text=summary_text, 
                                bg=self.colors['surface'], fg=self.colors['text_secondary'], font=('Segoe UI', 9),
                                justify=tk.LEFT)
        summary_label.pack(pady=8, anchor=tk.W)
        
    def show_installation(self):
        """Installation progress page"""
        install_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        install_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        title = tk.Label(install_frame, text="Ready to Install", 
                        bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 14, 'bold'))
        title.pack(pady=(0, 20))
        
        # Installation summary
        summary_text = f"""VirtuKey will be installed with the following settings:

Installation Directory: {self.install_path.get()}
Desktop Shortcut: {'Yes' if self.create_desktop_shortcut.get() else 'No'}
Start Menu Shortcuts: {'Yes' if self.create_startmenu_shortcut.get() else 'No'}
Auto-start with Windows: {'Yes' if self.auto_start.get() else 'No'}

Click Install to begin the installation."""
        
        summary_label = tk.Label(install_frame, text=summary_text, 
                                bg=self.colors['surface'], fg=self.colors['text_secondary'], font=('Segoe UI', 10),
                                justify=tk.LEFT)
        summary_label.pack(pady=10, anchor=tk.W)
        
    def show_complete(self):
        """Installation/Uninstallation complete page"""
        complete_frame = tk.Frame(self.content_frame, bg=self.colors['surface'])
        complete_frame.pack(expand=True, fill=tk.BOTH, padx=20, pady=20)
        
        # Success icon (using text)
        success_icon = tk.Label(complete_frame, text="✓", 
                               bg=self.colors['surface'], fg=self.colors['success'], font=('Segoe UI', 48, 'bold'))
        success_icon.pack(pady=(20, 10))
        
        if self.mode == "uninstall":
            title = tk.Label(complete_frame, text="Uninstallation Complete!", 
                            bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 16, 'bold'))
            title.pack(pady=(0, 20))
            
            desc = tk.Label(complete_frame, 
                           text="VirtuKey has been successfully removed from your computer.\n\n"
                                "All selected components have been uninstalled.\n\n"
                                "Thank you for using VirtuKey!",
                           bg=self.colors['surface'], fg=self.colors['text_secondary'], font=('Segoe UI', 11),
                           justify=tk.CENTER)
            desc.pack(pady=10)
            
        else:
            title = tk.Label(complete_frame, text="Installation Complete!", 
                            bg=self.colors['surface'], fg=self.colors['text_primary'], font=('Segoe UI', 16, 'bold'))
            title.pack(pady=(0, 20))
            
            desc = tk.Label(complete_frame, 
                           text="VirtuKey has been successfully installed on your computer.\n\n"
                                "You can now start using VirtuKey to manage your virtual desktops.\n\n"
                                "Thank you for choosing VirtuKey!",
                           bg=self.colors['surface'], fg=self.colors['text_secondary'], font=('Segoe UI', 11),
                           justify=tk.CENTER)
            desc.pack(pady=10)
            
            # Launch option only for install/reinstall
            self.launch_now = tk.BooleanVar(value=True)
            launch_cb = tk.Checkbutton(complete_frame, text="Launch VirtuKey now",
                                      variable=self.launch_now,
                                      bg=self.colors['surface'], font=('Segoe UI', 10, 'bold'))
            launch_cb.pack(pady=20)
        
    def browse_folder(self):
        """Browse for installation folder"""
        folder = filedialog.askdirectory(initialdir=self.install_path.get())
        if folder:
            self.install_path.set(os.path.join(folder, "VirtuKey"))
            
    def go_back(self):
        """Go to previous step"""
        if self.current_step > 0:
            self.show_step(self.current_step - 1)
            
    def go_next(self):
        """Go to next step"""
        if self.current_step < self.total_steps - 1:
            self.show_step(self.current_step + 1)
            
    def start_installation(self):
        """Start the actual installation process"""
        try:
            # Validate installation path first
            install_path = self.install_path.get().strip()
            if not install_path:
                messagebox.showerror("Error", "Please specify an installation directory.")
                return
                
            # Check if path is writable
            test_dir = Path(install_path)
            try:
                # Try to create the directory to test permissions
                test_dir.mkdir(parents=True, exist_ok=True)
                
                # Test write permissions
                test_file = test_dir / "permission_test.tmp"
                test_file.write_text("test")
                test_file.unlink()
                
            except PermissionError:
                messagebox.showerror("Permission Error", 
                    f"Cannot write to the selected directory:\n{install_path}\n\n"
                    f"Please choose a different location or run as administrator.")
                return
            except Exception as e:
                messagebox.showerror("Path Error", 
                    f"Cannot access the installation directory:\n{str(e)}\n\n"
                    f"Please choose a different location.")
                return
            
            # Proceed with installation
            self.perform_installation()
            # Move to completion step
            self.show_step(4)
        except Exception as e:
            messagebox.showerror("Installation Error", f"Failed to install VirtuKey:\n{str(e)}")
            
    def perform_installation(self):
        """Perform the actual installation"""
        try:
            # Create installation directory
            install_dir = Path(self.install_path.get())
            install_dir.mkdir(parents=True, exist_ok=True)
            
            # A running copy locks VirtuKey.exe and the DLL
            if not self.handle_running_virtukey():
                raise Exception("VirtuKey is still running.")

            # Copy main files (Icon.png is optional)
            files_to_copy = ["VirtuKey.exe", "VirtualDesktopAccessor.dll", "Icon.png"]
            files_copied = 0

            for file_name in files_to_copy:
                source_file = Path(self.get_resource_path(file_name))
                dest_file = install_dir / file_name

                if not source_file.exists() and file_name == "Icon.png":
                    continue
                if source_file.exists():
                    try:
                        shutil.copy2(source_file, dest_file)
                        files_copied += 1
                    except PermissionError:
                        raise Exception(f"Permission denied when copying {file_name}. Please check folder permissions.")
                    except Exception as e:
                        raise Exception(f"Error copying {file_name}: {str(e)}")
                else:
                    raise Exception(f"Source file not found: {source_file}")
            
            if files_copied == 0:
                raise Exception("No files were copied. Installation failed.")
            
            # Create shortcuts if requested
            if self.create_desktop_shortcut.get():
                self.create_desktop_shortcut_file()
                
            if self.create_startmenu_shortcut.get():
                self.create_startmenu_shortcut_file()
                
            # Add to startup if requested
            if self.auto_start.get():
                self.add_to_startup()
                
        except Exception as e:
            raise Exception(f"Installation failed: {str(e)}")
            
    def get_desktop_dir(self):
        """Real Desktop folder (Windows 11 often redirects it to OneDrive)"""
        try:
            buf = ctypes.create_unicode_buffer(260)
            ctypes.windll.shell32.SHGetFolderPathW(None, 0x10, None, 0, buf)  # CSIDL_DESKTOPDIRECTORY
            if buf.value:
                return Path(buf.value)
        except Exception:
            pass
        return Path.home() / "Desktop"
    
    def get_startmenu_dir(self):
        return Path(os.environ.get("APPDATA", Path.home() / "AppData" / "Roaming")) / "Microsoft" / "Windows" / "Start Menu" / "Programs" / "VirtuKey"
    
    def create_shortcuts(self, shortcuts):
        """Create several .lnk files with a single PowerShell call.
        shortcuts: list of (lnk_path, target, arguments, description)"""
        def ps_quote(value):
            return "'" + str(value).replace("'", "''") + "'"
        
        lines = ["$s = New-Object -ComObject WScript.Shell"]
        for lnk, target, args, desc in shortcuts:
            lines += [
                f"$l = $s.CreateShortcut({ps_quote(lnk)})",
                f"$l.TargetPath = {ps_quote(target)}",
                f"$l.Arguments = {ps_quote(args)}",
                f"$l.WorkingDirectory = {ps_quote(Path(target).parent)}",
                f"$l.Description = {ps_quote(desc)}",
                "$l.Save()",
            ]
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", "; ".join(lines)],
                       check=True, capture_output=True, creationflags=subprocess.CREATE_NO_WINDOW)
    
    def create_desktop_shortcut_file(self):
        """Create desktop shortcut"""
        try:
            exe_path = Path(self.install_path.get()) / "VirtuKey.exe"
            self.create_shortcuts([(self.get_desktop_dir() / "VirtuKey.lnk", exe_path, "",
                                    "VirtuKey - Virtual Desktop Manager")])
        except Exception as e:
            # Non-critical error - don't fail installation
            print(f"Warning: Could not create desktop shortcut: {e}")
        
    def create_startmenu_shortcut_file(self):
        """Create Start Menu shortcuts for the app and the uninstaller"""
        try:
            startmenu_path = self.get_startmenu_dir()
            startmenu_path.mkdir(parents=True, exist_ok=True)
            exe_path = Path(self.install_path.get()) / "VirtuKey.exe"
            
            # Frozen builds launch themselves; the script needs the interpreter
            if getattr(sys, "frozen", False):
                uninstall_target, uninstall_args = Path(sys.executable), ""
            else:
                uninstall_target, uninstall_args = Path(sys.executable), f'"{Path(__file__).resolve()}"'
            
            self.create_shortcuts([
                (startmenu_path / "VirtuKey.lnk", exe_path, "", "VirtuKey - Virtual Desktop Manager"),
                (startmenu_path / "Uninstall VirtuKey.lnk", uninstall_target, uninstall_args, "Uninstall VirtuKey"),
            ])
        except Exception as e:
            # Non-critical error - don't fail installation
            print(f"Warning: Could not create start menu shortcut: {e}")
        
    def add_to_startup(self):
        """Add to Windows startup using registry"""
        try:
            exe_path = Path(self.install_path.get()) / "VirtuKey.exe"
            
            # Add to registry for current user startup
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                winreg.SetValueEx(key, "VirtuKey", 0, winreg.REG_SZ, str(exe_path))
                
        except Exception as e:
            # Non-critical error - don't fail installation
            print(f"Warning: Could not add to startup: {e}")
        
    def is_virtukey_running(self):
        """Check if VirtuKey is currently running"""
        if PSUTIL_AVAILABLE:
            try:
                for proc in psutil.process_iter(['pid', 'name', 'exe']):
                    try:
                        if proc.info['name'] and 'VirtuKey.exe' in proc.info['name']:
                            return True, proc.info['pid']
                    except (psutil.NoSuchProcess, psutil.AccessDenied):
                        continue
                return False, None
            except Exception:
                return False, None
        else:
            # Fall back to tasklist command
            try:
                result = subprocess.run(['tasklist', '/FI', 'IMAGENAME eq VirtuKey.exe', '/FO', 'CSV'], 
                                      capture_output=True, text=True, check=True)
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:  # More than just the header
                    # Parse the CSV to get PID
                    for line in lines[1:]:
                        parts = line.split(',')
                        if len(parts) >= 2:
                            pid = parts[1].strip('"')
                            return True, int(pid)
                return False, None
            except Exception:
                return False, None
    
    def terminate_virtukey_process(self, pid):
        """Terminate VirtuKey process"""
        if PSUTIL_AVAILABLE:
            try:
                proc = psutil.Process(pid)
                proc.terminate()
                
                # Wait up to 5 seconds for process to terminate
                for _ in range(50):
                    if not proc.is_running():
                        return True
                    time.sleep(0.1)
                
                # Force kill if still running
                proc.kill()
                time.sleep(0.5)
                return not proc.is_running()
                
            except (psutil.NoSuchProcess, psutil.AccessDenied):
                return True  # Process already gone or no access
            except Exception:
                return False
        else:
            # Fall back to taskkill command
            try:
                # Try graceful termination first
                subprocess.run(['taskkill', '/PID', str(pid)], check=True, capture_output=True)
                
                # Wait a bit and check if process is gone
                time.sleep(1)
                is_running, _ = self.is_virtukey_running()
                if not is_running:
                    return True
                
                # Force kill if still running
                subprocess.run(['taskkill', '/F', '/PID', str(pid)], check=True, capture_output=True)
                time.sleep(0.5)
                
                # Final check
                is_running, _ = self.is_virtukey_running()
                return not is_running
                
            except subprocess.CalledProcessError:
                return True  # Process might already be gone
            except Exception:
                return False
    
    def handle_running_virtukey(self):
        """Handle running VirtuKey processes during uninstall"""
        is_running, pid = self.is_virtukey_running()
        
        if not is_running:
            return True
        
        # Ask user what to do
        result = messagebox.askyesnocancel(
            "VirtuKey is Running",
            "VirtuKey is currently running and must be closed before uninstalling.\n\n"
            "Would you like to:\n"
            "• Yes - Automatically close VirtuKey and continue\n"
            "• No - Cancel uninstallation (you can close it manually)\n"
            "• Cancel - Return to previous step"
        )
        
        if result is True:  # Yes - close automatically
            if self.terminate_virtukey_process(pid):
                # Double-check it's really closed
                time.sleep(1)
                is_running, _ = self.is_virtukey_running()
                if not is_running:
                    return True
                else:
                    messagebox.showerror("Error", "Could not close VirtuKey automatically. Please close it manually.")
                    return False
            else:
                messagebox.showerror("Error", "Failed to close VirtuKey. Please close it manually.")
                return False
                
        elif result is False:  # No - cancel uninstallation
            return False
            
        else:  # Cancel - return to previous step
            return False
        
    def start_uninstallation(self):
        """Start the uninstallation process"""
        try:
            # First check if VirtuKey is running and handle it
            if not self.handle_running_virtukey():
                return  # User cancelled or couldn't close VirtuKey
            
            # Proceed with uninstallation
            self.perform_uninstallation()
            # Move to completion step
            self.show_step(4)
        except Exception as e:
            messagebox.showerror("Uninstallation Error", f"Failed to uninstall VirtuKey:\n{str(e)}")
            
    def start_reinstallation(self):
        """Start the reinstallation process (uninstall then install)"""
        try:
            # First check if VirtuKey is running and handle it
            if not self.handle_running_virtukey():
                return  # User cancelled or couldn't close VirtuKey
            
            # First uninstall
            self.perform_uninstallation()
            # Then install
            self.perform_installation()
            # Move to completion step
            self.show_step(4)
        except Exception as e:
            messagebox.showerror("Reinstallation Error", f"Failed to reinstall VirtuKey:\n{str(e)}")
            
    def perform_uninstallation(self):
        """Perform the actual uninstallation"""
        try:
            install_dir = Path(self.install_path.get())
            
            # Remove main application files
            files_to_remove = ["VirtuKey.exe", "VirtualDesktopAccessor.dll", "Icon.png"]
            for file_name in files_to_remove:
                file_path = install_dir / file_name
                if file_path.exists():
                    try:
                        file_path.unlink()
                    except PermissionError:
                        raise Exception(f"Permission denied when removing {file_name}. Please close VirtuKey and try again.")
                    except Exception as e:
                        raise Exception(f"Error removing {file_name}: {str(e)}")
            
            # Remove shortcuts if requested
            if hasattr(self, 'remove_shortcuts') and self.remove_shortcuts.get():
                self.remove_desktop_shortcut()
                self.remove_startmenu_shortcut()
            
            # The startup entry would point at a missing exe, so always drop it
            self.remove_startup_entry()
            
            # Remove settings if requested
            if hasattr(self, 'remove_settings') and self.remove_settings.get():
                self.remove_user_settings()
            
            # Remove installation directory if empty
            try:
                if install_dir.exists() and not any(install_dir.iterdir()):
                    install_dir.rmdir()
            except:
                pass  # Don't fail if directory can't be removed
                
        except Exception as e:
            raise Exception(f"Uninstallation failed: {str(e)}")
            
    def remove_desktop_shortcut(self):
        """Remove desktop shortcut"""
        try:
            desktop_path = self.get_desktop_dir() / "VirtuKey.lnk"
            if desktop_path.exists():
                desktop_path.unlink()
        except Exception as e:
            print(f"Warning: Could not remove desktop shortcut: {e}")
        
    def remove_startmenu_shortcut(self):
        """Remove start menu shortcut"""
        try:
            startmenu_path = self.get_startmenu_dir()
            if startmenu_path.exists():
                shutil.rmtree(startmenu_path)
        except Exception as e:
            print(f"Warning: Could not remove start menu shortcuts: {e}")
        
    def remove_user_settings(self):
        """Remove user settings and configuration"""
        try:
            settings_dir = Path(os.environ.get("APPDATA", "")) / "VirtuKey"
            if os.environ.get("APPDATA") and settings_dir.exists():
                shutil.rmtree(settings_dir)
        except Exception as e:
            print(f"Warning: Could not remove settings: {e}")
    
    def remove_startup_entry(self):
        """Remove VirtuKey from Windows startup"""
        try:
            key_path = r"SOFTWARE\Microsoft\Windows\CurrentVersion\Run"
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER, key_path, 0, winreg.KEY_SET_VALUE) as key:
                try:
                    winreg.DeleteValue(key, "VirtuKey")
                except FileNotFoundError:
                    pass  # Value doesn't exist, that's fine
                    
        except Exception as e:
            print(f"Warning: Could not remove startup entry: {e}")
    
    def get_resource_path(self, filename):
        """Get path to resource file, works for both development and PyInstaller bundle"""
        if hasattr(sys, '_MEIPASS'):
            # Running as PyInstaller bundle
            return os.path.join(sys._MEIPASS, 'resource', filename)
        else:
            # Running as script
            return os.path.join(os.path.dirname(__file__), 'resource', filename)
    
    def finish_installation(self):
        """Finish the installation/uninstallation"""
        if self.mode != "uninstall" and hasattr(self, 'launch_now') and self.launch_now.get():
            # Launch the application only for install/reinstall
            exe_path = Path(self.install_path.get()) / "VirtuKey.exe"
            if exe_path.exists():
                os.startfile(str(exe_path))
                
        self.root.quit()
        
    def cancel_installation(self):
        """Cancel the installation/uninstallation"""
        action = "uninstallation" if self.mode == "uninstall" else "installation"
        result = messagebox.askyesno("Cancel", 
                                   f"Are you sure you want to cancel the {action}?")
        if result:
            self.root.quit()
            
    def run(self):
        """Start the installer"""
        # Center the window
        self.root.update_idletasks()
        w = min(self.px(700), self.root.winfo_screenwidth())
        h = min(self.px(600), self.root.winfo_screenheight() - self.px(48))
        x = (self.root.winfo_screenwidth() - w) // 2
        y = (self.root.winfo_screenheight() - h) // 2
        self.root.geometry(f"{w}x{h}+{x}+{y}")
        
        self.root.mainloop()

if __name__ == "__main__":
    installer = VirtuKeyInstaller()
    installer.run()
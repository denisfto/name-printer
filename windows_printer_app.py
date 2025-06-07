#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import tkinter as tk
from tkinter import ttk, messagebox
import win32print
import win32ui
import win32con
import json
import os
from PIL import Image, ImageDraw, ImageFont
import tempfile

class WindowsNamePrinter:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Name Printer - Windows")
        self.root.geometry("600x650")
        self.root.resizable(False, False)
        
        self.config_file = "printer_settings.json"
        self.selected_printer = None
        
        # Load saved printer settings
        self.load_printer_config()
        
        # Create interface
        self.create_widgets()
        
        # Check if printer is configured
        if not self.selected_printer:
            self.show_printer_setup()
    
    def load_printer_config(self):
        """Load saved printer configuration"""
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r') as f:
                    config = json.load(f)
                    self.selected_printer = config.get('printer')
                    print(f"Loaded printer: {self.selected_printer}")
        except Exception as e:
            print(f"Error loading config: {e}")
            self.selected_printer = None
    
    def save_printer_config(self):
        """Save printer configuration"""
        try:
            config = {'printer': self.selected_printer}
            with open(self.config_file, 'w') as f:
                json.dump(config, f, indent=2)
            print(f"Saved printer: {self.selected_printer}")
        except Exception as e:
            print(f"Error saving config: {e}")
    
    def get_available_printers(self):
        """Get list of available printers"""
        printers = []
        try:
            # Get all local and network printers
            printer_enum = win32print.EnumPrinters(
                win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            )
            for printer in printer_enum:
                printers.append(printer[2])  # Printer name
        except Exception as e:
            print(f"Error getting printers: {e}")
        return printers
    
    def show_printer_setup(self):
        """Show printer selection dialog"""
        printers = self.get_available_printers()
        if not printers:
            messagebox.showerror("Error", "No printers found in the system!")
            return
        
        # Create printer selection window
        printer_window = tk.Toplevel(self.root)
        printer_window.title("Select Printer")
        printer_window.geometry("500x400")
        printer_window.transient(self.root)
        printer_window.grab_set()
        
        # Center the window
        printer_window.geometry("+{}+{}".format(
            int(printer_window.winfo_screenwidth()/2 - 250),
            int(printer_window.winfo_screenheight()/2 - 200)
        ))
        
        # Title
        title_label = tk.Label(printer_window, text="Choose Printer", 
                              font=("Arial", 16, "bold"))
        title_label.pack(pady=20)
        
        info_label = tk.Label(printer_window, 
                             text="Select a printer from the list below.\nThis setting will be saved for future use.",
                             font=("Arial", 10))
        info_label.pack(pady=10)
        
        # Printer list
        listbox_frame = tk.Frame(printer_window)
        listbox_frame.pack(pady=20, padx=20, fill=tk.BOTH, expand=True)
        
        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        listbox = tk.Listbox(listbox_frame, font=("Arial", 11), 
                            yscrollcommand=scrollbar.set)
        for printer in printers:
            listbox.insert(tk.END, printer)
        listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=listbox.yview)
        
        # Select current printer if available
        if self.selected_printer and self.selected_printer in printers:
            index = printers.index(self.selected_printer)
            listbox.selection_set(index)
            listbox.see(index)
        
        # Buttons
        button_frame = tk.Frame(printer_window)
        button_frame.pack(pady=20)
        
        def confirm_selection():
            selection = listbox.curselection()
            if selection:
                self.selected_printer = listbox.get(selection[0])
                self.save_printer_config()
                self.update_printer_status()
                printer_window.destroy()
                messagebox.showinfo("Success", f"Printer '{self.selected_printer}' selected!")
            else:
                messagebox.showwarning("Warning", "Please select a printer!")
        
        def cancel_selection():
            printer_window.destroy()
        
        tk.Button(button_frame, text="Select Printer", command=confirm_selection,
                 font=("Arial", 12), bg="#4CAF50", fg="white", padx=20).pack(side=tk.LEFT, padx=10)
        tk.Button(button_frame, text="Cancel", command=cancel_selection,
                 font=("Arial", 12), bg="#f44336", fg="white", padx=20).pack(side=tk.LEFT, padx=10)
    
    def create_widgets(self):
        """Create the main interface"""
        # Title
        title_label = tk.Label(self.root, text="Name Printer for Windows", 
                              font=("Arial", 20, "bold"), fg="#2E7D32")
        title_label.pack(pady=15)
        
        subtitle_label = tk.Label(self.root, 
                                 text="Print names in large uppercase letters on A4 landscape",
                                 font=("Arial", 12), fg="#666")
        subtitle_label.pack(pady=5)
        
        # Printer status frame
        status_frame = tk.LabelFrame(self.root, text="Printer Status", 
                                    font=("Arial", 12, "bold"), padx=10, pady=10)
        status_frame.pack(pady=15, padx=20, fill=tk.X)
        
        self.printer_status_label = tk.Label(status_frame, 
                                           text="No printer selected", 
                                           font=("Arial", 11), fg="#f44336")
        self.printer_status_label.pack(side=tk.LEFT)
        
        setup_printer_btn = tk.Button(status_frame, text="Setup Printer",
                                     command=self.show_printer_setup,
                                     font=("Arial", 10), bg="#2196F3", fg="white")
        setup_printer_btn.pack(side=tk.RIGHT)
        
        # Input frame
        input_frame = tk.LabelFrame(self.root, text="Text Input", 
                                   font=("Arial", 12, "bold"), padx=15, pady=15)
        input_frame.pack(pady=15, padx=20, fill=tk.X)
        
        # Name input
        tk.Label(input_frame, text="First Name:", font=("Arial", 12)).grid(
            row=0, column=0, sticky=tk.W, pady=8)
        self.name_entry = tk.Entry(input_frame, font=("Arial", 14), width=35)
        self.name_entry.grid(row=0, column=1, pady=8, padx=(10, 0), sticky=tk.W)
        
        # Surname input
        tk.Label(input_frame, text="Last Name:", font=("Arial", 12)).grid(
            row=1, column=0, sticky=tk.W, pady=8)
        self.surname_entry = tk.Entry(input_frame, font=("Arial", 14), width=35)
        self.surname_entry.grid(row=1, column=1, pady=8, padx=(10, 0), sticky=tk.W)
        
        # Removed additional text input to simplify interface
        
        # Preview frame
        preview_frame = tk.LabelFrame(self.root, text="Preview", 
                                     font=("Arial", 12, "bold"), padx=15, pady=10)
        preview_frame.pack(pady=15, padx=20, fill=tk.BOTH, expand=True)
        
        self.preview_text = tk.Text(preview_frame, font=("Courier", 10), 
                                   height=8, state=tk.DISABLED, bg="#f8f8f8")
        self.preview_text.pack(fill=tk.BOTH, expand=True, pady=5)
        
        # Status frame
        self.status_frame = tk.Frame(self.root)
        self.status_frame.pack(pady=10, fill=tk.X, padx=20)
        
        self.status_label = tk.Label(self.status_frame, text="Ready to print", 
                                    font=("Arial", 11), fg="#4CAF50")
        self.status_label.pack(side=tk.LEFT)
        
        # Buttons frame
        button_frame = tk.Frame(self.root)
        button_frame.pack(pady=15)
        
        preview_btn = tk.Button(button_frame, text="Update Preview", 
                               command=self.update_preview, font=("Arial", 12), 
                               bg="#FF9800", fg="white", padx=20)
        preview_btn.pack(side=tk.LEFT, padx=10)
        
        self.print_btn = tk.Button(button_frame, text="PRINT NOW", 
                                  command=self.print_text, font=("Arial", 14, "bold"), 
                                  bg="#4CAF50", fg="white", padx=30, pady=5)
        self.print_btn.pack(side=tk.LEFT, padx=10)
        
        # Bind events for auto-preview
        self.name_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        self.surname_entry.bind('<KeyRelease>', lambda e: self.update_preview())
        
        # Initial setup
        self.update_printer_status()
        self.update_preview()
    
    def update_printer_status(self):
        """Update printer status display"""
        if self.selected_printer:
            self.printer_status_label.config(
                text=f"Selected: {self.selected_printer}", 
                fg="#4CAF50"
            )
            self.print_btn.config(state=tk.NORMAL)
        else:
            self.printer_status_label.config(
                text="No printer selected", 
                fg="#f44336"
            )
            self.print_btn.config(state=tk.DISABLED)
    
    def update_preview(self):
        """Update print preview"""
        name = self.name_entry.get().strip()
        surname = self.surname_entry.get().strip()
        
        preview_lines = []
        
        if name:
            preview_lines.append(f"FIRST NAME (large): {name.upper()}")
        if surname:
            preview_lines.append(f"LAST NAME (large):  {surname.upper()}")
        
        if not preview_lines:
            preview_lines = ["Enter first name or last name to see preview"]
        
        # Add printing info
        if name or surname:
            preview_lines.append("")
            preview_lines.append("PRINT SETTINGS:")
            longest = max([name, surname], key=len) if name and surname else (name or surname)
            preview_lines.append(f"• Longest text: '{longest.upper()}'")
            preview_lines.append(f"• Font size: Automatically maximized")
            preview_lines.append(f"• Layout: A4 Landscape, left-aligned")
        
        preview_text = '\n'.join(preview_lines)
        
        self.preview_text.config(state=tk.NORMAL)
        self.preview_text.delete(1.0, tk.END)
        self.preview_text.insert(1.0, preview_text)
        self.preview_text.config(state=tk.DISABLED)
    
    def create_print_image(self, name, surname):
        """Create image for printing with maximum font size"""
        # A4 landscape size at 300 DPI: 3508x2480 pixels
        width, height = 3508, 2480
        
        # Create white image
        image = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(image)
        
        # Margins (5% from each side)
        margin_x = int(width * 0.05)
        margin_y = int(height * 0.05)
        available_width = width - (margin_x * 2)
        available_height = height - (margin_y * 2)
        
        # Create main text lines (uppercase)
        main_lines = []
        if name.strip():
            main_lines.append(name.strip().upper())
        if surname.strip():
            main_lines.append(surname.strip().upper())
        
        if not main_lines:
            return None
        
        # Find longest main text
        longest_main_text = max(main_lines, key=len)
        
        # Find optimal font size for longest text
        best_font_size = 50
        best_font_obj = None
        
        # Try different font sizes from small to large
        for font_size in range(50, 800, 10):
            try:
                # Try to load Arial font
                font_paths = [
                    "C:\\Windows\\Fonts\\arial.ttf",
                    "C:\\Windows\\Fonts\\calibri.ttf",
                    "arial.ttf"
                ]
                
                font_obj = None
                for font_path in font_paths:
                    try:
                        if os.path.exists(font_path):
                            font_obj = ImageFont.truetype(font_path, font_size)
                            break
                    except:
                        continue
                
                if font_obj is None:
                    font_obj = ImageFont.load_default()
                
                # Measure text width
                bbox = draw.textbbox((0, 0), longest_main_text, font=font_obj)
                text_width = bbox[2] - bbox[0]
                
                # If text fits in 90% of available width, save this size
                if text_width <= available_width * 0.9:
                    best_font_size = font_size
                    best_font_obj = font_obj
                else:
                    # Too big, use previous size
                    break
                    
            except Exception:
                continue
        
        if best_font_obj is None:
            try:
                best_font_obj = ImageFont.truetype("C:\\Windows\\Fonts\\arial.ttf", 100)
            except:
                best_font_obj = ImageFont.load_default()
        
        # Calculate line heights and spacing
        line_heights = []
        line_spacing = int(best_font_size * 0.3)
        
        # Heights for main lines
        for line in main_lines:
            bbox = draw.textbbox((0, 0), line, font=best_font_obj)
            line_height = bbox[3] - bbox[1]
            line_heights.append(line_height)
        
        # Calculate total height
        total_height = sum(line_heights) + (len(main_lines) - 1) * line_spacing
        
        # Center vertically
        start_y = margin_y + (available_height - total_height) // 2
        current_y = start_y
        
        # Draw main text
        for i, line in enumerate(main_lines):
            x_pos = margin_x  # Left-aligned
            draw.text((x_pos, current_y), line, fill='black', font=best_font_obj)
            current_y += line_heights[i] + line_spacing
        
        return image
    
    def print_text(self):
        """Print the text directly to selected printer"""
        if not self.selected_printer:
            messagebox.showerror("Error", "Please select a printer first!")
            return
        
        name = self.name_entry.get().strip()
        surname = self.surname_entry.get().strip()
        
        if not name and not surname:
            messagebox.showwarning("Warning", "Please enter at least first name or last name!")
            return
        
        # Update status
        self.status_label.config(text="Creating print job...", fg="#FF9800")
        self.root.update()
        
        try:
            # Create image
            image = self.create_print_image(name, surname)
            if not image:
                raise Exception("Failed to create print image")
            
            # Save to temporary file
            temp_file = tempfile.mktemp(suffix='.png')
            image.save(temp_file, 'PNG', dpi=(300, 300))
            
            self.status_label.config(text="Sending to printer...", fg="#FF9800")
            self.root.update()
            
            # Print using Windows API
            import subprocess
            try:
                # Try to print directly to selected printer
                subprocess.run([
                    'rundll32.exe', 'shimgvw.dll,ImageView_PrintTo', 
                    temp_file, self.selected_printer
                ], check=True, timeout=30)
                
                self.status_label.config(text="Print job sent successfully!", fg="#4CAF50")
                messagebox.showinfo("Success", f"Print job sent to {self.selected_printer}")
                
            except subprocess.TimeoutExpired:
                # Fallback: open print dialog
                os.startfile(temp_file, 'print')
                self.status_label.config(text="Print dialog opened", fg="#4CAF50")
                messagebox.showinfo("Print", "Print dialog opened. Please select your printer and print.")
            
            except:
                # Another fallback
                os.startfile(temp_file, 'print')
                self.status_label.config(text="Print dialog opened", fg="#4CAF50")
                messagebox.showinfo("Print", "Print dialog opened. Please select your printer and print.")
            
            # Clean up temp file after delay
            self.root.after(10000, lambda: self.cleanup_temp_file(temp_file))
            
        except Exception as e:
            self.status_label.config(text="Print failed!", fg="#f44336")
            messagebox.showerror("Print Error", f"Failed to print: {str(e)}")
    
    def cleanup_temp_file(self, filepath):
        """Clean up temporary file"""
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass  # Ignore cleanup errors
    
    def run(self):
        """Run the application"""
        self.root.mainloop()

def main():
    """Main function"""
    print("Starting Windows Name Printer...")
    app = WindowsNamePrinter()
    app.run()

if __name__ == "__main__":
    main()

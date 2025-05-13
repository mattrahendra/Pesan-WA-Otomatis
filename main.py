import os
import time
import random
import threading
import pandas as pd
import pywhatkit as kit
import tkinter as tk
from tkinter import filedialog, scrolledtext, ttk, messagebox, Toplevel
from tkinter import ttk as ttk

class WhatsAppBulkSender:
    def __init__(self, root):
        self.root = root
        self.root.title("WhatsApp Bulk Messenger")
        self.root.geometry("800x700")
        self.root.resizable(True, True)
        
        self.df = None
        self.poster_path = None
        self.sending_thread = None
        self.is_running = False
        self.is_paused = False
        self.current_index = 0
        
        self.setup_ui()
    
    def setup_ui(self):
        # Frame for file inputs
        file_frame = ttk.LabelFrame(self.root, text="Data & Poster")
        file_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(file_frame, text="Data Excel:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.excel_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.excel_path_var, width=50).grid(row=0, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_excel).grid(row=0, column=2, padx=5, pady=5)
        self.preview_button = ttk.Button(file_frame, text="Preview Data", command=self.preview_excel, state="disabled")
        self.preview_button.grid(row=0, column=3, padx=5, pady=5)
        
        ttk.Label(file_frame, text="Poster Image:").grid(row=1, column=0, padx=5, pady=5, sticky="w")
        self.poster_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.poster_path_var, width=50).grid(row=1, column=1, padx=5, pady=5)
        ttk.Button(file_frame, text="Browse", command=self.browse_poster).grid(row=1, column=2, padx=5, pady=5)
        
        # Informasi variabel template
        ttk.Label(file_frame, text="Variabel yang tersedia: {nama}, {nomor}, {no}").grid(row=2, column=0, columnspan=4, padx=5, pady=5, sticky="w")
        
        # Frame for advanced settings
        settings_frame = ttk.LabelFrame(self.root, text="Pengaturan")
        settings_frame.pack(fill="x", padx=10, pady=5)
        
        ttk.Label(settings_frame, text="Delay minimum (detik):").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.min_delay_var = tk.StringVar(value="10")
        ttk.Entry(settings_frame, textvariable=self.min_delay_var, width=5).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(settings_frame, text="Delay maksimum (detik):").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.max_delay_var = tk.StringVar(value="15")
        ttk.Entry(settings_frame, textvariable=self.max_delay_var, width=5).grid(row=0, column=3, padx=5, pady=5)
        
        # Frame for range selection
        range_frame = ttk.LabelFrame(self.root, text="Range Selection")
        range_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(range_frame, text="Start from:").grid(row=0, column=0, padx=5, pady=5, sticky="w")
        self.start_index_var = tk.StringVar(value="1")
        ttk.Entry(range_frame, textvariable=self.start_index_var, width=10).grid(row=0, column=1, padx=5, pady=5)
        
        ttk.Label(range_frame, text="End at:").grid(row=0, column=2, padx=5, pady=5, sticky="w")
        self.end_index_var = tk.StringVar(value="100")
        ttk.Entry(range_frame, textvariable=self.end_index_var, width=10).grid(row=0, column=3, padx=5, pady=5)
        
        # Message template
        message_frame = ttk.LabelFrame(self.root, text="Custom Message Template")
        message_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.message_template = scrolledtext.ScrolledText(message_frame, wrap=tk.WORD, height=10)
        self.message_template.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Set default template
        default_template = """tes no {no}. kamu adalah {nama} dengan nomor {nomor}!"""
        self.message_template.insert(tk.END, default_template)
        
        # Frame for control buttons
        control_frame = ttk.Frame(self.root)
        control_frame.pack(fill="x", padx=10, pady=10)
        
        self.start_button = ttk.Button(control_frame, text="Start", command=self.start_sending)
        self.start_button.pack(side="left", padx=5)
        
        self.pause_button = ttk.Button(control_frame, text="Pause", command=self.pause_sending, state="disabled")
        self.pause_button.pack(side="left", padx=5)
        
        self.stop_button = ttk.Button(control_frame, text="Stop", command=self.stop_sending, state="disabled")
        self.stop_button.pack(side="left", padx=5)
        
        # Progress bar
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(self.root, variable=self.progress_var, maximum=100)
        self.progress_bar.pack(fill="x", padx=10, pady=10)
        
        # Log area
        log_frame = ttk.LabelFrame(self.root, text="Log")
        log_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        self.log_area = scrolledtext.ScrolledText(log_frame, wrap=tk.WORD, height=8)
        self.log_area.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Status bar
        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_bar.pack(fill="x", side="bottom", padx=10, pady=5)
    
    def browse_excel(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if file_path:
            self.excel_path_var.set(file_path)
            self.load_excel(file_path)
            self.preview_button.config(state="normal")
    
    def browse_poster(self):
        file_path = filedialog.askopenfilename(filetypes=[("Image files", "*.jpg *.jpeg *.png")])
        if file_path:
            self.poster_path_var.set(file_path)
            self.poster_path = file_path
    
    def load_excel(self, file_path):
        try:
            self.df = pd.read_excel(file_path, header=None)
            self.df.columns = ["No", "Nama", "NomorHP"]
            self.df["NomorFormatted"] = self.df["NomorHP"].apply(self.format_nomor)
            self.log(f"Data loaded: {len(self.df)} records found")
            self.end_index_var.set(str(len(self.df)))
        except Exception as e:
            self.log(f"Error loading data: {e}")
            messagebox.showerror("Error", f"Failed to load data: {e}")
            self.preview_button.config(state="disabled")
    
    def preview_excel(self):
        if self.df is None:
            messagebox.showerror("Error", "No data loaded to preview")
            return
        
        # Create a new popup window
        preview_window = Toplevel(self.root)
        preview_window.title("Excel Data Preview")
        preview_window.geometry("600x400")
        preview_window.resizable(True, True)
        
        # Create a frame for the treeview
        tree_frame = ttk.Frame(preview_window)
        tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create a treeview to display the data
        columns = list(self.df.columns)
        tree = ttk.Treeview(tree_frame, columns=columns, show="headings")
        
        # Set column headings
        for col in columns:
            tree.heading(col, text=col)
            tree.column(col, width=150, anchor=tk.W)
        
        # Add data to the treeview
        for _, row in self.df.iterrows():
            tree.insert("", tk.END, values=list(row))
        
        # Add scrollbars
        yscroll = ttk.Scrollbar(tree_frame, orient=tk.VERTICAL, command=tree.yview)
        xscroll = ttk.Scrollbar(tree_frame, orient=tk.HORIZONTAL, command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
        
        # Grid layout
        tree.grid(row=0, column=0, sticky="nsew")
        yscroll.grid(row=0, column=1, sticky="ns")
        xscroll.grid(row=1, column=0, sticky="ew")
        
        tree_frame.grid_rowconfigure(0, weight=1)
        tree_frame.grid_columnconfigure(0, weight=1)
        
        # Add a close button
        close_button = ttk.Button(preview_window, text="Close", command=preview_window.destroy)
        close_button.pack(pady=10)
    
    def format_nomor(self, nomor):
        if pd.isna(nomor):
            return ""
            
        # Convert to string and remove any hyphens
        nomor_str = str(nomor).replace("-", "")
        
        # Handle different formats
        if nomor_str.startswith("0"):
            return "+62" + nomor_str[1:]
        elif nomor_str.startswith("8"):
            return "+62" + nomor_str
        elif nomor_str.startswith("62"):
            return "+" + nomor_str
        else:
            return nomor_str
            
    def log(self, message):
        timestamp = time.strftime("%H:%M:%S")
        self.log_area.insert(tk.END, f"[{timestamp}] {message}\n")
        self.log_area.see(tk.END)
        
    def update_status(self, message):
        self.status_var.set(message)
        
    def update_progress(self, current, total):
        progress = (current / total) * 100
        self.progress_var.set(progress)
        
    def start_sending(self):
        if not self.df is not None or not self.poster_path:
            messagebox.showerror("Error", "Mohon pilih file data Excel dan gambar poster")
            return
            
        try:
            start_index = max(0, int(self.start_index_var.get()) - 1)  # Convert to 0-based index
            end_index = min(len(self.df), int(self.end_index_var.get()))
            
            if start_index >= end_index:
                messagebox.showerror("Error", "Indeks awal harus lebih kecil dari indeks akhir")
                return
                
            # Check if message template is not empty
            if not self.message_template.get("1.0", tk.END).strip():
                messagebox.showerror("Error", "Template pesan tidak boleh kosong")
                return
                
            self.current_index = start_index
            self.is_running = True
            self.is_paused = False
            
            # Update UI
            self.start_button.config(state="disabled")
            self.pause_button.config(state="normal")
            self.stop_button.config(state="normal")
            
            # Start sending thread
            self.sending_thread = threading.Thread(target=self.send_messages, args=(start_index, end_index))
            self.sending_thread.daemon = True
            self.sending_thread.start()
            
        except ValueError:
            messagebox.showerror("Error", "Mohon masukkan indeks yang valid")
            
    def pause_sending(self):
        if self.is_paused:
            self.is_paused = False
            self.pause_button.config(text="Pause")
            self.log("Resumed sending")
            self.update_status("Running")
        else:
            self.is_paused = True
            self.pause_button.config(text="Resume")
            self.log("Paused sending")
            self.update_status("Paused")
            
    def stop_sending(self):
        self.is_running = False
        self.is_paused = False
        self.log("Stopped sending")
        self.update_status("Stopped")
        
        # Update UI
        self.start_button.config(state="normal")
        self.pause_button.config(state="disabled", text="Pause")
        self.stop_button.config(state="disabled")
            
    def send_messages(self, start_index, end_index):
        total = end_index - start_index
        
        for i in range(start_index, end_index):
            # Check if stopped
            if not self.is_running:
                break
                
            # Check if paused
            while self.is_paused:
                time.sleep(0.5)
                if not self.is_running:
                    break
                    
            # Get data
            row = self.df.iloc[i]
            no = row["No"]
            nama = row["Nama"]
            nomor_asli = row["NomorHP"]
            nomor = row["NomorFormatted"]
            
            # Skip if no valid number
            if not nomor or pd.isna(nomor) or len(nomor) < 10:
                self.log(f"Skipping {nama}: Invalid number format")
                continue
                
            # Get custom message and apply template formatting
            template = self.message_template.get("1.0", tk.END)
            
            # Replace variables in template
            try:
                pesan = template.format(
                    nama=nama,
                    nomor=nomor_asli,
                    no=no
                )
            except KeyError as e:
                self.log(f"Template error: Unknown variable {e}")
                pesan = template
            except Exception as e:
                self.log(f"Template error: {e}")
                pesan = template

            # Send message
            try:
                self.log(f"Sending to {nama} ({nomor})...")
                self.update_status(f"Sending to {nama} ({i+1}/{end_index})")
                
                # Send image with caption
                kit.sendwhats_image(
                    receiver=nomor,
                    img_path=self.poster_path,
                    caption=pesan,
                    wait_time=30,
                    tab_close=True
                )
                
                self.log(f"Success: Message sent to {nama}")
                
                # Update progress
                self.current_index = i + 1
                self.root.after(0, lambda: self.update_progress(i - start_index + 1, total))
                
                # Random delay between min-max seconds
                try:
                    min_delay = int(self.min_delay_var.get())
                    max_delay = int(self.max_delay_var.get())
                    if min_delay > max_delay:
                        min_delay, max_delay = max_delay, min_delay
                except ValueError:
                    min_delay, max_delay = 10, 15
                    
                delay = random.randint(min_delay, max_delay)
                self.log(f"Waiting {delay} seconds before next message...")
                time.sleep(delay)
                
            except Exception as e:
                self.log(f"Failed to send to {nama} ({nomor}): {e}")
                
        # Completed
        if self.is_running:
            self.log("Process completed successfully!")
            self.update_status("Completed")
            self.root.after(0, lambda: self.stop_sending())


if __name__ == "__main__":
    root = tk.Tk()
    app = WhatsAppBulkSender(root)
    root.mainloop()
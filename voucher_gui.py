#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Voucher Processing GUI Application

A standalone GUI application that provides a user-friendly interface
for processing voucher templates and CSV data into PDF output files.
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import os
import shutil
import threading
from datetime import datetime
import tempfile
from pathlib import Path

# Import the converter functions
from converter import process_vouchers, convert_to_pdf


class VoucherProcessorGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Voucher Processor - Standalone Application")
        
        # Get screen dimensions for desktop-sized window
        screen_width = self.root.winfo_screenwidth()
        screen_height = self.root.winfo_screenheight()
        
        # Set window to 85% of screen size for desktop experience
        window_width = int(screen_width * 0.85)
        window_height = int(screen_height * 0.85)
        
        # Calculate center position
        x = (screen_width - window_width) // 2
        y = (screen_height - window_height) // 2
        
        self.root.geometry(f"{window_width}x{window_height}+{x}+{y}")
        self.root.minsize(1200, 800)
        self.root.resizable(True, True)
        
        # Optionally start maximized (uncomment next line if desired)
        # self.root.state('zoomed')  # Windows
        # self.root.attributes('-zoomed', True)  # Linux
        
        # Configure style
        self.setup_styles()
        
        # Initialize variables
        self.template_file = tk.StringVar()
        self.csv_file = tk.StringVar()
        self.processing = False
        
        # Create directories
        self.setup_directories()
        
        # Initialize history
        self.history_file = "processing_history.xlsx"
        self.setup_history()
        
        # Create GUI
        self.create_gui()
        
        # Load history on startup
        self.refresh_history()
    
    def setup_styles(self):
        """Setup custom styles for better visual appeal"""
        style = ttk.Style()
        
        # Configure colors and fonts
        style.configure('Title.TLabel', font=('Arial', 18, 'bold'), foreground='#2c3e50')
        style.configure('Heading.TLabel', font=('Arial', 12, 'bold'), foreground='#34495e')
        style.configure('TLabelframe', padding=15)
        style.configure('Process.TButton', font=('Arial', 11, 'bold'), padding=(20, 8))
        
        # Increase row height for better vertical spacing in history treeview
        style.configure('Treeview', rowheight=35)
    
    def setup_directories(self):
        """Create necessary directories"""
        os.makedirs('filesIn', exist_ok=True)
        os.makedirs('fileOut', exist_ok=True)
    
    def setup_history(self):
        """Setup history tracking spreadsheet"""
        if not os.path.exists(self.history_file):
            # Create initial history file
            df = pd.DataFrame(columns=[
                'Timestamp', 'Template_File', 'CSV_File', 'Records_Processed',
                'Pages_Generated', 'Output_PDF', 'Status', 'Notes'
            ])
            df.to_excel(self.history_file, index=False)
    
    def create_gui(self):
        """Create the main GUI interface with two-pane layout"""
        
        # Configure root grid
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Main container
        main_container = ttk.Frame(self.root)
        main_container.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=10, pady=10)
        main_container.columnconfigure(0, weight=1)
        main_container.rowconfigure(1, weight=1)
        
        # Title header
        title_frame = ttk.Frame(main_container)
        title_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        title_frame.columnconfigure(1, weight=1)
        
        title_label = ttk.Label(title_frame, text="🎫 Voucher Processing Application", 
                               style='Title.TLabel')
        title_label.grid(row=0, column=0, sticky=tk.W)
        
        # Status indicator
        self.status_indicator = ttk.Label(title_frame, text="● Ready", 
                                         foreground='green', font=('Arial', 10, 'bold'))
        self.status_indicator.grid(row=0, column=1, sticky=tk.E)
        
        # Create PanedWindow for two-pane layout
        paned_window = ttk.PanedWindow(main_container, orient=tk.HORIZONTAL)
        paned_window.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_container.rowconfigure(1, weight=1)
        
        # LEFT PANE - Input and Processing
        left_pane = ttk.Frame(paned_window)
        paned_window.add(left_pane, weight=1)
        left_pane.columnconfigure(0, weight=1)
        
        self.create_left_pane(left_pane)
        
        # RIGHT PANE - Status, Output, and History
        right_pane = ttk.Frame(paned_window)
        paned_window.add(right_pane, weight=1)
        right_pane.columnconfigure(0, weight=1)
        
        self.create_right_pane(right_pane)
        
        # Initial status message
        self.log_message("🚀 Application ready. Please upload template and CSV files to begin processing.")
    
    def create_left_pane(self, parent):
        """Create the left pane with file upload and processing controls"""
        
        # File Upload Section
        upload_frame = ttk.LabelFrame(parent, text="📁 File Upload", padding=15)
        upload_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        upload_frame.columnconfigure(1, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Template file upload
        ttk.Label(upload_frame, text="Template File (.docx):", style='Heading.TLabel').grid(
            row=0, column=0, sticky=tk.W, pady=(0, 8))
        
        template_frame = ttk.Frame(upload_frame)
        template_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        template_frame.columnconfigure(0, weight=1)
        
        self.template_entry = ttk.Entry(template_frame, textvariable=self.template_file, 
                                       state='readonly', font=('Arial', 10))
        self.template_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(template_frame, text="📄 Browse", command=self.browse_template).grid(
            row=0, column=1)
        
        # CSV file upload
        ttk.Label(upload_frame, text="Voucher Data (.csv):", style='Heading.TLabel').grid(
            row=2, column=0, sticky=tk.W, pady=(0, 8))
        
        csv_frame = ttk.Frame(upload_frame)
        csv_frame.grid(row=3, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        csv_frame.columnconfigure(0, weight=1)
        
        self.csv_entry = ttk.Entry(csv_frame, textvariable=self.csv_file, 
                                  state='readonly', font=('Arial', 10))
        self.csv_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))
        
        ttk.Button(csv_frame, text="📊 Browse", command=self.browse_csv).grid(
            row=0, column=1)
        
        # File info display
        self.file_info = ttk.Label(upload_frame, text="No files selected", 
                                  foreground='gray', font=('Arial', 9))
        self.file_info.grid(row=4, column=0, sticky=tk.W, pady=5)
        
        # Processing Section
        process_frame = ttk.LabelFrame(parent, text="⚡ Processing Controls", padding=15)
        process_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        process_frame.columnconfigure(0, weight=1)
        
        # Process button
        self.process_btn = ttk.Button(process_frame, text="🚀 Process Vouchers", 
                                     command=self.start_processing, style='Process.TButton', state='disabled')
        self.process_btn.grid(row=0, column=0, pady=10, sticky=(tk.W, tk.E))
        
        # Progress section
        progress_frame = ttk.Frame(process_frame)
        progress_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(5, 10))
        progress_frame.columnconfigure(0, weight=1)
        
        self.progress_label = ttk.Label(progress_frame, text="Ready to process", 
                                       font=('Arial', 9))
        self.progress_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        
        self.progress = ttk.Progressbar(progress_frame, mode='indeterminate', length=300)
        self.progress.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        # Quick Stats Section
        stats_frame = ttk.LabelFrame(parent, text="📈 Quick Statistics", padding=15)
        stats_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        parent.rowconfigure(2, weight=1)
        
        # Stats display
        self.stats_text = tk.Text(stats_frame, height=8, wrap=tk.WORD, font=('Courier', 9),
                                 background='#f8f9fa', state='disabled')
        stats_scrollbar = ttk.Scrollbar(stats_frame, orient=tk.VERTICAL, command=self.stats_text.yview)
        self.stats_text.configure(yscrollcommand=stats_scrollbar.set)
        
        self.stats_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        stats_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        stats_frame.columnconfigure(0, weight=1)
        stats_frame.rowconfigure(0, weight=1)
        
        self.update_stats()
    
    def create_right_pane(self, parent):
        """Create the right pane with status, output, and history"""
        
        # Status Section
        status_frame = ttk.LabelFrame(parent, text="📋 Processing Status", padding=15)
        status_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        status_frame.columnconfigure(0, weight=1)
        parent.columnconfigure(0, weight=1)
        
        # Status text area
        self.status_text = tk.Text(status_frame, height=10, wrap=tk.WORD, font=('Consolas', 9),
                                  background='#2c3e50', foreground='#ecf0f1')
        status_scrollbar = ttk.Scrollbar(status_frame, orient=tk.VERTICAL, command=self.status_text.yview)
        self.status_text.configure(yscrollcommand=status_scrollbar.set)
        
        self.status_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        status_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        status_frame.rowconfigure(0, weight=1)
        
        # Clear status button
        ttk.Button(status_frame, text="🗑️ Clear Log", command=self.clear_status).grid(
            row=1, column=0, pady=5, sticky=tk.W)
        
        # Output Section
        output_frame = ttk.LabelFrame(parent, text="📥 Output & Download", padding=15)
        output_frame.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=(0, 15))
        
        # Output controls
        output_controls = ttk.Frame(output_frame)
        output_controls.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=5)
        
        self.download_btn = ttk.Button(output_controls, text="💾 Download PDF", 
                                      command=self.download_pdf, state='disabled')
        self.download_btn.grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(output_controls, text="📁 Open Output Folder", 
                  command=self.open_output_folder).grid(row=0, column=1)
        
        # File status
        self.output_status = ttk.Label(output_frame, text="No output files generated yet", 
                                      foreground='gray', font=('Arial', 9))
        self.output_status.grid(row=1, column=0, sticky=tk.W, pady=5)
        
        # History Section
        history_frame = ttk.LabelFrame(parent, text="📝 Processing History", padding=15)
        history_frame.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 0))
        history_frame.columnconfigure(0, weight=1)
        history_frame.rowconfigure(0, weight=1)
        parent.rowconfigure(2, weight=1)
        
        # History treeview with better styling
        self.history_tree = ttk.Treeview(history_frame, 
                                        columns=('timestamp', 'template', 'records', 'status'), 
                                        show='headings', height=8)
        
        # Configure columns
        self.history_tree.heading('timestamp', text='🕒 Date/Time')
        self.history_tree.heading('template', text='📄 Template')
        self.history_tree.heading('records', text='📊 Records')
        self.history_tree.heading('status', text='🔍 Status')
        
        self.history_tree.column('timestamp', width=100, anchor='center')
        self.history_tree.column('template', width=150, anchor='w')
        self.history_tree.column('records', width=80, anchor='center')
        self.history_tree.column('status', width=100, anchor='center')
        
        # Style alternate rows with better contrast and padding effect
        self.history_tree.tag_configure('success', background='#d5f4e6')
        self.history_tree.tag_configure('failed', background='#ffeaa7')
        
        # Add some visual separation between rows
        self.history_tree.configure(style='Treeview')
        
        history_scrollbar = ttk.Scrollbar(history_frame, orient=tk.VERTICAL, command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=history_scrollbar.set)
        
        self.history_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        history_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        
        # History controls
        history_controls = ttk.Frame(history_frame)
        history_controls.grid(row=1, column=0, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Button(history_controls, text="🔄 Refresh History", 
                  command=self.refresh_history).grid(row=0, column=0, padx=(0, 10))
        
        ttk.Button(history_controls, text="📋 Export History", 
                  command=self.export_history).grid(row=0, column=1)
    
    def browse_template(self):
        """Browse for template DOCX file"""
        # Get user's home directory
        home_dir = os.path.expanduser("~")
        
        file_path = filedialog.askopenfilename(
            title="Select Template File",
            initialdir=home_dir,
            filetypes=[("Word Documents", "*.docx"), ("All Files", "*.*")]
        )
        if file_path:
            self.template_file.set(file_path)
            self.log_message(f"📄 Template file selected: {os.path.basename(file_path)}")
            self.update_file_info()
    
    def browse_csv(self):
        """Browse for CSV file"""
        # Get user's home directory
        home_dir = os.path.expanduser("~")
        
        file_path = filedialog.askopenfilename(
            title="Select CSV File",
            initialdir=home_dir,
            filetypes=[("CSV Files", "*.csv"), ("All Files", "*.*")]
        )
        if file_path:
            self.csv_file.set(file_path)
            self.log_message(f"📊 CSV file selected: {os.path.basename(file_path)}")
            self.update_file_info()
    
    def update_file_info(self):
        """Update file information display"""
        template = self.template_file.get()
        csv = self.csv_file.get()
        
        if template and csv:
            try:
                df = pd.read_csv(csv)
                record_count = len(df)
                pages_needed = int((record_count + 13) / 14)
                info = f"✅ Ready: {record_count} records → {pages_needed} pages"
                self.file_info.config(text=info, foreground='green')
                self.process_btn.config(state='normal')
            except Exception as e:
                info = f"⚠️ Error reading CSV: {str(e)}"
                self.file_info.config(text=info, foreground='red')
                self.process_btn.config(state='disabled')
        elif template or csv:
            missing = "CSV" if template else "Template"
            info = f"⏳ {missing} file needed"
            self.file_info.config(text=info, foreground='orange')
            self.process_btn.config(state='disabled')
        else:
            self.file_info.config(text="No files selected", foreground='gray')
            self.process_btn.config(state='disabled')
    
    def update_stats(self):
        """Update statistics display"""
        try:
            self.stats_text.config(state='normal')
            self.stats_text.delete(1.0, tk.END)
            
            stats = []
            stats.append("📊 APPLICATION STATISTICS")
            stats.append("=" * 30)
            
            # History stats
            if os.path.exists(self.history_file):
                df = pd.read_excel(self.history_file)
                total_processed = len(df)
                successful = len(df[df['Status'] == 'Success'])
                failed = len(df[df['Status'] == 'Failed'])
                
                stats.append(f"Total Processes: {total_processed}")
                stats.append(f"✅ Successful: {successful}")
                stats.append(f"❌ Failed: {failed}")
                
                if total_processed > 0:
                    success_rate = (successful / total_processed) * 100
                    stats.append(f"Success Rate: {success_rate:.1f}%")
                    
                    total_records = df['Records_Processed'].sum()
                    stats.append(f"Total Records: {total_records}")
            else:
                stats.append("No processing history yet")
            
            stats.append("")
            stats.append("📁 FILE SYSTEM")
            stats.append("-" * 20)
            
            # File system stats
            input_files = len([f for f in os.listdir('filesIn') if os.path.isfile(os.path.join('filesIn', f))]) if os.path.exists('filesIn') else 0
            output_files = len([f for f in os.listdir('fileOut') if os.path.isfile(os.path.join('fileOut', f))]) if os.path.exists('fileOut') else 0
            
            stats.append(f"Input Files: {input_files}")
            stats.append(f"Output Files: {output_files}")
            
            # Current files info
            if self.template_file.get():
                stats.append(f"Template: ✅")
            if self.csv_file.get():
                stats.append(f"CSV Data: ✅")
            
            self.stats_text.insert(tk.END, "\n".join(stats))
            self.stats_text.config(state='disabled')
            
        except Exception as e:
            pass
    
    def clear_status(self):
        """Clear the status log"""
        self.status_text.delete(1.0, tk.END)
        self.log_message("🗑️ Status log cleared")
    
    def export_history(self):
        """Export processing history"""
        try:
            if not os.path.exists(self.history_file):
                messagebox.showinfo("No History", "No processing history available to export.")
                return
            
            save_path = filedialog.asksaveasfilename(
                title="Export History As",
                defaultextension=".xlsx",
                filetypes=[("Excel Files", "*.xlsx"), ("CSV Files", "*.csv"), ("All Files", "*.*")],
                initialfile=f"processing_history_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )
            
            if save_path:
                if save_path.endswith('.csv'):
                    df = pd.read_excel(self.history_file)
                    df.to_csv(save_path, index=False)
                else:
                    shutil.copy2(self.history_file, save_path)
                
                messagebox.showinfo("Export Complete", f"History exported successfully to:\n{save_path}")
                self.log_message(f"📋 History exported to: {os.path.basename(save_path)}")
        except Exception as e:
            messagebox.showerror("Export Error", f"Failed to export history:\n{str(e)}")
    
    def log_message(self, message):
        """Add message to status log"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.status_text.insert(tk.END, f"[{timestamp}] {message}\n")
        self.status_text.see(tk.END)
        self.root.update_idletasks()
        
        # Update status indicator
        if "Error" in message or "Failed" in message:
            self.status_indicator.config(text="● Error", foreground='red')
        elif "Processing" in message:
            self.status_indicator.config(text="● Processing", foreground='orange')
        elif "completed" in message or "Success" in message:
            self.status_indicator.config(text="● Complete", foreground='green')
        elif "ready" in message.lower():
            self.status_indicator.config(text="● Ready", foreground='blue')
    
    def start_processing(self):
        """Start the processing in a separate thread"""
        if not self.template_file.get() or not self.csv_file.get():
            messagebox.showerror("Error", "Please select both template and CSV files.")
            return
        
        if self.processing:
            return
        
        # Update UI state
        self.progress_label.config(text="Preparing to process files...")
        self.status_indicator.config(text="● Starting", foreground='orange')
        
        # Start processing in separate thread to keep GUI responsive
        threading.Thread(target=self.process_files, daemon=True).start()
    
    def process_files(self):
        """Process the uploaded files"""
        self.processing = True
        self.process_btn.config(state='disabled')
        self.progress.start()
        
        # Update progress indicators
        self.root.after(0, lambda: self.progress_label.config(text="🚀 Processing vouchers..."))
        
        try:
            self.log_message("🚀 Starting file processing...")
            
            # Copy files to filesIn directory
            template_dest = os.path.join('filesIn', 'template.docx')
            csv_dest = os.path.join('filesIn', 'Gift Voucher.csv')
            
            # Copy template file (handle case where source and dest are the same)
            template_src = os.path.abspath(self.template_file.get())
            template_dst = os.path.abspath(template_dest)
            
            if template_src != template_dst:
                shutil.copy2(template_src, template_dst)
                self.log_message(f"📄 Template file copied to processing directory")
            else:
                self.log_message(f"📄 Template file already in processing directory")
            
            # Copy CSV file (handle case where source and dest are the same)
            csv_src = os.path.abspath(self.csv_file.get())
            csv_dst = os.path.abspath(csv_dest)
            
            if csv_src != csv_dst:
                shutil.copy2(csv_src, csv_dst)
                self.log_message(f"📊 CSV file copied to processing directory")
            else:
                self.log_message(f"📊 CSV file already in processing directory")
            
            # Read CSV to get record count
            try:
                df = pd.read_csv(csv_dest)
                record_count = len(df)
                pages_needed = int((record_count + 13) / 14)
                self.log_message(f"📊 CSV contains {record_count} records → {pages_needed} pages")
                self.root.after(0, lambda: self.progress_label.config(text=f"Processing {record_count} records..."))
            except Exception as e:
                self.log_message(f"❌ Error reading CSV: {e}")
                raise
            
            # Process vouchers using existing converter logic
            docx_output = os.path.join('fileOut', 'template_output.docx')
            pdf_output = os.path.join('fileOut', 'template_output.pdf')
            
            self.log_message("⚡ Processing vouchers into DOCX format...")
            self.root.after(0, lambda: self.progress_label.config(text="Creating DOCX document..."))
            success = process_vouchers(csv_dest, template_dest, docx_output)
            
            if success:
                self.log_message("✅ DOCX processing completed successfully")
                
                # Convert to PDF
                self.log_message("🔄 Converting DOCX to PDF...")
                self.root.after(0, lambda: self.progress_label.config(text="Converting to PDF..."))
                pdf_success = convert_to_pdf(docx_output, pdf_output)
                
                if pdf_success:
                    self.log_message("🎉 PDF conversion completed successfully!")
                    
                    # Remove intermediate DOCX file
                    try:
                        os.remove(docx_output)
                        self.log_message("🗑️ Intermediate DOCX file cleaned up")
                    except:
                        self.log_message("⚠️ Warning: Could not remove intermediate DOCX file")
                    
                    # Update UI
                    self.root.after(0, lambda: self.download_btn.config(state='normal'))
                    self.root.after(0, lambda: self.output_status.config(
                        text=f"✅ PDF ready: {os.path.basename(pdf_output)}", foreground='green'))
                    self.root.after(0, lambda: self.progress_label.config(text="✅ Processing complete!"))
                    
                    # Add to history
                    self.add_to_history(
                        template_file=os.path.basename(self.template_file.get()),
                        csv_file=os.path.basename(self.csv_file.get()),
                        records_processed=record_count,
                        pages_generated=pages_needed,
                        output_pdf=os.path.basename(pdf_output),
                        status="Success",
                        notes=f"Successfully processed {record_count} records into {pages_needed} pages"
                    )
                    
                    self.log_message("🎊 Processing completed! You can now download the PDF.")
                else:
                    raise Exception("PDF conversion failed")
            else:
                raise Exception("DOCX processing failed")
                
        except Exception as e:
            self.log_message(f"❌ Error during processing: {str(e)}")
            
            # Add error to history
            self.add_to_history(
                template_file=os.path.basename(self.template_file.get()) if self.template_file.get() else "N/A",
                csv_file=os.path.basename(self.csv_file.get()) if self.csv_file.get() else "N/A",
                records_processed=0,
                pages_generated=0,
                output_pdf="N/A",
                status="Failed",
                notes=str(e)
            )
            
            # Update UI
            self.root.after(0, lambda: self.progress_label.config(text="❌ Processing failed"))
            self.root.after(0, lambda: self.output_status.config(
                text="❌ Processing failed - no output generated", foreground='red'))
            
            messagebox.showerror("Processing Error", f"An error occurred during processing:\n{str(e)}")
        
        finally:
            self.processing = False
            self.root.after(0, lambda: self.process_btn.config(state='normal'))
            self.root.after(0, lambda: self.progress.stop())
            self.root.after(0, self.refresh_history)
            self.root.after(0, self.update_stats)
            self.root.after(0, self.update_file_info)
    
    def add_to_history(self, template_file, csv_file, records_processed, pages_generated, 
                      output_pdf, status, notes):
        """Add entry to processing history"""
        try:
            # Read existing history
            if os.path.exists(self.history_file):
                df = pd.read_excel(self.history_file)
            else:
                df = pd.DataFrame(columns=[
                    'Timestamp', 'Template_File', 'CSV_File', 'Records_Processed',
                    'Pages_Generated', 'Output_PDF', 'Status', 'Notes'
                ])
            
            # Add new entry
            new_entry = {
                'Timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'Template_File': template_file,
                'CSV_File': csv_file,
                'Records_Processed': records_processed,
                'Pages_Generated': pages_generated,
                'Output_PDF': output_pdf,
                'Status': status,
                'Notes': notes
            }
            
            df = pd.concat([df, pd.DataFrame([new_entry])], ignore_index=True)
            
            # Save updated history
            df.to_excel(self.history_file, index=False)
            
        except Exception as e:
            self.log_message(f"Error updating history: {e}")
    
    def refresh_history(self):
        """Refresh the history display"""
        try:
            # Clear existing entries first
            for item in self.history_tree.get_children():
                self.history_tree.delete(item)
            
            if os.path.exists(self.history_file):
                df = pd.read_excel(self.history_file)
                
                if len(df) > 0:
                    # Get last 20 entries in chronological order (oldest first)
                    recent_entries = df.tail(20)
                    
                    # Add entries one by one from oldest to newest
                    for idx, (_, row) in enumerate(recent_entries.iterrows()):
                        tag = 'success' if row['Status'] == 'Success' else 'failed'
                        
                        # Format timestamp for better readability
                        timestamp = row['Timestamp']
                        if isinstance(timestamp, str):
                            try:
                                # Parse and reformat timestamp
                                dt = pd.to_datetime(timestamp)
                                formatted_time = dt.strftime("%m/%d %H:%M")
                            except:
                                formatted_time = str(timestamp)[:10]  # First 10 chars
                        else:
                            formatted_time = str(timestamp)[:10]
                        
                        # Truncate template name if too long
                        template_name = str(row['Template_File'])
                        if len(template_name) > 15:
                            template_name = template_name[:12] + "..."
                        
                        # Insert each row separately with proper spacing
                        self.history_tree.insert('', 'end', values=(
                            formatted_time,
                            template_name,
                            str(row['Records_Processed']),
                            str(row['Status'])
                        ), tags=(tag,))
                        
                        # Allow GUI to update after each insertion
                        self.root.update_idletasks()
            
        except Exception as e:
            self.log_message(f"❌ Error refreshing history: {e}")
    
    def download_pdf(self):
        """Download the generated PDF"""
        pdf_path = os.path.join('fileOut', 'template_output.pdf')
        
        if not os.path.exists(pdf_path):
            messagebox.showerror("Error", "PDF file not found. Please process files first.")
            return
        
        # Get user's home directory
        home_dir = os.path.expanduser("~")
        
        # Ask user where to save the file
        save_path = filedialog.asksaveasfilename(
            title="Save PDF As",
            initialdir=home_dir,
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            initialfile=f"vouchers_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
        )
        
        if save_path:
            try:
                shutil.copy2(pdf_path, save_path)
                messagebox.showinfo("Success", f"PDF saved successfully to:\n{save_path}")
                self.log_message(f"💾 PDF downloaded to: {save_path}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save PDF:\n{str(e)}")
    
    def open_output_folder(self):
        """Open the output folder in file explorer"""
        output_path = os.path.abspath('fileOut')
        try:
            if os.name == 'nt':  # Windows
                os.startfile(output_path)
            elif os.name == 'posix':  # macOS and Linux
                os.system(f'xdg-open "{output_path}"')
        except Exception as e:
            self.log_message(f"Could not open output folder: {e}")
            messagebox.showinfo("Output Folder", f"Output folder location:\n{output_path}")


def main():
    """Main function to run the GUI application"""
    root = tk.Tk()
    
    # Set the application icon (if available)
    try:
        root.iconbitmap('icon.ico')  # You can add an icon file
    except:
        pass
    
    # Create and run the application
    app = VoucherProcessorGUI(root)
    
    # Window is already positioned and sized in __init__
    root.update_idletasks()
    
    # Start the GUI event loop
    root.mainloop()


if __name__ == "__main__":
    main()
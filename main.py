import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import json
import sys
import os
import threading
import subprocess
import platform

# Add the src directory to the Python path
sys.path.append(os.path.join(os.path.dirname(__file__), 'src'))

from src.excel_processor import ExcelProcessor
from src.tally_api_service import TallyApiService


class AccountSelectionDialog:
    """Dialog for selecting account name for payroll upload."""
    
    def __init__(self, parent):
        self.result = None
        
        # Create dialog window
        self.top = tk.Toplevel(parent)
        self.top.title("Select Account")
        self.top.geometry("400x200")
        self.top.resizable(False, False)
        
        # Center the dialog
        self.top.transient(parent)
        self.top.grab_set()
        
        # Create frame
        frame = ttk.Frame(self.top, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(frame, text="Select Account for Payroll", font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(frame, text="Enter the account name (PARTYLEDGERNAME) to be used for this payroll voucher:", wraplength=350)
        desc_label.pack(pady=(0, 10))
        
        # Account entry
        self.account_var = tk.StringVar(value="Cash")  # Default to "Cash"
        ttk.Label(frame, text="Account Name:").pack(anchor='w')
        self.account_entry = ttk.Entry(frame, textvariable=self.account_var, width=40)
        self.account_entry.pack(fill=tk.X, pady=(5, 15))
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.RIGHT)
        
        # Focus and bind enter
        self.account_entry.focus()
        self.account_entry.bind('<Return>', lambda e: self.ok())
        self.top.bind('<Escape>', lambda e: self.cancel())
        
        # Center the dialog on parent
        self.center_on_parent(parent)
    
    def center_on_parent(self, parent):
        """Center dialog on parent window."""
        self.top.update_idletasks()
        
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        dialog_width = self.top.winfo_width()
        dialog_height = self.top.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.top.geometry(f"+{x}+{y}")
    
    def ok(self):
        """Handle OK button."""
        account_name = self.account_var.get().strip()
        if account_name:
            self.result = account_name
            self.top.destroy()
        else:
            messagebox.showerror("Error", "Please enter an account name.")
    
    def cancel(self):
        """Handle Cancel button."""
        self.result = None
        self.top.destroy()


class NarrationDialog:
    """Dialog for entering narration for XML generation."""
    
    def __init__(self, parent, title="Enter Narration", default_text="", optional=False):
        self.result = None
        
        # Create dialog window
        self.top = tk.Toplevel(parent)
        self.top.title(title)
        self.top.geometry("450x250")
        self.top.resizable(False, False)
        
        # Center the dialog
        self.top.transient(parent)
        self.top.grab_set()
        
        # Create frame
        frame = ttk.Frame(self.top, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(frame, text=title, font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_text = "Enter the narration text to be included in the XML file:"
        if optional:
            desc_text += " (Optional - leave blank if not needed)"
        
        desc_label = ttk.Label(frame, text=desc_text, wraplength=400)
        desc_label.pack(pady=(0, 10))
        
        # Narration entry
        self.narration_var = tk.StringVar(value=default_text)
        ttk.Label(frame, text="Narration:").pack(anchor='w')
        self.narration_entry = tk.Text(frame, height=4, width=50, wrap=tk.WORD)
        self.narration_entry.pack(fill=tk.BOTH, expand=True, pady=(5, 15))
        self.narration_entry.insert('1.0', default_text)
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=(5, 0))
        if optional:
            ttk.Button(button_frame, text="Skip", command=self.skip).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.RIGHT)
        
        # Focus and bind enter
        self.narration_entry.focus()
        self.top.bind('<Escape>', lambda e: self.cancel())
        
        # Center the dialog on parent
        self.center_on_parent(parent)
    
    def center_on_parent(self, parent):
        """Center dialog on parent window."""
        self.top.update_idletasks()
        
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        dialog_width = self.top.winfo_width()
        dialog_height = self.top.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.top.geometry(f"+{x}+{y}")
    
    def ok(self):
        """Handle OK button."""
        narration_text = self.narration_entry.get('1.0', tk.END).strip()
        self.result = narration_text if narration_text else ""
        self.top.destroy()
    
    def skip(self):
        """Handle Skip button (for optional narration)."""
        self.result = ""
        self.top.destroy()
    
    def cancel(self):
        """Handle Cancel button."""
        self.result = None
        self.top.destroy()


class PayeAccountSelectionDialog:
    """Dialog for selecting account name for PAYE upload."""
    
    def __init__(self, parent):
        self.result = None
        
        # Create dialog window
        self.top = tk.Toplevel(parent)
        self.top.title("Select Bank Account")
        self.top.geometry("450x200")
        self.top.resizable(False, False)
        
        # Center the dialog
        self.top.transient(parent)
        self.top.grab_set()
        
        # Create frame
        frame = ttk.Frame(self.top, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(frame, text="Select Bank Account for PAYE Payment", font=('Arial', 12, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Description
        desc_label = ttk.Label(frame, text="Enter the bank account name to be used for PAYE/SDL payment:", wraplength=400)
        desc_label.pack(pady=(0, 10))
        
        # Account entry
        self.account_var = tk.StringVar(value="THE PEOLPE'S BANK OF ZANZIBAR LIMITED - TZS")  # Default bank account
        ttk.Label(frame, text="Bank Account Name:").pack(anchor='w')
        self.account_entry = ttk.Entry(frame, textvariable=self.account_var, width=50)
        self.account_entry.pack(fill=tk.X, pady=(5, 15))
        
        # Buttons
        button_frame = ttk.Frame(frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(button_frame, text="Cancel", command=self.cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=self.ok).pack(side=tk.RIGHT)
        
        # Focus and bind enter
        self.account_entry.focus()
        self.account_entry.bind('<Return>', lambda e: self.ok())
        self.top.bind('<Escape>', lambda e: self.cancel())
        
        # Center the dialog on parent
        self.center_on_parent(parent)
    
    def center_on_parent(self, parent):
        """Center dialog on parent window."""
        self.top.update_idletasks()
        
        parent_x = parent.winfo_rootx()
        parent_y = parent.winfo_rooty()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()
        
        dialog_width = self.top.winfo_width()
        dialog_height = self.top.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        self.top.geometry(f"+{x}+{y}")
    
    def ok(self):
        """Handle OK button."""
        account_name = self.account_var.get().strip()
        if account_name:
            self.result = account_name
            self.top.destroy()
        else:
            messagebox.showerror("Error", "Please enter a bank account name.")
    
    def cancel(self):
        """Handle Cancel button."""
        self.result = None
        self.top.destroy()


class ModernExcelProcessor:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Excel File Processor - Attendance & Payroll")
        self.root.geometry("1000x680")
        self.root.minsize(900, 600)
        
        # Configure style
        self.setup_style()
        
        # Initialize Excel processor and API service
        self.processor = ExcelProcessor()
        self.api_service = TallyApiService()
        self.selected_company = None
        
        # Show company selection screen first
        self.show_company_selection_screen()
        
        # Center the window
        self.center_window()
    
    def setup_style(self):
        """Set up modern styling for the application."""
        style = ttk.Style()
        
        # Configure colors
        self.colors = {
            'primary': '#2c3e50',
            'secondary': '#3498db',
            'success': '#27ae60',
            'warning': '#f39c12',
            'danger': '#e74c3c',
            'light': '#ecf0f1',
            'dark': '#34495e',
            'white': '#ffffff'
        }
        
        # Configure root window
        self.root.configure(bg=self.colors['light'])
        
        # Configure ttk styles
        style.theme_use('clam')
        
        # Enhanced button styles with optimized padding
        style.configure('Primary.TButton',
                       background=self.colors['primary'],
                       foreground=self.colors['white'],
                       font=('Arial', 11, 'bold'),
                       padding=(20, 12),
                       borderwidth=0)
        
        style.configure('Secondary.TButton',
                       background=self.colors['secondary'],
                       foreground=self.colors['white'],
                       font=('Arial', 11, 'bold'),
                       padding=(20, 12),
                       borderwidth=0)
        
        style.configure('Success.TButton',
                       background=self.colors['success'],
                       foreground=self.colors['white'],
                       font=('Arial', 11, 'bold'),
                       padding=(20, 12),
                       borderwidth=0)
        
        # Button hover effects
        style.map('Primary.TButton',
                 background=[('active', '#1a252f')])
        style.map('Secondary.TButton', 
                 background=[('active', '#2980b9')])
        style.map('Success.TButton',
                 background=[('active', '#229954')])
        
        # Configure label styles with bigger title
        style.configure('Title.TLabel',
                       background=self.colors['light'],
                       foreground=self.colors['primary'],
                       font=('Arial', 20, 'bold'))
        
        style.configure('Subtitle.TLabel',
                       background=self.colors['light'],
                       foreground=self.colors['dark'],
                       font=('Arial', 12))
        
        style.configure('Info.TLabel',
                       background=self.colors['light'],
                       foreground=self.colors['dark'],
                       font=('Arial', 10))
        
        # Configure combobox styles
        style.configure('TCombobox',
                       fieldbackground=self.colors['white'],
                       background=self.colors['light'],
                       foreground=self.colors['dark'],
                       font=('Arial', 12),
                       borderwidth=2,
                       relief='solid')
        
        style.map('TCombobox',
                 fieldbackground=[('readonly', self.colors['white']),
                                ('focus', self.colors['light'])],
                 bordercolor=[('focus', self.colors['secondary'])])
    
    def center_window(self):
        """Center the window on the screen."""
        self.root.update_idletasks()
        width = self.root.winfo_width()
        height = self.root.winfo_height()
        pos_x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        pos_y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f"{width}x{height}+{pos_x}+{pos_y}")
    
    def show_company_selection_screen(self):
        """Show company selection screen at startup."""
        # Clear any existing widgets
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create main container with centered content
        main_container = tk.Frame(self.root, bg=self.colors['light'])
        main_container.pack(fill=tk.BOTH, expand=True)
        
        # Center frame
        center_frame = tk.Frame(main_container, bg=self.colors['light'])
        center_frame.pack(expand=True)
        
        # Company selection card
        self.company_frame = tk.Frame(center_frame, bg=self.colors['white'], relief=tk.RAISED, bd=2)
        self.company_frame.pack(padx=50, pady=80, ipadx=40, ipady=40)
        
        # Header section
        header_frame = tk.Frame(self.company_frame, bg=self.colors['white'])
        header_frame.pack(fill=tk.X, pady=(0, 30))
        
        # Title with better spacing
        title_label = tk.Label(
            header_frame, 
            text="🏢 Company Selection", 
            font=('Arial', 24, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['primary']
        )
        title_label.pack(pady=(0, 10))
        
        # Subtitle
        subtitle_label = tk.Label(
            header_frame, 
            text="Please select a company from Tally to proceed", 
            font=('Arial', 14),
            bg=self.colors['white'],
            fg=self.colors['dark']
        )
        subtitle_label.pack()
        
        # Separator line
        separator = tk.Frame(self.company_frame, height=2, bg=self.colors['light'])
        separator.pack(fill=tk.X, pady=(20, 30))
        
        # Company selection section
        selection_frame = tk.Frame(self.company_frame, bg=self.colors['white'])
        selection_frame.pack(fill=tk.X, pady=(0, 30))
        
        # Company dropdown label
        dropdown_label = tk.Label(
            selection_frame, 
            text="📋 Select Company:", 
            font=('Arial', 14, 'bold'),
            bg=self.colors['white'],
            fg=self.colors['primary']
        )
        dropdown_label.pack(pady=(0, 15))
        
        # Dropdown container for better styling
        dropdown_container = tk.Frame(selection_frame, bg=self.colors['light'], relief=tk.SUNKEN, bd=1)
        dropdown_container.pack(pady=(0, 25), padx=20, fill=tk.X)
        
        # Company dropdown with better styling
        self.company_var = tk.StringVar(value="Loading companies...")
        self.company_dropdown = ttk.Combobox(
            dropdown_container,
            textvariable=self.company_var,
            state="readonly",
            font=('Arial', 14),
            width=35,
            height=10
        )
        self.company_dropdown.pack(pady=8, padx=8, fill=tk.X)
        
        # Buttons frame with centered layout
        button_frame = tk.Frame(selection_frame, bg=self.colors['white'])
        button_frame.pack(fill=tk.X, pady=(20, 0))
        
        # Button container for centering
        btn_container = tk.Frame(button_frame, bg=self.colors['white'])
        btn_container.pack()
        
        # Refresh button with better styling
        self.refresh_btn = ttk.Button(
            btn_container,
            text="🔄 Refresh Companies",
            command=self.load_companies,
            style='Secondary.TButton'
        )
        self.refresh_btn.pack(side=tk.LEFT, padx=(0, 30))
        
        # Proceed button with better styling
        self.proceed_btn = ttk.Button(
            btn_container,
            text="➡️ Proceed to Application",
            command=self.proceed_to_main,
            style='Success.TButton',
            state='disabled'
        )
        self.proceed_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        # Skip company selection button for testing
        self.skip_btn = ttk.Button(
            btn_container,
            text="🚀 Skip & Test Upload",
            command=self.skip_company_selection,
            style='Secondary.TButton'
        )
        self.skip_btn.pack(side=tk.LEFT)
        
        # Status section
        status_frame = tk.Frame(self.company_frame, bg=self.colors['white'])
        status_frame.pack(fill=tk.X, pady=(30, 0))
        
        # Status with icon
        self.company_status_label = tk.Label(
            status_frame,
            text="🔄 Loading companies from Tally...",
            font=('Arial', 12),
            bg=self.colors['white'],
            fg=self.colors['secondary']
        )
        self.company_status_label.pack()
        
        # Load companies automatically
        self.load_companies()
        
        # Bind selection change
        self.company_dropdown.bind('<<ComboboxSelected>>', self.on_company_selected)
    
    def load_companies(self):
        """Load companies from Tally API."""
        # Disable buttons during loading
        self.refresh_btn.configure(state='disabled', text='Loading...')
        self.proceed_btn.configure(state='disabled')
        self.company_status_label.configure(text="🔄 Connecting to Tally...", fg=self.colors['secondary'])
        self.company_var.set("Loading companies...")
        self.root.update()
        
        # Load companies in background thread
        thread = threading.Thread(target=self._load_companies_background)
        thread.daemon = True
        thread.start()
    
    def _load_companies_background(self):
        """Load companies in background thread."""
        try:
            result = self.api_service.get_companies()
            # Update UI in main thread
            self.root.after(0, self._handle_companies_result, result)
        except Exception as e:
            error_result = {"success": False, "error": str(e), "companies": []}
            self.root.after(0, self._handle_companies_result, error_result)
    
    def _handle_companies_result(self, result):
        """Handle companies loading result."""
        # Restore refresh button
        self.refresh_btn.configure(state='normal', text='🔄 Refresh Companies')
        
        if result.get("success", False):
            companies = result.get("companies", [])
            if companies:
                # Populate dropdown
                self.company_dropdown['values'] = companies
                self.company_var.set("-- Please Select a Company --")
                self.company_status_label.configure(
                    text=f"✅ Successfully loaded {len(companies)} companies from Tally",
                    fg=self.colors['success']
                )
            else:
                self.company_var.set("No companies available")
                self.company_status_label.configure(
                    text="⚠️ No companies found in your Tally database",
                    fg=self.colors['warning']
                )
        else:
            error_msg = result.get("error", "Unknown error")
            self.company_var.set("Connection failed - Click Refresh")
            self.company_status_label.configure(
                text=f"❌ Connection Error: {error_msg}",
                fg=self.colors['danger']
            )
    
    def on_company_selected(self, event=None):
        """Handle company selection."""
        selected = self.company_var.get()
        invalid_selections = [
            "Loading companies...", 
            "-- Please Select a Company --", 
            "No companies available", 
            "Connection failed - Click Refresh"
        ]
        
        if selected and selected not in invalid_selections:
            self.selected_company = selected
            self.proceed_btn.configure(state='normal')
            # Update status to show selection
            self.company_status_label.configure(
                text=f"🏢 Selected: {selected} - Ready to proceed",
                fg=self.colors['primary']
            )
        else:
            self.selected_company = None
            self.proceed_btn.configure(state='disabled')
    
    def proceed_to_main(self):
        """Proceed to main application with selected company."""
        if self.selected_company:
            # Clear company selection screen
            for widget in self.root.winfo_children():
                widget.destroy()
            
            # Create new main upload screen
            self.create_main_upload_screen()
        else:
            messagebox.showwarning("No Company Selected", "Please select a company to proceed.")
    
    def skip_company_selection(self):
        """Skip company selection for testing purposes."""
        # Set a default test company
        self.selected_company = "TEST COMPANY (No Tally Connection)"
        
        # Clear company selection screen
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Create main upload screen
        self.create_main_upload_screen()
    
    def create_main_upload_screen(self):
        """Create the main upload screen with attendance and payroll options."""
        # Main container (more compact)
        self.main_frame = tk.Frame(self.root, bg=self.colors['light'], padx=20, pady=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header section (more compact)
        header_frame = tk.Frame(self.main_frame, bg=self.colors['light'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Company info banner
        is_test_mode = "TEST COMPANY" in self.selected_company
        banner_color = self.colors['warning'] if is_test_mode else self.colors['primary']
        banner_icon = "🧪" if is_test_mode else "🏢"
        banner_text = f"{banner_icon} {'Test Mode - ' if is_test_mode else 'Connected to: '}{self.selected_company}"
        
        company_banner = tk.Frame(header_frame, bg=banner_color, relief=tk.RAISED, bd=1)
        company_banner.pack(fill=tk.X, pady=(0, 20))
        
        company_info = tk.Label(
            company_banner,
            text=banner_text,
            font=('Arial', 14, 'bold'),
            bg=banner_color,
            fg=self.colors['white'],
            pady=15
        )
        company_info.pack()
        
        # Title section with Download Example button
        title_section = tk.Frame(header_frame, bg=self.colors['light'])
        title_section.pack(fill=tk.X, pady=(0, 10))
        
        # Left side - Download Example button
        download_frame = tk.Frame(title_section, bg=self.colors['light'])
        download_frame.pack(side=tk.LEFT)
        
        download_btn = ttk.Button(
            download_frame,
            text="📥 Download Example",
            command=self.download_example_file,
            style='Primary.TButton'
        )
        download_btn.pack(pady=5)
        
        # Center - Title
        title_center = tk.Frame(title_section, bg=self.colors['light'])
        title_center.pack(expand=True)
        
        # Main title (smaller)
        title_label = tk.Label(
            title_center,
            text="📊 Excel File Processor & XML Generator",
            font=('Arial', 20, 'bold'),
            bg=self.colors['light'],
            fg=self.colors['primary']
        )
        title_label.pack()
        
        # Subtitle (smaller)
        subtitle_label = tk.Label(
            title_center,
            text="Upload Excel files and generate Tally-compatible XML for import",
            font=('Arial', 12),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        subtitle_label.pack()
        
        # Content container - Grid layout for multiple upload boxes
        content_frame = tk.Frame(self.main_frame, bg=self.colors['light'])
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # Create grid for upload cards (3 cards in a row)
        # Row 1: Current upload types
        row1_frame = tk.Frame(content_frame, bg=self.colors['light'])
        row1_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Left card - Attendance
        left_frame = tk.Frame(row1_frame, bg=self.colors['light'])
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Center card - Payroll  
        center_frame = tk.Frame(row1_frame, bg=self.colors['light'])
        center_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 5))
        
        # Right card - PAYE
        right_frame = tk.Frame(row1_frame, bg=self.colors['light'])
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Attendance Section
        self.create_upload_card(left_frame, "Attendance Upload", "attendance", "📊", 
                               "Upload attendance Excel files and generate XML for Tally import")
        
        # Payroll Section
        self.create_upload_card(center_frame, "Payroll Upload", "payroll", "💰", 
                               "Upload payroll Excel files and generate payment XML for Tally import")
        
        # PAYE Section
        self.create_upload_card(right_frame, "PAYE Upload", "paye", "🏛️", 
                               "Upload PAYE Excel sheet and generate PAYE/SDL payment XML for Tally import")
        
        # Row 2: Additional upload types
        row2_frame = tk.Frame(content_frame, bg=self.colors['light'])
        row2_frame.pack(fill=tk.X, pady=(10, 0))
        
        # Left card - ZSSF
        zssf_frame = tk.Frame(row2_frame, bg=self.colors['light'])
        zssf_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Right card - ZHSF
        zhsf_frame = tk.Frame(row2_frame, bg=self.colors['light'])
        zhsf_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Empty frame for balance (since we only have 2 cards in this row)
        empty_frame = tk.Frame(row2_frame, bg=self.colors['light'])
        empty_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # ZSSF Section
        self.create_upload_card(zssf_frame, "ZSSF Upload", "zssf", "🏦", 
                               "Upload Excel sheet 3 for ZSSF processing (XML generation disabled for now)")
        
        # ZHSF Section  
        self.create_upload_card(zhsf_frame, "ZHSF Upload", "zhsf", "🏥", 
                               "Upload Excel sheet 4 for ZHSF processing (XML generation disabled for now)")
        
        # Status section at bottom
        status_frame = tk.Frame(self.main_frame, bg=self.colors['light'])
        status_frame.pack(fill=tk.X, pady=(20, 0))
        
        status_title = tk.Label(
            status_frame,
            text="📊 Status",
            font=('Arial', 16, 'bold'),
            bg=self.colors['light'],
            fg=self.colors['primary']
        )
        status_title.pack(anchor=tk.W, pady=(0, 10))
        
        # Status display
        self.status_frame = tk.Frame(status_frame, bg=self.colors['white'], relief=tk.SOLID, bd=1)
        self.status_frame.pack(fill=tk.X, padx=5, pady=5)
        
        self.status_text = tk.Label(
            self.status_frame,
            text="✨ Ready to process Excel files and generate XML...",
            font=('Arial', 12),
            bg=self.colors['white'],
            fg=self.colors['dark'],
            wraplength=800,
            justify=tk.LEFT,
            pady=15,
            padx=20
        )
        self.status_text.pack(anchor=tk.W)
    
    def create_upload_card(self, parent, title, upload_type, icon, description):
        """Create a compact card for file upload functionality."""
        # Main card container (smaller)
        card_frame = tk.Frame(parent, bg=self.colors['white'], relief=tk.RAISED, bd=1)
        card_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        # Card header (smaller)
        header = tk.Frame(card_frame, bg=self.colors['primary'])
        header.pack(fill=tk.X)
        
        header_label = tk.Label(
            header,
            text=f"{icon} {title}",
            font=('Arial', 14, 'bold'),
            bg=self.colors['primary'],
            fg=self.colors['white'],
            pady=8
        )
        header_label.pack()
        
        # Card content (more compact)
        content = tk.Frame(card_frame, bg=self.colors['white'], padx=15, pady=15)
        content.pack(fill=tk.BOTH, expand=True)
        
        # Description (smaller)
        desc_label = tk.Label(
            content,
            text=description,
            font=('Arial', 10),
            bg=self.colors['white'],
            fg=self.colors['dark'],
            wraplength=250,
            justify=tk.LEFT
        )
        desc_label.pack(pady=(0, 15))
        
        # File selection area (compact)
        file_frame = tk.Frame(content, bg=self.colors['light'], relief=tk.SUNKEN, bd=1)
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # File path variable
        file_path_var = tk.StringVar(value="No file selected")
        setattr(self, f"{upload_type}_file_path", file_path_var)
        setattr(self, f"{upload_type}_result", None)
        
        file_label = tk.Label(
            file_frame,
            text="📁 Selected File:",
            font=('Arial', 9, 'bold'),
            bg=self.colors['light'],
            fg=self.colors['dark']
        )
        file_label.pack(anchor=tk.W, padx=8, pady=(8, 3))
        
        file_display = tk.Label(
            file_frame,
            textvariable=file_path_var,
            font=('Arial', 9),
            bg=self.colors['light'],
            fg=self.colors['dark'],
            wraplength=220,
            justify=tk.LEFT
        )
        file_display.pack(anchor=tk.W, padx=8, pady=(0, 8))
        
        # Buttons
        btn_frame = tk.Frame(content, bg=self.colors['white'])
        btn_frame.pack(fill=tk.X)
        
        # Browse button
        browse_btn = ttk.Button(
            btn_frame,
            text="📂 Browse File",
            command=lambda: self.browse_upload_file(upload_type),
            style='Primary.TButton'
        )
        browse_btn.pack(fill=tk.X, pady=(0, 8))
        
        # Process button
        process_btn = ttk.Button(
            btn_frame,
            text=f"⚙️ Process",
            command=lambda: self.process_upload_file(upload_type),
            style='Secondary.TButton',
            state='disabled'
        )
        process_btn.pack(fill=tk.X, pady=(0, 8))
        setattr(self, f"{upload_type}_process_btn", process_btn)
        
        # Generate and upload to telly button
        xml_btn = ttk.Button(
            btn_frame,
            text="� Generate and upload to telly",
            command=lambda: self.generate_xml_file(upload_type),
            style='Success.TButton',
            state='disabled'
        )
        xml_btn.pack(fill=tk.X, pady=(0, 5))
        setattr(self, f"{upload_type}_xml_btn", xml_btn)
    
    def browse_upload_file(self, upload_type):
        """Browse and select file for upload."""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
        
        file_path = filedialog.askopenfilename(
            title=f"Select {upload_type.title()} Excel File",
            filetypes=file_types
        )
        
        if file_path:
            file_path_var = getattr(self, f"{upload_type}_file_path")
            file_path_var.set(os.path.basename(file_path))
            setattr(self, f"{upload_type}_full_path", file_path)
            
            # Enable process button
            process_btn = getattr(self, f"{upload_type}_process_btn")
            process_btn.configure(state='normal')
            
            self.update_status(f"📁 Selected {upload_type} file: {os.path.basename(file_path)}")
    
    def process_upload_file(self, upload_type):
        """Process the uploaded file."""
        file_path = getattr(self, f"{upload_type}_full_path", None)
        if not file_path:
            messagebox.showerror("Error", "Please select a file first.")
            return
        
        # Disable buttons during processing
        process_btn = getattr(self, f"{upload_type}_process_btn")
        xml_btn = getattr(self, f"{upload_type}_xml_btn")
        process_btn.configure(state='disabled', text='Processing...')
        
        self.update_status(f"⚙️ Processing {upload_type} file...")
        
        # Process in background thread
        thread = threading.Thread(target=self._process_upload_background, args=(upload_type, file_path))
        thread.daemon = True
        thread.start()
    
    def _process_upload_background(self, upload_type, file_path):
        """Process file in background thread."""
        try:
            if upload_type == "attendance":
                result = self.processor.process_attendance_sheet(file_path)
            elif upload_type == "payroll":
                result = self.processor.process_payroll_sheet(file_path)
            elif upload_type == "paye":
                result = self.processor.process_paye_sheet(file_path)
            elif upload_type == "zssf":
                result = self.processor.process_zssf_sheet(file_path)
            elif upload_type == "zhsf":
                result = self.processor.process_zhsf_sheet(file_path)
            else:
                result = {"error": f"Unknown upload type: {upload_type}", "success": False}
            
            # Update UI in main thread
            self.root.after(0, lambda: self._handle_upload_processing_result(upload_type, result))
            
        except Exception as e:
            error_result = {"error": str(e), "success": False}
            self.root.after(0, lambda: self._handle_upload_processing_result(upload_type, error_result))
    
    def _handle_upload_processing_result(self, upload_type, result):
        """Handle the processing result."""
        process_btn = getattr(self, f"{upload_type}_process_btn")
        xml_btn = getattr(self, f"{upload_type}_xml_btn")
        
        if result.get("success", False):
            # Store result
            setattr(self, f"{upload_type}_result", result)
            
            # Update buttons
            process_btn.configure(state='normal', text=f"✅ Processed")
            xml_btn.configure(state='normal')
            
            # Update status
            employee_count = result.get("total_employees", 0)
            if upload_type == "payroll":
                total_amount = result.get("total_gross_salary", 0)
                self.update_status(f"✅ {upload_type.title()} processed successfully! "
                                 f"{employee_count} employees, Total: ₹{total_amount:,.2f}")
            elif upload_type == "paye":
                total_paye = result.get("total_paye", 0)
                total_sdl = result.get("total_sdl", 0)
                total_amount = result.get("total_amount", 0)
                self.update_status(f"✅ {upload_type.title()} processed successfully! "
                                 f"{employee_count} employees, PAYE: ₹{total_paye:,.2f}, SDL: ₹{total_sdl:,.2f}, Total: ₹{total_amount:,.2f}")
            elif upload_type == "zssf":
                total_zssf = result.get("total_zssf", 0)
                self.update_status(f"✅ {upload_type.upper()} processed successfully! "
                                 f"{employee_count} employees, Total ZSSF: ₹{total_zssf:,.2f}")
            elif upload_type == "zhsf":
                total_zhsf = result.get("total_zhsf", 0)
                self.update_status(f"✅ {upload_type.upper()} processed successfully! "
                                 f"{employee_count} employees, Total ZHSF: ₹{total_zhsf:,.2f}")
            else:
                self.update_status(f"✅ {upload_type.title()} processed successfully! "
                                 f"{employee_count} employees found")
        else:
            # Handle error
            process_btn.configure(state='normal', text=f"❌ Error - Retry")
            error_msg = result.get("error", "Unknown error occurred")
            self.update_status(f"❌ Error processing {upload_type}: {error_msg}")
            messagebox.showerror("Processing Error", f"Failed to process {upload_type} file:\n{error_msg}")
    
    def generate_xml_file(self, upload_type):
        """Generate XML file from processed data."""
        result = getattr(self, f"{upload_type}_result", None)
        if not result:
            messagebox.showerror("Error", f"Please process the {upload_type} file first.")
            return
        
        try:
            self.update_status(f"🔄 Generating {upload_type} XML...")
            
            # Use selected company name, or extract from result if testing
            company_name = self.selected_company
            if "TEST COMPANY" in company_name:
                company_name = result.get("company_name", "TEST COMPANY")
            
            # Ask for NARRATION with safe string concatenation
            if upload_type == "attendance":
                # For attendance, show popup to ask for narration
                date_str = result.get("date") or "current period"
                default_text = f"Attendance for {date_str}"
                
                narration_dialog = NarrationDialog(
                    self.root, 
                    title="Enter Attendance Narration",
                    default_text=default_text,
                    optional=False
                )
                self.root.wait_window(narration_dialog.top)
                narration = narration_dialog.result
                
                if narration is None:  # User cancelled
                    return
                
                xml_content = self.processor.generate_attendance_xml(result, company_name, narration)
            elif upload_type == "payroll":
                # For payroll, ask for account name
                account_name = self.get_account_name_for_payroll()
                if not account_name:
                    return
                
                # Ask for narration (optional for payroll)
                date_str = result.get("date") or "current period"
                default_text = f"Payroll for {date_str}"
                
                narration_dialog = NarrationDialog(
                    self.root, 
                    title="Enter Payroll Narration (Optional)",
                    default_text=default_text,
                    optional=True
                )
                self.root.wait_window(narration_dialog.top)
                narration = narration_dialog.result
                
                if narration is None:  # User cancelled
                    return
                
                xml_content = self.processor.generate_payroll_xml(result, company_name, account_name, narration)
            elif upload_type == "paye":
                # For PAYE, ask for account name
                account_name = self.get_account_name_for_paye()
                if not account_name:
                    return
                
                # Ask for narration (optional for PAYE) - safe string concatenation
                date_str = result.get("date") or "current period"
                default_text = f"PAYE and SDL for {date_str}"
                
                narration_dialog = NarrationDialog(
                    self.root, 
                    title="Enter PAYE Narration (Optional)",
                    default_text=default_text,
                    optional=True
                )
                self.root.wait_window(narration_dialog.top)
                narration = narration_dialog.result
                
                if narration is None:  # User cancelled
                    return
                
                xml_content = self.processor.generate_paye_xml(result, company_name, account_name, narration)
            elif upload_type == "zssf":
                # For ZSSF, ask for account name
                account_name = self.get_account_name_for_zssf()
                if not account_name:
                    return
                
                # Ask for narration (optional for ZSSF) - safe string concatenation
                date_str = result.get("date") or "current period"
                default_text = f"ZSSF for {date_str}"
                
                narration_dialog = NarrationDialog(
                    self.root, 
                    title="Enter ZSSF Narration (Optional)",
                    default_text=default_text,
                    optional=True
                )
                self.root.wait_window(narration_dialog.top)
                narration = narration_dialog.result
                
                if narration is None:  # User cancelled
                    return
                
                xml_content = self.processor.generate_zssf_xml(result, company_name, account_name, narration)
            elif upload_type == "zhsf":
                # For ZHSF, ask for account name
                account_name = self.get_account_name_for_zhsf()
                if not account_name:
                    return
                
                # Ask for narration (optional for ZHSF) - safe string concatenation
                date_str = result.get("date") or "current period"
                default_text = f"ZHSF for {date_str}"
                
                narration_dialog = NarrationDialog(
                    self.root, 
                    title="Enter ZHSF Narration (Optional)",
                    default_text=default_text,
                    optional=True
                )
                self.root.wait_window(narration_dialog.top)
                narration = narration_dialog.result
                
                if narration is None:  # User cancelled
                    return
                
                xml_content = self.processor.generate_zhsf_xml(result, company_name, account_name, narration)
            else:
                messagebox.showerror("Error", f"Unknown upload type: {upload_type}")
                return
            
            if xml_content:
                # Always save XML file first (regardless of Tally server status)
                output_path = self.processor.save_xml_file(xml_content, upload_type)
                if output_path:
                    # Show success message for XML generation
                    file_type_name = "PAYE" if upload_type == "paye" else upload_type.title()
                    self.update_status(f"✅ {file_type_name} XML file created: {output_path}")
                    
                    # Check if we're in test mode (no Tally server)
                    is_test_mode = "TEST COMPANY" in self.selected_company
                    
                    if is_test_mode:
                        # Test mode - skip Tally upload, just show XML success
                        messagebox.showinfo("XML Generated Successfully (Test Mode)", 
                                          f"{file_type_name} XML file has been generated and saved!\n\n"
                                          f"📁 File: {output_path}\n\n"
                                          f"🧪 Test Mode: Tally upload skipped\n"
                                          f"You can manually import this file into Tally when ready.")
                        self.update_status(f"✅ {file_type_name} XML generated successfully (Test Mode - No Tally Upload)")
                        
                        # Ask if user wants to open file location
                        if messagebox.askyesno("Open File Location?", 
                                             f"XML file saved successfully!\n\n"
                                             f"📁 Location: {output_path}\n\n"
                                             f"Would you like to open the file location?"):
                            import subprocess
                            subprocess.run(["open", "-R", output_path])  # macOS
                    else:
                        # Production mode - show success and attempt Tally upload
                        messagebox.showinfo("XML Generated Successfully", 
                                          f"{file_type_name} XML file has been generated and saved!\n\n"
                                          f"File: {output_path}\n\n"
                                          f"You can import this file directly into Tally or continue with automatic upload.")
                        
                        # Now attempt to upload to Tally
                        self.update_status(f"🚀 Attempting to upload {file_type_name} to Tally server...")
                        
                        try:
                            if upload_type == "attendance":
                                upload_result = self.api_service.upload_attendance_data(result, company_name)
                            elif upload_type == "payroll":
                                upload_result = self.api_service.upload_payroll_data(result, company_name, account_name)
                            else:  # paye
                                upload_result = self.api_service.upload_paye_data(result, company_name, account_name)
                        
                            if upload_result and upload_result.get("success"):
                                self.update_status(f"🎉 {file_type_name} successfully uploaded to Tally server!")
                                messagebox.showinfo("Upload Successful", 
                                                  f"{file_type_name} data has been successfully uploaded to Tally!\n\n"
                                                  f"✅ XML file saved: {output_path}\n"
                                                  f"✅ Data uploaded to Tally server\n\n"
                                                  f"The voucher should now be available in Tally.")
                            else:
                                error_msg = upload_result.get("error", "Unknown upload error") if upload_result else "Upload failed"
                                self.update_status(f"⚠️ {file_type_name} XML saved, but Tally upload failed: {error_msg}")
                                
                                # Show option to open file location
                                if messagebox.askyesno("Upload Failed - XML Saved", 
                                                     f"Tally upload failed: {error_msg}\n\n"
                                                     f"✅ XML file has been saved successfully: {output_path}\n\n"
                                                     f"You can manually import this file into Tally.\n"
                                                     f"Would you like to open the file location?"):
                                    import subprocess
                                    subprocess.run(["open", "-R", output_path])  # macOS
                                    
                        except Exception as upload_error:
                            self.update_status(f"⚠️ {file_type_name} XML saved, but Tally server error: {str(upload_error)}")
                            
                            # Show option to open file location
                            if messagebox.askyesno("Tally Server Error - XML Saved", 
                                                 f"Could not connect to Tally server: {str(upload_error)}\n\n"
                                                 f"✅ XML file has been saved successfully: {output_path}\n\n"
                                                 f"You can manually import this file into Tally when the server is available.\n"
                                                 f"Would you like to open the file location?"):
                                import subprocess
                                subprocess.run(["open", "-R", output_path])  # macOS
                else:
                    messagebox.showerror("Error", "Failed to save XML file.")
            else:
                messagebox.showerror("Error", "Failed to generate XML content.")
                
        except Exception as e:
            self.update_status(f"❌ Error generating XML: {str(e)}")
            messagebox.showerror("XML Generation Error", f"Failed to generate XML:\n{str(e)}")
    
    def get_account_name_for_payroll(self):
        """Get account name for payroll payment."""
        dialog = AccountSelectionDialog(self.root)
        self.root.wait_window(dialog.top)
        return dialog.result
    
    def get_account_name_for_paye(self):
        """Get account name for PAYE payment."""
        dialog = PayeAccountSelectionDialog(self.root)
        self.root.wait_window(dialog.top)
        return dialog.result
    
    def get_account_name_for_zssf(self):
        """Get account name for ZSSF processing."""
        dialog = PayeAccountSelectionDialog(self.root)  # Reuse the same dialog
        dialog.top.title("Select Bank Account for ZSSF")
        # Update dialog text for ZSSF
        for child in dialog.top.winfo_children():
            if hasattr(child, 'winfo_children'):
                for subchild in child.winfo_children():
                    if hasattr(subchild, 'configure') and hasattr(subchild, 'cget'):
                        try:
                            text = subchild.cget('text')
                            if 'PAYE' in text:
                                subchild.configure(text=text.replace('PAYE', 'ZSSF'))
                        except:
                            pass
        self.root.wait_window(dialog.top)
        return dialog.result
    
    def get_account_name_for_zhsf(self):
        """Get account name for ZHSF processing."""
        dialog = PayeAccountSelectionDialog(self.root)  # Reuse the same dialog
        dialog.top.title("Select Bank Account for ZHSF")
        # Update dialog text for ZHSF
        for child in dialog.top.winfo_children():
            if hasattr(child, 'winfo_children'):
                for subchild in child.winfo_children():
                    if hasattr(subchild, 'configure') and hasattr(subchild, 'cget'):
                        try:
                            text = subchild.cget('text')
                            if 'PAYE' in text:
                                subchild.configure(text=text.replace('PAYE', 'ZHSF'))
                        except:
                            pass
        self.root.wait_window(dialog.top)
        return dialog.result
    
    def download_example_file(self):
        """Download the example Excel file (STAFF SALARY 2025-12.xlsx) to Downloads folder."""
        try:
            import shutil
            import os
            
            # Source file path
            source_file = "/Users/goku/Documents/excel_processor/src/Payroll Voucher/STAFF SALARY 2025-12.xlsx"
            
            # Destination path (Downloads folder)
            downloads_folder = os.path.expanduser("~/Downloads")
            destination_file = os.path.join(downloads_folder, "STAFF SALARY 2025-12.xlsx")
            
            # Copy the file
            shutil.copy2(source_file, destination_file)
            
            # Show success message
            messagebox.showinfo(
                "Example Downloaded", 
                f"Example file downloaded successfully!\n\n"
                f"File: STAFF SALARY 2025-12.xlsx\n"
                f"Location: {destination_file}\n\n"
                f"This is the real format used for payroll processing."
            )
            
            # Update status
            self.update_status(f"📁 Example file downloaded: STAFF SALARY 2025-12.xlsx")
            
            # Open the file location
            import subprocess
            subprocess.run(["open", "-R", destination_file])  # macOS
            
        except Exception as e:
            messagebox.showerror("Download Error", f"Error downloading example file: {str(e)}")
    
    def update_status(self, message):
        """Update the status display."""
        self.status_text.configure(text=message)
        self.root.update_idletasks()
    
    def create_gui(self):
        """Create the main GUI interface."""
        # Main container with optimized spacing
        self.main_frame = tk.Frame(self.root, bg=self.colors['light'], padx=20, pady=15)
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame = self.main_frame  # Keep local reference for existing code
        
        # Header section with reduced spacing
        header_frame = tk.Frame(main_frame, bg=self.colors['light'])
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Company info section at top
        if self.selected_company:
            company_info_frame = tk.Frame(header_frame, bg=self.colors['success'], relief=tk.SOLID, bd=1)
            company_info_frame.pack(fill=tk.X, pady=(0, 15), ipady=10)
            
            # Left side - Company info
            company_left = tk.Frame(company_info_frame, bg=self.colors['success'])
            company_left.pack(side=tk.LEFT, padx=15, pady=5)
            
            company_label = tk.Label(
                company_left,
                text=f"🏢 Selected Company: {self.selected_company}",
                font=('Arial', 14, 'bold'),
                bg=self.colors['success'],
                fg=self.colors['white']
            )
            company_label.pack(anchor=tk.W)
            
            # Right side - Change company button
            company_right = tk.Frame(company_info_frame, bg=self.colors['success'])
            company_right.pack(side=tk.RIGHT, padx=15, pady=5)
            
            change_company_btn = ttk.Button(
                company_right,
                text="🔄 Change Company",
                command=self.show_company_selection_screen,
                style='Secondary.TButton'
            )
            change_company_btn.pack(anchor=tk.E)
        
        # Title section with Download Example button
        title_section = tk.Frame(header_frame, bg=self.colors['light'])
        title_section.pack(fill=tk.X, pady=(0, 5))
        
        # Left side - Download Example button
        download_frame = tk.Frame(title_section, bg=self.colors['light'])
        download_frame.pack(side=tk.LEFT)
        
        download_btn = ttk.Button(
            download_frame,
            text="📥 Download Example",
            command=self.download_example_file,
            style='Primary.TButton'
        )
        download_btn.pack(pady=5)
        
        # Center - Title
        title_center = tk.Frame(title_section, bg=self.colors['light'])
        title_center.pack(expand=True)
        
        title_label = ttk.Label(title_center, text="📊 Excel File Processor & XML Generator", style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(title_center, text="Upload Excel files and generate Tally-compatible XML for import", style='Subtitle.TLabel')
        subtitle_label.pack()
        
        # # Description line
        # desc_label = ttk.Label(header_frame, text="Supports header extraction, employee data processing, and JSON export", style='Info.TLabel')
        # desc_label.pack()
        
        # Upload sections container with reduced spacing
        upload_frame = tk.Frame(main_frame, bg=self.colors['light'])
        upload_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Attendance upload section
        self.create_upload_section(upload_frame, "Attendance Sheet", "attendance", 0)
        
        # Professional separator with text - reduced spacing
        separator_frame = tk.Frame(upload_frame, bg=self.colors['light'])
        separator_frame.pack(fill=tk.X, pady=15)
        
        sep_line1 = tk.Frame(separator_frame, height=1, bg=self.colors['dark'])
        sep_line1.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        sep_label = ttk.Label(separator_frame, text=" OR ", style='Info.TLabel')
        sep_label.pack(side=tk.LEFT, padx=10)
        
        sep_line2 = tk.Frame(separator_frame, height=1, bg=self.colors['dark'])
        sep_line2.pack(side=tk.RIGHT, fill=tk.X, expand=True)
        
        # Payroll upload section
        self.create_upload_section(upload_frame, "Payroll Sheet", "payroll", 1)
        
        # Status area with optimized spacing
        status_container = tk.Frame(main_frame, bg=self.colors['light'])
        status_container.pack(fill=tk.BOTH, expand=True, pady=(15, 0))
        
        status_label = ttk.Label(status_container, text="📊 Status:", style='Subtitle.TLabel')
        status_label.pack(anchor=tk.W, pady=(0, 8))
        
        # Status display area with optimized styling
        self.status_frame = tk.Frame(status_container, bg=self.colors['white'], relief=tk.SOLID, bd=1)
        self.status_frame.pack(fill=tk.BOTH, expand=True, padx=5, ipady=10)
        
        # Status icon and text container
        status_content = tk.Frame(self.status_frame, bg=self.colors['white'])
        status_content.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        self.status_text = tk.Label(
            status_content,
            text="✨ Ready to process Excel files...",
            font=('Arial', 12),
            bg=self.colors['white'],
            fg=self.colors['dark'],
            wraplength=700,
            justify=tk.LEFT
        )
        self.status_text.pack(anchor=tk.W)
    
    def create_upload_section(self, parent, title, section_type, row):
        """Create an upload section for either Attendance or Payroll."""
        # Main section container with optimized spacing
        section_container = tk.Frame(parent, bg=self.colors['light'])
        section_container.pack(fill=tk.X, pady=10, padx=15)
        
        section_frame = tk.Frame(section_container, bg=self.colors['white'], relief=tk.SOLID, bd=1)
        section_frame.pack(fill=tk.X, ipady=15, ipadx=20)
        
        # Section title with icon - reduced spacing
        title_frame = tk.Frame(section_frame, bg=self.colors['white'])
        title_frame.pack(fill=tk.X, pady=(0, 10))
        
        icon = "📊" if section_type == "attendance" else "💰"
        title_text = f"{icon} {title}"
        title_label = ttk.Label(title_frame, text=title_text, style='Subtitle.TLabel')
        title_label.configure(background=self.colors['white'])
        title_label.pack(anchor=tk.W)
        
        # File path display with optimized spacing
        file_frame = tk.Frame(section_frame, bg=self.colors['white'])
        file_frame.pack(fill=tk.X, pady=(0, 12))
        
        path_label = ttk.Label(file_frame, text="📁 Selected file:", style='Info.TLabel')
        path_label.configure(background=self.colors['white'])
        path_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Store file path variable
        file_path_var = tk.StringVar(value="No file selected")
        setattr(self, f"{section_type}_file_path", file_path_var)
        
        # File path display with background
        path_container = tk.Frame(file_frame, bg=self.colors['light'], relief=tk.SUNKEN, bd=1)
        path_container.pack(fill=tk.X, pady=(0, 5))
        
        path_display = ttk.Label(
            path_container,
            textvariable=file_path_var,
            style='Info.TLabel',
            width=70
        )
        path_display.configure(background=self.colors['light'])
        path_display.pack(anchor=tk.W, padx=10, pady=8)
        
        # Buttons frame
        btn_frame = tk.Frame(section_frame, bg=self.colors['white'])
        btn_frame.pack(fill=tk.X)
        
        # Browse button with icon
        browse_btn = ttk.Button(
            btn_frame,
            text="📂 Browse File",
            command=lambda: self.browse_file(section_type),
            style='Primary.TButton'
        )
        browse_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        # Process button with icon
        process_icon = "⚙️" if section_type == "attendance" else "💵"
        process_btn = ttk.Button(
            btn_frame,
            text=f"{process_icon} Process {title}",
            command=lambda: self.process_file(section_type),
            style='Success.TButton'
        )
        process_btn.pack(side=tk.LEFT, padx=(0, 20))
        
        # Visual separator
        separator = tk.Frame(btn_frame, width=2, height=30, bg=self.colors['dark'])
        separator.pack(side=tk.LEFT, padx=(0, 20))
        
        # Download Sample button with icon
        sample_btn = ttk.Button(
            btn_frame,
            text="📄 Download Sample",
            command=lambda: self.download_sample(section_type),
            style='Secondary.TButton'
        )
        sample_btn.pack(side=tk.LEFT)
    
    def browse_file(self, section_type):
        """Open file dialog to select Excel file."""
        file_types = [
            ("Excel files", "*.xlsx *.xls"),
            ("All files", "*.*")
        ]
        
        filename = filedialog.askopenfilename(
            title=f"Select {section_type.title()} Excel File",
            filetypes=file_types,
            initialdir=os.path.expanduser("~")
        )
        
        if filename:
            file_path_var = getattr(self, f"{section_type}_file_path")
            file_path_var.set(filename)
            self.add_result(f"✓ {section_type.title()} file selected: {os.path.basename(filename)}")
    
    def process_file(self, section_type):
        """Process the selected Excel file."""
        file_path_var = getattr(self, f"{section_type}_file_path")
        file_path = file_path_var.get()
        
        if file_path == "No file selected" or not file_path:
            messagebox.showwarning("No File Selected", f"Please select a {section_type} Excel file first.")
            return
        
        if not os.path.exists(file_path):
            messagebox.showerror("File Not Found", "The selected file no longer exists.")
            return
        
        # Get the process button and show loading
        process_btn = None
        for widget in self.root.winfo_children():
            for child in self._get_all_widgets(widget):
                if isinstance(child, ttk.Button) and f"Process {section_type.title()}" in child.cget('text'):
                    process_btn = child
                    break
        
        if process_btn:
            process_btn.configure(text="Processing...", state='disabled')
            self.root.update()
        
        # Update status
        self.status_text.configure(text=f"Processing {section_type} file: {os.path.basename(file_path)}...")
        self.root.update()
        
        # Process in background thread
        thread = threading.Thread(target=self._process_file_background, args=(section_type, file_path, process_btn))
        thread.daemon = True
        thread.start()
    
    def _get_all_widgets(self, widget):
        """Recursively get all widgets from a parent widget."""
        widgets = [widget]
        for child in widget.winfo_children():
            widgets.extend(self._get_all_widgets(child))
        return widgets
    
    def _process_file_background(self, section_type, file_path, process_btn):
        """Process file in background thread."""
        try:
            if section_type == "attendance":
                result = self.processor.process_attendance_sheet(file_path)
            else:
                result = self.processor.process_payroll_sheet(file_path)
            
            # Update UI in main thread
            self.root.after(0, self._handle_processing_result, section_type, result, process_btn)
            
        except Exception as e:
            error_result = {"error": str(e), "success": False}
            self.root.after(0, self._handle_processing_result, section_type, error_result, process_btn)
    
    def _handle_processing_result(self, section_type, result, process_btn):
        """Handle the processing result - show data screen or error popup."""
        print(f"🎯 DEBUG: Handling processing result for {section_type}")
        print(f"🎯 DEBUG: Result success: {result.get('success', False)}")
        print(f"🎯 DEBUG: Result keys: {list(result.keys()) if result else 'None'}")
        
        # Restore button state
        if process_btn:
            button_text = f"Process {section_type.title()} Sheet"
            process_btn.configure(text=button_text, state='normal')
        
        if result.get("success", False):
            # Valid Excel - show data screen
            print(f"✅ DEBUG: Valid Excel, showing data screen")
            self.status_text.configure(text=f"{section_type.title()} processing completed successfully!")
            self.show_data_screen(section_type, result)
        else:
            # Invalid Excel - show error popup
            print(f"❌ DEBUG: Invalid Excel, showing error popup")
            self.status_text.configure(text="Ready to process Excel files...")
            error_msg = result.get('error', 'Unknown error occurred')
            
            # Create detailed error message
            if 'validation_errors' in result:
                validation_errors = result['validation_errors']
                detailed_msg = "Invalid Excel format detected!\n\n"
                detailed_msg += "Errors found:\n"
                for error in validation_errors:
                    detailed_msg += f"• {error}\n"
                detailed_msg += "\n" + self._get_format_help_text()
            else:
                detailed_msg = f"Error processing file:\n{error_msg}\n\nPlease check your Excel file and try again."
            
            messagebox.showerror("Invalid Excel File", detailed_msg)
    
    def _upload_attendance_to_tally(self, attendance_result):
        """Upload attendance data to Tally."""
        # Update status
        self.status_text.configure(text="🔄 Uploading attendance data to Tally...")
        self.root.update()
        
        # Upload in background thread
        thread = threading.Thread(target=self._upload_attendance_background, args=(attendance_result,))
        thread.daemon = True
        thread.start()
    
    def _upload_attendance_background(self, attendance_result):
        """Upload attendance data in background thread."""
        try:
            # Use selected company name
            company_name = self.selected_company
            
            # Call upload API
            upload_result = self.api_service.upload_attendance_data(attendance_result, company_name)
            
            # Update UI in main thread
            self.root.after(0, self._handle_upload_result, upload_result)
            
        except Exception as e:
            error_result = {"success": False, "error": str(e)}
            self.root.after(0, self._handle_upload_result, error_result)
    
    def _handle_upload_result(self, upload_result):
        """Handle the upload result."""
        # Log the complete response for debugging
        print("="*60)
        print("🔍 TALLY API RESPONSE:")
        print("="*60)
        print(f"Success: {upload_result.get('success', False)}")
        print(f"Message: {upload_result.get('message', 'No message')}")
        print(f"Error: {upload_result.get('error', 'No error')}")
        
        # Log the actual response from Tally
        response_text = upload_result.get("response", "No response")
        if response_text:
            print("📄 Raw Tally Response:")
            print("-" * 40)
            print(response_text)
            print("-" * 40)
        print("="*60)
        
        if upload_result.get("success", False):
            # Success - no errors in Tally response
            created_count = upload_result.get("created_count", 0)
            self.status_text.configure(text="✅ Attendance data successfully uploaded to Tally!")
            
            messagebox.showinfo(
                "Upload Successful", 
                f"✅ Success!\n\nAttendance data has been successfully uploaded to {self.selected_company} in Tally!\n\n📊 Records Created: {created_count}\n\nReturning to main screen..."
            )
            
            # Clear file selections and return to main Excel processor screen
            self.attendance_file_path.set("No file selected")
            if hasattr(self, 'payroll_file_path'):
                self.payroll_file_path.set("No file selected")
            
            # Return to main Excel processor screen
            self.show_main_screen()
            
        else:
            # Error - Tally reported errors
            error_msg = upload_result.get("error", "Unknown error")
            errors_count = upload_result.get("errors_count", 0)
            error_message = upload_result.get("error_message", "")
            
            self.status_text.configure(text="❌ Failed to upload attendance data to Tally")
            
            # Create detailed error message
            if error_message:
                full_error_msg = f"❌ Upload Failed!\n\nTally Error: {error_message}\n\nThis usually means there's a data mismatch. Please check your Excel file and ensure:\n• Employee names exist in Tally\n• Attendance types are valid\n• Data format is correct"
            else:
                full_error_msg = f"❌ Upload Failed!\n\n{error_msg}\n\nErrors Found: {errors_count}\n\nPlease check your Excel data and try again."
            
            messagebox.showerror("Upload Failed", full_error_msg)
    
    def _upload_payroll_to_tally(self, payroll_result):
        """Upload payroll data to Tally."""
        # First, ask user for account information
        account_dialog = AccountSelectionDialog(self.root)
        self.root.wait_window(account_dialog.top)
        
        if account_dialog.result:
            account_name = account_dialog.result
            
            # Update status
            self.status_text.configure(text="🔄 Uploading payroll data to Tally...")
            self.root.update()
            
            # Upload in background thread
            thread = threading.Thread(target=self._upload_payroll_background, args=(payroll_result, account_name))
            thread.daemon = True
            thread.start()
        else:
            # User cancelled account selection
            self.status_text.configure(text="Upload cancelled")
    
    def _upload_payroll_background(self, payroll_result, account_name):
        """Upload payroll data in background thread."""
        try:
            # Use selected company name
            company_name = self.selected_company
            
            # Call upload API
            upload_result = self.api_service.upload_payroll_data(payroll_result, company_name, account_name)
            
            # Update UI in main thread
            self.root.after(0, self._handle_payroll_upload_result, upload_result)
            
        except Exception as e:
            error_result = {"success": False, "error": str(e)}
            self.root.after(0, self._handle_payroll_upload_result, error_result)
    
    def _handle_payroll_upload_result(self, upload_result):
        """Handle the payroll upload result."""
        # Log the complete response for debugging
        print("="*60)
        print("🔍 TALLY PAYROLL API RESPONSE:")
        print("="*60)
        print(f"Success: {upload_result.get('success', False)}")
        print(f"Message: {upload_result.get('message', 'No message')}")
        print(f"Error: {upload_result.get('error', 'No error')}")
        
        # Log the actual response from Tally
        response_text = upload_result.get("response", "No response")
        if response_text:
            print("📄 Raw Tally Response:")
            print("-" * 40)
            print(response_text)
            print("-" * 40)
        print("="*60)
        
        if upload_result.get("success", False):
            # Success - no errors in Tally response
            created_count = upload_result.get("created_count", 0)
            self.status_text.configure(text="✅ Payroll data successfully uploaded to Tally!")
            
            messagebox.showinfo(
                "Upload Successful", 
                f"✅ Success!\n\nPayroll data has been successfully uploaded to {self.selected_company} in Tally!\n\n📊 Records Created: {created_count}\n\nReturning to main screen..."
            )
            
            # Clear file selections and return to main Excel processor screen
            self.attendance_file_path.set("No file selected")
            if hasattr(self, 'payroll_file_path'):
                self.payroll_file_path.set("No file selected")
            
            # Return to main Excel processor screen
            self.show_main_screen()
            
        else:
            # Error - Tally reported errors
            error_msg = upload_result.get("error", "Unknown error")
            errors_count = upload_result.get("errors_count", 0)
            error_message = upload_result.get("error_message", "")
            
            self.status_text.configure(text="❌ Failed to upload payroll data to Tally")
            
            # Create detailed error message
            if error_message:
                full_error_msg = f"❌ Payroll Upload Failed!\n\nTally Error: {error_message}\n\nThis usually means there's a data mismatch. Please check your Excel file and ensure:\n• Employee names exist in Tally\n• Payhead names are valid\n• Data format is correct\n• Account name exists in Tally"
            else:
                full_error_msg = f"❌ Payroll Upload Failed!\n\n{error_msg}\n\nErrors Found: {errors_count}\n\nPlease check your Excel data and try again."
            
            messagebox.showerror("Upload Failed", full_error_msg)
    
    def add_result(self, text):
        """Update status text (replacing old results functionality)."""
        self.status_text.configure(text=text)
    
    def show_loading(self, section_type, file_name):
        """Show loading indicator in results."""
        self.results_text.insert(tk.END, f"🔄 Processing {section_type} file: {file_name}\n")
        self.results_text.insert(tk.END, "⏳ Loading... Please wait while we analyze your Excel file.\n")
        self.results_text.insert(tk.END, "   • Reading file structure...\n")
        self.results_text.insert(tk.END, "   • Analyzing data types...\n")
        self.results_text.insert(tk.END, "   • Processing columns and rows...\n\n")
        self.results_text.see(tk.END)
        self.root.update_idletasks()
        self.root.update()
    
    def clear_loading(self):
        """Clear loading messages."""
        # This will be handled by the results display
        pass
    
    def display_results(self, section_type, result):
        """Display the processing results directly."""
        print(f"🎯 display_results called with section_type: {section_type}")
        print(f"📊 Result success: {result.get('success', 'No success key')}")
        
        # Clear and start fresh
        self.results_text.delete(1.0, tk.END)
        self.results_text.insert(tk.END, f"✅ PROCESSING COMPLETED FOR {section_type.upper()}!\n\n")
        
        if result.get("success", False):
            # File information
            self.results_text.insert(tk.END, f"📊 FILE INFORMATION:\n")
            self.results_text.insert(tk.END, f"   📁 File: {result['file_name']}\n")
            self.results_text.insert(tk.END, f"   📋 Type: {result['sheet_type']}\n")
            self.results_text.insert(tk.END, f"   📏 Total Rows: {result['total_rows']}\n")
            self.results_text.insert(tk.END, f"   📐 Total Columns: {result['total_columns']}\n\n")
            
            if section_type == "attendance":
                # Header Information (Date, Company, Narration)
                self.results_text.insert(tk.END, f"📅 ATTENDANCE SHEET HEADER:\n")
                self.results_text.insert(tk.END, f"   📅 Date: {result.get('date', 'Not found')}\n")
                self.results_text.insert(tk.END, f"   🏢 Company Name: {result.get('company_name', 'Not found')}\n")
                self.results_text.insert(tk.END, f"   📝 Narration: {result.get('narration', 'Not found')}\n\n")
                
                # Employee Data
                employee_data = result.get('employee_data', [])
                self.results_text.insert(tk.END, f"👥 EMPLOYEE DATA ({result.get('total_employees', 0)} employees):\n")
                
                if employee_data:
                    # Show first 10 employees
                    for i, emp in enumerate(employee_data[:10], 1):
                        self.results_text.insert(tk.END, f"   {i}. ")
                        self.results_text.insert(tk.END, f"ID: {emp.get('employee_no', 'N/A')}, ")
                        self.results_text.insert(tk.END, f"Name: {emp.get('employee_name', 'N/A')}, ")
                        self.results_text.insert(tk.END, f"Type: {emp.get('attendance_type', 'N/A')}, ")
                        self.results_text.insert(tk.END, f"Days: {emp.get('attendance_days', 'N/A')}\n")
                    
                    if len(employee_data) > 10:
                        self.results_text.insert(tk.END, f"   ... and {len(employee_data) - 10} more employees\n")
                else:
                    self.results_text.insert(tk.END, "   No employee data found\n")
                
                self.results_text.insert(tk.END, f"\n   📍 Employee data starts at row: {result.get('employee_data_start_row', 'Unknown')}\n")
                
                # Validation warnings
                warnings = result.get('validation_warnings', [])
                if warnings:
                    self.results_text.insert(tk.END, f"\n⚠️ WARNINGS:\n")
                    for warning in warnings:
                        self.results_text.insert(tk.END, f"   • {warning}\n")
            
            else:  # payroll
                # Basic file info for payroll
                columns = result.get('columns', [])
                self.results_text.insert(tk.END, f"📋 COLUMNS ({len(columns)}):\n")
                for i, col in enumerate(columns[:10], 1):
                    self.results_text.insert(tk.END, f"   {i}. {col}\n")
                if len(columns) > 10:
                    self.results_text.insert(tk.END, f"   ... and {len(columns) - 10} more columns\n")
                
                # Sample data
                sample_data = result.get('sample_data', [])
                if sample_data:
                    self.results_text.insert(tk.END, f"\n📄 SAMPLE DATA (First Row):\n")
                    first_row = sample_data[0]
                    for key, value in list(first_row.items())[:5]:
                        self.results_text.insert(tk.END, f"   {key}: {value}\n")
        
        else:
            # Error handling
            self.results_text.insert(tk.END, f"❌ ERROR PROCESSING FILE:\n")
            self.results_text.insert(tk.END, f"   {result.get('error', 'Unknown error')}\n\n")
            
            # Show validation errors if available
            validation_errors = result.get('validation_errors', [])
            if validation_errors:
                self.results_text.insert(tk.END, f"� VALIDATION ERRORS:\n")
                for error in validation_errors:
                    self.results_text.insert(tk.END, f"   • {error}\n")
                
                self.results_text.insert(tk.END, f"\n💡 EXPECTED FORMAT FOR ATTENDANCE SHEET:\n")
                self.results_text.insert(tk.END, f"   • Row 1: Date in column B (e.g., 09-10-2025)\n")
                self.results_text.insert(tk.END, f"   • Row 2: Company Name in column B (e.g., LIGHT)\n")
                self.results_text.insert(tk.END, f"   • Row 3: Narration in column B (e.g., Test attendance)\n")
                self.results_text.insert(tk.END, f"   • Row 5+: Employee data with headers:\n")
    

    

    
    def show_main_screen(self):
        """Show the main screen (hide data screen)."""
        # Remove data frame if it exists
        if hasattr(self, 'data_frame'):
            self.data_frame.destroy()
            delattr(self, 'data_frame')
        
        # Clear all widgets from root
        for widget in self.root.winfo_children():
            widget.destroy()
        
        # Recreate main interface
        self.create_gui()
    
    def export_data_from_screen(self, section_type, result):
        """Export data from the data screen."""
        if not hasattr(self, 'current_result') or not self.current_result:
            messagebox.showerror("Export Error", "No data available to export.")
            return
            
        result = self.current_result
        section_type = self.current_section_type
        
        file_types = [("JSON files", "*.json"), ("All files", "*.*")]
        default_name = f"{os.path.splitext(result['file_name'])[0]}_processed_{section_type}.json"
        
        filename = filedialog.asksaveasfilename(
            title=f"Save {section_type.title()} Data",
            defaultextension=".json",
            filetypes=file_types,
            initialfile=default_name
        )
        
        if filename:
            try:
                if section_type == "attendance":
                    export_path = self.processor.export_attendance_data(result, filename)
                elif section_type == "payroll":
                    export_path = self.processor.export_payroll_data(result, filename)
                else:
                    import json
                    with open(filename, 'w', encoding='utf-8') as f:
                        json.dump(result, f, indent=2, ensure_ascii=False)
                    export_path = filename
                
                if export_path:
                    messagebox.showinfo("Export Successful", 
                                      f"Data exported successfully to:\n{export_path}")
                else:
                    messagebox.showerror("Export Failed", "Failed to export data.")
                    
            except Exception as e:
                messagebox.showerror("Export Error", f"Error exporting data: {str(e)}")
    
    def _get_format_help_text(self):
        """Get help text for expected format."""
        return """Expected format for Excel Sheets:

For Attendance Sheet:
• Row 1: Date in column B (e.g., 09-10-2025)
• Row 2: Company Name in column B (e.g., LIGHT)
• Row 3: Narration in column B (e.g., Test attendance)
• Row 5+: Employee data with headers:
  - EMPL NO, EMPLOYEE NAME, Attendance/Production Types, Attendance Days

For Payroll Sheet:
• Row 1: Date in column B (e.g., 09-10-2025)
• Row 2: Company Name in column B (e.g., LIGHT)
• Row 3: Account in column B (e.g., PEAPULE BANK OF UNITE)
• Row 6+: Employee data with headers:
  - EMPL NO, EMPLOYEE NAME, BASIC, HRA, MEDICAL, etc."""
    
    def show_data_screen(self, section_type, result):
        """Show data screen in the same window."""
        # Hide main content
        for widget in self.root.winfo_children():
            widget.pack_forget()
        
        # Create data screen frame with better layout
        self.data_frame = tk.Frame(self.root, bg=self.colors['light'])
        self.data_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Title
        title_text = f"{section_type.title()} Data Extracted Successfully"
        title_label = ttk.Label(self.data_frame, text=title_text, style='Title.TLabel')
        title_label.pack(pady=(0, 15))
        
        # Create main content frame (no scrollbar for now to fix layout)
        content_frame = tk.Frame(self.data_frame, bg=self.colors['light'])
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Add content directly to content frame
        self._populate_data_screen(content_frame, section_type, result)
        
        # Bottom buttons with better layout
        btn_frame = tk.Frame(self.data_frame, bg=self.colors['light'])
        btn_frame.pack(fill=tk.X, side=tk.BOTTOM, pady=(15, 5))
        
        # Left side buttons
        left_btn_frame = tk.Frame(btn_frame, bg=self.colors['light'])
        left_btn_frame.pack(side=tk.LEFT)
        
        # Back button
        back_btn = ttk.Button(
            left_btn_frame,
            text="← Back",
            command=self.show_main_screen,
            style='Secondary.TButton'
        )
        back_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Export/Upload button - conditional based on section type
        if section_type == "attendance":
            action_btn = ttk.Button(
                left_btn_frame,
                text="📤 Upload to Tally",
                command=lambda: self._upload_attendance_to_tally(result),
                style='Success.TButton'
            )
        elif section_type == "payroll":
            action_btn = ttk.Button(
                left_btn_frame,
                text="📤 Upload to Tally",
                command=lambda: self._upload_payroll_to_tally(result),
                style='Success.TButton'
            )
        else:
            action_btn = ttk.Button(
                left_btn_frame,
                text="Export to JSON",
                command=lambda: self.export_data_from_screen(section_type, result),
                style='Success.TButton'
            )
        action_btn.pack(side=tk.LEFT)
        
        # Right side buttons
        right_btn_frame = tk.Frame(btn_frame, bg=self.colors['light'])
        right_btn_frame.pack(side=tk.RIGHT)
        
        # Cancel button
        cancel_btn = ttk.Button(
            right_btn_frame,
            text="Cancel",
            command=self.show_main_screen,
            style='Secondary.TButton'
        )
        cancel_btn.pack(side=tk.RIGHT)
        
        # Store current result for export
        self.current_result = result
        self.current_section_type = section_type
    
    def _populate_data_screen(self, parent_frame, section_type, result):
        """Populate the data screen with extracted information."""
        # Configure parent frame to be responsive
        parent_frame.columnconfigure(0, weight=1)
        
        # Debug: Print what we're displaying
        print(f"🎯 DEBUG: Populating {section_type} screen")
        print(f"🎯 DEBUG: Result keys: {list(result.keys()) if result else 'None'}")
        print(f"🎯 DEBUG: Success: {result.get('success', False)}")
        
        if section_type == "attendance":
            # Combined File & Header Information Section (Side by Side)
            info_section = tk.LabelFrame(parent_frame, text="📊 File & Header Information", 
                                       font=('Arial', 11, 'bold'), bg=self.colors['light'],
                                       fg=self.colors['primary'])
            info_section.pack(fill=tk.X, pady=(0, 10), padx=5)
            
            # Create main container for left and right columns
            info_container = tk.Frame(info_section, bg=self.colors['light'])
            info_container.pack(fill=tk.X, padx=10, pady=8)
            
            # Left column - File Information
            left_frame = tk.Frame(info_container, bg=self.colors['light'])
            left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
            
            tk.Label(left_frame, text="📊 File Information", 
                    font=('Arial', 10, 'bold'), bg=self.colors['light'], fg=self.colors['primary']).pack(anchor=tk.W, pady=(0, 5))
            tk.Label(left_frame, text=f"� File: {result.get('file_name', 'N/A')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            tk.Label(left_frame, text=f"� Total Rows: {result.get('total_rows', 0)}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            tk.Label(left_frame, text=f"� Total Columns: {result.get('total_columns', 0)}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            
            # Right column - Header Information
            right_frame = tk.Frame(info_container, bg=self.colors['light'])
            right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
            
            # Debug print header values
            print(f"🔍 DEBUG: Date: {result.get('date', 'Not found')}")
            print(f"🔍 DEBUG: Company: {result.get('company_name', 'Not found')}")
            print(f"🔍 DEBUG: Narration: {result.get('narration', 'Not found')}")
            
            tk.Label(right_frame, text="📋 Header Information", 
                    font=('Arial', 10, 'bold'), bg=self.colors['light'], fg=self.colors['primary']).pack(anchor=tk.W, pady=(0, 5))
            tk.Label(right_frame, text=f"📅 Date: {result.get('date', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            tk.Label(right_frame, text=f"🏢 Company: {result.get('company_name', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            tk.Label(right_frame, text=f"📝 Narration: {result.get('narration', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            
            # Employee Data Section
            employee_data = result.get('employee_data', [])
            print(f"🔍 DEBUG: Employee data count: {len(employee_data)}")
            if employee_data:
                print(f"🔍 DEBUG: First employee: {employee_data[0]}")
            
            emp_section = tk.LabelFrame(parent_frame, text=f"👥 Employee Data ({len(employee_data)} employees)", 
                                      font=('Arial', 11, 'bold'), bg=self.colors['light'],
                                      fg=self.colors['primary'])
            emp_section.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
            
            # Employee data table
            if employee_data:
                # Create frame for table with better sizing
                table_frame = tk.Frame(emp_section, bg=self.colors['white'])
                table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                
                # Create treeview for employee data with better sizing
                columns = ('ID', 'Name', 'Attendance Type', 'Days')
                tree = ttk.Treeview(table_frame, columns=columns, show='headings', height=10)
                
                # Define headings
                tree.heading('ID', text='Employee ID')
                tree.heading('Name', text='Employee Name')
                tree.heading('Attendance Type', text='Attendance Type')
                tree.heading('Days', text='Days')
                
                # Configure column widths to ensure all columns are visible
                total_width = 650  # Approximate table width
                tree.column('ID', width=80, anchor='center', minwidth=60)
                tree.column('Name', width=160, anchor='w', minwidth=120)
                tree.column('Attendance Type', width=280, anchor='center', minwidth=200)
                tree.column('Days', width=80, anchor='center', minwidth=60)
                
                # Set minimum height for table visibility
                tree.configure(height=8)
                
                # Add data to tree
                for emp in employee_data:
                    tree.insert('', tk.END, values=(
                        emp.get('employee_no', 'N/A'),
                        emp.get('employee_name', 'N/A'),
                        emp.get('attendance_type', 'N/A'),
                        emp.get('attendance_days', 'N/A')
                    ))
                
                # Add scrollbar to tree
                tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
                tree.configure(yscrollcommand=tree_scroll.set)
                
                # Pack tree and scrollbar properly
                tree.pack(side="left", fill="both", expand=True)
                tree_scroll.pack(side="right", fill="y")
                
                # Configure table to use full width
                table_frame.columnconfigure(0, weight=1)
                table_frame.rowconfigure(0, weight=1)
            else:
                no_data_label = tk.Label(emp_section, text="❌ No employee data found", 
                        font=('Arial', 12, 'bold'), bg=self.colors['light'], 
                        fg=self.colors['danger'])
                no_data_label.pack(pady=20)
        
        else:  # payroll
            print(f"🎯 DEBUG: Processing payroll display")
            
            # Combined File & Header Information Section (Side by Side)
            info_section = tk.LabelFrame(parent_frame, text="📊 File & Header Information", 
                                       font=('Arial', 11, 'bold'), bg=self.colors['light'],
                                       fg=self.colors['primary'])
            info_section.pack(fill=tk.X, pady=(0, 10), padx=5)
            
            # Create main container for left and right columns
            info_container = tk.Frame(info_section, bg=self.colors['light'])
            info_container.pack(fill=tk.X, padx=10, pady=8)
            
            # Left column - File Information
            left_frame = tk.Frame(info_container, bg=self.colors['light'])
            left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 10))
            
            tk.Label(left_frame, text="📊 File Information", 
                    font=('Arial', 10, 'bold'), bg=self.colors['light'], fg=self.colors['primary']).pack(anchor=tk.W, pady=(0, 5))
            tk.Label(left_frame, text=f"� File: {result.get('file_name', 'N/A')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            tk.Label(left_frame, text=f"� Total Rows: {result.get('total_rows', 0)}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            tk.Label(left_frame, text=f"� Total Columns: {result.get('total_columns', 0)}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['dark']).pack(anchor=tk.W, pady=1)
            
            # Right column - Header Information
            right_frame = tk.Frame(info_container, bg=self.colors['light'])
            right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(10, 0))
            
            print(f"🎯 DEBUG: Payroll Header - Date: {result.get('date')}, Company: {result.get('company_name')}")
            
            tk.Label(right_frame, text="📋 Header Information", 
                    font=('Arial', 10, 'bold'), bg=self.colors['light'], fg=self.colors['primary']).pack(anchor=tk.W, pady=(0, 5))
            tk.Label(right_frame, text=f"📅 Date: {result.get('date', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            tk.Label(right_frame, text=f"🏢 Company: {result.get('company_name', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            tk.Label(right_frame, text=f"🏦 Account: {result.get('account', 'Not found')}", 
                    font=('Arial', 10), bg=self.colors['light'], fg=self.colors['success']).pack(anchor=tk.W, pady=1)
            
            # Payroll Summary Section
            summary_section = tk.LabelFrame(parent_frame, text="💰 Payroll Summary", 
                                          font=('Arial', 11, 'bold'), bg=self.colors['light'],
                                          fg=self.colors['primary'])
            summary_section.pack(fill=tk.X, pady=(0, 10), padx=5)
            
            summary_info = tk.Frame(summary_section, bg=self.colors['light'])
            summary_info.pack(fill=tk.X, padx=10, pady=8)
            
            print(f"🎯 DEBUG: Payroll Summary - Employees: {result.get('total_employees')}, Total: {result.get('total_gross_salary')}")
            
            tk.Label(summary_info, text=f"� Total Employees: {result.get('total_employees', 0)}", 
                    font=('Arial', 11, 'bold'), bg=self.colors['light'], fg=self.colors['warning']).pack(anchor=tk.W, pady=2)
            tk.Label(summary_info, text=f"💵 Total Gross Salary: {result.get('total_gross_salary', 0):,.2f}", 
                    font=('Arial', 11, 'bold'), bg=self.colors['light'], fg=self.colors['warning']).pack(anchor=tk.W, pady=2)
            tk.Label(summary_info, text=f"📈 Average Salary: {result.get('average_salary', 0):,.2f}", 
                    font=('Arial', 11, 'bold'), bg=self.colors['light'], fg=self.colors['warning']).pack(anchor=tk.W, pady=2)
            
            # Employee Payroll Data Section
            employee_data = result.get('employee_data', [])
            print(f"🎯 DEBUG: Employee data count: {len(employee_data)}")
            
            emp_section = tk.LabelFrame(parent_frame, text=f"👥 Employee Payroll Data ({len(employee_data)} employees)", 
                                      font=('Arial', 11, 'bold'), bg=self.colors['light'],
                                      fg=self.colors['primary'])
            emp_section.pack(fill=tk.BOTH, expand=True, pady=(0, 10), padx=5)
            
            # Employee payroll data table
            if employee_data:
                print(f"🎯 DEBUG: Creating payroll table with {len(employee_data)} employees")
                
                # Create frame for table with better sizing
                table_frame = tk.Frame(emp_section, bg=self.colors['white'])
                table_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
                
                # Get COMPLETELY dynamic column headers from first employee
                first_emp = employee_data[0] if employee_data else {}
                column_headers = first_emp.get('column_headers', [])
                
                if not column_headers:
                    # Fallback: derive from salary components
                    salary_components_keys = list(first_emp.get('salary_components', {}).keys())
                    column_headers = ['Employee ID', 'Employee Name'] + salary_components_keys
                
                print(f"🎯 UI DEBUG: Using dynamic columns: {column_headers}")
                
                # Create completely dynamic columns + Total Gross
                dynamic_columns = column_headers + ['Total Gross']
                tree = ttk.Treeview(table_frame, columns=dynamic_columns, show='headings', height=8)
                
                # Define headings dynamically using ACTUAL Excel column names
                for col_name in column_headers:
                    tree.heading(col_name, text=col_name)  # Use exact column name from Excel
                    
                    # Determine column width and alignment based on position
                    if 'ID' in col_name.upper() or 'NO' in col_name.upper():
                        tree.column(col_name, width=80, anchor='center', minwidth=60)
                    elif 'NAME' in col_name.upper():
                        tree.column(col_name, width=160, anchor='w', minwidth=120)
                    else:
                        tree.column(col_name, width=100, anchor='e', minwidth=80)
                
                # Total Gross column
                tree.heading('Total Gross', text='Total Gross')
                tree.column('Total Gross', width=120, anchor='e', minwidth=100)
                
                # Add data to tree
                for emp in employee_data:
                    all_data = emp.get('all_data', {})
                    
                    # Build values list using actual column headers
                    values = []
                    for col_name in column_headers:
                        cell_value = all_data.get(col_name, '')
                        
                        # Format based on data type
                        if isinstance(cell_value, (int, float)):
                            values.append(f"{cell_value:,.0f}")
                        else:
                            values.append(str(cell_value))
                    
                    # Add total gross
                    values.append(f"{emp.get('total_gross_salary', 0):,.0f}")
                    
                    tree.insert('', tk.END, values=tuple(values))
                
                # Add scrollbar to tree
                tree_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=tree.yview)
                tree.configure(yscrollcommand=tree_scroll.set)
                
                # Pack tree and scrollbar properly
                tree.pack(side="left", fill="both", expand=True)
                tree_scroll.pack(side="right", fill="y")
                
                # Configure table to use full width
                table_frame.columnconfigure(0, weight=1)
                table_frame.rowconfigure(0, weight=1)
            else:
                no_data_label = tk.Label(emp_section, text="❌ No employee payroll data found", 
                        font=('Arial', 12, 'bold'), bg=self.colors['light'], 
                        fg=self.colors['danger'])
                no_data_label.pack(pady=20)
    
    def run(self):
        """Start the application."""
        self.root.mainloop()


if __name__ == "__main__":
    app = ModernExcelProcessor()
    app.run()
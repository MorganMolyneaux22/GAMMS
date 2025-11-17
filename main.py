import tkinter as tk
from tkinter import ttk, scrolledtext, messagebox, filedialog
import subprocess
import threading
import os
from datetime import datetime
import re
import csv
import io

class GAMSimplifiedGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("GAM Made Simple")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f0f0')
        
        # Initialize data storage
        self.organizational_units = []
        self.formatted_ous = []  # Store formatted display versions of OUs
        self.is_loading_ous = True
        self.parsed_events = []  # Store parsed calendar events
        self.captured_groups = []  # Store captured groups for reporting
        self.group_report_data = []  # Store generated report data
        self.selected_ous = []  # Store multiple selected OUs
        self.user_report_data = []  # Store user report data
        
        # Create main style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Configure colors
        self.style.configure('Title.TLabel', font=('Arial', 16, 'bold'), background='#f0f0f0')
        self.style.configure('Header.TLabel', font=('Arial', 12, 'bold'), background='#f0f0f0')
        
        self.setup_gui()
        
        # Load organizational units on startup
        self.load_organizational_units()
        
    def setup_gui(self):
        # Main title
        title_label = ttk.Label(self.root, text="GAM Made Simple", style='Title.TLabel')
        title_label.pack(pady=10)
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.root)
        self.notebook.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Create tabs
        self.create_user_management_tab()
        self.create_drive_management_tab()
        self.create_group_management_tab()
        self.create_delegate_management_tab()
        self.create_calendar_management_tab()
        self.create_custom_command_tab()
        self.create_output_tab()
        
        # Status bar
        self.status_var = tk.StringVar()
        self.status_var.set("Ready")
        status_bar = ttk.Label(self.root, textvariable=self.status_var, relief='sunken')
        status_bar.pack(side='bottom', fill='x')
        
    def create_user_management_tab(self):
        # User Management Tab
        user_frame = ttk.Frame(self.notebook)
        self.notebook.add(user_frame, text="User Management")
        
        # User operations section
        user_ops_frame = ttk.LabelFrame(user_frame, text="User Operations", padding=10)
        user_ops_frame.pack(fill='x', padx=10, pady=5)
        
        # Get user info
        ttk.Label(user_ops_frame, text="Email:").grid(row=0, column=0, sticky='w', padx=5)
        self.user_email_entry = ttk.Entry(user_ops_frame, width=40)
        self.user_email_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Button(user_ops_frame, text="Get User Info", 
                  command=lambda: self.run_gam_command(f"gam info user {self.user_email_entry.get()}")).grid(row=0, column=2, padx=5)
        
        # List users in OU - Single Selection
        ttk.Label(user_ops_frame, text="Organizational Unit:").grid(row=1, column=0, sticky='w', padx=5)
        self.ou_combobox = ttk.Combobox(user_ops_frame, width=45, state="readonly", font=('Arial', 9))
        self.ou_combobox.grid(row=1, column=1, padx=5, pady=2)
        self.ou_combobox.set("🔄 Loading organizational units...")
        
        ttk.Button(user_ops_frame, text="List Users in OU", 
                  command=lambda: self.list_users_in_single_ou()).grid(row=1, column=2, padx=5)
        
        ttk.Button(user_ops_frame, text="Refresh OUs", 
                  command=lambda: self.refresh_organizational_units()).grid(row=1, column=3, padx=5)
        
        # Force sign out user
        ttk.Label(user_ops_frame, text="Force Sign Out:").grid(row=2, column=0, sticky='w', padx=5)
        self.signout_email_entry = ttk.Entry(user_ops_frame, width=40)
        self.signout_email_entry.grid(row=2, column=1, padx=5, pady=2)
        
        ttk.Button(user_ops_frame, text="Force Sign Out", 
                  command=lambda: self.force_sign_out_user()).grid(row=2, column=2, padx=5)
        
        # Multi-OU Selection Section
        multi_ou_frame = ttk.LabelFrame(user_frame, text="Multi-OU User Management", padding=10)
        multi_ou_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Available and Selected OUs
        selection_frame = ttk.Frame(multi_ou_frame)
        selection_frame.pack(fill='both', expand=True, pady=2)
        
        # Available OUs
        available_frame = ttk.Frame(selection_frame)
        available_frame.pack(side='left', fill='both', expand=True, padx=(0, 5))
        
        ttk.Label(available_frame, text="Available OUs:").pack(anchor='w')
        
        available_listbox_frame = ttk.Frame(available_frame)
        available_listbox_frame.pack(fill='both', expand=True, pady=(2, 0))
        
        self.available_ous_listbox = tk.Listbox(available_listbox_frame, height=6, selectmode=tk.EXTENDED)
        ou_scrollbar = ttk.Scrollbar(available_listbox_frame, orient='vertical', command=self.available_ous_listbox.yview)
        self.available_ous_listbox.configure(yscrollcommand=ou_scrollbar.set)
        
        self.available_ous_listbox.pack(side='left', fill='both', expand=True)
        ou_scrollbar.pack(side='right', fill='y')
        
        # Control buttons
        control_frame = ttk.Frame(selection_frame)
        control_frame.pack(side='left', padx=10)
        
        ttk.Button(control_frame, text="Add >>", 
                  command=lambda: self.add_selected_ous()).pack(pady=(0, 5))
        ttk.Button(control_frame, text="<< Remove", 
                  command=lambda: self.remove_selected_ous()).pack(pady=(0, 5))
        ttk.Button(control_frame, text="Clear All", 
                  command=lambda: self.clear_selected_ous()).pack()
        
        # Selected OUs
        selected_frame = ttk.Frame(selection_frame)
        selected_frame.pack(side='left', fill='both', expand=True, padx=(5, 0))
        
        ttk.Label(selected_frame, text="Selected OUs:").pack(anchor='w')
        
        selected_listbox_frame = ttk.Frame(selected_frame)
        selected_listbox_frame.pack(fill='both', expand=True, pady=(2, 0))
        
        self.selected_ous_listbox = tk.Listbox(selected_listbox_frame, height=6, selectmode=tk.EXTENDED)
        selected_scrollbar = ttk.Scrollbar(selected_listbox_frame, orient='vertical', command=self.selected_ous_listbox.yview)
        self.selected_ous_listbox.configure(yscrollcommand=selected_scrollbar.set)
        
        self.selected_ous_listbox.pack(side='left', fill='both', expand=True)
        selected_scrollbar.pack(side='right', fill='y')
        
        # Multi-OU actions
        multi_action_frame = ttk.Frame(multi_ou_frame)
        multi_action_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Button(multi_action_frame, text="List Users from Selected OUs", 
                  command=lambda: self.list_users_from_multiple_ous()).pack(side='left', padx=(0, 10))
        
        self.selected_ous_count = ttk.Label(multi_action_frame, text="0 OUs selected", font=('Arial', 9, 'italic'))
        self.selected_ous_count.pack(side='left')
        
        # User Reporting Section
        report_frame = ttk.LabelFrame(user_frame, text="User Reporting", padding=10)
        report_frame.pack(fill='x', padx=10, pady=5)
        
        # Report options
        options_frame = ttk.Frame(report_frame)
        options_frame.pack(fill='x', pady=2)
        
        ttk.Label(options_frame, text="Include in Report:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5)
        
        # User info options
        self.include_basic_info = tk.BooleanVar(value=True)
        self.include_groups = tk.BooleanVar(value=True)
        self.include_last_login = tk.BooleanVar(value=False)
        self.include_ou_path = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame, text="Basic Info", variable=self.include_basic_info).grid(row=0, column=1, padx=5)
        ttk.Checkbutton(options_frame, text="Group Memberships", variable=self.include_groups).grid(row=0, column=2, padx=5)
        ttk.Checkbutton(options_frame, text="Last Login", variable=self.include_last_login).grid(row=1, column=1, padx=5)
        ttk.Checkbutton(options_frame, text="OU Path", variable=self.include_ou_path).grid(row=1, column=2, padx=5)
        
        # Report action buttons
        report_action_frame = ttk.Frame(report_frame)
        report_action_frame.pack(fill='x', pady=(10, 0))
        
        ttk.Button(report_action_frame, text="Generate User Report", 
                  command=lambda: self.generate_user_report()).pack(side='left', padx=(0, 10))
        
        ttk.Button(report_action_frame, text="Save Report to CSV", 
                  command=lambda: self.save_user_report()).pack(side='left', padx=(0, 10))
        
        self.user_report_status = ttk.Label(report_action_frame, text="Ready to generate report", 
                                          font=('Arial', 9, 'italic'))
        self.user_report_status.pack(side='left')

    def create_drive_management_tab(self):
        # Drive Management Tab
        drive_frame = ttk.Frame(self.notebook)
        self.notebook.add(drive_frame, text="Drive Management")
        
        # Drive operations section
        drive_ops_frame = ttk.LabelFrame(drive_frame, text="Drive Operations", padding=10)
        drive_ops_frame.pack(fill='x', padx=10, pady=5)
        
        # Transfer ownership
        ttk.Label(drive_ops_frame, text="From User:").grid(row=0, column=0, sticky='w', padx=5)
        self.from_user_entry = ttk.Entry(drive_ops_frame, width=30)
        self.from_user_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(drive_ops_frame, text="To User:").grid(row=0, column=2, sticky='w', padx=5)
        self.to_user_entry = ttk.Entry(drive_ops_frame, width=30)
        self.to_user_entry.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Button(drive_ops_frame, text="Transfer All Files", 
                  command=lambda: self.transfer_drive_ownership()).grid(row=0, column=4, padx=5)
        
        # Show drive usage
        ttk.Label(drive_ops_frame, text="User for Drive Usage:").grid(row=1, column=0, sticky='w', padx=5)
        self.drive_usage_entry = ttk.Entry(drive_ops_frame, width=30)
        self.drive_usage_entry.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Button(drive_ops_frame, text="Show Drive Usage", 
                  command=lambda: self.run_gam_command(f"gam user {self.drive_usage_entry.get()} show drivefileacl")).grid(row=1, column=2, padx=5)
        
    def create_group_management_tab(self):
        # Group Management Tab
        group_frame = ttk.Frame(self.notebook)
        self.notebook.add(group_frame, text="Group Management")
        
        # Basic group operations section
        group_ops_frame = ttk.LabelFrame(group_frame, text="Basic Group Operations", padding=10)
        group_ops_frame.pack(fill='x', padx=10, pady=5)
        
        # List groups for OU
        ttk.Label(group_ops_frame, text="OU for Groups:").grid(row=0, column=0, sticky='w', padx=5)
        self.group_ou_combobox = ttk.Combobox(group_ops_frame, width=45, state="readonly", font=('Arial', 9))
        self.group_ou_combobox.grid(row=0, column=1, padx=5, pady=2)
        self.group_ou_combobox.set("🔄 Loading organizational units...")
        
        ttk.Button(group_ops_frame, text="List Groups in OU", 
                  command=lambda: self.list_groups_in_ou()).grid(row=0, column=2, padx=5)
        
        # Get group members
        ttk.Label(group_ops_frame, text="Group Email:").grid(row=1, column=0, sticky='w', padx=5)
        self.group_email_entry = ttk.Entry(group_ops_frame, width=40)
        self.group_email_entry.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Button(group_ops_frame, text="Get Group Members", 
                  command=lambda: self.run_gam_command(f"gam print group-members group {self.group_email_entry.get()}")).grid(row=1, column=2, padx=5)
        
        # Advanced group reporting section
        report_frame = ttk.LabelFrame(group_frame, text="Advanced Group Reporting", padding=10)
        report_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # Step 1: Capture groups from OU
        capture_frame = ttk.LabelFrame(report_frame, text="Step 1: Capture Groups from OU", padding=10)
        capture_frame.pack(fill='x', pady=5)
        
        capture_info = ttk.Label(capture_frame, 
                               text="First, list groups in an OU above, then capture them for detailed reporting:",
                               font=('Arial', 9, 'italic'))
        capture_info.pack(anchor='w', pady=(0, 5))
        
        capture_buttons = ttk.Frame(capture_frame)
        capture_buttons.pack(fill='x')
        
        ttk.Button(capture_buttons, text="📋 Capture Listed Groups", 
                  command=lambda: self.capture_groups_from_output()).pack(side='left', padx=5)
        
        self.captured_groups_label = ttk.Label(capture_buttons, text="No groups captured yet", 
                                             font=('Arial', 9, 'italic'))
        self.captured_groups_label.pack(side='left', padx=10)
        
        # Step 2: Configure report options
        config_frame = ttk.LabelFrame(report_frame, text="Step 2: Configure Report Options", padding=10)
        config_frame.pack(fill='x', pady=5)
        
        # Report options
        options_frame = ttk.Frame(config_frame)
        options_frame.pack(fill='x', pady=5)
        
        ttk.Label(options_frame, text="Include in Report:", font=('Arial', 10, 'bold')).grid(row=0, column=0, sticky='w', padx=5)
        
        self.include_members = tk.BooleanVar(value=True)
        self.include_owners = tk.BooleanVar(value=True)
        self.include_managers = tk.BooleanVar(value=True)
        
        ttk.Checkbutton(options_frame, text="👥 Members", variable=self.include_members).grid(row=0, column=1, padx=5)
        ttk.Checkbutton(options_frame, text="👑 Owners", variable=self.include_owners).grid(row=0, column=2, padx=5)
        ttk.Checkbutton(options_frame, text="🔧 Managers", variable=self.include_managers).grid(row=0, column=3, padx=5)
        
        # Additional options
        additional_frame = ttk.Frame(config_frame)
        additional_frame.pack(fill='x', pady=5)
        
        ttk.Label(additional_frame, text="Report Format:").grid(row=0, column=0, sticky='w', padx=5)
        
        self.report_format = tk.StringVar(value="detailed")
        ttk.Radiobutton(additional_frame, text="📊 Detailed (one row per member)", 
                       variable=self.report_format, value="detailed").grid(row=0, column=1, sticky='w', padx=5)
        ttk.Radiobutton(additional_frame, text="📝 Summary (one row per group)", 
                       variable=self.report_format, value="summary").grid(row=0, column=2, sticky='w', padx=5)
        
        # Step 3: Generate report
        generate_frame = ttk.LabelFrame(report_frame, text="Step 3: Generate Report", padding=10)
        generate_frame.pack(fill='x', pady=5)
        
        generate_buttons = ttk.Frame(generate_frame)
        generate_buttons.pack(fill='x')
        
        ttk.Button(generate_buttons, text="📈 Generate Group Report", 
                  command=lambda: self.generate_group_report()).pack(side='left', padx=5)
        
        ttk.Button(generate_buttons, text="💾 Save Report to CSV", 
                  command=lambda: self.save_group_report()).pack(side='left', padx=5)
        
        self.report_status_label = ttk.Label(generate_buttons, text="Ready to generate report", 
                                           font=('Arial', 9, 'italic'))
        self.report_status_label.pack(side='left', padx=10)
        
    def create_delegate_management_tab(self):
        # Delegate Management Tab
        delegate_frame = ttk.Frame(self.notebook)
        self.notebook.add(delegate_frame, text="Delegate Management")
        
        # Delegate operations section
        delegate_ops_frame = ttk.LabelFrame(delegate_frame, text="Delegate Operations", padding=10)
        delegate_ops_frame.pack(fill='x', padx=10, pady=5)
        
        # Add delegate
        ttk.Label(delegate_ops_frame, text="User Email:").grid(row=0, column=0, sticky='w', padx=5)
        self.delegate_user_entry = ttk.Entry(delegate_ops_frame, width=30)
        self.delegate_user_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(delegate_ops_frame, text="Delegate Email:").grid(row=0, column=2, sticky='w', padx=5)
        self.delegate_email_entry = ttk.Entry(delegate_ops_frame, width=30)
        self.delegate_email_entry.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Button(delegate_ops_frame, text="Add Delegate", 
                  command=lambda: self.run_gam_command(f"gam user {self.delegate_user_entry.get()} add delegate {self.delegate_email_entry.get()}")).grid(row=0, column=4, padx=5)
        
        # Show delegates
        ttk.Label(delegate_ops_frame, text="Show Delegates for:").grid(row=1, column=0, sticky='w', padx=5)
        self.show_delegate_entry = ttk.Entry(delegate_ops_frame, width=30)
        self.show_delegate_entry.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Button(delegate_ops_frame, text="Show Delegates", 
                  command=lambda: self.run_gam_command(f"gam user {self.show_delegate_entry.get()} print delegates")).grid(row=1, column=2, padx=5)
        
    def create_calendar_management_tab(self):
        # Calendar Management Tab
        calendar_frame = ttk.Frame(self.notebook)
        self.notebook.add(calendar_frame, text="Calendar Management")
        
        # Step 1: Find Events Section
        search_frame = ttk.LabelFrame(calendar_frame, text="Step 1: Find Calendar Events", padding=10)
        search_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        # User and search criteria
        search_criteria_frame = ttk.Frame(search_frame)
        search_criteria_frame.pack(fill='x', pady=5)
        
        ttk.Label(search_criteria_frame, text="User Email:").grid(row=0, column=0, sticky='w', padx=5)
        self.calendar_user_entry = ttk.Entry(search_criteria_frame, width=30)
        self.calendar_user_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(search_criteria_frame, text="Search Query:").grid(row=0, column=2, sticky='w', padx=5)
        self.event_query_entry = ttk.Entry(search_criteria_frame, width=30)
        self.event_query_entry.grid(row=0, column=3, padx=5, pady=2)
        
        # Date range search
        ttk.Label(search_criteria_frame, text="Start Date:").grid(row=1, column=0, sticky='w', padx=5)
        self.start_date_entry = ttk.Entry(search_criteria_frame, width=12)
        self.start_date_entry.grid(row=1, column=1, padx=5, pady=2, sticky='w')
        
        ttk.Label(search_criteria_frame, text="End Date:").grid(row=1, column=2, sticky='w', padx=5)
        self.end_date_entry = ttk.Entry(search_criteria_frame, width=12)
        self.end_date_entry.grid(row=1, column=3, padx=5, pady=2, sticky='w')
        
        # Search buttons
        button_frame = ttk.Frame(search_frame)
        button_frame.pack(fill='x', pady=5)
        
        ttk.Button(button_frame, text="Search by Query", 
                  command=lambda: self.search_events_by_query()).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Search by Date Range", 
                  command=lambda: self.search_events_by_date()).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="List All Events (Today)", 
                  command=lambda: self.list_today_events()).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="List User's Calendars", 
                  command=lambda: self.list_user_calendars()).pack(side='left', padx=5)
        
        # Step 2: Select and Delete Event Section
        delete_frame = ttk.LabelFrame(calendar_frame, text="Step 2: Select and Delete Calendar Event", padding=10)
        delete_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        delete_info_label = ttk.Label(delete_frame, 
                                     text="Found events will appear below. Select an event to delete:",
                                     font=('Arial', 9, 'italic'))
        delete_info_label.pack(anchor='w', pady=(0, 5))
        
        # Event selection section
        selection_frame = ttk.Frame(delete_frame)
        selection_frame.pack(fill='both', expand=True, pady=5)
        
        # Event listbox with scrollbar
        listbox_frame = ttk.Frame(selection_frame)
        listbox_frame.pack(fill='both', expand=True)
        
        # Create listbox with scrollbars
        self.event_listbox = tk.Listbox(listbox_frame, height=6, font=('Consolas', 9))
        scrollbar_y = ttk.Scrollbar(listbox_frame, orient='vertical', command=self.event_listbox.yview)
        scrollbar_x = ttk.Scrollbar(listbox_frame, orient='horizontal', command=self.event_listbox.xview)
        
        self.event_listbox.configure(yscrollcommand=scrollbar_y.set, xscrollcommand=scrollbar_x.set)
        
        # Pack scrollbars and listbox
        scrollbar_y.pack(side='right', fill='y')
        scrollbar_x.pack(side='bottom', fill='x')
        self.event_listbox.pack(side='left', fill='both', expand=True)
        
        # Bind selection event
        self.event_listbox.bind('<<ListboxSelect>>', self.on_event_select)
        
        # Selected event details
        details_frame = ttk.Frame(delete_frame)
        details_frame.pack(fill='x', pady=5)
        
        ttk.Label(details_frame, text="Selected Event Details:", font=('Arial', 10, 'bold')).pack(anchor='w')
        
        detail_fields_frame = ttk.Frame(details_frame)
        detail_fields_frame.pack(fill='x', pady=2)
        
        ttk.Label(detail_fields_frame, text="User Email:").grid(row=0, column=0, sticky='w', padx=5)
        self.delete_user_entry = ttk.Entry(detail_fields_frame, width=30, state='readonly')
        self.delete_user_entry.grid(row=0, column=1, padx=5, pady=2)
        
        ttk.Label(detail_fields_frame, text="Event Title:").grid(row=0, column=2, sticky='w', padx=5)
        self.selected_event_title = ttk.Entry(detail_fields_frame, width=30, state='readonly')
        self.selected_event_title.grid(row=0, column=3, padx=5, pady=2)
        
        ttk.Label(detail_fields_frame, text="Event ID:").grid(row=1, column=0, sticky='w', padx=5)
        self.delete_event_id_entry = ttk.Entry(detail_fields_frame, width=30, state='readonly')
        self.delete_event_id_entry.grid(row=1, column=1, padx=5, pady=2)
        
        ttk.Label(detail_fields_frame, text="Date/Time:").grid(row=1, column=2, sticky='w', padx=5)
        self.selected_event_datetime = ttk.Entry(detail_fields_frame, width=30, state='readonly')
        self.selected_event_datetime.grid(row=1, column=3, padx=5, pady=2)
        
        # Delete button
        delete_button_frame = ttk.Frame(delete_frame)
        delete_button_frame.pack(fill='x', pady=5)
        
        ttk.Button(delete_button_frame, text="🗑️ Delete Selected Event", 
                  command=lambda: self.delete_selected_calendar_event()).pack(side='left', padx=5)
        
        ttk.Button(delete_button_frame, text="Clear Selection", 
                  command=lambda: self.clear_event_selection()).pack(side='left', padx=5)
        
        # Helper text
        helper_frame = ttk.LabelFrame(calendar_frame, text="💡 Quick Tips", padding=10)
        helper_frame.pack(fill='x', padx=10, pady=5)
        
        helper_text = ttk.Label(helper_frame, 
                               text="• Use specific search terms like 'Budget Meeting' or 'Team Standup'\n"
                                    "• Date format: YYYY-MM-DD (e.g., 2025-11-10)\n"
                                    "• Event IDs are long strings like '4a1b2c3d4e5f6g7h8i9j0k'\n"
                                    "• Copy/paste Event IDs from the Output tab to avoid typos",
                               justify='left')
        helper_text.pack(anchor='w')
        
    def create_custom_command_tab(self):
        # Custom Command Tab
        custom_frame = ttk.Frame(self.notebook)
        self.notebook.add(custom_frame, text="Custom Commands")
        
        # Custom command section
        custom_ops_frame = ttk.LabelFrame(custom_frame, text="Custom GAM Command", padding=10)
        custom_ops_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        ttk.Label(custom_ops_frame, text="Enter GAM Command:").pack(anchor='w', padx=5)
        self.custom_command_entry = scrolledtext.ScrolledText(custom_ops_frame, height=5, width=80)
        self.custom_command_entry.pack(fill='x', padx=5, pady=5)
        
        button_frame = ttk.Frame(custom_ops_frame)
        button_frame.pack(fill='x', padx=5, pady=5)
        
        ttk.Button(button_frame, text="Execute Command", 
                  command=lambda: self.run_custom_command()).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Clear", 
                  command=lambda: self.custom_command_entry.delete(1.0, tk.END)).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Save Output", 
                  command=lambda: self.save_output()).pack(side='left', padx=5)
        
    def create_output_tab(self):
        # Output Tab
        output_frame = ttk.Frame(self.notebook)
        self.notebook.add(output_frame, text="Output")
        
        # Output section
        output_ops_frame = ttk.LabelFrame(output_frame, text="Command Output", padding=10)
        output_ops_frame.pack(fill='both', expand=True, padx=10, pady=5)
        
        self.output_text = scrolledtext.ScrolledText(output_ops_frame, height=25, width=100)
        self.output_text.pack(fill='both', expand=True, padx=5, pady=5)
        
        # Clear output button
        ttk.Button(output_ops_frame, text="Clear Output", 
                  command=lambda: self.output_text.delete(1.0, tk.END)).pack(pady=5)
        
    def load_organizational_units(self):
        """Load organizational units from GAM on startup"""
        self.status_var.set("Loading organizational units...")
        
        def fetch_ous():
            try:
                # Execute GAM command to get all OUs
                result = subprocess.run("gam print orgs", shell=True, capture_output=True, text=True, timeout=60)
                
                # Parse the results in main thread
                self.root.after(0, lambda: self.process_ou_results(result))
                
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self.handle_ou_error("Timeout loading organizational units"))
            except Exception as e:
                self.root.after(0, lambda: self.handle_ou_error(f"Error loading OUs: {str(e)}"))
        
        # Run in separate thread to avoid blocking GUI
        thread = threading.Thread(target=fetch_ous)
        thread.daemon = True
        thread.start()
    
    def process_ou_results(self, result):
        """Process the GAM organizational units results"""
        self.organizational_units = []
        self.formatted_ous = []  # Store formatted display versions
        
        if result.returncode == 0:
            # Parse the CSV output from GAM
            lines = result.stdout.strip().split('\n')
            if len(lines) > 1:  # Skip header line
                for line in lines[1:]:
                    # Extract OU path from CSV (usually first column)
                    parts = line.split(',')
                    if parts:
                        ou_path = parts[0].strip('"')
                        if ou_path and ou_path != '/':  # Skip root OU
                            self.organizational_units.append(ou_path)
                            # Create formatted version for display
                            formatted_ou = self.format_ou_for_display(ou_path)
                            self.formatted_ous.append(formatted_ou)
            
            # Sort both lists together by the original OU path
            combined = list(zip(self.organizational_units, self.formatted_ous))
            combined.sort(key=lambda x: x[0])
            self.organizational_units, self.formatted_ous = zip(*combined) if combined else ([], [])
            self.organizational_units = list(self.organizational_units)
            self.formatted_ous = list(self.formatted_ous)
            
            # Update comboxes
            self.update_ou_comboboxes()
            self.status_var.set(f"✅ Loaded {len(self.organizational_units)} organizational units")
            
        else:
            self.handle_ou_error(f"GAM error: {result.stderr}")
    
    def format_ou_for_display(self, ou_path):
        """Format OU path for better display in dropdown"""
        if not ou_path or ou_path == '/':
            return "📁 Root Organization"
        
        # Remove leading slash and split into parts
        parts = ou_path.strip('/').split('/')
        
        # Create indentation based on depth
        depth = len(parts) - 1
        indent = "  " * depth
        
        # Get the last part (actual OU name)
        ou_name = parts[-1]
        
        # Add appropriate icons and formatting
        if depth == 0:
            # Top level OU
            icon = "🏢"
        elif depth == 1:
            # Second level
            icon = "📂"
        else:
            # Deeper levels
            icon = "📁"
        
        return f"{indent}{icon} {ou_name}"
    
    def handle_ou_error(self, error_message):
        """Handle errors when loading OUs"""
        self.organizational_units = []
        self.update_ou_comboboxes()
        self.status_var.set("Ready (Could not load OUs)")
        
        # Optionally show error in output tab
        if hasattr(self, 'output_text'):
            self.output_text.insert(tk.END, f"\n[{datetime.now().strftime('%H:%M:%S')}] {error_message}\n")
            self.output_text.insert(tk.END, "You can still enter OU paths manually or try refreshing.\n")
            self.output_text.insert(tk.END, "-" * 50 + "\n")
    
    def update_ou_comboboxes(self):
        """Update all OU comboboxes and listboxes with loaded data"""
        self.is_loading_ous = False
        
        if self.organizational_units:
            # Update user management single OU combobox
            self.ou_combobox['values'] = self.organizational_units
            self.ou_combobox.set("Select an OU...")
            
            # Update available OUs listbox for multi-selection
            self.available_ous_listbox.delete(0, tk.END)
            for formatted_ou in self.formatted_ous:
                self.available_ous_listbox.insert(tk.END, formatted_ou)
            
            # Update group management OU combobox  
            if hasattr(self, 'group_ou_combobox'):
                self.group_ou_combobox['values'] = self.organizational_units
                self.group_ou_combobox.set("Select an OU...")
        else:
            # No OUs found
            self.ou_combobox['values'] = []
            self.ou_combobox.set("No OUs found")
            
            self.available_ous_listbox.delete(0, tk.END)
            self.available_ous_listbox.insert(0, "No OUs found")
            
            if hasattr(self, 'group_ou_combobox'):
                self.group_ou_combobox['values'] = []
                self.group_ou_combobox.set("No OUs found")
    
    def refresh_organizational_units(self):
        """Refresh the organizational units list"""
        if not self.is_loading_ous:
            self.is_loading_ous = True
            self.ou_combobox.set("🔄 Loading OUs...")
            if hasattr(self, 'group_ou_combobox'):
                self.group_ou_combobox.set("🔄 Loading OUs...")
            self.available_ous_listbox.delete(0, tk.END)
            self.available_ous_listbox.insert(0, "🔄 Loading organizational units...")
            self.load_organizational_units()

    def run_gam_command(self, command):
        """Run a GAM command in the background"""
        self.status_var.set(f"Running: {command}")
        self.output_text.insert(tk.END, f"\n[{datetime.now().strftime('%H:%M:%S')}] Executing: {command}\n")
        self.output_text.insert(tk.END, "-" * 50 + "\n")
        self.output_text.see(tk.END)
        
        # Switch to output tab
        self.notebook.select(6)  # Output tab index
        
        def execute_command():
            try:
                # Execute the GAM command
                result = subprocess.run(command, shell=True, capture_output=True, text=True, timeout=300)
                
                # Update GUI in main thread
                self.root.after(0, lambda: self.display_command_result(result, command))
                
            except subprocess.TimeoutExpired:
                self.root.after(0, lambda: self.display_error("Command timed out after 5 minutes"))
            except Exception as e:
                self.root.after(0, lambda: self.display_error(f"Error executing command: {str(e)}"))
        
        # Run command in separate thread
        thread = threading.Thread(target=execute_command)
        thread.daemon = True
        thread.start()
        
    def display_command_result(self, result, command):
        """Display the result of a command execution"""
        if result.returncode == 0:
            self.output_text.insert(tk.END, result.stdout)
            self.status_var.set("Command completed successfully")
            
            # Check if this was a calendar event search command and parse results
            if "print events" in command or "print calendar-events" in command:
                self.parse_calendar_events(result.stdout, command)
            
            # Check if this was a group listing command and capture available groups
            if "print groups" in command and "ou " in command:
                self.last_group_output = result.stdout
                
        else:
            self.output_text.insert(tk.END, f"ERROR: {result.stderr}")
            self.status_var.set("Command failed")
            
        self.output_text.insert(tk.END, "\n" + "=" * 50 + "\n\n")
        self.output_text.see(tk.END)
        
    def display_error(self, error_message):
        """Display an error message"""
        self.output_text.insert(tk.END, f"ERROR: {error_message}\n")
        self.output_text.insert(tk.END, "\n" + "=" * 50 + "\n\n")
        self.output_text.see(tk.END)
        self.status_var.set("Error occurred")
        
    def run_custom_command(self):
        """Run a custom GAM command"""
        command = self.custom_command_entry.get(1.0, tk.END).strip()
        if command:
            self.run_gam_command(command)
        else:
            messagebox.showwarning("Warning", "Please enter a GAM command")
            
    def transfer_drive_ownership(self):
        """Transfer drive ownership from one user to another"""
        from_user = self.from_user_entry.get().strip()
        to_user = self.to_user_entry.get().strip()
        
        if from_user and to_user:
            response = messagebox.askyesno("Confirm", f"Transfer all files from {from_user} to {to_user}?")
            if response:
                # This is a complex operation that might need multiple commands
                self.run_gam_command(f"gam user {from_user} transfer drive {to_user}")
        else:
            messagebox.showwarning("Warning", "Please enter both from and to user emails")
            
    def list_groups_in_ou(self):
        """List all groups for users in a specific OU"""
        ou = self.group_ou_combobox.get().strip()
        if ou and ou != "Loading OUs..." and ou != "No OUs found":
            # This might require a more complex command depending on GAM version
            self.run_gam_command(f"gam ou \"{ou}\" print groups")
        else:
            messagebox.showwarning("Warning", "Please select an organizational unit")
            
    def save_output(self):
        """Save the output to a file"""
        output_content = self.output_text.get(1.0, tk.END)
        if output_content.strip():
            filename = filedialog.asksaveasfilename(
                defaultextension=".txt",
                filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
                title="Save Output"
            )
            if filename:
                try:
                    with open(filename, 'w', encoding='utf-8') as f:
                        f.write(output_content)
                    messagebox.showinfo("Success", f"Output saved to {filename}")
                except Exception as e:
                    messagebox.showerror("Error", f"Could not save file: {str(e)}")
        else:
            messagebox.showwarning("Warning", "No output to save")

    def list_user_calendars(self):
        """List all calendars that a user has access to"""
        user_email = self.calendar_user_entry.get().strip()
        if user_email:
            self.run_gam_command(f"gam user {user_email} print calendars")
        else:
            messagebox.showwarning("Warning", "Please enter a user email")
    
    def search_events_by_query(self):
        """Search calendar events by query text"""
        user_email = self.calendar_user_entry.get().strip()
        query = self.event_query_entry.get().strip()
        
        if not user_email:
            messagebox.showwarning("Warning", "Please enter a user email")
            return
        
        if not query:
            messagebox.showwarning("Warning", "Please enter a search query (e.g., 'Budget Meeting')")
            return
        
        # Use the correct GAM command for searching events by query
        self.run_gam_command(f'gam user {user_email} print events query "{query}"')
    
    def search_events_by_date(self):
        """Search calendar events within a date range"""
        user_email = self.calendar_user_entry.get().strip()
        start_date = self.start_date_entry.get().strip()
        end_date = self.end_date_entry.get().strip()
        
        if not user_email:
            messagebox.showwarning("Warning", "Please enter a user email")
            return
        
        if not start_date and not end_date:
            messagebox.showwarning("Warning", "Please enter at least a start date or end date")
            return
        
        # Build the GAM command with date parameters
        base_command = f"gam user {user_email} print calendar-events"
        date_params = []
        
        if start_date:
            try:
                # Validate and format start date
                datetime.strptime(start_date, '%Y-%m-%d')
                date_params.append(f"starttime {start_date}T00:00:00")
            except ValueError:
                messagebox.showerror("Error", "Start date must be in YYYY-MM-DD format")
                return
        
        if end_date:
            try:
                # Validate and format end date
                datetime.strptime(end_date, '%Y-%m-%d')
                date_params.append(f"endtime {end_date}T23:59:59")
            except ValueError:
                messagebox.showerror("Error", "End date must be in YYYY-MM-DD format")
                return
        
        if date_params:
            command = f"{base_command} {' '.join(date_params)}"
        else:
            command = base_command
        
        self.run_gam_command(command)
    
    def list_today_events(self):
        """List all events for today for the specified user"""
        user_email = self.calendar_user_entry.get().strip()
        if not user_email:
            messagebox.showwarning("Warning", "Please enter a user email")
            return
        
        # Get today's date
        today = datetime.now().strftime('%Y-%m-%d')
        
        # Set the dates in the fields for user reference
        self.start_date_entry.delete(0, tk.END)
        self.start_date_entry.insert(0, today)
        self.end_date_entry.delete(0, tk.END)
        self.end_date_entry.insert(0, today)
        
        # Run the search for today
        self.run_gam_command(f"gam user {user_email} print calendar-events starttime {today}T00:00:00 endtime {today}T23:59:59")
    

    
    def force_sign_out_user(self):
        """Force sign out a user from all devices"""
        email = self.signout_email_entry.get().strip()
        if email:
            response = messagebox.askyesno("Confirm Force Sign Out", 
                                         f"Force sign out {email} from all devices?\n\n"
                                         f"This will:\n"
                                         f"• Sign out the user from all devices\n"
                                         f"• Invalidate all active sessions\n"
                                         f"• Require the user to sign in again")
            if response:
                self.run_gam_command(f"gam user {email} signout")
        else:
            messagebox.showwarning("Warning", "Please enter a user email")
    
    def parse_calendar_events(self, output, command):
        """Parse calendar event output and populate the selection list"""
        self.parsed_events = []
        
        try:
            lines = output.strip().split('\n')
            if len(lines) < 2:  # No events found
                self.event_listbox.delete(0, tk.END)
                self.event_listbox.insert(0, "No events found")
                return
            
            # Skip informational lines like "Getting Events for..."
            header_line = None
            data_start = 0
            
            for i, line in enumerate(lines):
                if line.startswith('primaryEmail,') or 'id,' in line:
                    header_line = line
                    data_start = i + 1
                    break
            
            if not header_line:
                self.event_listbox.delete(0, tk.END)
                self.event_listbox.insert(0, "Could not parse event data")
                return
            
            # Parse header to get column indices
            headers = [h.strip() for h in header_line.split(',')]
            
            # Find required column indices
            try:
                email_idx = headers.index('primaryEmail')
                id_idx = headers.index('id')
                summary_idx = headers.index('summary')
                start_idx = headers.index('start.dateTime')
                status_idx = headers.index('status')
            except ValueError as e:
                self.event_listbox.delete(0, tk.END)
                self.event_listbox.insert(0, f"Missing required columns: {e}")
                return
            
            # Parse event data
            for line_num, line in enumerate(lines[data_start:], data_start + 1):
                if not line.strip():
                    continue
                
                try:
                    # Parse CSV line (handling commas in quoted fields)
                    import csv
                    import io
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    if len(fields) > max(email_idx, id_idx, summary_idx, start_idx, status_idx):
                        event_data = {
                            'email': fields[email_idx],
                            'id': fields[id_idx],
                            'summary': fields[summary_idx],
                            'start_datetime': fields[start_idx],
                            'status': fields[status_idx],
                            'line_number': line_num
                        }
                        
                        # Format datetime for display
                        try:
                            if event_data['start_datetime']:
                                # Parse and format the datetime
                                from datetime import datetime
                                dt = datetime.fromisoformat(event_data['start_datetime'].replace('Z', '+00:00'))
                                formatted_dt = dt.strftime('%Y-%m-%d %I:%M %p')
                                event_data['formatted_datetime'] = formatted_dt
                            else:
                                event_data['formatted_datetime'] = 'No date'
                        except:
                            event_data['formatted_datetime'] = event_data['start_datetime']
                        
                        self.parsed_events.append(event_data)
                        
                except Exception as e:
                    continue  # Skip malformed lines
            
            # Update the listbox
            self.update_event_listbox()
            
        except Exception as e:
            self.event_listbox.delete(0, tk.END)
            self.event_listbox.insert(0, f"Error parsing events: {str(e)}")
    
    def update_event_listbox(self):
        """Update the event listbox with parsed events"""
        self.event_listbox.delete(0, tk.END)
        
        if not self.parsed_events:
            self.event_listbox.insert(0, "No events found")
            return
        
        # Add events to listbox
        for i, event in enumerate(self.parsed_events):
            # Format: "[Date] Event Title (Status)"
            display_text = f"[{event['formatted_datetime']}] {event['summary']} ({event['status']})"
            self.event_listbox.insert(i, display_text)
        
        # Auto-select first event if only one found
        if len(self.parsed_events) == 1:
            self.event_listbox.selection_set(0)
            self.on_event_select(None)
    
    def on_event_select(self, event):
        """Handle event selection in the listbox"""
        selection = self.event_listbox.curselection()
        if not selection or not self.parsed_events:
            return
        
        selected_idx = selection[0]
        if selected_idx >= len(self.parsed_events):
            return
            
        selected_event = self.parsed_events[selected_idx]
        
        # Update the detail fields
        self.update_entry_field(self.delete_user_entry, selected_event['email'])
        self.update_entry_field(self.selected_event_title, selected_event['summary'])
        self.update_entry_field(self.delete_event_id_entry, selected_event['id'])
        self.update_entry_field(self.selected_event_datetime, selected_event['formatted_datetime'])
    
    def update_entry_field(self, entry_widget, value):
        """Update a readonly entry field with a new value"""
        entry_widget.config(state='normal')
        entry_widget.delete(0, tk.END)
        entry_widget.insert(0, value)
        entry_widget.config(state='readonly')
    
    def clear_event_selection(self):
        """Clear the event selection and detail fields"""
        self.event_listbox.selection_clear(0, tk.END)
        self.update_entry_field(self.delete_user_entry, "")
        self.update_entry_field(self.selected_event_title, "")
        self.update_entry_field(self.delete_event_id_entry, "")
        self.update_entry_field(self.selected_event_datetime, "")
    
    def delete_selected_calendar_event(self):
        """Delete the currently selected calendar event"""
        user_email = self.delete_user_entry.get().strip()
        event_id = self.delete_event_id_entry.get().strip()
        event_title = self.selected_event_title.get().strip()
        event_datetime = self.selected_event_datetime.get().strip();
        
        if not user_email or not event_id:
            messagebox.showwarning("Warning", "Please select an event from the list above")
            return
        
        # Show detailed confirmation dialog
        response = messagebox.askyesno("⚠️ Confirm Delete Event", 
                                     f"Delete calendar event:\n\n"
                                     f"📧 User: {user_email}\n"
                                     f"📅 Event: {event_title}\n"
                                     f"🕐 Date/Time: {event_datetime}\n"
                                     f"🆔 Event ID: {event_id}\n\n"
                                     f"⚠️ This action CANNOT be undone!\n\n"
                                     f"Are you sure you want to delete this event?")
        
        if response:
            # Use the correct GAM command format for deleting calendar events
            self.run_gam_command(f"gam calendar {user_email} delete events eventId {event_id} doit")
            
            # Clear the selection after deletion attempt
            self.clear_event_selection()
            
            # Clear the event list to encourage a new search
            self.event_listbox.delete(0, tk.END)
            self.parsed_events = []
    
    def capture_groups_from_output(self):
        """Capture groups from the last group listing output"""
        if not hasattr(self, 'last_group_output') or not self.last_group_output:
            messagebox.showwarning("Warning", "No group output to capture. Please list groups in an OU first.")
            return
        
        try:
            self.captured_groups = []
            lines = self.last_group_output.strip().split('\n')
            
            # Debug: Show output in the GUI for troubleshooting
            self.output_text.insert(tk.END, f"\n[DEBUG] Analyzing group output format:\n")
            self.output_text.insert(tk.END, f"Total lines: {len(lines)}\n")
            
            # Show first few lines
            debug_lines = lines[:8] if len(lines) >= 8 else lines
            for i, line in enumerate(debug_lines):
                self.output_text.insert(tk.END, f"Line {i:2d}: {repr(line)}\n")
            if len(lines) > 8:
                self.output_text.insert(tk.END, f"... and {len(lines) - 8} more lines\n")
            self.output_text.insert(tk.END, "-" * 50 + "\n")
            
            # Find the header line with multiple strategies
            header_line = None
            data_start = 0
            
            # Strategy 1: Look for lines with common group headers
            for i, line in enumerate(lines):
                if not line.strip() or line.startswith('Getting') or line.startswith('Got'):
                    continue
                
                line_lower = line.lower()
                # Check for CSV format with group-related headers
                if ',' in line and any(keyword in line_lower for keyword in ['email', 'group', 'name', 'description']):
                    # Additional check: make sure it's not just data that happens to contain these words
                    parts = line.split(',')
                    if len(parts) >= 2:  # At least 2 columns
                        header_line = line
                        data_start = i + 1
                        self.output_text.insert(tk.END, f"[DEBUG] Found header (Strategy 1) at line {i}: {line}\n")
                        break
            
            # Strategy 2: If no header found, look for the first CSV-like line
            if not header_line:
                for i, line in enumerate(lines):
                    if not line.strip() or line.startswith('Getting') or line.startswith('Got'):
                        continue
                    if ',' in line and len(line.split(',')) >= 2:
                        # Check if this looks like it could be group data
                        parts = [p.strip() for p in line.split(',')]
                        if any('@' in part for part in parts):  # Contains email addresses
                            header_line = line
                            data_start = i
                            self.output_text.insert(tk.END, f"[DEBUG] Using first data line as header (Strategy 2) at line {i}: {line}\n")
                            break
            
            if not header_line:
                error_msg = "Could not find group data in output.\n\nTroubleshooting:\n"
                error_msg += f"- Total lines: {len(lines)}\n"
                error_msg += f"- Non-empty lines: {len([l for l in lines if l.strip()])}\n"
                error_msg += "- Check the Output tab for detailed debug information"
                
                self.output_text.insert(tk.END, f"[DEBUG] No header found. All non-empty lines:\n")
                for i, line in enumerate(lines):
                    if line.strip():
                        self.output_text.insert(tk.END, f"  {i:2d}: {line}\n")
                
                messagebox.showerror("Error", error_msg)
                return
            
            # Parse header to get column indices
            headers = [h.strip().strip('"') for h in header_line.split(',')]
            self.output_text.insert(tk.END, f"[DEBUG] Parsed headers: {headers}\n")
            
            # Find email and name columns with flexible matching
            email_idx = None
            name_idx = None
            
            # Define possible column name variations
            email_patterns = ['email', 'group', 'groupemail', 'primaryemail', 'address']
            name_patterns = ['name', 'displayname', 'groupname', 'description', 'title']
            
            for i, header in enumerate(headers):
                header_lower = header.lower().strip()
                
                # Check for email column
                if email_idx is None:
                    for pattern in email_patterns:
                        if pattern in header_lower:
                            email_idx = i
                            self.output_text.insert(tk.END, f"[DEBUG] Found email column at index {i}: '{header}'\n")
                            break
                
                # Check for name column
                if name_idx is None:
                    for pattern in name_patterns:
                        if pattern in header_lower:
                            name_idx = i
                            self.output_text.insert(tk.END, f"[DEBUG] Found name column at index {i}: '{header}'\n")
                            break
            
            # Fallback: if no email column found, try to detect it from data
            if email_idx is None:
                self.output_text.insert(tk.END, f"[DEBUG] No email column found in headers, trying to detect from data...\n")
                # Look at the first data line to find column with @ symbol
                for i, line in enumerate(lines[data_start:], data_start):
                    if line.strip() and not line.startswith('Getting') and not line.startswith('Got'):
                        try:
                            fields = next(csv.reader(io.StringIO(line)))
                            for j, field in enumerate(fields):
                                if '@' in field.strip():
                                    email_idx = j
                                    self.output_text.insert(tk.END, f"[DEBUG] Detected email column at index {j} from data\n")
                                    break
                            if email_idx is not None:
                                break
                        except:
                            continue
            
            if email_idx is None:
                messagebox.showerror("Error", 
                    f"Could not find email column in output.\n\n"
                    f"Headers found: {headers}\n\n"
                    f"Expected headers containing: {', '.join(email_patterns)}")
                return
            
            # Parse group data
            groups_found = 0
            for line_num, line in enumerate(lines[data_start:], data_start):
                if not line.strip() or line.startswith('Getting') or line.startswith('Got'):
                    continue
                
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    if len(fields) > email_idx and fields[email_idx].strip():
                        email = fields[email_idx].strip()
                        name = fields[name_idx].strip() if name_idx and len(fields) > name_idx else email
                        
                        # Validate that this looks like a group email
                        if '@' in email and '.' in email:
                            group_data = {
                                'email': email,
                                'name': name
                            }
                            self.captured_groups.append(group_data)
                            groups_found += 1
                            
                            # Show first few captured groups for verification
                            if groups_found <= 3:
                                self.output_text.insert(tk.END, f"[DEBUG] Captured: {email} ({name})\n")
                        
                except Exception as e:
                    self.output_text.insert(tk.END, f"[DEBUG] Error parsing line {line_num}: {str(e)} - Line: {line}\n")
                    continue
            
            self.output_text.insert(tk.END, f"[DEBUG] Total groups captured: {len(self.captured_groups)}\n")
            self.output_text.insert(tk.END, "=" * 50 + "\n")
            self.output_text.see(tk.END)
            
            # Update the label and show result
            if self.captured_groups:
                self.captured_groups_label.config(text=f"✅ Captured {len(self.captured_groups)} groups")
                self.report_status_label.config(text="Ready to generate report")
                messagebox.showinfo("Success", 
                    f"Successfully captured {len(self.captured_groups)} groups!\n\n"
                    f"Sample groups:\n" + 
                    "\n".join([f"• {g['email']}" for g in self.captured_groups[:5]]) +
                    (f"\n... and {len(self.captured_groups) - 5} more" if len(self.captured_groups) > 5 else ""))
            else:
                self.captured_groups_label.config(text="❌ No groups found in output")
                messagebox.showwarning("Warning", 
                    "No valid groups found in the output.\n\n"
                    "Check the Output tab for detailed debug information.\n"
                    "Make sure the GAM command returned group data in CSV format.")
                
        except Exception as e:
            self.output_text.insert(tk.END, f"[DEBUG] Exception in capture_groups_from_output: {str(e)}\n")
            self.output_text.see(tk.END)
            messagebox.showerror("Error", f"Error capturing groups: {str(e)}")
    
    def generate_group_report(self):
        """Generate detailed group report with members, owners, and managers"""
        if not self.captured_groups:
            messagebox.showwarning("Warning", "No groups captured. Please capture groups first.")
            return
        
        if not (self.include_members.get() or self.include_owners.get() or self.include_managers.get()):
            messagebox.showwarning("Warning", "Please select at least one type to include in the report.")
            return
        
        self.report_status_label.config(text="🔄 Generating report...")
        self.group_report_data = []
        
        # Generate report in background thread
        def generate_report():
            try:
                total_groups = len(self.captured_groups)
                
                for i, group in enumerate(self.captured_groups):
                    group_email = group['email']
                    group_name = group['name']
                    
                    # Update status in main thread
                    self.root.after(0, lambda i=i: self.report_status_label.config(
                        text=f"🔄 Processing group {i+1}/{total_groups}: {group_email}"))
                    
                    # Collect data for this group
                    group_data = {
                        'group_email': group_email,
                        'group_name': group_name,
                        'members': [],
                        'owners': [],
                        'managers': []
                    }
                    
                    # Get group members if requested
                    if self.include_members.get():
                        try:
                            result = subprocess.run(
                                f"gam print group-members group {group_email} role member",
                                shell=True, capture_output=True, text=True, timeout=30
                            )
                            if result.returncode == 0:
                                group_data['members'] = self.parse_group_members(result.stdout)
                        except Exception as e:
                            group_data['members'] = [f"Error: {str(e)}"]
                    
                    # Get group owners if requested
                    if self.include_owners.get():
                        try:
                            result = subprocess.run(
                                f"gam print group-members group {group_email} role owner",
                                shell=True, capture_output=True, text=True, timeout=30
                            )
                            if result.returncode == 0:
                                group_data['owners'] = self.parse_group_members(result.stdout)
                        except Exception as e:
                            group_data['owners'] = [f"Error: {str(e)}"]
                    
                    # Get group managers if requested
                    if self.include_managers.get():
                        try:
                            result = subprocess.run(
                                f"gam print group-members group {group_email} role manager",
                                shell=True, capture_output=True, text=True, timeout=30
                            )
                            if result.returncode == 0:
                                group_data['managers'] = self.parse_group_members(result.stdout)
                        except Exception as e:
                            group_data['managers'] = [f"Error: {str(e)}"]
                    
                    self.group_report_data.append(group_data)
                
                # Update UI in main thread
                self.root.after(0, lambda: self.report_generation_complete())
                
            except Exception as e:
                self.root.after(0, lambda: self.report_status_label.config(
                    text=f"❌ Error generating report: {str(e)}"))
        
        # Run in background thread
        thread = threading.Thread(target=generate_report)
        thread.daemon = True
        thread.start()
    
    def parse_group_members(self, output):
        """Parse group member output and return list of emails"""
        members = []
        try:
            lines = output.strip().split('\n')
            
            # Find header line
            header_line = None
            data_start = 0
            
            for i, line in enumerate(lines):
                if 'email' in line.lower():
                    header_line = line
                    data_start = i + 1
                    break
            
            if not header_line:
                return members
            
            # Parse members
            headers = [h.strip() for h in header_line.split(',')]
            email_idx = headers.index('email') if 'email' in headers else 0
            
            for line in lines[data_start:]:
                if not line.strip():
                    continue
                
                try:
                    reader = csv.reader(io.StringIO(line))
                    fields = next(reader)
                    
                    if len(fields) > email_idx:
                        members.append(fields[email_idx])
                        
                except Exception:
                    continue
                    
        except Exception:
            pass
        
        return members
    
    def report_generation_complete(self):
        """Handle completion of report generation"""
        if self.group_report_data:
            total_groups = len(self.group_report_data)
            total_members = sum(len(g['members']) + len(g['owners']) + len(g['managers']) 
                              for g in self.group_report_data)
            
            self.report_status_label.config(
                text=f"✅ Report ready: {total_groups} groups, {total_members} total entries")
        else:
            self.report_status_label.config(text="❌ No data generated")
    
    def save_group_report(self):
        """Save the generated group report to a CSV file"""
        if not self.group_report_data:
            messagebox.showwarning("Warning", "No report data to save. Please generate a report first.")
            return
        
        filename = filedialog.asksaveasfilename(
            defaultextension=".csv",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")],
            title="Save Group Report"
        )
        
        if not filename:
            return
        
        try:
            with open(filename, 'w', newline='', encoding='utf-8') as csvfile:
                if self.report_format.get() == "detailed":
                    self.save_detailed_report(csvfile)
                else:
                    self.save_summary_report(csvfile)
            
            messagebox.showinfo("Success", f"Group report saved to {filename}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Could not save report: {str(e)}")
    
    def save_detailed_report(self, csvfile):
        """Save detailed report (one row per member)"""
        writer = csv.writer(csvfile)
        
        # Write header
        header = ['Group Email', 'Group Name', 'Member Email', 'Role']
        writer.writerow(header)
        
        # Write data
        for group in self.group_report_data:
            group_email = group['group_email']
            group_name = group['group_name']
            
            # Write members
            if self.include_members.get():
                for member in group['members']:
                    writer.writerow([group_email, group_name, member, 'Member'])
            
            # Write owners
            if self.include_owners.get():
                for owner in group['owners']:
                    writer.writerow([group_email, group_name, owner, 'Owner'])
            
            # Write managers
            if self.include_managers.get():
                for manager in group['managers']:
                    writer.writerow([group_email, group_name, manager, 'Manager'])
    
    def save_summary_report(self, csvfile):
        """Save summary report (one row per group)"""
        writer = csv.writer(csvfile)
        
        # Build header
        header = ['Group Email', 'Group Name']
        if self.include_members.get():
            header.extend(['Member Count', 'Members'])
        if self.include_owners.get():
            header.extend(['Owner Count', 'Owners'])
        if self.include_managers.get():
            header.extend(['Manager Count', 'Managers'])
        
        writer.writerow(header)
        
        # Write data
        for group in self.group_report_data:
            row = [group['group_email'], group['group_name']]
            
            if self.include_members.get():
                members = group['members']
                row.extend([len(members), '; '.join(members)])
            
            if self.include_owners.get():
                owners = group['owners']
                row.extend([len(owners), '; '.join(owners)])
            
            if self.include_managers.get():
                managers = group['managers']
                row.extend([len(managers), '; '.join(managers)])
            
            writer.writerow(row)
    


def main():
    root = tk.Tk()
    app = GAMSimplifiedGUI(root)
    root.mainloop()

if __name__ == "__main__":
    main()


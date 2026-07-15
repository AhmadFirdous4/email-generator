import os
import sys
import re
import pandas as pd
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

# ==========================================
# 1. NAVIGATION EXCEPTIONS (For CLI Mode)
# ==========================================
class GoToMainMenu(Exception): pass
class RestartMenu(Exception): pass

def prompt_input(prompt_text):
    """CLI input listener to handle instant menu routing."""
    val = input(prompt_text).strip()
    if val.lower() == 'm': raise GoToMainMenu()
    if val.lower() == 'r': raise RestartMenu()
    return val

# ==========================================
# 2. SHARED CORE ENGINE (Business Logic)
# ==========================================
def read_file_with_encoding(filepath):
    encodings = ['utf-8', 'windows-1252', 'latin-1']
    for enc in encodings:
        try:
            with open(filepath, 'r', encoding=enc) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    return None

def load_templates(script_dir, log_fn=print):
    templates_dir = os.path.join(script_dir, 'templates')
    if not os.path.exists(templates_dir):
        log_fn(f"⚠️  Warning: 'templates' folder not found at {templates_dir}")
        return {}
    
    templates = {}
    for filename in os.listdir(templates_dir):
        if filename.endswith(('.html', '.htm')):
            cat_name = filename.rsplit('.', 1)[0].replace('_', ' ')
            filepath = os.path.join(templates_dir, filename)
            content = read_file_with_encoding(filepath)
            if content is None:
                log_fn(f"✗ Error loading {filename}: Encoding failure.")
                continue
                
            # Generic HTML/formatting cleanup
            content = re.sub(r'\s+src=""', '', content)
            content = re.sub(r'mso-[^:]*:[^;]*;?', '', content)
            content = re.sub(r'<meta\s+http-equiv="Content-Type"[^>]*>', '', content, flags=re.IGNORECASE)
            content = re.sub(r'<meta\s+charset=["\']?[^"\'>\s]+["\']?>', '', content, flags=re.IGNORECASE)
            
            templates[cat_name] = content
    return templates

def get_display_name(group_df, company_name):
    if len(group_df) > 1:
        return f"{company_name} Team"
    row = group_df.iloc[0]
    if 'First Name' in group_df.columns and pd.notna(row['First Name']):
        raw_name = str(row['First Name']).strip()
        if raw_name: return raw_name
    if 'Employee Name' in group_df.columns and pd.notna(row['Employee Name']):
        raw_name = str(row['Employee Name']).strip().split()[0]
        if raw_name: return raw_name
    return company_name

def gather_unique_emails(group_df):
    raw_emails = []
    for _, row in group_df.iterrows():
        if 'Company Mail' in group_df.columns and pd.notna(row['Company Mail']):
            raw_emails.append(str(row['Company Mail']))
        if 'Employee Mail' in group_df.columns and pd.notna(row['Employee Mail']):
            raw_emails.append(str(row['Employee Mail']))
    combined = " ".join(raw_emails)
    return list(dict.fromkeys([e.strip() for e in re.split(r'[,\s]+', combined) if e.strip() and '@' in e]))

def process_batch(input_file, templates, output_dir, log_fn=print):
    """Core execution engine for parsing files and saving outputs."""
    if not os.path.exists(input_file):
        raise FileNotFoundError(f"File not found: {input_file}")
        
    if input_file.lower().endswith(('.xlsx', '.xls')):
        df_raw = pd.read_excel(input_file, header=None)
    else:
        df_raw = pd.read_csv(input_file, header=None)
        
    header_idx = 0
    found_header = False
    for idx, row in df_raw.head(15).iterrows():
        row_values_lower = row.astype(str).str.lower().tolist()
        if any('category' in val or 'company' in val for val in row_values_lower):
            header_idx = idx
            found_header = True
            break
            
    if found_header:
        log_fn(f"→ Smart Scan: Found headers on row {header_idx + 1}")
        df = pd.read_excel(input_file, skiprows=header_idx) if input_file.lower().endswith(('.xlsx', '.xls')) else pd.read_csv(input_file, skiprows=header_idx)
    else:
        df = df_raw.copy()
        df.columns = df.iloc[0]
        df = df[1:]

    df.columns = df.columns.astype(str).str.strip()
    column_mapping = {
        'company': 'Company', 'company name': 'Company',      
        'category': 'Category', 'first name': 'First Name',
        'employee name': 'Employee Name', 'person': 'First Name',          
        'company mail': 'Company Mail', 'company email': 'Company Mail',  
        'employee mail': 'Employee Mail', 'employee email': 'Employee Mail' 
    }
    rename_dict = {c: column_mapping[c.lower()] for c in df.columns if c.lower() in column_mapping}
    df = df.rename(columns=rename_dict)

    if 'Category' not in df.columns:
        raise KeyError(f"Missing required 'Category' column. Found: {df.columns.tolist()}")

    df['Company'] = df['Company'].fillna("Unknown Company").astype(str).str.strip().replace("", "Unknown Company")
    df['Category'] = df['Category'].fillna("Unknown").astype(str).str.strip()

    grouped_data = df.groupby(['Company', 'Category'])
    count, skipped = 0, 0

    for (company_name, cat), group in grouped_data:
        if cat not in templates:
            log_fn(f"⚠️  Skipping: Category template missing for '{cat}'")
            skipped += 1
            continue
            
        template = templates[cat]
        display_name = get_display_name(group, company_name)
        unique_emails = gather_unique_emails(group)
        to_block = f"<b>To:</b> {', '.join(unique_emails)}<br>" if unique_emails else ""
        
        personalized = template.replace('{To_Block}', to_block).replace('{First_Name}', display_name).replace('{Company}', company_name)
        personalized = re.sub(r'<!DOCTYPE[^>]*>', '', personalized, flags=re.IGNORECASE)
        personalized = '<!DOCTYPE html>\n<meta charset="UTF-8">\n' + personalized
        
        safe_company = "".join(c for c in company_name if c.isalnum() or c in (' ', '_')).rstrip()
        safe_cat = cat.replace(" ", "_")
        
        with open(os.path.join(output_dir, f"{safe_company}_{safe_cat}.html"), "w", encoding="utf-8") as f:
            f.write(personalized)
        count += 1
        
    return count, skipped

def process_manual(company_name, first_name, email_input, selected_cat, templates, output_dir):
    """Core execution engine for crafting a single manual layout entry."""
    if not selected_cat or selected_cat not in templates:
        raise ValueError("Invalid template selected.")
        
    display_name = first_name if first_name else f"{company_name} Team"
    if email_input:
        email_list = [e.strip() for e in re.split(r'[,;]+', email_input) if e.strip()]
        to_block = f"<b>To:</b> {', '.join(email_list)}<br>"
    else:
        to_block = ""
        
    template = templates[selected_cat]
    personalized = template.replace('{To_Block}', to_block).replace('{First_Name}', display_name).replace('{Company}', company_name)
    personalized = re.sub(r'<!DOCTYPE[^>]*>', '', personalized, flags=re.IGNORECASE)
    personalized = '<!DOCTYPE html>\n<meta charset="UTF-8">\n' + personalized
    
    safe_company = "".join(c for c in company_name if c.isalnum() or c in (' ', '_')).rstrip()
    safe_cat = selected_cat.replace(" ", "_")
    
    filename = os.path.join(output_dir, f"{safe_company}_{safe_cat}_Manual.html")
    with open(filename, "w", encoding="utf-8") as f:
        f.write(personalized)
    return filename

# ==========================================
# 3. INTERFACE A: TEXT MENU (CLI FRONTEND)
# ==========================================
def run_cli_batch(templates, output_dir, script_dir):
    while True:
        try:
            print("\n--- Batch File Entry ---")
            print("(Navigation: 'm' = Main Menu, 'r' = Restart this menu)")
            input_file = prompt_input("Enter the filename or full path: ")
            if not input_file:
                print("❌ No file name entered.")
                continue
            
            full_path = input_file if os.path.isabs(input_file) else os.path.join(script_dir, input_file)
            print(f"\nProcessing data from: {full_path}")
            
            count, skipped = process_batch(full_path, templates, output_dir, print)
            print(f"\n✅ Success! Created {count} formatted emails in: {output_dir}")
            if skipped > 0: print(f"⚠️  Skipped {skipped} rows due to unmapped templates.")
            break
        except RestartMenu:
            print("\n🔄 Restarting Batch Menu...")
        except Exception as e:
            print(f"\n❌ Error processing tracking: {e}")

def run_cli_manual(templates, output_dir):
    while True:
        try:
            print("\n--- Manual Email Entry ---")
            print("(Navigation: 'm' = Main Menu, 'r' = Restart this menu)")
            
            company = prompt_input("Enter Company Name: ") or "Unknown Company"
            first = prompt_input("Enter First Name (leave blank for Team): ")
            emails = prompt_input("Enter Recipient Email(s) separated by commas: ")
            
            print("\nAvailable Templates:")
            keys = list(templates.keys())
            for i, name in enumerate(keys, start=1): print(f"[{i}] {name}")
                
            while True:
                choice_input = prompt_input("\nSelect template by number: ")
                try:
                    choice = int(choice_input)
                    if 1 <= choice <= len(keys):
                        selected_cat = keys[choice - 1]
                        break
                    print("Invalid index selection.")
                except ValueError:
                    print("Please supply a valid integer index.")
            
            filepath = process_manual(company, first, emails, selected_cat, templates, output_dir)
            print(f"\n✅ Success! Manual template created at:\n→ {filepath}")
            break
        except RestartMenu:
            print("\n🔄 Restarting Manual Menu...")

def run_cli_mode(templates, output_dir, script_dir):
    while True:
        try:
            print("\n===============================")
            print(" EMAIL GENERATOR MENU (CLI Mode)")
            print("===============================")
            print("[1] Batch process from an Excel/CSV file")
            print("[2] Manual entry (Single Email)")
            print("[3] Quit Code")
            
            choice = input("\nEnter your choice (1, 2, or 3): ").strip()
            if choice == '1': run_cli_batch(templates, output_dir, script_dir)
            elif choice == '2': run_cli_manual(templates, output_dir)
            elif choice == '3':
                print("\nExiting. Goodbye!")
                break
            else:
                print("\n❌ Invalid choice.")
        except GoToMainMenu:
            print("\n🔙 Back tracking to Main Menu loop...")

# ==========================================
# 4. INTERFACE B: WINDOW FRAME (GUI FRONTEND)
# ==========================================
class UnifiedGUI:
    def __init__(self, root, templates, output_dir, script_dir):
        self.root = root
        self.templates = templates
        self.output_dir = output_dir
        self.script_dir = script_dir
        
        self.root.title("Automated Email Suite")
        self.root.geometry("1000x700")
        self.root.minsize(900, 600)
        
        self.setup_ui_styles()
        self.build_ui_layout()
        
    def setup_ui_styles(self):
        """Configure modern, clean color scheme"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Color palette
        bg_main = '#f8fafc'
        bg_card = '#ffffff'
        text_primary = '#1e293b'
        text_secondary = '#64748b'
        accent = '#1e40af'
        accent_light = '#3b82f6'
        border = '#e2e8f0'
        sidebar_bg = '#0f172a'
        sidebar_text = '#f1f5f9'
        
        style.configure('TFrame', background=bg_main)
        style.configure('Sidebar.TFrame', background=sidebar_bg)
        style.configure('Card.TFrame', background=bg_card, relief='flat', borderwidth=0)
        style.configure('StatusBar.TFrame', background=bg_card, relief='flat', borderwidth=1)
        
        style.configure('TLabel', background=bg_main, font=('Segoe UI', 10), foreground=text_primary)
        style.configure('Header.TLabel', background=bg_main, font=('Segoe UI', 16, 'bold'), foreground=accent)
        style.configure('Subheader.TLabel', background=bg_card, font=('Segoe UI', 11, 'bold'), foreground=text_primary)
        style.configure('Secondary.TLabel', background=bg_card, font=('Segoe UI', 9), foreground=text_secondary)
        style.configure('Status.TLabel', background=bg_card, font=('Segoe UI', 9), foreground=text_secondary)
        
        style.configure('Sidebar.TLabel', background=sidebar_bg, foreground=sidebar_text, font=('Segoe UI', 13, 'bold'))
        
        style.configure('TButton', font=('Segoe UI', 10), borderwidth=0, relief='flat')
        style.configure('TEntry', fieldbackground=bg_card, background=bg_card, foreground=text_primary, font=('Segoe UI', 10), borderwidth=1, relief='solid', padding=6)
        
        style.configure('Primary.TButton', font=('Segoe UI', 11, 'bold'), foreground='white', background=accent, relief='flat')
        style.map('Primary.TButton', background=[('active', accent_light), ('pressed', accent)])
        
        style.configure('Sidebar.TButton', font=('Segoe UI', 10, 'bold'), foreground='white', background='#334155', width=20, relief='flat', padding=10)
        style.map('Sidebar.TButton', background=[('active', accent), ('pressed', accent_light)])
        
        style.configure('Secondary.TButton', font=('Segoe UI', 10), foreground=text_primary, background='#e2e8f0', relief='flat')
        style.map('Secondary.TButton', background=[('active', '#cbd5e1'), ('pressed', '#d4dce8')])

    def build_ui_layout(self):
        """Build the complete GUI layout"""
        # ===== TOP STATUS BAR =====
        top_bar = ttk.Frame(self.root, style='StatusBar.TFrame')
        top_bar.pack(side='top', fill='x', padx=0, pady=0)
        
        # Brand area
        brand_frame = ttk.Frame(top_bar, style='StatusBar.TFrame')
        brand_frame.pack(side='left', padx=20, pady=12, fill='x', expand=True)
        
        ttk.Label(brand_frame, text="📧 Email Suite", font=('Segoe UI', 14, 'bold'), foreground='#1e40af', background='#ffffff').pack(anchor='w')
        ttk.Label(brand_frame, text=f"Templates: {len(self.templates)} loaded  •  Output: {os.path.basename(self.output_dir)}", 
                 style='Status.TLabel').pack(anchor='w', pady=(2, 0))
        
        # ===== MAIN CONTAINER =====
        main_container = ttk.Frame(self.root, style='TFrame')
        main_container.pack(fill='both', expand=True, padx=0, pady=0)
        
        # ===== SIDEBAR =====
        sidebar = ttk.Frame(main_container, style='Sidebar.TFrame', width=200)
        sidebar.pack(side='left', fill='y', padx=0, pady=0)
        sidebar.pack_propagate(False)
        
        ttk.Label(sidebar, text="TOOLS", style='Sidebar.TLabel').pack(pady=20, padx=15)
        
        ttk.Button(sidebar, text="📁 Batch Process", style='Sidebar.TButton', 
                  command=lambda: self.show_view(self.view_batch)).pack(pady=8, padx=12, fill='x')
        ttk.Button(sidebar, text="✏️  Manual Entry", style='Sidebar.TButton', 
                  command=lambda: self.show_view(self.view_manual)).pack(pady=8, padx=12, fill='x')
        
        ttk.Frame(sidebar, style='Sidebar.TFrame').pack(fill='both', expand=True)
        
        ttk.Button(sidebar, text="❌ Close", style='Sidebar.TButton', 
                  command=self.root.quit).pack(pady=12, padx=12, fill='x')

        # ===== WORKSPACE PANEL =====
        self.workspace = ttk.Frame(main_container, style='TFrame')
        self.workspace.pack(side='right', fill='both', expand=True, padx=0, pady=0)
        
        # Workspace content area
        content_frame = ttk.Frame(self.workspace, style='TFrame')
        content_frame.pack(fill='both', expand=True, padx=25, pady=25)
        
        # Views container
        self.views_container = ttk.Frame(content_frame, style='TFrame')
        self.views_container.pack(fill='both', expand=True)
        
        # Build views
        self.build_batch_view()
        self.build_manual_view()
        
        # ===== LOG AREA =====
        log_label_frame = ttk.Frame(self.workspace, style='TFrame')
        log_label_frame.pack(fill='x', padx=25, pady=(10, 5))
        ttk.Label(log_label_frame, text="Activity Log", font=('Segoe UI', 11, 'bold'), foreground='#1e293b').pack(anchor='w')
        
        log_frame = ttk.Frame(self.workspace, style='TFrame')
        log_frame.pack(fill='both', expand=True, padx=25, pady=(0, 25), side='bottom')
        
        self.log_box = scrolledtext.ScrolledText(log_frame, height=8, font=('Consolas', 9), 
                                                 bg='#1e293b', fg='#e2e8f0', insertbackground='white',
                                                 relief='solid', borderwidth=1)
        self.log_box.pack(fill='both', expand=True)
        self.log_box.config(state='disabled')
        
        # Init state
        self.show_view(self.view_batch)
        self.log(f"✅ Application ready  •  {len(self.templates)} templates loaded")

    def build_batch_view(self):
        """Batch processing view"""
        self.view_batch = ttk.Frame(self.views_container, style='TFrame')
        
        ttk.Label(self.view_batch, text="Batch Process", style='Header.TLabel').pack(anchor='w', pady=(0, 20))
        
        # File selection card
        card = ttk.Frame(self.view_batch, style='Card.TFrame', padding=20)
        card.pack(fill='x', pady=5)
        card.columnconfigure(1, weight=1)
        
        ttk.Label(card, text="Data File", style='Subheader.TLabel').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        
        ttk.Label(card, text="Select Excel or CSV:", font=('Segoe UI', 10), background='#ffffff', foreground='#64748b').grid(row=1, column=0, sticky='w', padx=(0, 15), pady=6)
        
        path_frame = ttk.Frame(card, style='Card.TFrame')
        path_frame.grid(row=1, column=1, sticky='ew', pady=6)
        
        self.entry_path = ttk.Entry(path_frame, font=('Segoe UI', 10))
        self.entry_path.pack(side='left', fill='x', expand=True, padx=(0, 10))
        ttk.Button(path_frame, text="Browse", style='Secondary.TButton', command=self.browse).pack(side='right')
        
        ttk.Button(card, text="Process File", style='Primary.TButton', command=self.run_batch_gui).grid(row=2, column=1, sticky='e', pady=(20, 0))

    def build_manual_view(self):
        """Manual entry view"""
        self.view_manual = ttk.Frame(self.views_container, style='TFrame')
        
        ttk.Label(self.view_manual, text="Manual Entry", style='Header.TLabel').pack(anchor='w', pady=(0, 20))
        
        # Manual entry card
        card = ttk.Frame(self.view_manual, style='Card.TFrame', padding=20)
        card.pack(fill='x', pady=5)
        card.columnconfigure(1, weight=1)
        
        # Company Name
        ttk.Label(card, text="Company Name", style='Subheader.TLabel').grid(row=0, column=0, columnspan=2, sticky='w', pady=(0, 10))
        ttk.Label(card, text="Company:", background='#ffffff', foreground='#64748b').grid(row=1, column=0, sticky='w', padx=(0, 15), pady=6)
        self.e_company = ttk.Entry(card, font=('Segoe UI', 10))
        self.e_company.grid(row=1, column=1, sticky='ew', pady=6)
        
        # First Name
        ttk.Label(card, text="Contact Name", style='Subheader.TLabel').grid(row=2, column=0, columnspan=2, sticky='w', pady=(15, 10))
        ttk.Label(card, text="First Name:", background='#ffffff', foreground='#64748b').grid(row=3, column=0, sticky='w', padx=(0, 15), pady=6)
        self.e_first = ttk.Entry(card, font=('Segoe UI', 10))
        self.e_first.grid(row=3, column=1, sticky='ew', pady=6)
        ttk.Label(card, text="Leave blank for 'Team'", style='Secondary.TLabel').grid(row=4, column=1, sticky='w', pady=(0, 0))
        
        # Emails
        ttk.Label(card, text="Recipients", style='Subheader.TLabel').grid(row=5, column=0, columnspan=2, sticky='w', pady=(15, 10))
        ttk.Label(card, text="Email(s):", background='#ffffff', foreground='#64748b').grid(row=6, column=0, sticky='w', padx=(0, 15), pady=6)
        self.e_emails = ttk.Entry(card, font=('Segoe UI', 10))
        self.e_emails.grid(row=6, column=1, sticky='ew', pady=6)
        ttk.Label(card, text="Separate multiple with commas", style='Secondary.TLabel').grid(row=7, column=1, sticky='w')
        
        # Template Selection
        ttk.Label(card, text="Template", style='Subheader.TLabel').grid(row=8, column=0, columnspan=2, sticky='w', pady=(15, 10))
        ttk.Label(card, text="Category:", background='#ffffff', foreground='#64748b').grid(row=9, column=0, sticky='w', padx=(0, 15), pady=6)
        self.combo = ttk.Combobox(card, values=list(self.templates.keys()), state="readonly", font=('Segoe UI', 10))
        self.combo.grid(row=9, column=1, sticky='ew', pady=6)
        if self.templates: 
            self.combo.current(0)
        
        # Button
        ttk.Button(card, text="Generate Email", style='Primary.TButton', command=self.run_manual_gui).grid(row=10, column=1, sticky='e', pady=(25, 0))

    def show_view(self, target):
        """Switch between batch and manual views"""
        self.view_batch.pack_forget()
        self.view_manual.pack_forget()
        target.pack(fill='both', expand=True)

    def log(self, text):
        """Append message to activity log"""
        self.log_box.config(state='normal')
        self.log_box.insert(tk.END, text + "\n")
        self.log_box.see(tk.END)
        self.log_box.config(state='disabled')

    def browse(self):
        """Open file browser"""
        p = filedialog.askopenfilename(
            initialdir=self.script_dir, 
            title="Select Excel or CSV file", 
            filetypes=(("Excel", "*.xlsx *.xls"), ("CSV", "*.csv"), ("All Files", "*.*"))
        )
        if p:
            self.entry_path.delete(0, tk.END)
            self.entry_path.insert(0, p)
            self.log(f"📂 Selected: {os.path.basename(p)}")

    def run_batch_gui(self):
        """Execute batch processing"""
        f = self.entry_path.get().strip()
        if not f or not os.path.exists(f):
            messagebox.showerror("Error", "Please select a valid data file first.")
            return
        try:
            self.log(f"\n🔄 Processing: {os.path.basename(f)}")
            count, skipped = process_batch(f, self.templates, self.output_dir, self.log)
            self.log(f"✅ Complete! Created {count} emails. Skipped: {skipped}")
            messagebox.showinfo("Success", f"✅ Batch processing complete!\n\nCreated: {count} emails\nSkipped: {skipped} rows\n\nOutput: {self.output_dir}")
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Processing failed:\n{str(e)}")

    def run_manual_gui(self):
        """Execute manual email generation"""
        company = self.e_company.get().strip() or "Unknown Company"
        first = self.e_first.get().strip()
        emails = self.e_emails.get().strip()
        template = self.combo.get()
        
        if not template:
            messagebox.showerror("Error", "Please select a template.")
            return
        
        try:
            self.log(f"\n🔄 Generating email for {company}")
            filename = process_manual(company, first, emails, template, self.templates, self.output_dir)
            self.log(f"✅ Created: {os.path.basename(filename)}")
            messagebox.showinfo("Success", f"✅ Email generated!\n\n{os.path.basename(filename)}")
            
            # Clear form
            self.e_company.delete(0, tk.END)
            self.e_first.delete(0, tk.END)
            self.e_emails.delete(0, tk.END)
        except Exception as e:
            self.log(f"❌ Error: {str(e)}")
            messagebox.showerror("Error", f"Generation failed:\n{str(e)}")

# ==========================================
# 5. INITIALIZATION DISTRIBUTION ROUTER
# ==========================================
def main():
    script_dir = os.path.dirname(os.path.abspath(__file__))
    output_dir = os.path.join(script_dir, 'Generated_Emails')
    if not os.path.exists(output_dir): 
        os.makedirs(output_dir)

    # Load templates
    templates = load_templates(script_dir, log_fn=lambda x: None)
    if not templates:
        print("❌ Error: No HTML templates found in 'templates' folder.")
        return

    # CLI or GUI mode
    if len(sys.argv) > 1 and sys.argv[1].lower() == '--cli':
        run_cli_mode(templates, output_dir, script_dir)
    else:
        root = tk.Tk()
        app = UnifiedGUI(root, templates, output_dir, script_dir)
        root.mainloop()

if __name__ == "__main__":
    main()
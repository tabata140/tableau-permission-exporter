import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import threading
from datetime import datetime
from tableau_client import TableauCloudClient


class LoginWindow:
    """Login Window"""
    
    def __init__(self, root, on_login_success):
        self.root = root
        self.on_login_success = on_login_success
        self.root.title("Tableau Permission Exporter - Login")
        self.root.geometry("400x250")
        
        self.frame = ttk.Frame(root, padding="20")
        self.frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.create_login_form()
    
    def create_login_form(self):
        # Clear existing widgets
        for widget in self.frame.winfo_children():
            widget.destroy()
        
        ttk.Label(self.frame, text="Site ID:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.site_name_entry = ttk.Entry(self.frame, width=30)
        self.site_name_entry.grid(row=0, column=1, pady=5)
        ttk.Label(self.frame, text="(Leave empty for default site)", font=('Arial', 8), foreground='gray').grid(row=1, column=1, sticky=tk.W)
        
        ttk.Label(self.frame, text="PAT Name:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.pat_name_entry = ttk.Entry(self.frame, width=30)
        self.pat_name_entry.grid(row=2, column=1, pady=5)
        
        ttk.Label(self.frame, text="PAT Secret:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.pat_secret_entry = ttk.Entry(self.frame, width=30, show="*")
        self.pat_secret_entry.grid(row=3, column=1, pady=5)
        
        ttk.Label(self.frame, text="Pod Name (e.g., 10ax):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.pod_name_entry = ttk.Entry(self.frame, width=30)
        self.pod_name_entry.grid(row=4, column=1, pady=5)
        self.pod_name_entry.insert(0, "10ax")
        
        self.login_button = ttk.Button(self.frame, text="Login", command=self.login)
        self.login_button.grid(row=5, column=0, columnspan=2, pady=20)
        
        self.status_label = ttk.Label(self.frame, text="", foreground="blue")
        self.status_label.grid(row=6, column=0, columnspan=2)
    
    def login(self):
        site_name = self.site_name_entry.get().strip()
        pat_name = self.pat_name_entry.get().strip()
        pat_secret = self.pat_secret_entry.get().strip()
        pod_name = self.pod_name_entry.get().strip()
        
        if not pat_name or not pat_secret or not pod_name:
            messagebox.showerror("Error", "Please enter PAT Name, PAT Secret, and Pod Name")
            return
        
        self.login_button.config(state='disabled')
        self.status_label.config(text="Authenticating...", foreground="blue")
        
        def authenticate():
            try:
                client = TableauCloudClient(pod_name=pod_name)
                client.sign_in(pat_name=pat_name, pat_secret=pat_secret, site_name=site_name)
                
                self.root.after(0, lambda: self.on_login_success(client))
                self.root.after(0, self.root.destroy)
            except Exception as e:
                error_msg = str(e)
                # Show simple error message
                self.root.after(0, lambda: messagebox.showerror("Authentication Failed", 
                    "Authentication failed. Please check:\n"
                    "- Site ID\n"
                    "- PAT Name\n"
                    "- PAT Secret\n"
                    "- Pod Name"))
                # Re-enable form for retry
                self.root.after(0, lambda: self.status_label.config(text="", foreground="red"))
                self.root.after(0, lambda: self.login_button.config(state='normal'))
        
        thread = threading.Thread(target=authenticate)
        thread.daemon = True
        thread.start()


class MainWindow:
    """Main Window"""
    
    def __init__(self, root, client):
        self.root = root
        self.client = client
        self.root.title("Tableau Permission Exporter")
        self.root.geometry("900x600")
        
        self.projects = []
        self.project_items = {}
        self.content_loaded = {}
        self.suppress_content_type_callback = False
        
        self.content_types = {
            'workbook': tk.BooleanVar(master=root, value=False),
            'datasource': tk.BooleanVar(master=root, value=False),
            'view': tk.BooleanVar(master=root, value=False),
            'flow': tk.BooleanVar(master=root, value=False),
        }
        
        # Export options
        self.include_group_members = tk.BooleanVar(master=root, value=False)
        
        for var in self.content_types.values():
            var.trace_add('write', self.on_content_type_changed)
        
        self.setup_ui()
        self.load_projects()
    
    def setup_ui(self):
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(2, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="Tableau Permission Exporter", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=2, pady=10)
        
        # Step 1: Content Type Selection
        step1_frame = ttk.LabelFrame(main_frame, text="Step 1: Select Content Types", padding="10")
        step1_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        desc_label = ttk.Label(step1_frame, text="Select which content types to display in the tree below:", 
                              font=('Arial', 9))
        desc_label.pack(anchor=tk.W, pady=(0, 5))
        
        types_frame = ttk.Frame(step1_frame)
        types_frame.pack(fill="x")
        
        ttk.Checkbutton(types_frame, text="Workbooks", variable=self.content_types['workbook']).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(types_frame, text="Data Sources", variable=self.content_types['datasource']).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(types_frame, text="Views (Sheets)", variable=self.content_types['view']).pack(side=tk.LEFT, padx=5)
        ttk.Checkbutton(types_frame, text="Flows", variable=self.content_types['flow']).pack(side=tk.LEFT, padx=5)
        
        # Step 2: Content Selection (Left) and Export (Right)
        step2_container = ttk.Frame(main_frame)
        step2_container.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E, tk.N, tk.S))
        step2_container.columnconfigure(0, weight=1)
        step2_container.rowconfigure(0, weight=1)
        
        # Left side: Content tree
        left_frame = ttk.LabelFrame(step2_container, text="Step 2: Select Content to Export", padding="10")
        left_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 5))
        
        # Status label showing selected types
        self.selected_types_label = ttk.Label(left_frame, text="Showing: (none selected)", 
                                              font=('Arial', 9), foreground='blue')
        self.selected_types_label.pack(anchor=tk.W, pady=(0, 5))
        
        help_label = ttk.Label(left_frame, text="Expand projects (click +) to load content", 
                              font=('Arial', 8), foreground='gray')
        help_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Treeview
        tree_frame = ttk.Frame(left_frame)
        tree_frame.pack(fill="both", expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame, orient="vertical")
        self.content_tree = ttk.Treeview(tree_frame, yscrollcommand=scrollbar.set, selectmode='none')
        scrollbar.config(command=self.content_tree.yview)
        
        self.content_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        self.content_tree["columns"] = ("checkbox",)
        self.content_tree.column("#0", width=500)
        self.content_tree.column("checkbox", width=60, anchor="center")
        self.content_tree.heading("#0", text="Project / Content Name")
        self.content_tree.heading("checkbox", text="Select")
        
        # Configure tag for selected items (light gray background)
        self.content_tree.tag_configure('selected', background='#E0E0E0')
        
        self.content_tree.bind('<Button-1>', self.on_tree_click)
        self.content_tree.bind('<<TreeviewOpen>>', self.on_tree_expand)
        
        # Right side: Buttons
        right_frame = ttk.Frame(step2_container)
        right_frame.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        step3_frame = ttk.LabelFrame(right_frame, text="Step 3: Export", padding="10")
        step3_frame.pack(fill="x", pady=(0, 10))
        
        # Export options
        ttk.Checkbutton(
            step3_frame, 
            text="Include group membership info",
            variable=self.include_group_members
        ).pack(anchor=tk.W, pady=(0, 5))
        
        self.export_button = ttk.Button(step3_frame, text="Export Permissions", command=self.export_permissions)
        self.export_button.pack(fill="x", pady=5)
        
        button_frame = ttk.LabelFrame(right_frame, text="Selection Tools", padding="10")
        button_frame.pack(fill="x")
        
        ttk.Button(button_frame, text="Select All", command=self.select_all_projects).pack(fill="x", pady=5)
        ttk.Button(button_frame, text="Deselect All", command=self.deselect_all_projects).pack(fill="x", pady=5)
        
        # Status
        self.status_label = ttk.Label(right_frame, text="", wraplength=300)
        self.status_label.pack(fill="x", pady=10)
        
        self.progress = ttk.Progressbar(right_frame, mode='indeterminate')
    
    def update_selected_types_label(self):
        """Update the label showing which content types are selected"""
        selected = [k.replace('_', ' ').title() for k, v in self.content_types.items() if v.get()]
        if selected:
            self.selected_types_label.config(text=f"Showing: {', '.join(selected)}")
        else:
            self.selected_types_label.config(text="Showing: (none selected - select types above)")
    
    def on_content_type_changed(self, *args):
        """Called when content type selection changes"""
        # Skip if callback is suppressed (during Deselect All)
        if self.suppress_content_type_callback:
            return
        
        self.update_selected_types_label()
        
        # Clear the loaded cache
        self.content_loaded.clear()
        
        # Check if any types are selected
        selected_types = [k for k, v in self.content_types.items() if v.get()]
        
        if not selected_types:
            # No types selected - close all projects
            def close_all_projects(parent=''):
                for item in self.content_tree.get_children(parent):
                    tags = self.content_tree.item(item, 'tags')
                    if tags and tags[0].startswith('project_'):
                        # Close this project and recurse
                        self.content_tree.item(item, open=False)
                        close_all_projects(item)
            close_all_projects()
            return
        
        # Reload content for all expanded projects
        def reload_expanded_projects(parent=''):
            for item in self.content_tree.get_children(parent):
                # Check if this item is expanded (open)
                is_open = self.content_tree.item(item, 'open')
                
                if is_open:
                    # This is an expanded item
                    tags = self.content_tree.item(item, 'tags')
                    if tags and tags[0].startswith('project_'):
                        proj_id = tags[0].replace('project_', '')
                        self.load_content_for_project(item, proj_id)
                
                # Recursively check children
                reload_expanded_projects(item)
        
        reload_expanded_projects()
    
    def get_all_tree_items(self, parent=''):
        """Get all tree items recursively"""
        items = []
        for item in self.content_tree.get_children(parent):
            items.append(item)
            items.extend(self.get_all_tree_items(item))
        return items
    
    def on_tree_click(self, event):
        """Handle tree item click for checkbox toggle"""
        region = self.content_tree.identify("region", event.x, event.y)
        if region == "cell":
            column = self.content_tree.identify_column(event.x)
            item = self.content_tree.identify_row(event.y)
            
            if column == "#1" and item:
                self.toggle_selection(item)
                # Prevent default behavior (expand/collapse) by returning "break"
                return "break"
    
    def on_tree_expand(self, event):
        """Handle tree item expansion"""
        item = self.content_tree.focus()
        tags = self.content_tree.item(item, 'tags')
        
        if tags and tags[0].startswith('project_'):
            proj_id = tags[0].replace('project_', '')
            
            if proj_id not in self.content_loaded:
                self.load_content_for_project(item, proj_id)
    
    def toggle_selection(self, item):
        """Toggle item selection - checkbox always visible"""
        current_value = self.content_tree.set(item, "checkbox")
        # Toggle between checked and unchecked (always show checkbox)
        if current_value == "☑":
            new_value = "☐"
        else:
            new_value = "☑"
        self.content_tree.set(item, "checkbox", new_value)
        
        # Update background color
        self._update_selection_tags(item, new_value)
        
        # Toggle all children recursively
        self._toggle_children(item, new_value)
    
    def _update_selection_tags(self, item, checkbox_value):
        """Update tags to show/hide background color"""
        current_tags = list(self.content_tree.item(item, 'tags'))
        
        if checkbox_value == "☑":
            # Add 'selected' tag if not present
            if 'selected' not in current_tags:
                current_tags.append('selected')
        else:
            # Remove 'selected' tag if present
            if 'selected' in current_tags:
                current_tags.remove('selected')
        
        self.content_tree.item(item, tags=current_tags)
    
    def _toggle_children(self, item, value):
        """Recursively toggle all children"""
        for child in self.content_tree.get_children(item):
            self.content_tree.set(child, "checkbox", value)
            self._update_selection_tags(child, value)
            self._toggle_children(child, value)
    
    def load_projects(self):
        self.status_label.config(text="Loading projects...")
        self.progress.pack(fill="x", pady=5)
        self.progress.start()
        
        def load():
            try:
                self.projects = self.client.get_projects()
                self.root.after(0, self.display_projects)
            except Exception as e:
                self.root.after(0, lambda: messagebox.showerror("Error", f"Failed to load projects: {str(e)}"))
            finally:
                self.root.after(0, self.progress.stop)
                self.root.after(0, lambda: self.progress.pack_forget())
                self.root.after(0, lambda: self.status_label.config(text=""))
        
        thread = threading.Thread(target=load)
        thread.daemon = True
        thread.start()
    
    def display_projects(self):
        """Display projects in tree view with empty checkboxes"""
        for item in self.content_tree.get_children():
            self.content_tree.delete(item)
        
        self.project_items.clear()
        
        project_by_id = {p['id']: p for p in self.projects}
        top_level_projects = [p for p in self.projects if not p.get('parentProjectId')]
        top_level_projects.sort(key=lambda x: x['name'])
        
        for project in top_level_projects:
            self._add_project_to_tree("", project, project_by_id)
    
    def _add_project_to_tree(self, parent, project, project_by_id):
        """Recursively add project and its children to tree"""
        proj_id = project['id']
        proj_name = project['name']
        
        # Insert with unchecked checkbox
        item = self.content_tree.insert(parent, "end", text=proj_name, values=("☐",), 
                                       tags=(f'project_{proj_id}',))
        self.project_items[proj_id] = item
        
        children = [p for p in self.projects if p.get('parentProjectId') == proj_id]
        children.sort(key=lambda x: x['name'])
        
        for child in children:
            self._add_project_to_tree(item, child, project_by_id)
        
        # Always add a dummy child to ensure + button appears
        # This will be removed when the project is expanded
        self.content_tree.insert(item, "end", text="", values=("",), tags=('dummy',))
    
    def load_content_for_project(self, item, proj_id):
        """Load content for a project when expanded"""
        selected_types = [k for k, v in self.content_types.items() if v.get()]
        
        # Remove dummy elements and existing content items (but keep sub-projects)
        for child in list(self.content_tree.get_children(item)):
            tags = self.content_tree.item(child, 'tags')
            # Keep sub-projects, remove everything else (dummy and content items)
            if not tags or tags[0] == 'dummy' or not tags[0].startswith('project_'):
                self.content_tree.delete(child)
        
        if not selected_types:
            # No types selected, just return
            self.content_loaded[proj_id] = True
            return
        
        self.content_loaded[proj_id] = True
        
        # Show loading indicator
        loading_item = self.content_tree.insert(item, 0, text="  Loading...", values=("",), tags=('loading',))
        
        def load():
            try:
                content_items = []
                
                # Note: 'project' type shows sub-projects which are already in the tree
                # We only need to load other content types
                
                if 'workbook' in selected_types:
                    try:
                        workbooks = self.client.get_workbooks(proj_id)
                        for wb in workbooks:
                            content_items.append(('workbook', wb['id'], wb['name']))
                    except Exception as wb_error:
                        print(f"Error loading workbooks for project {proj_id}: {wb_error}")
                        import traceback
                        traceback.print_exc()
                
                if 'datasource' in selected_types:
                    try:
                        datasources = self.client.get_datasources(proj_id)
                        for ds in datasources:
                            content_items.append(('datasource', ds['id'], ds['name']))
                    except Exception as ds_error:
                        print(f"Error loading datasources for project {proj_id}: {ds_error}")
                
                if 'flow' in selected_types:
                    try:
                        flows = self.client.get_flows(proj_id)
                        for flow in flows:
                            content_items.append(('flow', flow['id'], flow['name']))
                    except Exception as flow_error:
                        print(f"Error loading flows for project {proj_id}: {flow_error}")
                
                if 'view' in selected_types:
                    try:
                        workbooks = self.client.get_workbooks(proj_id)
                        for wb in workbooks:
                            wb_detail = self.client.get_workbook_by_id(wb['id'])
                            if wb_detail and 'views' in wb_detail:
                                for view in wb_detail['views']:
                                    content_items.append(('view', view['id'], f"{wb['name']} / {view['name']}", wb['name']))
                    except Exception as view_error:
                        print(f"Error loading views for project {proj_id}: {view_error}")
                
                # Remove loading indicator and add content
                def update_tree():
                    try:
                        self.content_tree.delete(loading_item)
                    except:
                        pass
                    self._add_content_to_tree(item, content_items)
                
                self.root.after(0, update_tree)
            except Exception as e:
                print(f"General error in load_content_for_project: {e}")
                def show_error():
                    try:
                        self.content_tree.delete(loading_item)
                    except:
                        pass
                    messagebox.showerror("Error", f"Failed to load content: {str(e)}")
                self.root.after(0, show_error)
        
        thread = threading.Thread(target=load)
        thread.daemon = True
        thread.start()
    
    def _add_content_to_tree(self, parent_item, content_items):
        """Add content items to tree under project (before sub-projects)"""
        # First, find all sub-projects (to keep them at the bottom)
        sub_project_items = []
        content_item_ids = []
        
        for child in list(self.content_tree.get_children(parent_item)):
            tags = self.content_tree.item(child, 'tags')
            if tags and tags[0].startswith('project_'):
                # This is a sub-project - keep it
                sub_project_items.append(child)
            else:
                # This is a content item or dummy - mark for deletion
                content_item_ids.append(child)
        
        # Remove old content items and dummies
        for item_id in content_item_ids:
            self.content_tree.delete(item_id)
        
        # If no content items found, show a message
        if not content_items:
            self.content_tree.insert(parent_item, 0, 
                                    text="  (No content of selected types in this project)", 
                                    values=("",),  # No checkbox for this message
                                    tags=('no_content',))
            return
        
        # Add new content items at the beginning (position 0, before sub-projects)
        # Insert in reverse order so they appear in the correct order
        for i, content_item in enumerate(reversed(content_items)):
            content_type = content_item[0]
            content_id = content_item[1]
            content_name = content_item[2]
            workbook_name = content_item[3] if len(content_item) > 3 else None
            
            tag = f'{content_type}_{content_id}'
            if workbook_name:
                tag += f'_wb_{workbook_name}'
            
            # Insert at position 0 (before sub-projects)
            self.content_tree.insert(parent_item, 0, text=f"  [{content_type.upper()}] {content_name}", 
                                    values=("☐",), tags=(tag,))
    
    def select_all_projects(self):
        """Select all items"""
        for item in self.content_tree.get_children():
            self._select_all_recursive(item)
    
    def _select_all_recursive(self, item):
        """Recursively select all items"""
        self.content_tree.set(item, "checkbox", "☑")
        self._update_selection_tags(item, "☑")
        for child in self.content_tree.get_children(item):
            self._select_all_recursive(child)
    
    def deselect_all_projects(self):
        """Deselect all items and clear content type selections"""
        # Set flag to prevent callback from reloading content
        self.suppress_content_type_callback = True
        
        # Clear Step 1 content type selections
        for var in self.content_types.values():
            var.set(False)
        
        # Update the label
        self.update_selected_types_label()
        
        # Clear content loaded cache
        self.content_loaded.clear()
        
        # Deselect all tree items and close all projects
        for item in self.content_tree.get_children():
            self._deselect_all_recursive(item)
            self._close_all_projects_recursive(item)
        
        # Re-enable callback
        self.suppress_content_type_callback = False
    
    def _close_all_projects_recursive(self, item):
        """Recursively close all projects"""
        # First, close all children
        for child in self.content_tree.get_children(item):
            self._close_all_projects_recursive(child)
        
        # Then close this item
        tags = self.content_tree.item(item, 'tags')
        if tags and tags[0].startswith('project_'):
            self.content_tree.item(item, open=False)
    
    def _clear_content_recursive(self, item):
        """Recursively clear content items but keep projects"""
        children_to_delete = []
        has_subprojects = False
        has_dummy = False
        
        for child in list(self.content_tree.get_children(item)):
            tags = self.content_tree.item(child, 'tags')
            if tags and tags[0].startswith('project_'):
                # This is a sub-project, keep it and recurse
                has_subprojects = True
                self._clear_content_recursive(child)
            elif tags and tags[0] == 'dummy':
                # Dummy already exists, keep it
                has_dummy = True
            else:
                # This is content or loading - mark for deletion
                children_to_delete.append(child)
        
        # Delete marked children
        for child in children_to_delete:
            self.content_tree.delete(child)
        
        # If no sub-projects exist and no dummy, add dummy to ensure + button appears
        if not has_subprojects and not has_dummy:
            self.content_tree.insert(item, "end", text="", values=("",), tags=('dummy',))
    
    def _deselect_all_recursive(self, item):
        """Recursively deselect all items"""
        self.content_tree.set(item, "checkbox", "☐")
        self._update_selection_tags(item, "☐")
        for child in self.content_tree.get_children(item):
            self._deselect_all_recursive(child)
    
    def get_selected_items(self):
        """Get list of selected items"""
        selected = {'projects': [], 'workbooks': [], 'datasources': [], 'views': [], 'flows': []}
        
        def check_item(item):
            if self.content_tree.set(item, "checkbox") == "☑":
                tags = self.content_tree.item(item, 'tags')
                if tags:
                    tag = tags[0]
                    if tag.startswith('project_'):
                        proj_id = tag.replace('project_', '')
                        selected['projects'].append(proj_id)
                    elif tag.startswith('workbook_'):
                        wb_id = tag.replace('workbook_', '')
                        # Get workbook name from tree text
                        wb_text = self.content_tree.item(item, 'text')
                        wb_name = wb_text.strip()
                        if wb_name.startswith('[WORKBOOK]'):
                            wb_name = wb_name[10:].strip()
                        selected['workbooks'].append({'id': wb_id, 'name': wb_name})
                    elif tag.startswith('datasource_'):
                        ds_id = tag.replace('datasource_', '')
                        # Get datasource name from tree text
                        ds_text = self.content_tree.item(item, 'text')
                        ds_name = ds_text.strip()
                        if ds_name.startswith('[DATASOURCE]'):
                            ds_name = ds_name[12:].strip()
                        selected['datasources'].append({'id': ds_id, 'name': ds_name})
                    elif tag.startswith('view_'):
                        parts = tag.split('_wb_')
                        view_id = parts[0].replace('view_', '')
                        workbook_name = parts[1] if len(parts) > 1 else ''
                        # Get view name from tree text (format: "  [VIEW] workbook / view_name")
                        view_text = self.content_tree.item(item, 'text')
                        # Remove "  [VIEW] " prefix more carefully
                        view_name = view_text.strip()
                        if view_name.startswith('[VIEW]'):
                            view_name = view_name[6:].strip()
                        selected['views'].append({'id': view_id, 'name': view_name, 'workbook_name': workbook_name})
                    elif tag.startswith('flow_'):
                        flow_id = tag.replace('flow_', '')
                        # Get flow name from tree text
                        flow_text = self.content_tree.item(item, 'text')
                        flow_name = flow_text.strip()
                        if flow_name.startswith('[FLOW]'):
                            flow_name = flow_name[6:].strip()
                        selected['flows'].append({'id': flow_id, 'name': flow_name})
            
            for child in self.content_tree.get_children(item):
                check_item(child)
        
        for item in self.content_tree.get_children():
            check_item(item)
        
        return selected
    
    def export_permissions(self):
        selected_items = self.get_selected_items()
        
        has_selection = any(selected_items[k] for k in selected_items)
        if not has_selection:
            messagebox.showwarning("Warning", "Please select at least one project or content item")
            return
        
        folder_path = filedialog.askdirectory(title="Select output folder")
        if not folder_path:
            return
        
        self.export_button.config(state='disabled')
        self.status_label.config(text="Exporting...")
        self.progress.pack(fill="x", pady=5)
        self.progress.start()
        
        def export():
            try:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                
                group_map = self.client.get_groups()
                user_map = self.client.get_users()
                
                total_files = 0
                
                # Export selected projects
                if selected_items['projects']:
                    all_project_perms = []
                    
                    for proj_id in selected_items['projects']:
                        try:
                            proj_name = next((p['name'] for p in self.projects if p['id'] == proj_id), 'Unknown')
                            
                            perms = self.client.get_project_permissions(proj_id)
                            
                            if not perms:
                                all_project_perms.append({
                                    'content_type': 'project',
                                    'content_id': proj_id,
                                    'content_name': proj_name,
                                    'project_name': proj_name,
                                    'asset_permissions': '',
                                    'grantee_type': '',
                                    'grantee_id': '',
                                    'grantee_name': 'No permissions set',
                                })
                            else:
                                for perm in perms:
                                    perm_data = {
                                        'content_type': 'project',
                                        'content_id': proj_id,
                                        'content_name': proj_name,
                                        'project_name': proj_name,
                                        'asset_permissions': perm.get('asset_permissions', ''),
                                        'grantee_type': perm['grantee_type'],
                                        'grantee_id': perm['grantee_id'],
                                        'grantee_name': group_map.get(perm['grantee_id'], '') if perm['grantee_type'] == 'group' else user_map.get(perm['grantee_id'], ''),
                                    }
                                    for cap in perm['capabilities']:
                                        perm_data[cap['name']] = cap['mode']
                                    all_project_perms.append(perm_data)
                        except Exception as proj_error:
                            # Log error but continue with other projects
                            print(f"WARNING: Failed to export project '{proj_name}': {str(proj_error)}")
                    
                    # Write all projects to a single CSV file
                    if all_project_perms:
                        filename = f"{folder_path}/project_permissions_{timestamp}.csv"
                        self.client.export_permissions_to_csv(all_project_perms, filename)
                        total_files += 1
                
                # Export selected workbooks
                if selected_items['workbooks']:
                    all_perms = []
                    for wb_info in selected_items['workbooks']:
                        wb_id = wb_info['id']
                        wb_name = wb_info.get('name', 'Unknown')
                        perms = self.client.get_workbook_permissions(wb_id)
                        
                        if not perms:
                            perm_data = {
                                'content_type': 'workbook',
                                'content_id': wb_id,
                                'content_name': wb_name,
                                'grantee_name': 'No permissions set',
                            }
                            all_perms.append(perm_data)
                        else:
                            for perm in perms:
                                perm_data = {
                                    'content_type': 'workbook',
                                    'content_id': wb_id,
                                    'content_name': wb_name,
                                    'grantee_type': perm['grantee_type'],
                                    'grantee_id': perm['grantee_id'],
                                    'grantee_name': group_map.get(perm['grantee_id'], '') if perm['grantee_type'] == 'group' else user_map.get(perm['grantee_id'], ''),
                                }
                                for cap in perm['capabilities']:
                                    perm_data[cap['name']] = cap['mode']
                                all_perms.append(perm_data)
                    
                    if all_perms:
                        filename = f"{folder_path}/workbook_permissions_{timestamp}.csv"
                        self.client.export_permissions_to_csv(all_perms, filename)
                        total_files += 1
                
                # Export selected datasources
                if selected_items['datasources']:
                    all_perms = []
                    for ds_info in selected_items['datasources']:
                        ds_id = ds_info['id']
                        ds_name = ds_info.get('name', 'Unknown')
                        perms = self.client.get_datasource_permissions(ds_id)
                        
                        if not perms:
                            perm_data = {
                                'content_type': 'datasource',
                                'content_id': ds_id,
                                'content_name': ds_name,
                                'grantee_name': 'No permissions set',
                            }
                            all_perms.append(perm_data)
                        else:
                            for perm in perms:
                                perm_data = {
                                    'content_type': 'datasource',
                                    'content_id': ds_id,
                                    'content_name': ds_name,
                                    'grantee_type': perm['grantee_type'],
                                    'grantee_id': perm['grantee_id'],
                                    'grantee_name': group_map.get(perm['grantee_id'], '') if perm['grantee_type'] == 'group' else user_map.get(perm['grantee_id'], ''),
                                }
                                for cap in perm['capabilities']:
                                    perm_data[cap['name']] = cap['mode']
                                all_perms.append(perm_data)
                    
                    if all_perms:
                        filename = f"{folder_path}/datasource_permissions_{timestamp}.csv"
                        self.client.export_permissions_to_csv(all_perms, filename)
                        total_files += 1
                
                # Export selected views
                if selected_items['views']:
                    all_perms = []
                    for view_info in selected_items['views']:
                        try:
                            view_id = view_info['id']
                            view_name = view_info.get('name', 'Unknown')
                            workbook_name = view_info['workbook_name']
                            perms = self.client.get_view_permissions(view_id)
                            
                            if not perms:
                                perm_data = {
                                    'content_type': 'view',
                                    'content_id': view_id,
                                    'content_name': view_name,
                                    'workbook_name': workbook_name,
                                    'grantee_name': 'No permissions set',
                                }
                                all_perms.append(perm_data)
                            else:
                                for perm in perms:
                                    perm_data = {
                                        'content_type': 'view',
                                        'content_id': view_id,
                                        'content_name': view_name,
                                        'workbook_name': workbook_name,
                                        'grantee_type': perm['grantee_type'],
                                        'grantee_id': perm['grantee_id'],
                                        'grantee_name': group_map.get(perm['grantee_id'], '') if perm['grantee_type'] == 'group' else user_map.get(perm['grantee_id'], ''),
                                    }
                                    for cap in perm['capabilities']:
                                        perm_data[cap['name']] = cap['mode']
                                    all_perms.append(perm_data)
                        except Exception as view_error:
                            print(f"WARNING: Failed to export view {view_info.get('id', 'unknown')}: {view_error}")
                    
                    if all_perms:
                        filename = f"{folder_path}/view_permissions_{timestamp}.csv"
                        self.client.export_permissions_to_csv(all_perms, filename)
                        total_files += 1
                
                # Export selected flows
                if selected_items['flows']:
                    all_perms = []
                    for flow_info in selected_items['flows']:
                        flow_id = flow_info['id']
                        flow_name = flow_info.get('name', 'Unknown')
                        perms = self.client.get_flow_permissions(flow_id)
                        
                        if not perms:
                            perm_data = {
                                'content_type': 'flow',
                                'content_id': flow_id,
                                'content_name': flow_name,
                                'grantee_name': 'No permissions set',
                            }
                            all_perms.append(perm_data)
                        else:
                            for perm in perms:
                                perm_data = {
                                    'content_type': 'flow',
                                    'content_id': flow_id,
                                    'content_name': flow_name,
                                    'grantee_type': perm['grantee_type'],
                                    'grantee_id': perm['grantee_id'],
                                    'grantee_name': group_map.get(perm['grantee_id'], '') if perm['grantee_type'] == 'group' else user_map.get(perm['grantee_id'], ''),
                                }
                                for cap in perm['capabilities']:
                                    perm_data[cap['name']] = cap['mode']
                                all_perms.append(perm_data)
                    
                    if all_perms:
                        filename = f"{folder_path}/flow_permissions_{timestamp}.csv"
                        self.client.export_permissions_to_csv(all_perms, filename)
                        total_files += 1
                
                # Export group membership if requested
                if self.include_group_members.get():
                    group_file = self.client.export_group_members(folder_path, timestamp)
                    if group_file:
                        total_files += 1
                
                self.root.after(0, lambda: messagebox.showinfo(
                    "Complete",
                    f"Exported {total_files} files\nLocation: {folder_path}"
                ))
                self.root.after(0, lambda: self.status_label.config(text="Export complete"))
            except Exception as e:
                error_msg = str(e)
                self.root.after(0, lambda: messagebox.showerror("Export Error", f"Export failed:\n{error_msg}"))
                self.root.after(0, lambda: self.status_label.config(text=""))
            finally:
                self.root.after(0, self.progress.stop)
                self.root.after(0, lambda: self.progress.pack_forget())
                self.root.after(0, lambda: self.export_button.config(state='normal'))
        
        thread = threading.Thread(target=export)
        thread.daemon = True
        thread.start()


def main():
    """Launch application"""
    def on_login_success(client):
        root = tk.Tk()
        app = MainWindow(root, client)
        root.mainloop()
    
    root = tk.Tk()
    login_window = LoginWindow(root, on_login_success)
    root.mainloop()


if __name__ == "__main__":
    main()

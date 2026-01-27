"""
Translator Manager GUI
Window for managing translators
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from pathlib import Path


class TranslatorManagerWindow:
    def __init__(self, parent, translator_manager, refresh_callback):
        self.parent = parent
        self.translator_manager = translator_manager
        self.refresh_callback = refresh_callback

        # Create window
        self.window = tk.Toplevel(parent)
        self.window.title("Translator Management")
        self.window.geometry("900x600")
        self.window.resizable(True, True)

        # Make window modal
        self.window.transient(parent)
        self.window.grab_set()

        # Current selection
        self.selected_translator = None
        self.signature_path = None

        # Setup UI
        self.setup_ui()

        # Load translators
        self.refresh_list()

        # Center window
        self.center_window()

    def center_window(self):
        """Center the window on screen"""
        self.window.update_idletasks()
        width = self.window.winfo_width()
        height = self.window.winfo_height()
        x = (self.window.winfo_screenwidth() // 2) - (width // 2)
        y = (self.window.winfo_screenheight() // 2) - (height // 2)
        self.window.geometry(f'{width}x{height}+{x}+{y}')

    def setup_ui(self):
        """Setup the user interface"""

        # Main container
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        self.window.columnconfigure(0, weight=1)
        self.window.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(1, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Translator Management",
            font=("Arial", 14, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))

        # Two-panel layout
        paned_window = ttk.PanedWindow(main_frame, orient=tk.HORIZONTAL)
        paned_window.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Left panel - Translator list
        left_frame = ttk.Frame(paned_window, padding="5")
        paned_window.add(left_frame, weight=1)

        # Right panel - Details/Edit
        right_frame = ttk.Frame(paned_window, padding="5")
        paned_window.add(right_frame, weight=2)

        self.create_list_panel(left_frame)
        self.create_detail_panel(right_frame)

        # Bottom buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=2, column=0, pady=(10, 0))

        close_btn = ttk.Button(
            button_frame,
            text="Close",
            command=self.close_window,
            width=15
        )
        close_btn.pack()

    def create_list_panel(self, parent):
        """Create translator list panel"""

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(1, weight=1)

        # Search frame
        search_frame = ttk.Frame(parent)
        search_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(0, weight=1)

        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search)

        search_entry = ttk.Entry(search_frame, textvariable=self.search_var)
        search_entry.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        ttk.Label(search_frame, text="🔍").grid(row=0, column=1)

        # Translator list
        list_frame = ttk.Frame(parent)
        list_frame.grid(row=1, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)

        # Treeview with scrollbar
        self.tree = ttk.Treeview(
            list_frame,
            columns=('name',),
            show='tree',
            selectmode='browse'
        )
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Bind selection event
        self.tree.bind('<<TreeviewSelect>>', self.on_select)

        # List buttons
        list_button_frame = ttk.Frame(parent)
        list_button_frame.grid(row=2, column=0, pady=(10, 0))

        ttk.Button(
            list_button_frame,
            text="New",
            command=self.new_translator,
            width=10
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            list_button_frame,
            text="Delete",
            command=self.delete_translator,
            width=10
        ).pack(side=tk.LEFT, padx=2)

        ttk.Button(
            list_button_frame,
            text="Refresh",
            command=self.refresh_list,
            width=10
        ).pack(side=tk.LEFT, padx=2)

    def create_detail_panel(self, parent):
        """Create translator detail/edit panel"""

        parent.columnconfigure(0, weight=1)
        parent.rowconfigure(0, weight=1)

        # Detail frame
        detail_frame = ttk.LabelFrame(parent, text="Translator Details", padding="10")
        detail_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        detail_frame.columnconfigure(1, weight=1)

        current_row = 0

        # Name
        ttk.Label(detail_frame, text="Name:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(detail_frame, textvariable=self.name_var, width=40)
        self.name_entry.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        current_row += 1

        # Language Pairs
        ttk.Label(detail_frame, text="Language Pairs:").grid(
            row=current_row, column=0, sticky=(tk.W, tk.N), pady=5, padx=(0, 10)
        )

        lang_frame = ttk.Frame(detail_frame)
        lang_frame.grid(row=current_row, column=1, sticky=(tk.W, tk.E, tk.N, tk.S), pady=5)
        lang_frame.columnconfigure(0, weight=1)
        lang_frame.rowconfigure(0, weight=1)

        self.lang_pairs_text = tk.Text(lang_frame, height=6, width=40, wrap=tk.WORD)
        self.lang_pairs_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        lang_scroll = ttk.Scrollbar(lang_frame, orient=tk.VERTICAL, command=self.lang_pairs_text.yview)
        lang_scroll.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.lang_pairs_text.configure(yscrollcommand=lang_scroll.set)

        ttk.Label(
            detail_frame,
            text="(One per line, e.g., 'Spanish > English')",
            font=("Arial", 8),
            foreground="gray"
        ).grid(row=current_row + 1, column=1, sticky=tk.W)
        current_row += 2

        # Signature
        ttk.Label(detail_frame, text="Signature:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )

        sig_frame = ttk.Frame(detail_frame)
        sig_frame.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        sig_frame.columnconfigure(0, weight=1)

        self.signature_label = ttk.Label(sig_frame, text="No signature file", foreground="gray")
        self.signature_label.grid(row=0, column=0, sticky=tk.W)

        ttk.Button(
            sig_frame,
            text="Browse...",
            command=self.browse_signature,
            width=12
        ).grid(row=0, column=1, padx=(5, 0))
        current_row += 1

        # Buttons
        button_frame = ttk.Frame(detail_frame)
        button_frame.grid(row=current_row, column=0, columnspan=2, pady=(20, 0))

        self.save_btn = ttk.Button(
            button_frame,
            text="Save",
            command=self.save_translator,
            width=15,
            state=tk.DISABLED
        )
        self.save_btn.pack(side=tk.LEFT, padx=5)

        self.cancel_btn = ttk.Button(
            button_frame,
            text="Cancel",
            command=self.cancel_edit,
            width=15,
            state=tk.DISABLED
        )
        self.cancel_btn.pack(side=tk.LEFT, padx=5)

        # Initially disable editing
        self.set_edit_mode(False)

    def refresh_list(self):
        """Refresh the translator list"""

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        # Get all translators
        translators = self.translator_manager.get_all_translators()

        # Sort by name
        translators.sort(key=lambda t: t['name'])

        # Add to tree
        for translator in translators:
            self.tree.insert(
                '',
                tk.END,
                text=translator['name'],
                values=(translator['name'],)
            )

    def on_search(self, *args):
        """Handle search input"""
        query = self.search_var.get().strip()

        # Clear tree
        for item in self.tree.get_children():
            self.tree.delete(item)

        if not query:
            # Show all if search is empty
            self.refresh_list()
            return

        # Search and display results
        results = self.translator_manager.search_translators(query)
        results.sort(key=lambda t: t['name'])

        for translator in results:
            self.tree.insert(
                '',
                tk.END,
                text=translator['name'],
                values=(translator['name'],)
            )

    def on_select(self, event):
        """Handle translator selection"""
        selection = self.tree.selection()

        if not selection:
            return

        item = self.tree.item(selection[0])
        translator_name = item['text']

        # Load translator details
        self.load_translator(translator_name)

    def load_translator(self, name):
        """Load translator details into form"""

        translator = self.translator_manager.get_translator(name)

        if not translator:
            return

        self.selected_translator = name

        # Populate fields
        self.name_var.set(translator['name'])

        # Language pairs
        self.lang_pairs_text.delete('1.0', tk.END)
        lang_pairs = translator.get('language_pairs', [])
        self.lang_pairs_text.insert('1.0', '\n'.join(lang_pairs))

        # Signature
        self.signature_path = translator.get('signature_path')
        if self.signature_path and Path(self.signature_path).exists():
            self.signature_label.config(
                text=Path(self.signature_path).name,
                foreground="black"
            )
        else:
            self.signature_label.config(text="No signature file", foreground="gray")
            self.signature_path = None

        # Enable editing
        self.set_edit_mode(True)

    def new_translator(self):
        """Create new translator"""

        self.selected_translator = None

        # Clear fields
        self.name_var.set("")
        self.lang_pairs_text.delete('1.0', tk.END)
        self.signature_path = None
        self.signature_label.config(text="No signature file", foreground="gray")

        # Enable editing
        self.set_edit_mode(True)

        # Focus name field
        self.name_entry.focus()

    def save_translator(self):
        """Save translator"""

        # Get values
        name = self.name_var.get().strip()
        lang_pairs_text = self.lang_pairs_text.get('1.0', tk.END).strip()

        # Validate
        if not name:
            messagebox.showerror("Validation Error", "Please enter a translator name")
            return

        # Parse language pairs
        lang_pairs = [line.strip() for line in lang_pairs_text.split('\n') if line.strip()]

        if not lang_pairs:
            messagebox.showerror("Validation Error", "Please enter at least one language pair")
            return

        # Save
        if self.selected_translator:
            # Update existing
            success, message = self.translator_manager.update_translator(
                self.selected_translator,
                name,
                lang_pairs,
                self.signature_path
            )
        else:
            # Add new
            success, message = self.translator_manager.add_translator(
                name,
                lang_pairs,
                self.signature_path
            )

        if success:
            messagebox.showinfo("Success", message)
            self.refresh_list()
            self.refresh_callback()  # Refresh main window dropdown
            self.cancel_edit()
        else:
            messagebox.showerror("Error", message)

    def delete_translator(self):
        """Delete selected translator"""

        selection = self.tree.selection()

        if not selection:
            messagebox.showwarning("No Selection", "Please select a translator to delete")
            return

        item = self.tree.item(selection[0])
        translator_name = item['text']

        # Confirm deletion
        result = messagebox.askyesno(
            "Confirm Deletion",
            f"Are you sure you want to delete translator '{translator_name}'?\n\n"
            "This action cannot be undone."
        )

        if not result:
            return

        # Delete
        success, message = self.translator_manager.delete_translator(translator_name)

        if success:
            messagebox.showinfo("Success", message)
            self.refresh_list()
            self.refresh_callback()  # Refresh main window dropdown
            self.cancel_edit()
        else:
            messagebox.showerror("Error", message)

    def cancel_edit(self):
        """Cancel editing"""

        self.selected_translator = None

        # Clear fields
        self.name_var.set("")
        self.lang_pairs_text.delete('1.0', tk.END)
        self.signature_path = None
        self.signature_label.config(text="No signature file", foreground="gray")

        # Disable editing
        self.set_edit_mode(False)

        # Clear selection
        self.tree.selection_remove(self.tree.selection())

    def browse_signature(self):
        """Browse for signature file"""

        file_path = filedialog.askopenfilename(
            title="Select Signature Image",
            filetypes=[
                ("Image Files", "*.png *.jpg *.jpeg *.gif *.bmp"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            self.signature_path = file_path
            self.signature_label.config(text=Path(file_path).name, foreground="black")

    def set_edit_mode(self, enabled):
        """Enable or disable edit mode"""

        state = tk.NORMAL if enabled else tk.DISABLED
        button_state = tk.NORMAL if enabled else tk.DISABLED

        self.name_entry.config(state=state)
        self.lang_pairs_text.config(state=state)
        self.save_btn.config(state=button_state)
        self.cancel_btn.config(state=button_state)

    def close_window(self):
        """Close the window"""
        self.window.destroy()

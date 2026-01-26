"""
Translation Processor Desktop Application
Main GUI Application
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import DND_FILES, TkinterDnD
from datetime import datetime
from pathlib import Path
import json
import sys

from translator_manager import TranslatorManager
from document_processor import DocumentProcessor


class TranslationProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Translation Processor - Park Evaluation Services")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        # Initialize managers
        self.translator_manager = TranslatorManager()
        self.document_processor = DocumentProcessor()

        # Current file
        self.selected_file = None

        # Setup UI
        self.setup_ui()

    def setup_ui(self):
        """Setup the main user interface"""

        # Main container with padding
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights for responsiveness
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)

        # Title
        title_label = ttk.Label(
            main_frame,
            text="Translation Processor",
            font=("Arial", 18, "bold")
        )
        title_label.grid(row=0, column=0, pady=(0, 10))

        # Subtitle
        subtitle_label = ttk.Label(
            main_frame,
            text="Park Evaluation Services",
            font=("Arial", 10)
        )
        subtitle_label.grid(row=1, column=0, pady=(0, 20))

        # File Selection Section
        self.create_file_section(main_frame, row=2)

        # Metadata Section
        self.create_metadata_section(main_frame, row=3)

        # Action Buttons
        self.create_action_buttons(main_frame, row=4)

        # Status Bar
        self.create_status_bar(main_frame, row=5)

    def create_file_section(self, parent, row):
        """Create file selection section with drag-drop support"""

        # File section frame
        file_frame = ttk.LabelFrame(parent, text="Document Selection", padding="10")
        file_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=10)
        file_frame.columnconfigure(0, weight=1)

        # Drag and drop area
        self.drop_area = tk.Label(
            file_frame,
            text="Drag & Drop Word Document Here\n\nOR",
            relief=tk.RIDGE,
            bg="#f0f0f0",
            font=("Arial", 12),
            height=5
        )
        self.drop_area.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))

        # Enable drag and drop
        self.drop_area.drop_target_register(DND_FILES)
        self.drop_area.dnd_bind('<<Drop>>', self.on_file_drop)

        # Browse button
        browse_btn = ttk.Button(
            file_frame,
            text="Browse for File...",
            command=self.browse_file
        )
        browse_btn.grid(row=1, column=0, columnspan=2, pady=(0, 10))

        # Selected file label
        self.file_label = ttk.Label(
            file_frame,
            text="No file selected",
            foreground="gray"
        )
        self.file_label.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E))

    def create_metadata_section(self, parent, row):
        """Create metadata input form"""

        # Metadata section frame
        metadata_frame = ttk.LabelFrame(parent, text="Case Information", padding="10")
        metadata_frame.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=10)
        metadata_frame.columnconfigure(1, weight=1)

        current_row = 0

        # Case Number
        ttk.Label(metadata_frame, text="Case Number:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )
        self.case_number_var = tk.StringVar()
        case_entry = ttk.Entry(metadata_frame, textvariable=self.case_number_var, width=30)
        case_entry.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        current_row += 1

        # Source Language
        ttk.Label(metadata_frame, text="Source Language:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )
        self.source_lang_var = tk.StringVar()
        source_combo = ttk.Combobox(
            metadata_frame,
            textvariable=self.source_lang_var,
            values=self.get_common_languages(),
            width=28
        )
        source_combo.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        current_row += 1

        # Target Language
        ttk.Label(metadata_frame, text="Target Language:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )
        self.target_lang_var = tk.StringVar(value="English")
        target_combo = ttk.Combobox(
            metadata_frame,
            textvariable=self.target_lang_var,
            values=self.get_common_languages(),
            width=28
        )
        target_combo.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        current_row += 1

        # Translator
        ttk.Label(metadata_frame, text="Translator:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )

        translator_frame = ttk.Frame(metadata_frame)
        translator_frame.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)
        translator_frame.columnconfigure(0, weight=1)

        self.translator_var = tk.StringVar()
        self.translator_combo = ttk.Combobox(
            translator_frame,
            textvariable=self.translator_var,
            values=self.translator_manager.get_translator_names(),
            width=23,
            state="readonly"
        )
        self.translator_combo.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 5))

        manage_btn = ttk.Button(
            translator_frame,
            text="Manage",
            command=self.open_translator_manager,
            width=10
        )
        manage_btn.grid(row=0, column=1)
        current_row += 1

        # Date
        ttk.Label(metadata_frame, text="Date:").grid(
            row=current_row, column=0, sticky=tk.W, pady=5, padx=(0, 10)
        )
        self.date_var = tk.StringVar(value=datetime.now().strftime("%B %d, %Y"))
        date_entry = ttk.Entry(metadata_frame, textvariable=self.date_var, width=30)
        date_entry.grid(row=current_row, column=1, sticky=(tk.W, tk.E), pady=5)

    def create_action_buttons(self, parent, row):
        """Create action buttons"""

        button_frame = ttk.Frame(parent)
        button_frame.grid(row=row, column=0, pady=20)

        # Process button
        self.process_btn = ttk.Button(
            button_frame,
            text="Process Translation",
            command=self.process_document,
            state=tk.DISABLED,
            width=20
        )
        self.process_btn.grid(row=0, column=0, padx=5)

        # Clear button
        clear_btn = ttk.Button(
            button_frame,
            text="Clear",
            command=self.clear_form,
            width=15
        )
        clear_btn.grid(row=0, column=1, padx=5)

    def create_status_bar(self, parent, row):
        """Create status bar"""

        self.status_var = tk.StringVar(value="Ready")
        status_bar = ttk.Label(
            parent,
            textvariable=self.status_var,
            relief=tk.SUNKEN,
            anchor=tk.W
        )
        status_bar.grid(row=row, column=0, sticky=(tk.W, tk.E), pady=(10, 0))

    def get_common_languages(self):
        """Return list of common languages"""
        return [
            "English", "Spanish", "French", "German", "Italian",
            "Portuguese", "Chinese", "Japanese", "Korean", "Arabic",
            "Russian", "Hindi", "Bengali", "Polish", "Dutch",
            "Turkish", "Vietnamese", "Thai", "Greek", "Hebrew"
        ]

    def on_file_drop(self, event):
        """Handle file drop event"""
        # Get file path (remove curly braces if present)
        file_path = event.data.strip('{}')
        self.load_file(file_path)

    def browse_file(self):
        """Open file browser dialog"""
        file_path = filedialog.askopenfilename(
            title="Select Word Document",
            filetypes=[
                ("Word Documents", "*.docx"),
                ("All Files", "*.*")
            ]
        )

        if file_path:
            self.load_file(file_path)

    def load_file(self, file_path):
        """Load selected file"""
        path = Path(file_path)

        # Validate file
        if not path.exists():
            messagebox.showerror("Error", "File does not exist")
            return

        if path.suffix.lower() != '.docx':
            messagebox.showerror("Error", "Please select a Word document (.docx)")
            return

        self.selected_file = path
        self.file_label.config(text=f"Selected: {path.name}", foreground="black")
        self.drop_area.config(bg="#d4edda", text=f"✓ {path.name}\n\nDrop another file to replace")
        self.process_btn.config(state=tk.NORMAL)
        self.update_status(f"Loaded: {path.name}")

    def open_translator_manager(self):
        """Open translator management window"""
        from translator_manager_gui import TranslatorManagerWindow
        TranslatorManagerWindow(self.root, self.translator_manager, self.refresh_translator_list)

    def refresh_translator_list(self):
        """Refresh the translator dropdown list"""
        self.translator_combo['values'] = self.translator_manager.get_translator_names()

    def validate_inputs(self):
        """Validate all required inputs"""
        errors = []

        if not self.selected_file:
            errors.append("Please select a Word document")

        if not self.case_number_var.get().strip():
            errors.append("Please enter a case number")

        if not self.source_lang_var.get().strip():
            errors.append("Please select a source language")

        if not self.target_lang_var.get().strip():
            errors.append("Please select a target language")

        if not self.translator_var.get().strip():
            errors.append("Please select a translator")

        if not self.date_var.get().strip():
            errors.append("Please enter a date")

        return errors

    def process_document(self):
        """Process the translation document"""

        # Validate inputs
        errors = self.validate_inputs()
        if errors:
            messagebox.showerror("Validation Error", "\n".join(errors))
            return

        # Confirm processing
        result = messagebox.askyesno(
            "Confirm Processing",
            f"Process translation for case #{self.case_number_var.get()}?"
        )

        if not result:
            return

        try:
            self.update_status("Processing document...")
            self.process_btn.config(state=tk.DISABLED)
            self.root.update()

            # Get translator info
            translator = self.translator_manager.get_translator(self.translator_var.get())

            # Prepare metadata
            metadata = {
                'case_number': self.case_number_var.get().strip(),
                'source_language': self.source_lang_var.get().strip(),
                'target_language': self.target_lang_var.get().strip(),
                'language_pair': f"{self.source_lang_var.get()} > {self.target_lang_var.get()} Translation",
                'translator_name': self.translator_var.get().strip(),
                'translator_signature': translator.get('signature_path', '') if translator else '',
                'date': self.date_var.get().strip(),
            }

            # Ask where to save output
            output_path = filedialog.asksaveasfilename(
                title="Save Processed PDF As",
                defaultextension=".pdf",
                initialfile=f"Park_Case_{metadata['case_number']}_Certified.pdf",
                filetypes=[("PDF Files", "*.pdf")]
            )

            if not output_path:
                self.update_status("Processing cancelled")
                self.process_btn.config(state=tk.NORMAL)
                return

            # Process the document
            success = self.document_processor.process_translation(
                self.selected_file,
                output_path,
                metadata
            )

            if success:
                messagebox.showinfo(
                    "Success",
                    f"Translation processed successfully!\n\nSaved to:\n{output_path}"
                )
                self.update_status("Processing complete")

                # Ask if user wants to open the file
                if messagebox.askyesno("Open File", "Would you like to open the PDF?"):
                    import os
                    os.startfile(output_path)

                # Clear form for next document
                self.clear_form()
            else:
                messagebox.showerror("Error", "Failed to process document. Check the console for details.")
                self.update_status("Processing failed")

        except Exception as e:
            messagebox.showerror("Error", f"An error occurred:\n{str(e)}")
            self.update_status("Error occurred")
            import traceback
            traceback.print_exc()
        finally:
            self.process_btn.config(state=tk.NORMAL)

    def clear_form(self):
        """Clear all form fields"""
        self.selected_file = None
        self.file_label.config(text="No file selected", foreground="gray")
        self.drop_area.config(
            bg="#f0f0f0",
            text="Drag & Drop Word Document Here\n\nOR"
        )
        self.case_number_var.set("")
        self.source_lang_var.set("")
        self.target_lang_var.set("English")
        self.translator_var.set("")
        self.date_var.set(datetime.now().strftime("%B %d, %Y"))
        self.process_btn.config(state=tk.DISABLED)
        self.update_status("Ready")

    def update_status(self, message):
        """Update status bar message"""
        self.status_var.set(message)
        self.root.update()


def main():
    """Main entry point"""
    try:
        root = TkinterDnD.Tk()
        app = TranslationProcessorApp(root)
        root.mainloop()
    except Exception as e:
        print(f"Error starting application: {e}")
        import traceback
        traceback.print_exc()
        input("Press Enter to exit...")


if __name__ == "__main__":
    main()

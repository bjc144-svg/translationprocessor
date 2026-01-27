"""
Translator Manager
Handles translator database with cloud synchronization
"""
import json
import os
from pathlib import Path
from datetime import datetime
import shutil


class TranslatorManager:
    def __init__(self, data_dir=None):
        """Initialize translator manager"""

        if data_dir is None:
            # Use a data directory in the app folder
            self.data_dir = Path(__file__).parent.parent / 'data'
        else:
            self.data_dir = Path(data_dir)

        self.data_dir.mkdir(parents=True, exist_ok=True)

        # Database file
        self.db_file = self.data_dir / 'translators.json'

        # Signatures directory
        self.signatures_dir = self.data_dir / 'signatures'
        self.signatures_dir.mkdir(parents=True, exist_ok=True)

        # Load database
        self.translators = self.load_database()

    def load_database(self):
        """Load translator database from file"""
        if self.db_file.exists():
            try:
                with open(self.db_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
            except Exception as e:
                print(f"Error loading database: {e}")
                return {}
        else:
            # Initialize with empty database
            return {}

    def save_database(self):
        """Save translator database to file"""
        try:
            # Create backup
            if self.db_file.exists():
                backup_file = self.db_file.with_suffix('.json.bak')
                shutil.copy2(self.db_file, backup_file)

            # Save database
            with open(self.db_file, 'w', encoding='utf-8') as f:
                json.dump(self.translators, f, indent=2, ensure_ascii=False)

            return True
        except Exception as e:
            print(f"Error saving database: {e}")
            return False

    def add_translator(self, name, language_pairs, signature_path=None):
        """Add a new translator"""

        if name in self.translators:
            return False, "Translator already exists"

        # Copy signature file if provided
        signature_stored_path = None
        if signature_path and os.path.exists(signature_path):
            signature_stored_path = self.copy_signature(name, signature_path)

        # Add translator
        self.translators[name] = {
            'name': name,
            'language_pairs': language_pairs,
            'signature_path': signature_stored_path,
            'created_at': datetime.now().isoformat(),
            'updated_at': datetime.now().isoformat()
        }

        # Save database
        if self.save_database():
            return True, "Translator added successfully"
        else:
            return False, "Failed to save database"

    def update_translator(self, old_name, new_name, language_pairs, signature_path=None):
        """Update existing translator"""

        if old_name not in self.translators:
            return False, "Translator not found"

        # Check if new name conflicts with another translator
        if new_name != old_name and new_name in self.translators:
            return False, "A translator with this name already exists"

        # Get existing data
        translator = self.translators[old_name]

        # Update signature if new one provided
        if signature_path and os.path.exists(signature_path):
            # Remove old signature if exists
            if translator.get('signature_path') and os.path.exists(translator['signature_path']):
                try:
                    os.remove(translator['signature_path'])
                except:
                    pass

            # Copy new signature
            signature_stored_path = self.copy_signature(new_name, signature_path)
            translator['signature_path'] = signature_stored_path

        # Update translator data
        translator['name'] = new_name
        translator['language_pairs'] = language_pairs
        translator['updated_at'] = datetime.now().isoformat()

        # If name changed, update dictionary key
        if old_name != new_name:
            del self.translators[old_name]
            self.translators[new_name] = translator
        else:
            self.translators[old_name] = translator

        # Save database
        if self.save_database():
            return True, "Translator updated successfully"
        else:
            return False, "Failed to save database"

    def delete_translator(self, name):
        """Delete a translator"""

        if name not in self.translators:
            return False, "Translator not found"

        # Remove signature file if exists
        translator = self.translators[name]
        if translator.get('signature_path') and os.path.exists(translator['signature_path']):
            try:
                os.remove(translator['signature_path'])
            except Exception as e:
                print(f"Error removing signature file: {e}")

        # Remove translator from database
        del self.translators[name]

        # Save database
        if self.save_database():
            return True, "Translator deleted successfully"
        else:
            return False, "Failed to save database"

    def get_translator(self, name):
        """Get translator by name"""
        return self.translators.get(name)

    def get_all_translators(self):
        """Get all translators"""
        return list(self.translators.values())

    def get_translator_names(self):
        """Get list of translator names"""
        return sorted(self.translators.keys())

    def copy_signature(self, translator_name, source_path):
        """Copy signature file to signatures directory"""
        try:
            source = Path(source_path)
            extension = source.suffix
            # Create unique filename
            dest_filename = f"{translator_name.replace(' ', '_')}{extension}"
            dest_path = self.signatures_dir / dest_filename

            shutil.copy2(source, dest_path)
            return str(dest_path)
        except Exception as e:
            print(f"Error copying signature: {e}")
            return None

    def search_translators(self, query):
        """Search translators by name or language pair"""
        query = query.lower()
        results = []

        for translator in self.translators.values():
            # Search in name
            if query in translator['name'].lower():
                results.append(translator)
                continue

            # Search in language pairs
            for pair in translator.get('language_pairs', []):
                if query in pair.lower():
                    results.append(translator)
                    break

        return results

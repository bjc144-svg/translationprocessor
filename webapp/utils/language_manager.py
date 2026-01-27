"""
Language Manager
Manages source and target languages for translation processing
"""
import json
from pathlib import Path


class LanguageManager:
    def __init__(self, data_dir):
        """
        Initialize language manager

        Args:
            data_dir: Directory to store language data
        """
        self.data_dir = Path(data_dir)
        self.data_dir.mkdir(parents=True, exist_ok=True)
        self.languages_file = self.data_dir / 'languages.json'

        # Initialize with default languages if file doesn't exist
        if not self.languages_file.exists():
            self._init_default_languages()

    def _init_default_languages(self):
        """Initialize with default common languages"""
        default_languages = [
            "English", "Spanish", "French", "German", "Italian",
            "Portuguese", "Chinese", "Japanese", "Korean", "Arabic",
            "Russian", "Hindi", "Bengali", "Polish", "Dutch",
            "Turkish", "Vietnamese", "Thai", "Greek", "Hebrew"
        ]

        data = {
            'languages': default_languages,
            'version': '1.0'
        }

        with open(self.languages_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)

    def get_languages(self):
        """
        Get list of available languages

        Returns:
            List of language names
        """
        try:
            with open(self.languages_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
                return sorted(data.get('languages', []))
        except Exception as e:
            print(f"Error loading languages: {e}")
            return []

    def add_language(self, language):
        """
        Add a new language

        Args:
            language: Language name to add

        Returns:
            tuple: (success: bool, message: str)
        """
        language = language.strip()

        if not language:
            return False, "Language name cannot be empty"

        try:
            data = self._load_data()
            languages = data.get('languages', [])

            # Check if language already exists (case-insensitive)
            if any(lang.lower() == language.lower() for lang in languages):
                return False, f"Language '{language}' already exists"

            # Add language
            languages.append(language)
            data['languages'] = languages

            # Save
            with open(self.languages_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            return True, f"Language '{language}' added successfully"

        except Exception as e:
            return False, f"Error adding language: {str(e)}"

    def delete_language(self, language):
        """
        Delete a language

        Args:
            language: Language name to delete

        Returns:
            tuple: (success: bool, message: str)
        """
        try:
            data = self._load_data()
            languages = data.get('languages', [])

            # Find and remove language (case-insensitive)
            found = False
            for i, lang in enumerate(languages):
                if lang.lower() == language.lower():
                    languages.pop(i)
                    found = True
                    break

            if not found:
                return False, f"Language '{language}' not found"

            data['languages'] = languages

            # Save
            with open(self.languages_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            return True, f"Language '{language}' deleted successfully"

        except Exception as e:
            return False, f"Error deleting language: {str(e)}"

    def _load_data(self):
        """Load language data from file"""
        try:
            with open(self.languages_file, 'r', encoding='utf-8') as f:
                return json.load(f)
        except:
            return {'languages': [], 'version': '1.0'}

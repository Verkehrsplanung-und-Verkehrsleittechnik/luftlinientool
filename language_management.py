import logging
from pathlib import Path
import pandas as pd

## @class Translator
# @brief Translation class to allow the user to choose the language
class Translator:
    ## @brief Initializes the translator with the specified language.
    # @param excel_path Path to the Excel file containing translations. Default: "Translations.xlsx"
    # @param language Language code to use for translations. Default: "en"
    def __init__(self, excel_path= Path.cwd() / "Translations.json", language: str="en"): #edit the translation excel path here

        self.translations = dict()
        self.excel_path = excel_path
        self.load_translations()

        self._selected_language = None

        self.update_selected_language(language)

    ## @brief Loads translations from the Excel file.
    def load_translations(self):
        if self.excel_path.suffix == ".json":
            df = pd.read_json(self.excel_path, orient="index")
        elif self.excel_path.suffix ==".xlsx":
            df = pd.read_excel(self.excel_path, header=0, index_col=0) # visum-python can't open excel files
        else:
            logging.error(f"Unsupported file format for translations: {self.excel_path}")
            return

        self.translations = df.to_dict()


    ## @brief Translates a key to the selected language.
    # @param key The key to translate.
    # @return The translated text or the original key if no translation is found.
    def translate(self, key):
        return self.translations.get(self._selected_language, {}).get(key, key)

    ## @brief Updates the selected language.
    # @param language The language code to set as the selected language.
    def update_selected_language(self, language):
        if language in self.translations:
            self._selected_language = language
        else:
            logging.error(f"Language {language} not found in translations.")

        # todo: add further adjustements here if necessary

    ## @brief Getter method for the selected language (optional, but good practice).
    # @return The currently selected language code.
    def get_selected_language(self):
        return self._selected_language

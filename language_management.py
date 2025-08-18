import logging

import pandas as pd

## @class Translator
# @brief Translation class to allow the user to choose the language
class Translator:
    ## @brief Initializes the translator with the specified language.
    # @param excel_path Path to the Excel file containing translations. Default: "Translations.xlsx"
    # @param language Language code to use for translations. Default: "en"
    def __init__(self, excel_path="Translations.xlsx", language: str="en"): #edit the translation excel path here
        self.excel_path = excel_path
        self.load_translations()

        self._selected_language = None

        self.update_selected_language(language)

    ## @brief Loads translations from the Excel file.
    def load_translations(self):
        df = pd.read_excel(self.excel_path, header=0, index_col=0)
        # languages = df.columns[1:]  # skip key column --> integrated in  import excel file
        self.translations = df.to_dict()
        # for _, row in df.iterrows(): # try to not to use iterrows whenever possible (worst pandas solution)
        #     key = row[0]
        #     for lang in languages:
        #         self.translations[lang][key] = row[lang]

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

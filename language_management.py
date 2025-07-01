import logging

import pandas as pd

# translation class to allow the user to choose the language
class Translator:
    def __init__(self, excel_path="Translations.xlsx", language: str="en"): #edit the translation excel path here
        self.excel_path = excel_path
        self.load_translations()

        self._selected_language = None

        self.update_selected_language(language)

    def load_translations(self):
        df = pd.read_excel(self.excel_path, header=0, index_col=0)
        # languages = df.columns[1:]  # skip key column --> integrated in  import excel file
        self.translations = df.to_dict()
        # for _, row in df.iterrows(): # try to not to use iterrows whenever possible (worst pandas solution)
        #     key = row[0]
        #     for lang in languages:
        #         self.translations[lang][key] = row[lang]

    def translate(self, key):
        return self.translations.get(self._selected_language, {}).get(key, key)

    def update_selected_language(self, language):
        if language in self.translations:
            self._selected_language = language
        else:
            logging.error(f"Language {language} not found in translations.")

        # todo: add further adjustements here if necessary

    # Getter-Methode für die ausgewählte Sprache (optional, aber gute Praxis)
    def get_selected_language(self):
        return self._selected_language


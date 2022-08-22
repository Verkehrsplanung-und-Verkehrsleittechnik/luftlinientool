import pandas as pd
import logging
import math
import sys
import numpy as np


# todo update nach Änderung sofort in GUI Event, hier nicht nochmaliges Update


class LuftlinienCalculator:

    ## Konstruktor
    def __init__(self, source,
                 attr_vfs: str = "TypeNo",
                 list_vfs: list = ["VFS 0", "VFS I", "VFS II", "VFS III", "VFS IV", "VFS V"],
                 use_gui: bool = False):

        self.attr_zones = ["No", "Name", "XCoord", "YCoord"]
        self.attr_vfs = attr_vfs

        # Liste der VFS, die bearbeitet werden sollen todo Preprocessing, das Liste nur diese VFS enthält
        self.vfs = list_vfs

        # todo Fallunterscheidungen Input
        # Ziel: dict mit VFS: Wert
        self.param_austauschfkt_vfs = dict(zip(list_vfs, 1 * np.ones(len(list_vfs), dtype=int)))
        self.param_austauschfkt_vfs = dict(zip(list_vfs, 1 * np.ones(len(list_vfs), dtype=int)))

        # Erhöhen des Rekursionslimit des python Interpreters
        self.recursion_limit = 6000

        # Jerknüpfung mit GUI: True/False
        self.use_gui = use_gui
        
        # Einlesen der Bezirksdaten
        if not isinstance(source, str):
            self.zones = pd.DataFrame(source.Net.Zones.GetMultipleAttributes(self.attr_zones), columns=self.attr_zones)
            logging.info("%s Bezirke eingelesen", len(self.zones))
        else:
            logging.warning("Einlesen der Bezirksdaten ist fehlgeschlagen, Inputformat ist nicht implementiert")

        # Init VFS Matrizen
        # Dict mit Matrix je VFS: Anzahl Bezirke x Anzahl Bezirke
        self.matrizen_VFS = self.init_results()


    def calculate_main(self):
        # Schleife über alle vfs
        for vfs in self.vfs:
            self.calculate_vfs(vfs)

            # todo Idee Aktivierung Outputexportbuttions in GUI
            if self.use_gui:
                a=1
            

    # entspricht Funktion Program.LLCalc
    def calculate_vfs(self, vfs, austauschfkt, versorgungsfkt):
        a = 1
        # Test, ob Bezirke mit gleichen Koordinaten existieren --> Abbruch

        # Filtere Bezirksdaten, die die Bedingungen erfüllen
        # Sind Aktiv
        # TypNr < VFS

    def init_results(self):
        dict_vfs = {}
        for vfs in self.vfs:
            dict_vfs[vfs] = np.zeros(len(self.zones))

        self.matrizen_VFS = dict_vfs








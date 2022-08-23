import pandas as pd
import logging
import math
import sys
import numpy as np
from scipy.spatial import Delaunay



# todo update nach Änderung sofort in GUI Event, hier nicht nochmaliges Update
["VFS 0", "VFS I", "VFS II", "VFS III", "VFS IV", "VFS V"]


class LuftlinienCalculator:

    ## Konstruktor
    def __init__(self, source,
                 attr_vfs: str = "TypeNo",
                 dict_vfs: dict = {"VFS 0": 1, "VFS I": 2, "VFS II": 3, "VFS III": 4, "VFS IV": 5, "VFS V": 6},
                 use_gui: bool = False):

        self.attr_zones = ["No", "Name", "XCoord", "YCoord"]
        self.attr_central_level = attr_vfs
        self.attr_zones.append(self.attr_central_level)

        # Liste der VFS, die bearbeitet werden sollen todo Preprocessing, das Liste nur diese VFS enthält
        self.vfs = dict_vfs

        # todo Fallunterscheidungen Input
        # Ziel: dict mit VFS: Wert
        self.param_austauschfkt_vfs = dict(zip(dict_vfs.keys(), 1 * np.ones(len(dict_vfs), dtype=int)))
        self.param_versorgungsfkt_vfs = dict(zip(dict_vfs, 1 * np.ones(len(dict_vfs), dtype=int)))

        # Erhöhen des Rekursionslimit des python Interpreters
        self.recursion_limit = 6000

        # Jerknüpfung mit GUI: True/False
        self.use_gui = use_gui
        
        # Einlesen der Bezirksdaten
        # Wichtig: Index der Tabelle = 0...n
        if not isinstance(source, str):
            self.zones = pd.DataFrame(source.Net.Zones.GetMultipleAttributes(self.attr_zones, OnlyActive=False), columns=self.attr_zones)
            logging.info("%s Bezirke eingelesen", len(self.zones))
        else:
            logging.warning("Einlesen der Bezirksdaten ist fehlgeschlagen, Inputformat ist nicht implementiert")

        # Init VFS Matrizen
        # Dict mit Matrix je VFS: Anzahl Bezirke x Anzahl Bezirke
        self.init_results()


    def calculate_main(self):
        # Init Ergebnisse
        logging.info(f"Berechnung über alle VFS wird gestartet")
        self.init_results()
        logging.info(f"Adjazenzmatrizen wurden initialisiert")

        # Schleife über alle vfs
        for vfs in self.vfs:
            self.calculate_vfs(vfs)

            # todo Idee Aktivierung Outputexportbuttions in GUI
            if self.use_gui:
                a=1
            

    # entspricht Funktion Program.LLCalc
    def calculate_vfs(self, vfs):
        a = 1
        value_vfs = self.vfs[vfs]

        austauschfkt = self.param_austauschfkt_vfs[vfs]
        versorgungsfkt = self.param_versorgungsfkt_vfs[vfs]

        # todo Test, ob Bezirke mit gleichen Koordinaten existieren --> Abbruch

        # Filtere Bezirksdaten, die die Bedingungen erfüllen
        # Sind Aktiv todo Erweiterung Filterung nach attr_filter
        # TypNr <= VFS
        active_zones = self.zones
        active_zones = active_zones.loc[active_zones[self.attr_central_level] <= value_vfs, :]
        
        logging.info(f"{vfs}: Delauney Triangulation wird für {len(active_zones)} Bezirke durchgeführt")

        # Delaunay Triangulation
        tri = Delaunay(active_zones[["XCoord", "YCoord"]])
        zone_orig_idx_triangles = active_zones.index.values[tri.simplices]
        logging.info(f"{vfs}: es wurden {len(zone_orig_idx_triangles)} Dreiecke gebildet")

        # Adjazenzmatrix ausfüllen

        # Schleife über Dreiecke
        for p1, p2, p3 in zone_orig_idx_triangles:
            # die drei Punkte des Dreiecks
            # folgende Abhängigkeiten sind einzufügen:
            # p1 - p2, p2 - p1, p1 - p3, p3 - p1, p3 - p2, p2 - p3
            self.matrizen_VFS[vfs][p1, p2] = 1
            self.matrizen_VFS[vfs][p1, p3] = 1
            self.matrizen_VFS[vfs][p2, p1] = 1
            self.matrizen_VFS[vfs][p2, p3] = 1
            self.matrizen_VFS[vfs][p3, p1] = 1
            self.matrizen_VFS[vfs][p3, p2] = 1

        logging.info(f"{vfs}: Die Ergebnisse wurden als Adjazenzmatrix festgehalten")

        # Nachbarschaften Grad n bestimmen
        if austauschfkt > 0:
            logging.info(f"{vfs}: Der Nachbarschaftsgrad muss berechnet werden")


        # Versorgungsfunktion

    def init_results(self):
        dict_vfs = {}
        for vfs in self.vfs:
            dict_vfs[vfs] = np.zeros([len(self.zones), len(self.zones)], dtype=bool)

        self.matrizen_VFS = dict_vfs

    def export_matrix(self):
        todo=1

    def export_net(self):
        todo=1








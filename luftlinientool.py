## @package luftlinientool
# todo Beschreibung
# @brief todo Kurzbeschreibung

import pandas as pd
import logging
import numpy as np
from scipy.spatial import Delaunay
from pathlib import Path
from math import sin, cos, sqrt, atan2, radians


# todo update nach Änderung sofort in GUI Event, hier nicht nochmaliges Update

# ====== allgemeine, nützliche FUnktionen =====

## writes data of defined object type to .net file
# @param[in] f: target file opened in "write" or "append" mode
# @param[in]:
def write_object_to_net(object, df_object_attributes_to_write, file):
    header = ["*", "*"]
    header.insert(1, "* Table: " + object + "s")
    header.append(
        ("$" + object.upper().replace(" ", "") + ":" + ";".join(df_object_attributes_to_write.columns)).upper() + "\n")

    file.write("\n".join(header))
    df_object_attributes_to_write.to_csv(file, header=False, sep=";", index=False)


## check if a matrix is symmetric
#     @param matrix:
#     @param tol:
#     @return:
def is_symmetric(matrix, tol=1e-8):
    # Anwendung der Maximums-Norm für die Diff zwischen der Matrix und der Transponierten
    # Norm > 0 -> keine Symmetrie
    return np.linalg.norm(matrix.astype(int) - matrix.T.astype(int), np.Inf) < tol


## Berechnung der Distanz zwischen Koordinaten
# Implementation der Haversine Formel
# @param x1:
# @param y1:
# @param vec_x2:
# @param vec_y2:
# @return:
def calculate_distance_coordinates(x1, y1, vec_x2, vec_y2):
    # approximate radius of earth in km
    R = 6373.0

    lat1 = radians(y1)
    lon1 = radians(x1)
    vec_lat2 = np.radians(vec_y2)
    vec_lon2 = np.radians(vec_x2)

    # todo Fallunterscheidung für negative Koordinaten
    diff_lon = vec_lon2 - lon1
    diff_lat = vec_lat2 - lat1

    # Haversine formula
    tmp = np.sin(diff_lat / 2) ** 2 + np.cos(lat1) * np.cos(vec_lat2) * np.sin(diff_lon / 2) ** 2
    distances_km = R * 2 * np.arcsin(np.sqrt(tmp))

    return distances_km


## identifiziert die nächsten n Punkte aus einer gegebenen Tabelle
def get_nearest_points_from_set(x_point, y_point, array_points, n=None):
    if (n is not None) and (n >= len(array_points)):
        # es werden alle möglichen Punkte zurückgegeben
        return list(range(0, n))

    # Berechne Entfernungen
    distances = calculate_distance_coordinates(x1=x_point, y1=y_point,
                                               vec_x2=array_points[:, 0], vec_y2=array_points[:, 1])

    # Index der n niedrigsten Werte
    list_indizes = np.argpartition(distances, n)[:n]

    return list_indizes


# ===== Klassendefinition ======
## Klasse LuftlinienCalculator
class LuftlinienCalculator:

    ## Konstruktor
    # @param source: Dateiname (str) oder Visuminstanz
    # @param attr_vfs: Name des Bezirkattributs, das die Kategorisierung in OZ,MZ,UZ ... enthält
    # @param dict_vfs: Dictionary, das die Attributwerte für die jeweiligen VFS enthält
    # @param max_entfernung: Angabe, bis zu welcher Entfernung, Nachbar angebunden werden
    # @param anz_versorger: Angabe, an wie viele höherrangige Zentren ein Bezirk angebunden werden soll
    # @param attr_quelle: Name des Attributs, das angibt, ob der Bezirk als Quelle berücksichtigt wird
    # @param attr_ziel:
    # @param use_gui:
    # @param path_output:
    def __init__(self, source,
                 attr_vfs: str = "TypeNo",
                 dict_vfs: dict = {"VFS 0": 0, "VFS I": 1, "VFS II": 2, "VFS III": 3, "VFS IV": 4, "VFS V": 5},
                 max_entfernung=1,
                 anz_versorger=0,
                 attr_quelle=None,
                 attr_ziel=None,
                 use_gui: bool = False,
                 path_output=None):

        self.attr_zones = ["No", "Name", "XCoord", "YCoord"]
        self.attr_central_level = attr_vfs
        self.attr_zones.append(self.attr_central_level)

        self.debug_mode = False

        # Pfade
        self.path_output = path_output

        if attr_quelle is not None:
            self.attr_zones.append(attr_quelle)
        if attr_ziel is not None:
            self.attr_zones.append(attr_ziel)

        self.attr_is_from_zone = attr_quelle
        self.attr_is_to_zone = attr_ziel

        # Liste der VFS, die bearbeitet werden sollen todo Preprocessing, das Liste nur diese VFS enthält
        self.vfs = dict_vfs

        # gibt an, bis zu welchem "Nachbarschaftsgrad" gleichrangige verbindungen verfolgt werden sollen
        # (ehemals Austauschfkt)
        # Ziel: dict mit VFS: Wert
        if isinstance(max_entfernung, int):
            # Umwandlung in dict mit Skalar je VFS
            self.nachbarschaftsgrad_vfs = dict(zip(dict_vfs.keys(), max_entfernung * np.ones(len(dict_vfs), dtype=int)))
        elif isinstance(max_entfernung, dict):
            self.nachbarschaftsgrad_vfs = max_entfernung
        else:
            raise TypeError("Type Übergabeparameter nicht implementiert")

        # gibt an, wie viele (höherrangige) Versorger verbunden werden sollen
        # Ziel: dict mit VFS: Wert
        if isinstance(anz_versorger, int):
            # Umwandlung in dict mit Skalar je VFS
            self.anz_versorger_vfs = dict(
                zip(dict_vfs.keys(), anz_versorger * np.ones(len(dict_vfs), dtype=int)))
        elif isinstance(anz_versorger, dict):
            self.anz_versorger_vfs = anz_versorger
        else:
            raise TypeError("Type Übergabeparameter nicht implementiert")

        # Verknüpfung mit GUI: True/False
        self.use_gui = use_gui

        # Einlesen der Bezirksdaten
        # Wichtig: Index der Tabelle = 0...n
        if not isinstance(source, str):
            self.visum = source
            self.zones = pd.DataFrame(source.Net.Zones.GetMultipleAttributes(self.attr_zones, OnlyActive=False),
                                      columns=self.attr_zones)
            logging.info("%s Bezirke eingelesen", len(self.zones))
        else:
            self.visum = None
            logging.warning("Einlesen der Bezirksdaten ist fehlgeschlagen, Inputformat ist nicht implementiert")

        # Init VFS Matrizen
        # Dict mit Matrix je VFS: Anzahl Bezirke x Anzahl Bezirke
        self.init_results()

    def adj_matrix_to_links(self, list_vfs=None):
        if list_vfs is None:
            list_vfs = self.vfs.keys()

        list_df_edges = []
        for vfs in list_vfs:
            if not is_symmetric(self.matrizen_VFS[vfs]):
                logging.warning(f"{vfs}: Adjazenzmatrix ist nicht symmetrisch")

            df_edges = pd.DataFrame(self.matrizen_VFS[vfs]).stack().reset_index()
            df_edges.columns = ["FromNodeNo", "ToNodeNo", "TypeNo"]

            # Filtere Strecken mit True
            df_edges = df_edges.loc[df_edges["TypeNo"] == True, :]

            # setze Attribut VFS
            df_edges.loc[:, "TypeNo"] = vfs

            list_df_edges.append(df_edges)

        df_edges = pd.concat(list_df_edges)

        return df_edges

    def adj_matrix_to_set_of_connected_zones(self, vfs, use_zone_names=True):
        if use_zone_names:
            df = pd.DataFrame(self.matrizen_VFS[vfs], index=self.zones["Name"], columns=self.zones["Name"])
        else:
            df = pd.DataFrame(self.matrizen_VFS[vfs])
        df_set_zones = df.mul(df.columns.tolist()).apply(lambda x: set(zone for zone in x if zone), axis=1).to_frame(
            name="set zones")
        df_set_zones["no zones"] = df_set_zones["set zones"].apply(len)

        return df_set_zones

    def calculate_reachability_max_steps(self, max_steps, vfs):
        matrix = np.linalg.matrix_power(self.matrizen_VFS[vfs], max_steps)
        np.fill_diagonal(matrix, 0)

        return matrix

    ## a
    #
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
                a = 1

    ## entspricht Funktion Program.LLCalc
    def calculate_vfs(self, vfs):

        # Attributswert der Bezirke für die gewählte VFS
        value_vfs = self.vfs[vfs]

        k_nachbar = self.nachbarschaftsgrad_vfs[vfs]
        anz_versorger = self.anz_versorger_vfs[vfs]

        # todo Test, ob Bezirke mit gleichen Koordinaten existieren --> Abbruch

        # Filtere Bezirksdaten, die die Bedingungen erfüllen
        # Sind Aktiv todo Erweiterung Filterung nach attr_filter
        # TypNr <= VFS
        active_zones = self.zones
        active_zones = active_zones.loc[active_zones[self.attr_central_level] <= value_vfs,
                       :]

        if len(active_zones) < 3:
            logging.info(f"{vfs}: es sind zu wenige Bezirke aktiv")
        else:
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

            logging.info(
                f"{vfs}: die initiale Adjazenzmatrix ohne Berücksichtigung der zusätzlichen Bedingungen, wurde erstellt")

            # Nachbarschaften Grad n bestimmen
            if k_nachbar > 1:
                logging.info(f"{vfs}: der Nachbarschaftsgrad muss berechnet werden")
                adj_k_steps = self.calculate_reachability_max_steps(k_nachbar, vfs)
                self.matrizen_VFS[vfs] = adj_k_steps

            # Verbindungen mit Versorgungsfunktion
            if anz_versorger > 0:
                # Erstelle für jeden Bezirk eine Liste der verbundenen Bezirke
                df_list_zones = self.adj_matrix_to_set_of_connected_zones(vfs, use_zone_names=False)

                # Tabelle der möglichen Versorgungszentren
                provider = self.zones.loc[(self.zones[self.attr_central_level] < self.vfs[vfs])
                                          & (self.zones[self.attr_is_from_zone] > 0), :]
                # Menge der möglichen Versorgungszentren
                set_names_provider = set(provider.index)

                # Bestimme für jeden aktiven Bezirk, ob dieser bereits an ein Versorgungszentrum angeschlossen ist
                df_list_zones = df_list_zones.loc[df_list_zones.index.isin(
                    active_zones.loc[active_zones[self.attr_is_from_zone] > 0, :].index), :]
                df_list_zones["no_provider"] = df_list_zones["set zones"].apply(set_names_provider.intersection).apply(len)
                df_list_zones["provider"] = (df_list_zones.index.isin(set_names_provider)) \
                                            | (df_list_zones["no_provider"] >= anz_versorger)

                # Für alle Bezirke, die die Bedingung nich erfüllen: Verbinde die nächsten k Versorgungszentren
                for zone in df_list_zones.index[df_list_zones["provider"] < True]:
                    zone_data = self.zones.loc[zone, :]

                    # falls bereits mit einem Versorgungszentrum verbunden -> Lösche das Zentrum aus der Menge der Punkte
                    tmp_set_provider = set_names_provider - df_list_zones.loc[zone, "set zones"]
                    provider_tmp = provider.loc[list(tmp_set_provider), :]

                    # Bestimme die fehlende Anzahl an Versorgungszentren
                    # Auswahlkriterium: nächstgelegen
                    list_idx_provider = get_nearest_points_from_set(x_point=zone_data.loc["XCoord"],
                                                                    y_point=zone_data.loc["YCoord"],
                                                                    n=anz_versorger - df_list_zones.loc[
                                                                        zone, "no_provider"],
                                                                    array_points=provider_tmp[["XCoord", "YCoord"]].values)
                    self.matrizen_VFS[vfs][zone, provider_tmp.index[list_idx_provider]] = 1
                    self.matrizen_VFS[vfs][provider_tmp.index[list_idx_provider], zone] = 1
                    # debugbefehl Entfernungen
                    # distances = calculate_distance_coordinates(x1=zone_data.loc["XCoord"], y1=zone_data.loc["YCoord"],
                    #                                            vec_x2=provider_tmp.loc[:, "XCoord"].values,
                    #                                            vec_y2=provider_tmp.loc[:, "YCoord"].values)

            # inaktive Quelle oder Ziel
            if (self.attr_is_from_zone is not None) & (self.attr_is_to_zone is not None):
                # Aufbau Maske mit aktiven und inaktiven OD Paaren
                # Quelle und Ziel müssen aktiv sein und die transponierte Matrix davon
                idx_inactive = np.matmul(self.zones[self.attr_is_from_zone].values.reshape(-1, 1),
                                         self.zones[self.attr_is_to_zone].values.reshape(1, -1))

                idx_inactive = (idx_inactive + idx_inactive.transpose()).astype(bool)

                # Adjazenzmatrix wird mit Maske multipliziert, um die Werte der aktiven Paare zu enthalten
                self.matrizen_VFS[vfs] = self.matrizen_VFS[vfs] * idx_inactive
            elif (self.attr_is_from_zone is not None) & (self.attr_is_to_zone is not None):
                raise ValueError("Fall ist nicht implementiert: Quell oder Zielattribut gegeben aber nicht beides")

            # debugzwecke
            if self.debug_mode:
                # zeigt an, mit welchen Bezirken ein Bezirk verbunden ist (=benachbarte Zentren)
                list_zones = self.adj_matrix_to_set_of_connected_zones(vfs)

                # Zeigt das Ergebnis in Visum an
                self.export_net(visum=self.visum, list_vfs=[vfs], links_additive=False)

                logging.info(f"{vfs}: das Ergebnis kann in Visum bestaunt werden")

    def delete_unused_nodes(self):
        # Lösche Punkte ohne Strecke

        # Filter anpassen
        self.visum.Filters.NodeFilter().Init()
        self.visum.Filters.NodeFilter().AddCondition("OP_NONE", True, "Count:InLinks", "GreaterVal", 0)
        self.visum.Filters.NodeFilter().UseFilter = True

        # Löschen
        self.visum.Net.Nodes.RemoveAll(OnlyActive=True)

        # Filter initialisieren
        self.visum.Filters.NodeFilter().Init()

    def export_matrix(self, visum=None, list_vfs=None):
        if list_vfs is None:
            list_vfs = self.vfs.keys()

        for vfs in list_vfs:
            matrix = self.matrizen_VFS[vfs]
            if visum is not None:
                name_matrix = f"{vfs}_{self.nachbarschaftsgrad_vfs[vfs]}_{self.anz_versorger_vfs[vfs]}"
                matrix_instances = self.visum.Net.Matrices.ItemsByRef(f'''Matrix([CODE]= "{name_matrix}") ''')

                if matrix_instances.Count < 1:
                    # Erstelle Matrix
                    matrix_instance = visum.Net.AddMatrix(-1, 2, 3)
                    matrix_instance.SetAttValue("CODE", name_matrix)
                    matrix_instance.SetAttValue("NAME", name_matrix)
                elif matrix_instances.Count < 1:
                    logging.warning("Matrixcode ist mehrfach vorhanden")
                    matrix_instance = matrix_instances.Iterator.Item
                else:
                    matrix_instance = matrix_instances.Iterator.Item

                matrix_instance.SetValues(matrix)

            else:
                if self.path_output is None:
                    path_mat = Path.cwd()
                else:
                    path_mat = self.path_output

                path_mat = path_mat / f"{vfs}_max_nachbar_{self.nachbarschaftsgrad_vfs[vfs]}_anz_versorgungszentren_{self.anz_versorger_vfs[vfs]}.mtx"
                df_mat = pd.DataFrame(self.matrizen_VFS[vfs],
                                      columns=self.zones["No"].values.astype(int), index=self.zones["No"].values.astype(int)
                                      , dtype=int
                                      ).stack().reset_index()

                with open(path_mat, "w", newline='\n') as f:
                    str_header = '''$O
* Universität Stuttgart
*
* Verbindungsfunktionsstufe 5
*
* symmetrische Matrix
*
* Parameter
* Nachbarschaftsniveau Z-Z:
* Nachbarschaftsniveau Z-Z+:
*
* Zeitbereich
0 24
*
* Faktor
*
1.0
*
* vonBezirk NachBezirk Matrixwert
'''

                    f.write(str_header)
                    df_mat.to_csv(f, header=False, sep=" ", index=False)


        # Falls Visuminstanz erkannt: erstelle & exportiere Daten in Visum

        # Sonst: Speichere .mtx Datei
        todo = 1

    def export_net(self, visum=None, links_additive=True, list_vfs=None):
        if self.path_output is None:
            # falls kein Dateipfad übergeben ist: Verwende Visumdateipfad, falls eine Visuminstanz existiert, ansonsten verwende den aktuellen Pfad
            if visum is not None:
                path_net = Path(visum.GetPath(1))
            else:
                path_net = Path.cwd()
        else:
            path_net = self.path_output

        if list_vfs is None:
            list_vfs = self.vfs.keys()

        path_net = path_net / f"{'_'.join(list_vfs)}.net"

        # Erstelle eine Knotenliste
        df_nodes = self.zones

        # Erstelle Streckenliste
        df_edges = self.adj_matrix_to_links(list_vfs=list_vfs)
        if len(df_edges) < 1:
            logging.warning("es existieren keine Strecken")
            return

        # Erstelle Liste mit Streckentypen
        df_linktypes = df_edges["TypeNo"].drop_duplicates().to_frame(name="Name")
        df_linktypes["No"] = df_linktypes["Name"].replace(self.vfs)
        df_linktypes["Rank"] = df_linktypes["No"]

        list_tsys_net = pd.DataFrame(visum.Net.TSystems.GetMultipleAttributes(["Code"])).squeeze().values.tolist()
        df_linktypes["TSysSet"] = ",".join(list_tsys_net)

        # Übersetze VFS in TypeNo
        df_edges["TypeNo"].replace(self.vfs, inplace=True)

        # Übersetze Id in Bezirksnummer
        df_edges["FromNodeNo"].replace(self.zones["No"], inplace=True)
        df_edges["ToNodeNo"].replace(self.zones["No"], inplace=True)

        # Hinzufügen einer Nummer
        # 1. Identifikation der Hin- & Gegenrichtung
        df_edges["No"] = df_edges[["FromNodeNo", "ToNodeNo"]].min(axis=1).astype(str) + "_" + df_edges[
            ["FromNodeNo", "ToNodeNo"]].max(axis=1).astype(str)

        # 2. Nummerierung
        if links_additive:
            no_start = visum.Net.AttValue(r"Max:Links\No") + 1
        else:
            no_start = 1

        dict_no = dict(zip(df_edges["No"].drop_duplicates(), range(no_start, int(len(df_edges) / 2) + 1)))
        df_edges["No"].replace(dict_no, inplace=True)

        # Lösche Strecken, die in unterschiedlichen VFS mehrmals vorkommen
        # höchste Stufe wird behalten (Sortierung nach aufsteigender Nummer & Löschen der Duplikate)
        df_edges.sort_values("TypeNo", inplace=True)
        df_edges.drop_duplicates(["FromNodeNo", "ToNodeNo"], inplace=True)

        # Abgleich Knotennummern/Namen
        # if visum is not None:
        #     node_no_max_existing = visum.Net.AttValue(r"Max:Nodes\No")
        # Wenn möglich: Knotennummern == Bezirksnummern
        # Sonst nächste freie Nummern

        # Schreibe .net Datei
        with open(path_net, mode="w", newline="\n") as f:
            header = '''$VISION
* Universität Stuttgart Fakultät 2 Bau+Umweltingenieurwissenschaften Stuttgart
* 08/23/22
* 
* Table: Version block
* 
$VERSION:VERSNR;FILETYPE;LANGUAGE;UNIT
13;Net;ENG;KM

'''

            f.write(header)
            write_object_to_net("Node", df_nodes[["No", "Name", "TypeNo", "XCoord", "YCoord"]], f)
            write_object_to_net("Link type", df_linktypes, f)
            write_object_to_net("Link", df_edges[["No", "FromNodeNo", "ToNodeNo", "TypeNo"]], f)

        # Falls Visuminstanz übergeben: lade die .net Datei
        if visum is not None:
            if links_additive is not True:
                visum.Net.Links.RemoveAll(OnlyActive=True)

            visum.IO.LoadNet(str(path_net), ReadAdditive=True)

            if visum.Net.Links.Count < len(df_edges):
                logging.warning("Da hat beim Import der Netzdatei etwas nicht geklappt")

        logging.info(f"die Netzdatei wurde erfolgreich erstellt")

    ## initialisiert die Adjazenzmatrizen
    def init_results(self):
        dict_vfs = {}
        for vfs in self.vfs:
            dict_vfs[vfs] = np.zeros([len(self.zones), len(self.zones)], dtype=bool)

        self.matrizen_VFS = dict_vfs

## @package luftlinientool
# todo Beschreibung
# @brief todo Kurzbeschreibung

import pandas as pd
import logging
import numpy as np
from scipy.spatial import Delaunay
from pathlib import Path
from math import radians
import win32com.client as com
import webbrowser


# todo update nach Änderung sofort in GUI Event, hier nicht nochmaliges Update

# ====== allgemeine, nützliche FUnktionen =====

## öffnet eine Visuminstanz falls nicht bereits offen
# ermöglicht simultanes Aufrufen der Datei Visumintern und -extern
# @param path: Dateipfad (Path/str) einer Visumversionsdatei
# @param version: Visumversion, default 22
# @return: Visuminstanz
def open_visum(path, version=220):
    try:
        # testet ob die Variable Visum existiert
        global Visum
        Visum
        name = Visum.UserPreferences.DocumentName
    except NameError:
        # falls nicht - Öffne eine Visuminstanz
        logging.info('initialize visum instance')
        Visum = com.Dispatch(f"Visum.Visum.{version}")
        logging.info('open visum file: {}'.format(path))
        Visum.LoadVersion(path)
        logging.info('erfolgreich geladen')
    return Visum


## Exportiert die Daten eines Visumobjekttyps in Netzdateiformat
# @param[in] object: Visumobjekttyp (Singular), z.B. 'link'
# @param[in] df_object_attributes_to_write: Datentabelle des Objekts. Tabelle enthält nur Attribute, die in Visum
# importiert werden können (insbesondere die notwendigen Attribute)
# @param[in] file: Zieldatei, im Schreib- oder Erweiterungsmodus (w oder a)
def write_object_to_net(object, df_object_attributes_to_write, file):
    header = ["*", "*"]
    header.insert(1, "* Table: " + object + "s")
    header.append(
        ("$" + object.upper().replace(" ", "") + ":" + ";".join(df_object_attributes_to_write.columns)).upper() + "\n")

    # Header wird geschrieben
    file.write("\n".join(header))
    # Tabelle wird geschrieben
    df_object_attributes_to_write.to_csv(file, header=False, sep=";", index=False)


## check if a matrix is symmetric
# @param matrix: Matrix, die auf Symmetrie getestet werden soll
# @param tol: Toleranz für erlaubte Abweichung, default 1e-8
# @return: True oder False
def is_symmetric(matrix, tol=1e-8):
    # Anwendung der Maximums-Norm für die Diff zwischen der Matrix und der Transponierten
    # Norm > 0 -> keine Symmetrie
    return np.linalg.norm(matrix.astype(int) - matrix.T.astype(int), np.Inf) < tol


## Berechnung der Distanz zwischen Koordinaten (Lat, Lon)
# Implementation der Haversine Formel
# @param x1: x-Koordinate Punkt 1
# @param y1: y-Koordinate Punkt 1
# @param vec_x2: x-Koordinate Punktevektor
# @param vec_y2: y-Koordinate Punktevektor
# @return Vektor mit den Distanzen aller Punkte des Punktevektors zu Punkt 1
def calculate_distance_coordinates_haversine(x1, y1, vec_x2, vec_y2):
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

## Berechnung der Distanz zwischen Koordinaten (x, y)
# Euklidische Distanzberechnung!
# @param x1: x-Koordinate Punkt 1
# @param y1: y-Koordinate Punkt 1
# @param vec_x2: x-Koordinate Punktevektor
# @param vec_y2: y-Koordinate Punktevektor
# @return Vektor mit den Distanzen aller Punkte des Punktevektors zu Punkt 1
def calculate_eucl_distance_coordinates(x1, y1, vec_x2, vec_y2):
    diff_x = vec_x2 - x1
    diff_y = vec_y2 - y1

    # todo Fallunterscheidung für negative Koordinaten

    distances = np.sqrt(np.square(diff_x) + np.square(diff_y))

    return distances


## identifiziert die nächsten n Punkte aus einer gegebenen Punktemenge zu einem einzelnen Punkt
# zuerst werden die Distanzen aller Punkte zu dem einzelnen Punkt berechnet (Haversine)
# danach werden die n am kürzesten entfernten Punkte gefiltert und deren Indizes zurückgegeben
# @param x_point: x-Koordinate des Referenzpunktes
# @param y_point: y-Koordinate des Referenzpunktes
# @param array_points: Array mit den x- & y-Koordinaten der Punkte
# @param n: gewünschte Punkteanzahl
# @return list_indizes: Liste der Indizes der nächstgelegenen n Punkte
def get_nearest_points_from_set(x_point, y_point, array_points, formula, n=None):
    # Falls keine Auswahl existiert
    if (n is not None) and (n >= len(array_points)):
        # es werden alle möglichen Punkte zurückgegeben
        return list(range(0, len(array_points)))

    # Berechne Entfernungen
    if formula =="haversine":
        distances = calculate_distance_coordinates_haversine(x1=x_point, y1=y_point, vec_x2=array_points[:, 0],
                                                             vec_y2=array_points[:, 1])
    elif formula == "euclidean":
        distances = calculate_eucl_distance_coordinates(x1=x_point, y1=y_point, vec_x2=array_points[:, 0],
                                                        vec_y2=array_points[:, 1])
    else:
        logging.warning("Fall Abstandsberechnung ist nicht implementiert")

    # Index der n niedrigsten Werte
    list_indizes = np.argpartition(distances, n)[:n]

    return list_indizes


## Öffnet die Readme Datei

def show_info(path_scripts: Path = Path.cwd()):
    webbrowser.open(str(path_scripts / "README.md"), new=2)


# ===== Klassendefinition ======
## Klasse LuftlinienCalculator
# eine Instanz der Klasse enthält Attribute und Berechnungsmöglichkeiten um die VFS zwischen Bezirken zu ermitteln
class LuftlinienCalculator:

    ## Konstruktor
    # @param source: Dateiname (str) oder Visuminstanz
    # @param attr_vfs: Name des Bezirkattributs, das die Kategorisierung in OZ,MZ,UZ ... enthält. Default: TypeNr
    # @param dict_vfs: Dictionary, das die Attributwerte für die jeweiligen VFS enthält
    # @param max_entfernung: Angabe, bis zu welcher Entfernung, Nachbar angebunden werden
    # @param anz_versorger: Angabe, an wie viele höherrangige Zentren ein Bezirk angebunden werden soll
    # @param attr_quelle: Name des Attributs, das angibt, ob der Bezirk als Quelle berücksichtigt wird. Default: None
    # @param attr_ziel: Name des Attributs, das angibt, ob der Bezirk als Ziel berücksichtigt wird. Default: None
    # @param use_gui: gibt an, ob die Instanz mit einer GUI vernetzt ist. Default: False
    # @param use_filter: gibt an, ob nur aktive Bezirke berücksichtigt werden. Kann nur verwendet werden, wenn source = Visuminstanz
    # @param path_output: optionale Möglichkeit einen Pfad für den Dateiexport anzugeben. Default: None. Dann wird bei bedarf der aktuelle Ordner verwendet.
    def __init__(self, source,
                 attr_vfs: str = "TypeNo",
                 dict_vfs: dict = {"VFS 0": 0, "VFS 1": 1, "VFS 2": 2, "VFS 3": 3, "VFS 4": 4, "VFS 5": 5},
                 max_entfernung=1,
                 anz_versorger=0,
                 attr_quelle=None,
                 attr_ziel=None,
                 use_gui: bool = False,
                 use_filter: bool = False,
                 formula_distance: str="euclidean",
                 path_output=None):

        # True, wenn am Debuggen. Ermöglicht die Durchführung von Zwischenanalysen, die im normalen Programmablauf nicht berücksichtigt werden
        self.debug_mode = False

        # Benötigte Bezirksattribute
        # verarbeiten der Übergabeparameter
        self.attr_zones = ["No", "Name", "XCoord", "YCoord"]
        self.attr_central_level = attr_vfs
        self.attr_zones.append(self.attr_central_level)
        if attr_quelle is not None:
            self.attr_zones.append(attr_quelle)
        if attr_ziel is not None:
            self.attr_zones.append(attr_ziel)

        # Pfade
        self.path_output = path_output

        # Liste der VFS, die bearbeitet werden sollen
        self.vfs = dict_vfs

        # Abstandsberechnung
        self.formula_dist = formula_distance

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
            attr_zones = self.attr_zones
            self.zones = pd.DataFrame(source.Net.Zones.GetMultipleAttributes(attr_zones, OnlyActive=False),
                                      columns=attr_zones)
            set_active_zones = set(np.array(source.Net.Zones.GetMultiAttValues("No", OnlyActive=True), dtype=int)[:, 1])
            self.zones["IsActive"] = self.zones["No"].isin(set_active_zones)

            logging.info("%s Bezirke eingelesen", len(self.zones))
        else:
            self.visum = None
            logging.warning("Einlesen der Bezirksdaten ist fehlgeschlagen, Inputformat ist nicht implementiert")

        if attr_quelle is None:
            attr_quelle = 'quelle'
            self.zones[attr_quelle] = 1

        if attr_ziel is None:
            attr_ziel = 'ziel'
            self.zones[attr_ziel] = 1

        # Abfangen attr_ziel=attr_quelle: Lösche Spaltenduplikat
        if attr_ziel == attr_quelle:
            self.zones = self.zones.T.drop_duplicates().T

        self.attr_is_from_zone = attr_quelle
        self.attr_is_to_zone = attr_ziel

        # Init VFS Matrizen
        # Dict mit Matrix je VFS: Anzahl Bezirke x Anzahl Bezirke
        self.init_results()

        # Sprache
        self.language = self.visum.GetCurrentLanguage()

    ## übersetzt die Adjazenzmatrizen der gewünschten VFS in eine Streckenliste
    # @param list_vfs: Liste der VFS. Falls nicht gegeben, werden alle VFS der Instanz verwendet
    # @return: DataFrame mit allen Strecken und ihrer VFS. Achtung: Duplikate werden nicht entfernt
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

    ## wandelt die Adjazenzmatrix in eine Liste der verbundenen Bezirke je Bezirk um
    # @param vfs: str, Name der zu betrachtenden VFS
    # @param use_zone_names: bool, falls True werden die hitnerlegten Bezirksnamen verwendet
    # @return: DataFrame mit list Objekt je Bezirk und einer Spalte, die die Anzahl enthält
    def adj_matrix_to_set_of_connected_zones(self, vfs, use_zone_names=True):
        # Matrix zu DataFrame
        if use_zone_names:
            # Falls Namen verwendet werden sollen, werden die Zeilen & Spalten benannt
            df = pd.DataFrame(self.matrizen_VFS[vfs], index=self.zones["Name"], columns=self.zones["Name"])
        else:
            df = pd.DataFrame(self.matrizen_VFS[vfs])

        # Erstellt einen DataFrame, der für jede Zeile der Matrix die Spaltennamen enthält, für die der Eintrag True ist
        df_set_zones = df.mul(df.columns.tolist()).apply(lambda x: set(zone for zone in x if zone), axis=1).to_frame(
            name="set zones")
        # Ermittelt die Länge jeder Liste
        df_set_zones["no zones"] = df_set_zones["set zones"].apply(len)

        return df_set_zones

    # Berechnet, welche Nachbarn innerhalb von n Schritten erreicht werden können
    # @param max_steps: maximale Entfernung (Schritte)
    # @param vfs: zu untersuchende VFS
    # @return matrix: Adjazenzmatrix für die Erreichbare Nachbarn innerhalb der max-steps
    def calculate_reachability_max_steps(self, max_steps, vfs):

        matrix = np.linalg.matrix_power(self.matrizen_VFS[vfs], max_steps)
        np.fill_diagonal(matrix, 0)

        return matrix

    ## berechnet für jede hinterlegte VFS der Instanz die Adjazenzmatrix
    def calculate_main(self):
        # Init Ergebnisse
        logging.info(f"Berechnung über alle VFS wird gestartet")
        self.init_results()
        logging.info(f"Adjazenzmatrizen wurden initialisiert")

        # Schleife über alle vfs
        for vfs in self.vfs:
            # Berechne die Werte für die VFS
            self.calculate_vfs(vfs)

            # # todo Idee Aktivierung Outputexportbuttions in GUI
            # if self.use_gui:
            #     a = 1

        logging.info("Die Berechnung über alle VFS ist abgeschlossen")

    ## berechnet die Verbindungen einer VFS
    # entspricht Funktion Program.LLCalc
    # @param vfs: die Verbindungsfunktionsstufe, für die Verbindungen ermittel werden
    def calculate_vfs(self, vfs):

        # Attributswert der Bezirke für die gewählte VFS
        value_vfs = self.vfs[vfs]

        # Attribute der VFS
        k_nachbar = self.nachbarschaftsgrad_vfs[vfs]
        anz_versorger = self.anz_versorger_vfs[vfs]

        # Filtere Bezirksdaten, die die Bedingungen erfüllen
        # Sind Aktiv todo Erweiterung Filterung nach attr_filter
        # TypNr <= VFS
        active_zones = self.zones
        active_zones = active_zones.loc[(active_zones[self.attr_central_level] <= value_vfs)
                                        & (active_zones["IsActive"] > 0),
                       :]

        # Abfangen, falls es Bezirke mit identischen Koordinaten gibt, dann funktioniert DeLauney nicht zuverlässig
        if len(active_zones) > len(active_zones[["XCoord", "YCoord"]].drop_duplicates()):
            duplicate_zones = active_zones[active_zones.duplicated(subset=["XCoord", "YCoord"], keep=False)]
            duplicate_zones_string = ', '.join(duplicate_zones["No"].apply(lambda x: str(int(x))) + "/" + duplicate_zones["Name"])
            raise ValueError(f"Abbruch: Bezirke mit den identischen Koordinaten (NUMMER/NAME):{duplicate_zones_string}")
        elif len(active_zones) < 3:
            logging.info(f"{vfs}: es sind zu wenige Bezirke aktiv")
        else:
            logging.info(f"{vfs}: Delauney Triangulation wird für {len(active_zones)} Bezirke durchgeführt")

            if k_nachbar > 0:

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
                df_list_zones["no_provider"] = df_list_zones["set zones"].apply(set_names_provider.intersection).apply(
                    len)
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
                                                                    array_points=provider_tmp[
                                                                        ["XCoord", "YCoord"]].values,
                                                                    formula=self.formula_dist)
                    self.matrizen_VFS[vfs][zone, provider_tmp.index[list_idx_provider]] = 1
                    self.matrizen_VFS[vfs][provider_tmp.index[list_idx_provider], zone] = 1
                    # debugbefehl Entfernungen
                    # distances = calculate_distance_coordinates(x1=zone_data.loc["XCoord"], y1=zone_data.loc["YCoord"],
                    #                                            vec_x2=provider_tmp.loc[:, "XCoord"].values,
                    #                                            vec_y2=provider_tmp.loc[:, "YCoord"].values)

            # inaktive Quelle oder Ziel

            # Aufbau Maske mit aktiven und inaktiven OD Paaren
            # Quelle und Ziel müssen aktiv sein und die transponierte Matrix davon (Symmetrie)
            # Logik: Filtere OD-Paare mit Quelle & Ziel aktiv...
            #
            #  Quelle * Ziel  = Matrix
            # (1 0).T * (1 1) = (1  1
            #                    0  0)
            #
            # und symmetrisiere diese
            # (1  1
            #  1  0)


            # Attribute Quelle und Ziel
            vector_is_from_zone = self.zones[self.attr_is_from_zone].values
            vector_is_to_zone = self.zones[self.attr_is_to_zone].values
            # über dyadisches Produkt ("outer product") verknüpfen
            # Logik als Maske über existierende Matrix legen
            idx_active = np.outer(vector_is_from_zone, vector_is_to_zone).astype(bool)
            # symmetrisieren der Matrix (Bool Oder-Verknüpfung mit transponierter Matrix)
            # Wo OD-Relation, da DO-Relation
            idx_active_symm = idx_active + idx_active.T

            # Adjazenzmatrix wird mit Maske multipliziert, um die Werte der aktiven Paare zu enthalten
            self.matrizen_VFS[vfs] = self.matrizen_VFS[vfs] * idx_active_symm.astype(int)

            # Symmetrietest
            if np.sum(self.matrizen_VFS[vfs] - self.matrizen_VFS[vfs].T) > 0:
                raise ValueError("Matrix ist nicht symmetrisch")

            # debugzwecke
            if self.debug_mode:
                # zeigt an, mit welchen Bezirken ein Bezirk verbunden ist (=benachbarte Zentren)
                list_zones = self.adj_matrix_to_set_of_connected_zones(vfs)

                # Zeigt das Ergebnis in Visum an
                self.export_net(visum=self.visum, list_vfs=[vfs], links_additive=False)

                logging.info(f"{vfs}: das Ergebnis kann in Visum bestaunt werden")

            logging.info(f"Die Berechnung {vfs} ist abgeschlossen")

            df_zones_info = self.adj_matrix_to_set_of_connected_zones(vfs)
            df_zones_info["set zones"] = df_zones_info["set zones"].str.join(",")
            logging.info('\t'+ df_zones_info.to_string().replace('\n', '\n\t'))

    ## löschte Knoten in Visum, die keine Strecken anbinden
    # geht nur, wenn eine Visuminstanz enthalten ist
    # alle Knoten ohne Strecken werden gefiltert & die aktiven Knoten werden gelöscht
    # anschließend wird der Filter zurückgesetzt
    def delete_unused_nodes(self):
        if self.visum is None:
            logging.warning("Knoten löschen: es ist keine Visuminstanz verknüpft")
            return

        # Lösche Punkte ohne Strecke

        # Filter anpassen
        self.visum.Filters.NodeFilter().Init()
        self.visum.Filters.NodeFilter().AddCondition("OP_NONE", True, "Count:InLinks", "GreaterVal", 0)
        self.visum.Filters.NodeFilter().UseFilter = True

        n = self.visum.Net.Nodes.CountActive

        # Löschen
        self.visum.Net.Nodes.RemoveAll(OnlyActive=True)

        # Filter initialisieren
        self.visum.Filters.NodeFilter().Init()

        logging.info(f"{n} isolierte Knoten wurden gelöscht")

    ## exportiert die gewünschten Adjazenzmatrizen
    # entweder direkt nach Visum (falls Visuminstanz verknüpft)
    # oder als .mtx datei
    # Vorhandene matrizen werden überschrieben
    # @param visum: optionale Übergabe einer Visuminstanz. Default None
    # @param list_vfs: optionale Übergabe einer Menge an VFS. Default: None (alle des Objekts)
    def export_matrix(self, visum=None, list_vfs=None):

        # Falls Visuminstanz erkannt: erstelle & exportiere Daten in Visum
        # Sonst: Speichere .mtx Datei

        if list_vfs is None:
            list_vfs = self.vfs.keys()

        for vfs in list_vfs:
            matrix = self.matrizen_VFS[vfs]
            if visum is not None:
                # Benennung
                if self.anz_versorger_vfs[vfs] < 1:
                    # Term mit Versorgungsfkt wird weggelassen
                    name_matrix = f"RIN_{vfs}_n={self.nachbarschaftsgrad_vfs[vfs]}"
                else:
                    # Term mit Versorgungsfkt wird hinzugefügt
                    name_matrix = f"RIN_{vfs}_n={self.nachbarschaftsgrad_vfs[vfs]}_v={self.anz_versorger_vfs[vfs]}"

                if visum.Net.Matrices.Count < 1:
                    # Erstelle Matrix
                    matrix_instance = visum.Net.AddMatrix(-1, 2, 3)
                    matrix_instance.SetAttValue("CODE", name_matrix)
                    matrix_instance.SetAttValue("NAME", name_matrix)
                else:
                    # Suche existierende Matrizen mit der Benennung
                    matrix_instances = self.visum.Net.Matrices.ItemsByRef(f'''Matrix([CODE]= "{name_matrix}") ''')

                    if matrix_instances.Count < 1:
                        # Erstelle Matrix
                        matrix_instance = visum.Net.AddMatrix(-1, 2, 3)
                        matrix_instance.SetAttValue("CODE", name_matrix)
                        matrix_instance.SetAttValue("NAME", name_matrix)
                    elif matrix_instances.Count > 1:
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
                                      columns=self.zones["No"].values.astype(int),
                                      index=self.zones["No"].values.astype(int)
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

        logging.info(f"{len(list_vfs)} Matrizen wurden exportiert")

    ## exportiert eine Netzdatei
    # falls eine Visuminstanz übergeben wird, wird die Netdatei in Visum geladen
    # @param visum: optionale Übergabe einer Visuminstanz. Default None
    # @param links_additive: falls False werden die existierenden Strecken in Visum gelöscht
    # @param list_vfs: Liste der VFS, die berücksichtigt werden sollen. Default: Alle des Objekts
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
            # return

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
        if links_additive & (visum is not None):
            # Abgleich Knotennummern/Namen
            #Bsp und/oder python
            no_node_max = visum.Net.AttValue(r"Max:Nodes\No") or 0
            no_link_max = visum.Net.AttValue(r"Max:Links\No") or 0

            if (no_node_max is not None) & (no_node_max > df_nodes["No"].max()):
                no_node_start = no_node_max + 1
            else:
                no_node_start = 1
                
            no_link_start = no_link_max + 1

        else:
            no_link_start = 1
            no_node_start = 1

        dict_no_nodes = dict(zip(df_nodes["No"].drop_duplicates(), range(no_node_start, no_node_start + len(df_nodes) + 1)))
        dict_no_links = dict(zip(df_edges["No"].drop_duplicates(), range(no_link_start, no_link_start + int(len(df_edges) / 2) + 1)))

        # Test: für jede Strecke existiert eine Nummer
        if len(dict_no_links) != len(df_edges["No"].drop_duplicates()):
            logging.error("Streckennummerierung passt nicht zur Streckeanzahl")

        df_edges["No"].replace(dict_no_links, inplace=True)
        df_nodes["No"].replace(dict_no_nodes, inplace=True)
        df_edges["FromNodeNo"].replace(dict_no_nodes, inplace=True)
        df_edges["ToNodeNo"].replace(dict_no_nodes, inplace=True)

        # Lösche Strecken, die in unterschiedlichen VFS mehrmals vorkommen
        # höchste Stufe wird behalten (Sortierung nach aufsteigender Nummer & Löschen der Duplikate)
        df_edges.sort_values("TypeNo", inplace=True)
        df_edges.drop_duplicates(["FromNodeNo", "ToNodeNo"], inplace=True)

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
            write_object_to_net("Node", df_nodes[self.attr_zones], f)
            write_object_to_net("Link type", df_linktypes, f)
            write_object_to_net("Link", df_edges[["No", "FromNodeNo", "ToNodeNo", "TypeNo"]], f)

        # Falls Visuminstanz übergeben: lade die .net Datei
        if visum is not None:
            if links_additive is not True:
                visum.Net.Links.RemoveAll(OnlyActive=True)

            visum.IO.LoadNet(str(path_net), ReadAdditive=True)

            if visum.Net.Links.Count < len(df_edges):
                logging.warning("Da hat beim Import der Netzdatei etwas nicht geklappt")

        logging.info(f"die Netzdatei von {len(list_vfs)} VFS wurde exportiert")

    def export_zones_uda_connections(self, vfs):
        # Erstelle UDA wenn nicht vorhanden

        str_no_conn =f"RIN_Anz_Verbindungen_{vfs}".replace(" ", "")
        str_conn = f"RIN_Verbindungen_{vfs}".replace(" ", "")

        try:
            self.visum.Net.Zones.AddUserDefinedAttribute(str_no_conn,
                                                         str_no_conn,
                                                         str_no_conn, 5)
            self.visum.Net.Zones.AddUserDefinedAttribute(str_conn,
                                                         str_conn,
                                                         str_conn, 5)
        except:
            pass

        # Lade Verbindungen
        df_zones = self.adj_matrix_to_set_of_connected_zones(vfs).reset_index()
        df_zones["No"] = df_zones.Name.replace(self.zones.set_index("Name")["No"].astype(int).to_dict())
        df_zones.set_index("No", inplace=True)

        # Schreibe das Ergebnis nach Visum
        df_format = pd.DataFrame(self.visum.Net.Zones.GetMultiAttValues("No"), columns=["Idx","No"]).set_index("No")

        df_format = df_format.join(df_zones)
        df_format[str_conn] = df_format["set zones"].str.join(",")

        self.visum.Net.Zones.SetMultiAttValues(str_no_conn, df_format.loc[:, ["Idx", "no zones"]].values)
        self.visum.Net.Zones.SetMultiAttValues(str_conn, df_format.loc[:, ["Idx", str_conn]].values)


    ## initialisiert die Adjazenzmatrizen
    def init_results(self):
        dict_vfs = {}
        for vfs in self.vfs:
            dict_vfs[vfs] = np.zeros([len(self.zones), len(self.zones)], dtype=bool)

        self.matrizen_VFS = dict_vfs


if __name__ == '__main__':
    from pathlib import Path

    # Parameterübergabe
    path_source = Path(r"C:\Users\ac128405\Desktop\Software\Luftlinientool\Beispielnetz")
    file_source = 'ZentraleOrteBW_Bezirke.ver'

    # Settings Logging
    path_logfile = Path(__file__)
    path_logfile = path_source / path_logfile.name.replace(".py", ".log")

    root_logger = logging.getLogger()
    root_logger.setLevel(logging.INFO)
    logger_format = logging.Formatter("%(asctime)s %(levelname)s: %(message)s", datefmt="%d.%m.%Y %I:%M:%S %p")
    # Output in Konsole & Logfile
    stream_handler = logging.StreamHandler()
    stream_handler.setFormatter(logger_format)
    file_handler = logging.FileHandler(path_logfile, mode="w")
    file_handler.setFormatter(logger_format)
    # add handles to logger
    root_logger.addHandler(file_handler)
    root_logger.addHandler(stream_handler)

    source = path_source / file_source

    Visum = open_visum(source)
    llt1 = LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel", anz_versorger=2, max_entfernung=1)

    llt1.calculate_main()
    llt1.export_matrix(visum=llt1.visum)
    llt1.export_net(visum=llt1.visum, links_additive=False)

    llt1.delete_unused_nodes()

    del Visum

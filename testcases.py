# from VisumOverlay import *
import logging

'''
Ziel: dieses Skript enthält Testszenarien & erwartete Ergebnisse, um bei Änderungen/Weiterentwicklungen schnell die Auswirkungen zu checken
'''

# hier bitte die Testszenarien anlegen
# Erwartungswert in 'vergleich_mtx_sum' festlegen.
dict_testcases = {
    "VFS 0 n=0 a=1": {"vfs": "VFS 0", "n_entfernung": 0, "anz_versorger": 1, "is_quelle": None, "is_ziel": None,
                  "vergleich_mtx_sum": 0},
    "VFS 0 n=1 a=0": {"vfs": "VFS 0", "n_entfernung": 1, "anz_versorger": 0, "is_quelle": None, "is_ziel": None,
                      "vergleich_mtx_sum": 26},
    "VFS 2 n=1 a=0": {"vfs": "VFS II", "n_entfernung": 1, "anz_versorger": 0, "is_quelle": None, "is_ziel": None,
                      "vergleich_mtx_sum": 26},
    "VFS 2 n=0 a=1": {"vfs": "VFS II", "n_entfernung": 0, "anz_versorger": 1, "is_quelle": None, "is_ziel": None,
                      "vergleich_mtx_sum": 26},
    # Quelle & Ziel Filter

    # Bezirksfilter

    }

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # Parameterübergabe
    path_source = Path(r"S:\VuV-Tools\Fertige Tools\Luftlinientool\Python (Visumintegration)")
    file_source = 'Testnetz.ver'

    # Settings Logging
    path_logfile = Path(__file__)
    path_logfile = Path.cwd() / path_logfile.name.replace(".py", ".log")

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

    Visum = llt.open_visum(source)

    for case, param in dict_testcases.items():
        ltt1 = llt.LuftlinienCalculator(Visum,
                                        anz_versorger=1, max_entfernung=1,
                                        dict_vfs={"VFS 0": 0, "VFS I": 1, "VFS II": 2, "VFS III": 3, "VFS IV": 4, "VFS V": 5},)


        ltt1.calculate_main()
        ltt1.export_matrix(visum=ltt1.visum)


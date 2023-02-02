# from VisumOverlay import *
import logging

'''
Ziel: dieses Skript enthält Testszenarien & erwartete Ergebnisse, um bei Änderungen/Weiterentwicklungen schnell die Auswirkungen zu checken
'''

# hier bitte die Testszenarien anlegen
# Erwartungswert in 'vergleich_mtx_sum' festlegen.
dict_testcases = {
    # "VFS 0 n=0 a=0": {"vfs": {"VFS 0": 0}, "n_entfernung": 0, "anz_versorger": 0, "is_quelle": None, "is_ziel": None,
    #               "vergleich_mtx_sum": 0},
    # "VFS 0 n=0 a=1": {"vfs": {"VFS 0": 0}, "n_entfernung": 0, "anz_versorger": 1, "is_quelle": None, "is_ziel": None,
    #               "vergleich_mtx_sum": 0},
    # "VFS 0 n=1 a=0": {"vfs": {"VFS 2": 2}, "n_entfernung": 1, "anz_versorger": 0, "is_quelle": None, "is_ziel": None,
    #                   "vergleich_mtx_sum": 26},
    # "VFS 2 n=1 a=0": {"vfs": {"VFS 2": 2}, "n_entfernung": 1, "anz_versorger": 0, "is_quelle": None, "is_ziel": None,
    #                   "vergleich_mtx_sum": 26},
    "VFS 2 n=0 a=1": {"vfs": {"VFS 2": 2}, "n_entfernung": 0, "anz_versorger": 1, "is_quelle": None, "is_ziel": None,
                      "vergleich_mtx_sum": 26},
    # Quelle & Ziel Filter
    "VFS 2 n=1 a=0 quelle ziel": {"vfs": {"VFS 2": 2}, "n_entfernung": 1, "anz_versorger": 0, "is_quelle": "Quelle",
                      "is_ziel": "Ziel",
                      "vergleich_mtx_sum": 26},

    # Bezirksfilter

    }

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # Parameterübergabe
    path_source = Path(r"C:\Users\ac128405\Desktop\Software\Luftlinientool\Beispielnetz")
    file_source = 'BWNetz_V1.ver'

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
                                        anz_versorger=param["anz_versorger"],
                                        max_entfernung=param["n_entfernung"],
                                        attr_quelle=param["is_quelle"],
                                        attr_ziel=param["is_ziel"],
                                        dict_vfs=param["vfs"])
        ltt1.calculate_main()
        ltt1.export_matrix(visum=ltt1.visum)


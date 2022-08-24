from VisumOverlay import *
import win32com.client as com
import logging

# ========= functions =========
def open_visum(path, version=220):
    try:
        ''' VisumStart
        - enables the usage in the procedure sequence or as Visum instance
        - name of the Visum version that is opened
        '''
        global Visum
        Visum
        name = Visum.UserPreferences.DocumentName
    except NameError:
        logging.info('initialize visum instance')
        Visum = com.Dispatch(f"Visum.Visum.{version}")
        logging.info('open visum file: {}'.format(path))
        Visum.LoadVersion(path)
        logging.info('erfolgreich geladen')
    return Visum

if __name__ == '__main__':
    from pathlib import Path
    import luftlinientool as llt

    # Parameterübergabe
    path_source = Path(r"C:\Users\ac128405\Desktop\Software\Luftlinientool\Beispielnetz")
    file_source = 'ZentraleOrteBW_Bezirke_OZ.ver'

    # Settings Logging
    path_logfile = Path(__file__)
    path_logfile = Path(r"S:\Forschung\BASt_EmV\13_Skripte") / "Logfiles" / path_logfile.name.replace(".py", ".log")

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

    if source.suffix == ".ver":
        Visum = open_visum(source)
        ltt1 = llt.LuftlinienCalculator(Visum, attr_quelle="Quelle", attr_ziel="Ziel")
    else:
        print("nicht implementiert")

    ltt1.calculate_main()

    del Visum



even thought the file is in the directory its still not opening 
Traceback (most recent call last):
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 40, in <module>
    df1 = read_csv_flexible(file1_path, header=None, min_cols=11)
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 17, in read_csv_flexible
    df = pd.read_csv(path, header=header, dtype=str, sep=None, engine="python", encoding="utf-8-sig")
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\io\parsers\readers.py", line 1026, in read_csv
    return _read(filepath_or_buffer, kwds)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\io\parsers\readers.py", line 620, in _read
    parser = TextFileReader(filepath_or_buffer, **kwds)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\io\parsers\readers.py", line 1620, in __init__
    self._engine = self._make_engine(f, self.engine)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\io\parsers\readers.py", line 1880, in _make_engine
    self.handles = get_handle(
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\io\common.py", line 873, in get_handle
    handle = open(
FileNotFoundError: [Errno 2] No such file or directory: 'U:\\rlbnas1_rlb_bw_firw_z\\FC\\07 EDV-Projekte\\SMART Vorkalk\\Wartungstabellen\\Befüllte Wartungstabellen (ECHTDATEN)\\IMPORTASSISTENT_fuer_RK_und_EK\\Originaldateien\\Tab_EM_ICAAP.csv'

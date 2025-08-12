Traceback (most recent call last):
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\core\indexes\base.py", line 3805, in get_loc
    return self._engine.get_loc(casted_key)
  File "index.pyx", line 167, in pandas._libs.index.IndexEngine.get_loc
  File "index.pyx", line 196, in pandas._libs.index.IndexEngine.get_loc
  File "pandas\\_libs\\hashtable_class_helper.pxi", line 7081, in pandas._libs.hashtable.PyObjectHashTable.get_item
  File "pandas\\_libs\\hashtable_class_helper.pxi", line 7089, in pandas._libs.hashtable.PyObjectHashTable.get_item
KeyError: 'Column2'

The above exception was the direct cause of the following exception:

Traceback (most recent call last):
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 17, in <module>
    df1["Column2"] = df1["Column2"].astype(str)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\core\frame.py", line 4102, in __getitem__
    indexer = self.columns.get_loc(key)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\core\indexes\base.py", line 3812, in get_loc
    raise KeyError(key) from err
KeyError: 'Column2'

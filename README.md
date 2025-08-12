Traceback (most recent call last):
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 165, in <module>
    main()
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 151, in main
    df2 = read_ek_basis_primaerbanken(FILE2_PATH)
  File "u:\rlbnas1_rlb_bw_firw_z\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\PythonsckriptTP\EK.py", line 130, in read_ek_basis_primaerbanken
    df[col] = to_float_from_maybe_comma(df[col])
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\core\frame.py", line 4102, in __getitem__
    indexer = self.columns.get_loc(key)
  File "C:\Entwicklung\.conda\envs\MSR-reports\lib\site-packages\pandas\core\indexes\base.py", line 3812, in get_loc
    raise KeyError(key) from err
KeyError: 'Eigenkapitalkosten_Fix_(in_%)'

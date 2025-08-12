I need python code for the following power querys. At first open the file whose name is Tab_EM_ICAAP.csv          which is in directory 


U: \Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien

Please do the following transformation steps. Please note that in this file in row 1 there are no column names .so please add some dummy names in rows so that further processing gets smoothy.


= Table.TransformColumnTypes(Quelle,{{"Column1", Int64.Type}, {"Column2", type text}, {"Column7", Int64.Type}, {"Column8", Int64.Type}, {"Column9", type text}})

= Table.TransformColumnTypes(#"Geänderter Typ", {{"Column11", type number}, {"Column9", type number}}, "en-US")


= Table.RenameColumns(#"Geänderter Typ Zahl aus Englischem Zellenformat CCF",{{"Column1", "BLZ"}, {"Column2", "Rating_od_wNote"}, {"Column3", "Rating_Kategorie"}, {"Column4", "Forderungsklasse"}, {"Column5", "Risikokundengruppe"}})

= Table.DuplicateColumn(#"Umbenannte Spalten", "Rating_od_wNote", "Copy of Rating_od_wNote")

= Table.RenameColumns(#"Hilfsspalte einfuegen",{{"Rating_od_wNote", "Rating_od_wNote_Original"}})

= Table.ReorderColumns(#"Umbenannte Spalten2",{"BLZ", "Rating_od_wNote_Original", "Copy of Rating_od_wNote", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column6", "Column7", "Column8", "Column9", "Column10", "Column11"})

= Table.RenameColumns(#"Neu angeordnete Spalten",{{"Copy of Rating_od_wNote", "Rating_od_wNote_Hilfsspalte"}})

= Table.RemoveColumns(#"Umbenannte Spalten3",{"Column6"})

= Table.ReplaceValue(#"Entfernte Spalten",".",",",Replacer.ReplaceText,{"Rating_od_wNote_Hilfsspalte"})

= Table.AddColumn(#"Punkt Durch Komma in Hilfsspalte", "Rating_od_wNote", each if ([Rating_Kategorie]="10" or [Rating_Kategorie]="11" or [Rating_Kategorie]="12" or [Forderungsklasse] ="1" or [Forderungsklasse] ="2" or [Forderungsklasse] ="3" or [Forderungsklasse] ="4" or [Forderungsklasse] ="5") then [Rating_od_wNote_Hilfsspalte] else [Rating_od_wNote_Original])

= Table.ReorderColumns(#"Bei Ratingkategorie 10 bis 12 Werte aus Hilfsspalte",{"BLZ", "Rating_od_wNote", "Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte", "Rating_Kategorie", "Forderungsklasse", "Risikokundengruppe", "Column7", "Column8", "Column9", "Column10", "Column11"})

= Table.RenameColumns(#"Neue Ratingspalte nach vorne schieben",{{"Column7", "Laufzeit_Von_(in_Tagen)"}, {"Column8", "Laufzeit_Bis_(in_Tagen)"}, {"Column9", "Risikokostensatz_Fix_(in_%)"}, {"Column10", "Risikokostensatz_Variabel_(in_%)"}, {"Column11", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"}})


= Table.ReplaceValue(#"Spalten umbenennen","-2","",Replacer.ReplaceText,{"Risikokundengruppe"})

= Table.SelectRows(#"Ersetzter Wert", each ([Rating_od_wNote] <> "-1,0" and [Rating_od_wNote] <> "-1.0" and [Rating_od_wNote] <> "-2,0" and [Rating_od_wNote] <> "-2.0"))

= Table.RemoveColumns(#"Unnoetige Ratings entfernen",{"Rating_od_wNote_Original", "Rating_od_wNote_Hilfsspalte"})

= Table.ReorderColumns(#"Hilfsspalten fürs Rating löschen",{"BLZ", "Rating_Kategorie", "Rating_od_wNote", "Forderungsklasse", "Risikokundengruppe", "Laufzeit_Von_(in_Tagen)", "Laufzeit_Bis_(in_Tagen)", "Risikokostensatz_Fix_(in_%)", "Risikokostensatz_Variabel_(in_%)", "Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)"})

= Table.RemoveColumns(#"Kategorie vor Rating",{"Risikokundengruppe"})


= Table.RenameColumns(#"Entfernte Spalten1",{{"Risikokostensatz_Fix_(in_%)", "Eigenkapitalkosten_Fix_(in_%)"}, {"Risikokostensatz_Variabel_(in_%)", "Eigenkapitalkosten_Variabel_(in_%)"}, {"Risikokostensatz_nicht_ausgenutzter_Rahmen_(in_%)", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"}})


As next step open the file with name EK_Basis_Primaerbanken.csv   in the following directory

U:\ Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Originaldateien\Primär-bzw Raiffeisenbanken
Please do the following transformation steps. Please note that in this file in row 1 there are  column names .


= Table.PromoteHeaders(Quelle)


= Table.SelectRows(#"Erste Zeile als Header", each ([BLZ] = "34"))


= Table.TransformColumnTypes(#"Gefilterte Zeilen",{{"Laufzeit_Von_(in_Tagen)", Int64.Type}, {"Laufzeit_Bis_(in_Tagen)", Int64.Type}, {"Eigenkapitalkosten_Fix_(in_%)", type number}, {"Eigenkapitalkosten_Variabel_(in_%)", type number}, {"Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)", type number}, {"BLZ", Int64.Type}})


Then as next step, the processed two files , I need to combine and on to the combined file I want to perform following transformations

= Table.SelectRows(Quelle, each ([BLZ] <> 55000))

= Table.ReplaceValue(#"nur 34 und 34000",365,366,Replacer.ReplaceValue,{"Laufzeit_Bis_(in_Tagen)"})

= Table.ReplaceValue(#"365_durch 366 ersetzen",366,367,Replacer.ReplaceValue,{"Laufzeit_Von_(in_Tagen)"})

= Table.SelectRows(#"Ersetzter Wert", each ([Rating_Kategorie] <> "9"))

= Table.DuplicateColumn(#"Ratingkategorie 9 wegfiltern", "Eigenkapitalkosten_Variabel_(in_%)", "Eigenkapitalkosten_Variabel_(in_%) - Kopie")

= Table.RemoveColumns(#"Duplizierte Spalte",{"Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"})

= Table.RenameColumns(#"Entfernte Spalten",{{"Eigenkapitalkosten_Variabel_(in_%) - Kopie", "Eigenkapitalkosten_nicht_ausgenutzter_Rahmen_(in_%)"}})

= Table.SelectRows(NaRahmenSatz, each ([Rating_od_wNote] <> "-1" and [Rating_od_wNote] <> "-2"))






Finally save the file in name    EK_Basis_Final in the following directory

U:\Controlling\FC\07 EDV-Projekte\SMART Vorkalk\Wartungstabellen\Befüllte Wartungstabellen (ECHTDATEN)\IMPORTASSISTENT_fuer_RK_und_EK\Upload_Dateien_Vorkalk

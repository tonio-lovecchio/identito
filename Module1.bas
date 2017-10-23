Attribute VB_Name = "Module1"
Sub SNPidentito2P()
'
' SNPidentito2P Macro
' Ceci est très basique

' Selection des plaques et attribution des variables

    Dim Fichier1 As Variant
    Dim Fichier2 As Variant
    Dim Fichier1_nom As String
    Dim Fichier2_nom As String
    Dim Path As String
    Dim Sheet1 As Variant
    Dim Torrent As Variant
    Dim Torrent_nom As String
    
    
 ' Ouverture des plaques QPCR
 
    MsgBox "Veuillez selectionner la première plaque de résultat QPCR", vbOKCancel
    Fichier1 = Application.GetOpenFilename("fichiers texte (*.txt),*.txt")
    If Fichier1 = False Then Exit Sub
    Workbooks.Open Filename:=Fichier1
    Fichier1_nom = ActiveWorkbook.Name
    MsgBox "Veuillez selectionner la seconde plaque de résultat QPCR", vbOKCancel
    Fichier2 = Application.GetOpenFilename("fichiers texte (*.txt),*.txt")
    If Fichier2 = False Then Exit Sub
    Workbooks.Open Filename:=Fichier2
    Fichier2_nom = ActiveWorkbook.Name
    
 ' Ouverture plaque Torrent
 
    MsgBox "Veuillez selectionner le fichier de résultat du Torrent Server Identitovigilance"
    Torrent = Application.GetOpenFilename("fichiers excel (*.xls),*.xls")
    If Torrent = False Then Exit Sub
    Workbooks.Open Filename:=Torrent
    Path = ThisWorkbook.Path
    
 ' Traitement première plaque
    
    Windows(Fichier1_nom).Activate
    Sheet = ActiveSheet.Name
    Cells(1, 1).Value = Fichier1_nom
    ActiveWorkbook.SaveAs Filename:= _
    Range("A1"), _
    FileFormat:=xlText, CreateBackup:=False
    Sheets(Sheet).Name = "Feuille1"
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
   ' Bloc de calcul des valeurs SNP plaque 1

' P1-bloc 1.
    
    Dim SNP1
    SNP1 = Range("C20").Value
    Select Case SNP1
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        End Select
        Selection.AutoFill Destination:=Range("G20:G51"), Type:=xlFillDefault
        Range("G20:G51").Select
    
  ' P1-bloc 2.
  
  Dim SNP2
    SNP2 = Range("C55").Value
    Select Case SNP2
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G55:G86"), Type:=xlFillDefault
            Range("G55:G86").Select
            
 ' P1-bloc 3.
 
 Dim SNP3
    SNP3 = Range("C90").Value
    Select Case SNP3
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier1_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G90:G121"), Type:=xlFillDefault
            Range("G90:G121").Select
            
 ' Traitement deuxième plaque
 
    Windows(Fichier2_nom).Activate
    Sheet = ActiveSheet.Name
    Cells(1, 1).Value = Fichier2_nom
    ActiveWorkbook.SaveAs Filename:= _
    Range("A1"), _
    FileFormat:=xlText, CreateBackup:=False
    Sheets(Sheet).Name = "Feuille1"
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
   ' Bloc de calcul des valeurs SNP plaque 2

' P1-bloc 4.
 
 Dim SNP4
    SNP4 = Range("C20").Value
    Select Case SNP4
         Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        End Select
        Selection.AutoFill Destination:=Range("G20:G51"), Type:=xlFillDefault
        Range("G20:G51").Select
            
   ' P1-bloc 5.
 
 Dim SNP5
    SNP5 = Range("C55").Value
    Select Case SNP5
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G55").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G55:G86"), Type:=xlFillDefault
            Range("G55:G86").Select
            
     ' P1-bloc 6.
 
 Dim SNP6
    SNP6 = Range("C90").Value
    Select Case SNP6
       Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier2_nom).Activate
            Range("G90").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G90:G121"), Type:=xlFillDefault
            Range("G90:G121").Select
            
 ' Tri des echantillons
 
        'Bloc 1
    Windows(Fichier1_nom).Activate
    Range("A20:J51").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B20:B51"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A20:J51")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Bloc 2
    Range("A54:J86").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B55:B86"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A54:J86")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Bloc 3
    Range("A89:J121").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B90:B121"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A89:J121")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Bloc 4
    Windows(Fichier2_nom).Activate
    Range("A20:J51").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B20:B51"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A20:J51")
        .Header = xlGuess
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Bloc 5
    Range("A54:J86").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B55:B86"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A54:J86")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
        'Bloc 6
    Range("A89:J121").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B90:B121"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A89:J121")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
       
 ' Création Table de comparaison
    
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:= _
        "\\D:\IDENTITO-adeplaceretrenommer", _
        FileFormat:=xlExcel8, CreateBackup:=False
        
 ' Report de la liste patients et des valeurs déduites des SNP depuis Plaque 1
    
    Windows(Fichier1_nom).Activate
    Range("B20:B51").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("A2").Select
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.AutoFit
    Windows(Fichier1_nom).Activate
    Range("C20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("B1").Select
    ActiveSheet.Paste
    Columns("B:B").EntireColumn.AutoFit
    Windows(Fichier1_nom).Activate
    Range("G20:G51").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(Fichier1_nom).Activate
    Range("C55").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("C1").Select
    ActiveSheet.Paste
    Windows(Fichier1_nom).Activate
    Range("G55:G86").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("C2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(Fichier1_nom).Activate
    Range("C90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("D1").Select
    ActiveSheet.Paste
    Windows(Fichier1_nom).Activate
    Range("G90:G121").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
' Report de la liste patients et des valeurs déduites des SNP depuis Plaque 2
    
    Windows(Fichier2_nom).Activate
    Range("C20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("E1").Select
    ActiveSheet.Paste
    Windows(Fichier2_nom).Activate
    Range("C55").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("F1").Select
    ActiveSheet.Paste
    Windows(Fichier2_nom).Activate
    Range("G20:G51").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier2_nom).Activate
    Range("G55:G86").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier2_nom).Activate
    Range("C90").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("G1").Select
    ActiveSheet.Paste
    Windows(Fichier2_nom).Activate
    Range("G90:G121").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    
    
    Windows("MacroIdentito.xls").Activate
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("I2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I33"), Type:=xlFillDefault
    Range("I2:I33").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[-8],RC[-7],RC[-6],RC[-5],RC[-4],RC[-3])"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J33"), Type:=xlFillDefault
    Range("J2:J33").Select
    Columns("I:I").EntireColumn.AutoFit
    Range("L2").Select
    Windows("MacroIdentito.xls").Activate
    Range("H2").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("L2:L33"), Type:=xlFillDefault
    Range("L2:L33").Select
    Range("L2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=J2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlNotEqual, _
        Formula1:="=J2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.AutoFill Destination:=Range("L2:L33"), Type:=xlFillDefault
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Valeur QPCR"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Valeur Torrent"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Comparaison"
    Columns("L:L").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    Workbooks.Open Filename:=Torrent
    Range("C4:C35").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("K2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks.Open Filename:=Torrent
    Range("B4:B35").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xls").Activate
    Range("M2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("M2:M33").Select
    Range("L2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=I2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Columns("M:M").EntireColumn.AutoFit
    Range("N2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-5],4,3)"
    Range("O2").Select
    ActiveCell.FormulaR1C1 = "=MID(RC[-2],4,3)"
    Range("N2:O2").Select
    Selection.AutoFill Destination:=Range("N2:O33"), Type:=xlFillDefault
    Range("N2:O33").Select
    Range("N2:N33").Select
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 255
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Range("N2").Select
    Selection.FormatConditions.Add Type:=xlCellValue, Operator:=xlEqual, _
        Formula1:="=$O2"
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.AutoFill Destination:=Range("N2:N33"), Type:=xlFillDefault
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 5296274
        .TintAndShade = 0
    End With
    ActiveWorkbook.Save
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("I1:N33").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("I1:N1").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Selection.Borders(xlInsideHorizontal).LineStyle = xlNone
    Range("I2:I33").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    Range("L2:L33").Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    Selection.Borders(xlInsideVertical).LineStyle = xlNone
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    ActiveWorkbook.Save
    
End Sub


Attribute VB_Name = "Module6"
Sub SNPidentito1P()
Attribute SNPidentito1P.VB_ProcData.VB_Invoke_Func = " \n14"
'
' SNPidentito1P Macro
'

' Selection des plaques et attribution des variables

    MsgBox "Veuillez selectionner la plaque de résultat QPCR", vbOKCancel
    Dim Fichier As Variant
    Dim Fichier_nom As String
    Dim Path As String
    Dim Sheet As Variant
    Dim Torrent As Variant
    Dim Torrent_nom As String
    Fichier = Application.GetOpenFilename("fichiers texte (*.txt),*.txt")
    If Fichier = False Then Exit Sub
    Workbooks.Open Filename:=Fichier
    Fichier_nom = ActiveWorkbook.Name
    
    MsgBox "Veuillez selectionner le fichier de résultat du Torrent Server Identitovigilance"
    Torrent = Application.GetOpenFilename("fichiers excel (*.xls),*.xls")
    If Torrent = False Then Exit Sub
    Workbooks.Open Filename:=Torrent
    Path = ThisWorkbook.Path
    Windows(Fichier_nom).Activate
    Sheet = ActiveSheet.Name
    Cells(1, 1).Value = Fichier_nom
    ActiveWorkbook.SaveAs Filename:= _
    Range("A1"), _
    FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    Sheets(Sheet).Name = "Feuille1"
    Columns("G:G").Select
    Selection.Insert Shift:=xlToRight, CopyOrigin:=xlFormatFromLeftOrAbove
    
   ' Bloc de calcul des valeurs SNP

' P1-bloc 1.
    
    Dim SNP1
    SNP1 = Range("C20").Value
    Select Case SNP1
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G20").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G20:G35"), Type:=xlFillDefault
            Range("G20:G35").Select
    
  ' P1-bloc 2.
  
  Dim SNP2
    SNP2 = Range("C39").Value
    Select Case SNP2
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G39").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G39:G54"), Type:=xlFillDefault
            Range("G39:G54").Select
            
 ' P1-bloc 3.
 
 Dim SNP3
    SNP3 = Range("C58").Value
    Select Case SNP3
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G58").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G58:G73"), Type:=xlFillDefault
            Range("G58:G73").Select
            
 ' P1-bloc 4.
 
 Dim SNP4
    SNP4 = Range("C77").Value
    Select Case SNP4
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G77").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G77:G92"), Type:=xlFillDefault
            Range("G77:G92").Select
            
   ' P1-bloc 5.
 
 Dim SNP5
    SNP5 = Range("C96").Value
    Select Case SNP5
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G96").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G96:G111"), Type:=xlFillDefault
            Range("G96:G111").Select
            
     ' P1-bloc 6.
 
 Dim SNP6
    SNP6 = Range("C115").Value
    Select Case SNP6
        Case "SNP1-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B8").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP2-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B9").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP3-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B10").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP4-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B11").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP5-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B12").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP6-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B13").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case "SNP7-260215"
            Windows("MacroIdentito.xls").Activate
            Range("B14").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        Case Else
            Windows("MacroIdentito.xls").Activate
            Range("B15").Select
            Selection.Copy
            Windows(Fichier_nom).Activate
            Range("G115").Select
            ActiveSheet.Paste
        End Select
    Selection.AutoFill Destination:=Range("G115:G130"), Type:=xlFillDefault
            Range("G115:G130").Select
            
    Range("A20:G35").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B20:B35"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A19:G35")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A38:G54").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B39:B54"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A38:G54")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=27
    Range("A57:G73").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B58:B73"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A57:G73")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=24
    Range("A76:G92").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B77:B92"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A76:G92")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=18
    Range("A95:G111").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B96:B111"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A95:G111")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=15
    Range("A114:G130").Select
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Clear
    ActiveWorkbook.Worksheets("Feuille1").Sort.SortFields. _
        Add Key:=Range("B115:B130"), SortOn:=xlSortOnValues, Order:=xlAscending, _
        DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Feuille1").Sort
        .SetRange Range("A114:G130")
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    ActiveWindow.SmallScroll Down:=-138
    
    ' Création du fichier de Comparaison
    
    Workbooks.Add
    ActiveWorkbook.SaveAs Filename:= _
        "\\D:\IDENTITO-adeplaceretrenommer.xlsx", _
        FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
    
    ' Report de la liste patients et des valeurs SNP
    
    Windows(Fichier_nom).Activate
    Range("B20:B35").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("A2").Select
    ActiveSheet.Paste
    Columns("A:A").EntireColumn.AutoFit
    Windows(Fichier_nom).Activate
    Range("C20").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("B1").Select
    ActiveSheet.Paste
    Columns("B:B").EntireColumn.AutoFit
    
    Windows(Fichier_nom).Activate
    Range("G20:G35").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("B2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=9
    Range("C39").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("C1").Select
    ActiveSheet.Paste
    Range("C2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Selection.ClearContents
    Windows(Fichier_nom).Activate
    Range("G39:G54").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=24
    Range("C58").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("D1").Select
    ActiveSheet.Paste
    Windows(Fichier_nom).Activate
    Range("G58:G73").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("D2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=18
    Range("C77").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("E1").Select
    ActiveSheet.Paste
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=15
    Range("C96").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("F1").Select
    ActiveSheet.Paste
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=-12
    Range("G77:G92").Select
    Application.CutCopyMode = False
    Selection.Copy
    ActiveWindow.SmallScroll Down:=15
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("E2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier_nom).Activate
    Range("G96:G111").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("F2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows(Fichier_nom).Activate
    ActiveWindow.SmallScroll Down:=18
    Range("C115").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("G1").Select
    ActiveSheet.Paste
    Windows(Fichier_nom).Activate
    Range("G115:G130").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("G2").Select
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Windows("MacroIdentito.xls").Activate
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("I2").Select
    Application.CutCopyMode = False
    ActiveCell.FormulaR1C1 = "=RC[-8]"
    Range("I2").Select
    Selection.AutoFill Destination:=Range("I2:I17"), Type:=xlFillDefault
    Range("I2:I17").Select
    Range("J2").Select
    ActiveCell.FormulaR1C1 = _
        "=CONCATENATE(RC[-8],RC[-7],RC[-6],RC[-5],RC[-4],RC[-3])"
    Range("J2").Select
    Selection.AutoFill Destination:=Range("J2:J17"), Type:=xlFillDefault
    Range("J2:J17").Select
    Columns("I:I").EntireColumn.AutoFit
    Range("L2").Select
    Windows(Fichier_nom).Activate
    Windows("MacroIdentito.xls").Activate
    Range("H2").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Selection.PasteSpecial Paste:=xlPasteFormulas, Operation:=xlNone, _
        SkipBlanks:=False, Transpose:=False
    Application.CutCopyMode = False
    Selection.AutoFill Destination:=Range("L2:L17"), Type:=xlFillDefault
    Range("L2:L17").Select
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
    Selection.AutoFill Destination:=Range("L2:L17"), Type:=xlFillDefault
    Range("L2:L17").Select
    Range("J1").Select
    ActiveCell.FormulaR1C1 = "Valeur QPCR"
    Range("K1").Select
    ActiveCell.FormulaR1C1 = "Valeur Torrent"
    Range("L1").Select
    ActiveCell.FormulaR1C1 = "Comparaison"
    Range("I1:L17").Select
    Columns("L:L").EntireColumn.AutoFit
    Columns("K:K").EntireColumn.AutoFit
    Columns("J:J").EntireColumn.AutoFit
    Columns("I:I").EntireColumn.AutoFit
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    
    Workbooks.Open Filename:=Torrent
    Range("C4:C19").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("K2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Workbooks.Open Filename:=Torrent
    Range("B4:B19").Select
    Selection.Copy
    Windows("IDENTITO-adeplaceretrenommer.xlsx").Activate
    Range("M2").Select
    ActiveSheet.Paste
    Application.CutCopyMode = False
    Range("M2:M17").Select
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
    Selection.AutoFill Destination:=Range("N2:O17"), Type:=xlFillDefault
    Range("N2:O17").Select
    Range("N2:N17").Select
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
    Selection.AutoFill Destination:=Range("N2:N17"), Type:=xlFillDefault
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
    Range("I1:N17").Select
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
    Range("I2:N17").Select
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
    Range("L2:L17").Select
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
    Range("G20").Select
    ActiveWorkbook.Save
    
End Sub

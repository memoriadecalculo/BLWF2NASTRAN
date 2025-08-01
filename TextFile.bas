Rem Attribute VBA_ModuleType=VBAModule
Option VBASupport 1
Option Explicit
Public delOldData As Integer

Sub BLWF_export_script()

'Defining variables
Dim counter As Long
Dim nLine As Long
Dim BLWF_script As String

'Clear script
ThisWorkbook.Sheets("BLWF_bat").Cells.Clear

nLine = 0

'Clear script's output
ThisWorkbook.Activate
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "ECHO OFF"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "CLS"

'Changing directory
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "CD /D %1"

'Initialize nCase variable
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "set nCASE=1"

'Initialize function addCase
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":addCASE"
nLine = nLine + 2

'Add cases
counter = 1
Do While counter <= Range("BLWF_nCases")
 Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %nCASE% EQU " + Format(counter) + " set case=" + Range("BLWF_nCases").Offset(counter, 0)
 nLine = nLine + 1
 counter = counter + 1
Loop
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %nCASE% EQU " + Format(counter) + " GOTO GLOBALexit"

'Execute function addCase
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "copy ..\" + Range("BLWF_exe") + " %case%\"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %ERRORLEVEL% NEQ 0 GOTO errorCOPY"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "cd %case%"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = Range("BLWF_exe") + " " + Range("BLWF_FileIn") + ".dat"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %ERRORLEVEL% NEQ 0 GOTO errorRUN"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "del /F /Q " + Range("BLWF_exe")
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %ERRORLEVEL% NEQ 0 GOTO errorDEL"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "set /A nCase=%nCase%+1"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "cd .."
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "..\" + Range("ADJ_exe") + " %case%\" + Range("BLWF_FileIn") + ".dat"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "IF %ERRORLEVEL% NEQ 0 GOTO errorREPLACE"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "GOTO addCASE"
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":errorCOPY"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "echo ERROR COPYING BLWF TO %case%!!!"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "GOTO GLOBALexit"
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":errorRUN"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "echo ERROR RUNNING %case%!!!"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "GOTO GLOBALexit"
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":errorDEL"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "echo ERROR DELETING BLWF FROM %case%!!!"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "GOTO GLOBALexit"
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":errorREPLACE"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "echo ERROR REPLACING in pl4 file FROM %case%!!!"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "GOTO GLOBALexit"
nLine = nLine + 2
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = ":GLOBALexit"
nLine = nLine + 1
Range("BLWF_scriptIni").Offset(nLine, 0).FormulaR1C1 = "pause"

'Exporting script
ThisWorkbook.Activate
BLWF_script = ThisWorkbook.Path + "\" + Range("BLWF_script")
Call ExportToTextFile(BLWF_script, ThisWorkbook.Sheets("BLWF_bat"), "", False, False)

End Sub

Sub BLWF_import_lift()

'Defining variables
Dim counter As Long
Dim CaseName As String
Dim caseSheet As String
Dim I As Integer
Dim J As Integer
Dim graphSerie

counter = 0
Do While counter < Range("BLWF_nCases")
 counter = counter + 1
 CaseName = Range("BLWF_case1").Offset(counter - 1, 0)
 caseSheet = CaseName + Range("BLWF_FileIn") + ".lift"
 
 'Importing new data
 If counter = 1 Then Call BLWF_import_file(".lift")
 
 'Deleting old data
 For I = Charts("GRAPH_D").SeriesCollection.Count To 1 Step -1
  If UCase(CaseName) = UCase(Charts("GRAPH_D").SeriesCollection(I).Name) Then
   If delOldData Then
    Charts("GRAPH_D").SeriesCollection(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I
 For I = Charts("GRAPH_L").SeriesCollection.Count To 1 Step -1
  If UCase(CaseName) = UCase(Charts("GRAPH_L").SeriesCollection(I).Name) Then
   If delOldData Then
    Charts("GRAPH_L").SeriesCollection(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I
 For I = Charts("GRAPH_S").SeriesCollection.Count To 1 Step -1
  If UCase(CaseName) = UCase(Charts("GRAPH_S").SeriesCollection(I).Name) Then
   If delOldData Then
    Charts("GRAPH_S").SeriesCollection(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I
 For I = Charts("GRAPH_B").SeriesCollection.Count To 1 Step -1
  If UCase(CaseName) = UCase(Charts("GRAPH_B").SeriesCollection(I).Name) Then
   If delOldData Then
    Charts("GRAPH_B").SeriesCollection(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I
 For I = Charts("GRAPH_T").SeriesCollection.Count To 1 Step -1
  If UCase(CaseName) = UCase(Charts("GRAPH_T").SeriesCollection(I).Name) Then
   If delOldData = vbYes Then
    Charts("GRAPH_T").SeriesCollection(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I

 'Calculating data to graphics
 Sheets("BLWF_lift").Range("O4:AS36").Copy
 Sheets(caseSheet).Activate
 Range("O4").Select
 ActiveSheet.Paste
 Application.CutCopyMode = False
 
 'Adding drag serie
 Set graphSerie = Charts("GRAPH_D").SeriesCollection.NewSeries
 graphSerie.Name = CaseName
 graphSerie.XValues = Sheets(caseSheet).Range("O7:O27")
 graphSerie.Values = Sheets(caseSheet).Range("U7:U27")
 
 'Adding lift serie
 Set graphSerie = Charts("GRAPH_L").SeriesCollection.NewSeries
 graphSerie.Name = CaseName
 graphSerie.XValues = Sheets(caseSheet).Range("O7:O27")
 graphSerie.Values = Sheets(caseSheet).Range("V7:V27")
 
 'Adding shear serie
 Set graphSerie = Charts("GRAPH_S").SeriesCollection.NewSeries
 graphSerie.Name = CaseName
 graphSerie.XValues = Sheets(caseSheet).Range("AF6:AF36")
 graphSerie.Values = Sheets(caseSheet).Range("AN6:AN36")
 
 'Adding bending serie
 Set graphSerie = Charts("GRAPH_B").SeriesCollection.NewSeries
 graphSerie.Name = CaseName
 graphSerie.XValues = Sheets(caseSheet).Range("AF6:AF36")
 graphSerie.Values = Sheets(caseSheet).Range("AO6:AO36")
 
 'Adding torsion serie
 Set graphSerie = Charts("GRAPH_T").SeriesCollection.NewSeries
 graphSerie.Name = CaseName
 graphSerie.XValues = Sheets(caseSheet).Range("AF6:AF36")
 graphSerie.Values = Sheets(caseSheet).Range("AP6:AP36")
 
 'Hidding worksheet
 Sheets(caseSheet).Visible = False
 
EndImport:
Loop

End Sub

Sub BLWF_export_dat()

Dim counter As Long
Dim CaseName As String

'Add cases
counter = 1
Do While counter <= Range("BLWF_nCases")
 'Adding title
 CaseName = Range("BLWF_nCases").Offset(counter, 0)
 Do While Len(CaseName) < 62
  CaseName = CaseName + " "
 Loop
 ThisWorkbook.Sheets("BLWF_dat").Cells(1, 2) = CaseName
 
 'Adding numbers
 CaseName = Format(Range("CASES_mach").Offset(counter, 0), "0.00000000")
 CaseName = CaseName + Format(Range("CASES_alpha").Offset(counter, 0), "00.0000000;-00.000000")
 CaseName = CaseName + Format(Range("CASES_Re").Offset(counter, 0), "000000000.")
 CaseName = CaseName + "   2.1330 "
 CaseName = CaseName + Format(Range("Swing") / 2, "0000.00000")
 CaseName = CaseName + "   1.0000    2.0000    0.0000"
 CaseName = Replace(CaseName, ",", ".")
 Range("BLWF_datMach") = CaseName
 
 'Exporting dat
 CaseName = ThisWorkbook.Path + "\" + Range("BLWF_nCases").Offset(counter, 0)
 If Len(Dir(CaseName, vbDirectory)) = 0 Then
  MkDir CaseName
 End If
 ThisWorkbook.Activate
 Call ExportToTextFile(CaseName + "\" + Range("BLWF_FileIn") + ".dat", ThisWorkbook.Sheets("BLWF_dat"), "", False, False)
 
 counter = counter + 1
Loop

End Sub

Sub BLWF_import_pl4()

'Defining variables
Dim counter As Long
Dim CaseName As String
Dim caseSheet As String
Dim I As Long
Dim graphSerie

counter = 0
Do While counter < Range("BLWF_nCases")
 counter = counter + 1
 CaseName = Range("BLWF_case1").Offset(counter - 1, 0)
 caseSheet = CaseName + Range("BLWF_FileIn") + ".pl4"
 
 'Importing new data
 If counter = 1 Then Call BLWF_import_file(".pl4")
 
 'Deleting old data
 For I = Worksheets.Count To 1 Step -1
  If UCase(caseSheet) = UCase(Worksheets(I).Name) Then
   If delOldData = vbYes Then

   Else
    GoTo EndImport
   End If
  End If
 Next I
 
 'Calculating data to graphics
 Sheets("BLWF_pl4").Range("J1:AO1651").Copy
 Sheets(caseSheet).Activate
 Range("J1").Select
 ActiveSheet.Paste
 Application.CutCopyMode = False
 
 'Hidding worksheet
 Sheets(caseSheet).Visible = False
 
EndImport:
Loop

End Sub

Private Sub BLWF_import_file(Extension As String)

'Defining variables
Dim counter As Long
Dim CaseName As String
Dim caseSheetName As String
Dim mainName As String
Dim inputNameFull As String
Dim I As Integer
Dim caseSheet

counter = 0
Do While counter < Range("BLWF_nCases")
 counter = counter + 1
 CaseName = Range("BLWF_case1").Offset(counter - 1, 0)
 caseSheetName = CaseName + Range("BLWF_FileIn") + Extension
 mainName = Range("BLWF_FileIn")
 
 'Deleting old data
 For I = Worksheets.Count To 1 Step -1
  If UCase(caseSheetName) = UCase(Worksheets(I).Name) Then
   If delOldData = vbYes Then
    Worksheets(I).Delete
   Else
    GoTo EndImport
   End If
  End If
 Next I
 
 'Importing new data
 inputNameFull = ThisWorkbook.Path + "\" + CaseName + "\" + Range("BLWF_FileIn") + Extension
 Workbooks.OpenText fileName:=inputNameFull, _
     Origin:=xlMSDOS, StartRow:=1, DataType:=xlDelimited, _
     TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=True, Tab:=False, _
     Semicolon:=False, Comma:=False, Space:=True, Other:=True, OtherChar:="=", _
     DecimalSeparator:=".", ThousandsSeparator:="," ', TrailingMinusNumbers:=True
 With ThisWorkbook.Sheets.Add
  .Name = caseSheetName
 End With
 Cells.Copy
 ThisWorkbook.Sheets(caseSheetName).Paste
 Application.CutCopyMode = False
 ActiveWorkbook.Close
 
EndImport:
Loop

End Sub

Sub BLWF_interpol_pl4(Surface As String)

Dim I As Long
Dim J As Long
Dim lineIni As Long
Dim lineEnd As Long
Dim colIni As Long
Dim colEnd As Long
Dim COL_X As Range
Dim COL_Y As Range
Dim counter As Long
Dim CaseName As String
Dim caseSheet As String

lineIni = 15
lineEnd = 115
colIni = 12
colEnd = 25
 
If LCase(Surface) = "lower" Then
 colIni = colIni + 16
 colEnd = colEnd + 16
End If

counter = 0
Do While counter < Range("BLWF_nCases")
 counter = counter + 1
 CaseName = Range("BLWF_case1").Offset(counter - 1, 0)
 caseSheet = CaseName + Range("BLWF_FileIn") + ".pl4"
 
 With Worksheets(caseSheet)
  For I = lineIni To lineEnd
   For J = colIni To colEnd
    If (.Cells(I, colIni - 1) < .Cells(lineIni - 3, J)) Or (.Cells(I, colIni - 1) > .Cells(lineIni - 3, J) + .Cells(lineIni - 2, J)) Then
     .Cells(I, J) = 0
    Else
     If LCase(Surface) = "upper" Then
      Select Case J
       Case colIni
        Set COL_X = .Range(.Cells((J - colIni) * 137 + lineIni - 1, 10), .Cells((J - colIni) * 137 + lineIni + 63, 10))
        Set COL_Y = .Range(.Cells((J - colIni) * 137 + lineIni - 1, 4), .Cells((J - colIni) * 137 + lineIni + 63, 4))
       Case colEnd
        Set COL_X = .Range(.Cells((J - colIni - 2) * 137 + lineIni - 1, 10), .Cells((J - colIni - 2) * 137 + lineIni + 63, 10))
        Set COL_Y = .Range(.Cells((J - colIni - 2) * 137 + lineIni - 1, 4), .Cells((J - colIni - 2) * 137 + lineIni + 63, 4))
       Case Else
        Set COL_X = .Range(.Cells((J - colIni - 1) * 137 + lineIni - 1, 10), .Cells((J - colIni - 1) * 137 + lineIni + 63, 10))
        Set COL_Y = .Range(.Cells((J - colIni - 1) * 137 + lineIni - 1, 4), .Cells((J - colIni - 1) * 137 + lineIni + 63, 4))
      End Select
     Else
      Select Case J
       Case colIni
        Set COL_X = .Range(.Cells((J - colIni) * 137 + lineIni + 64, 10), .Cells((J - colIni) * 137 + lineIni + 128, 10))
        Set COL_Y = .Range(.Cells((J - colIni) * 137 + lineIni + 64, 4), .Cells((J - colIni) * 137 + lineIni + 128, 4))
       Case colEnd
        Set COL_X = .Range(.Cells((J - colIni - 2) * 137 + lineIni + 64, 10), .Cells((J - colIni - 2) * 137 + lineIni + 128, 10))
        Set COL_Y = .Range(.Cells((J - colIni - 2) * 137 + lineIni + 64, 4), .Cells((J - colIni - 2) * 137 + lineIni + 128, 4))
       Case Else
        Set COL_X = .Range(.Cells((J - colIni - 1) * 137 + lineIni + 64, 10), .Cells((J - colIni - 1) * 137 + lineIni + 128, 10))
        Set COL_Y = .Range(.Cells((J - colIni - 1) * 137 + lineIni + 64, 4), .Cells((J - colIni - 1) * 137 + lineIni + 128, 4))
      End Select
     End If
     .Cells(I, J) = Interpola(COL_X, COL_Y, .Cells(I, colIni - 1)) * .Range("L3") ' + .Range("K3") Somente se aplica o diferencial de pressao, pois a pressao no infinito tambem existe dentro da asa.
    End If
   Next J
  Next I
 End With

Loop

End Sub


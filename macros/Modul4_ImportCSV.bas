Attribute VB_Name = "Modul4_ImportCSV"
'===============================================================================
' MODUL: Modul4_ImportCSV
' ZWECK: Robuster CSV/.dat-Import mit transaktionalem Austausch des Daten-Sheets.
' VERSION: 2.1.0
'===============================================================================
Option Explicit

Private Const MODUL As String = "ImportCSV"

Public Function ImportCSV() As Boolean
    Dim wb As Workbook
    Dim wsTemp As Worksheet
    Dim qt As QueryTable
    Dim zeilenAnzahl As Long
    Dim spaltenAnzahl As Long

    ImportCSV = False
    Set wb = ThisWorkbook

    Call Log(LVL_SECTION, MODUL, "CSV-Import gestartet")
    Call Log(LVL_INFO, MODUL, "Quelldatei", PFAD_DAT)

    On Error GoTo EH

    If Not DateiExistiertSicher(PFAD_DAT) Then
        Call Log(LVL_ERR, MODUL, "Quelldatei nicht gefunden", PFAD_DAT)
        MsgBox "Datei nicht gefunden:" & vbCrLf & PFAD_DAT & vbCrLf & vbCrLf & _
               "Bitte Pfad in Modul1_Konstanten prüfen.", _
               vbExclamation, "CSV-Import - Datei fehlt"
        Exit Function
    End If

    Call WorksheetLoeschen(wb, WS_DATEN_TEMP)

    Set wsTemp = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    wsTemp.Name = WS_DATEN_TEMP
    wsTemp.Cells.NumberFormat = "@"

    Call Log(LVL_DEBUG, MODUL, "Temporäres Worksheet erstellt", wsTemp.Name)

    Set qt = wsTemp.QueryTables.Add(Connection:="TEXT;" & PFAD_DAT, Destination:=wsTemp.Range("A1"))
    With qt
        .Name = "PHI_Import"
        .TextFileParseType = xlDelimited
        .TextFileSemicolonDelimiter = True
        .TextFileCommaDelimiter = False
        .TextFileTabDelimiter = False
        .TextFileSpaceDelimiter = False
        .TextFileConsecutiveDelimiter = False
        .TextFileTextQualifier = xlTextQualifierDoubleQuote
        .TextFileColumnDataTypes = TextSpaltenTypen(PFAD_DAT)
        .TextFileTrailingMinusNumbers = True
        .RefreshOnFileOpen = False
        .RefreshStyle = xlInsertDeleteCells
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveFormatting = True
        .Refresh BackgroundQuery:=False
    End With

    qt.Delete
    Set qt = Nothing

    zeilenAnzahl = LetzteZeile(wsTemp, 1)
    spaltenAnzahl = LetzteSpalte(wsTemp, 1)

    If zeilenAnzahl < 2 Then
        Call Log(LVL_ERR, MODUL, "Import ergab keine Datenzeilen", _
            "Nur Header oder leer. Zeilen: " & zeilenAnzahl)
        MsgBox "Die importierte Datei enthält keine Datensätze." & vbCrLf & _
               "Bitte Datei prüfen: " & PFAD_DAT, _
               vbExclamation, "Import - Keine Daten"
        GoTo Rollback
    End If

    If SpalteFinden(wsTemp, C_MATRIX, MODUL, True) <= 0 Then
        Call Log(LVL_ERR, MODUL, "Pflicht-Spalte 'Matrix' fehlt - Import wird verworfen")
        GoTo Rollback
    End If

    Call WorksheetLoeschen(wb, WS_DATEN)
    wsTemp.Name = WS_DATEN

    Call Log(LVL_OK, MODUL, "CSV-Import erfolgreich", _
        "Zeilen: " & (zeilenAnzahl - 1) & " Datensätze | Spalten: " & spaltenAnzahl)

    ImportCSV = True
    Exit Function

Rollback:
    On Error Resume Next
    If Not qt Is Nothing Then qt.Delete
    If Not wsTemp Is Nothing Then WorksheetLoeschen wb, wsTemp.Name
    On Error GoTo 0
    Exit Function

EH:
    Call LogFehler(MODUL, "Kritischer Fehler beim CSV-Import", Err.Number, Err.Description)

    On Error Resume Next
    If Not qt Is Nothing Then qt.Delete
    If Not wsTemp Is Nothing Then WorksheetLoeschen wb, wsTemp.Name
    On Error GoTo 0

    MsgBox "CSV-Import fehlgeschlagen!" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Details im Log-Sheet '" & WS_LOG & "'.", _
           vbCritical, "Import - Kritischer Fehler"
End Function

Private Function TextSpaltenTypen(ByVal dateiPfad As String) As Variant
    Dim dateiNr As Integer
    Dim ersteZeile As String
    Dim spaltenAnzahl As Long
    Dim i As Long
    Dim typen() As Variant

    On Error GoTo Fallback

    dateiNr = FreeFile
    Open dateiPfad For Input As #dateiNr
    Line Input #dateiNr, ersteZeile
    Close #dateiNr
    dateiNr = 0

    spaltenAnzahl = UBound(Split(ersteZeile, ";")) + 1
    If spaltenAnzahl < 1 Then spaltenAnzahl = 1

    ReDim typen(0 To spaltenAnzahl - 1)
    For i = 0 To spaltenAnzahl - 1
        typen(i) = xlTextFormat
    Next i

    TextSpaltenTypen = typen
    Exit Function

Fallback:
    On Error Resume Next
    If dateiNr <> 0 Then Close #dateiNr
    On Error GoTo 0

    ReDim typen(0 To 0)
    typen(0) = xlTextFormat
    TextSpaltenTypen = typen
End Function

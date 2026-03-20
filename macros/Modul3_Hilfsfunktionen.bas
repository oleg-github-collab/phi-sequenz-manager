Attribute VB_Name = "Modul3_Hilfsfunktionen"
'===============================================================================
' MODUL: Modul3_Hilfsfunktionen
' ZWECK: Wiederverwendbare Hilfsfunktionen fuer robustes Excel/VBA-Verhalten.
' VERSION: 2.1.0
'===============================================================================
Option Explicit

Private Type AppStatusSnapshot
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    Calculation As XlCalculation
    DisplayAlerts As Boolean
    StatusBar As Variant
End Type

Private m_AppStatus As AppStatusSnapshot
Private m_AppStatusDepth As Long

'===============================================================================
' APPLICATION STATE
'===============================================================================

Public Sub AppOptimieren()
    If m_AppStatusDepth = 0 Then
        m_AppStatus.ScreenUpdating = Application.ScreenUpdating
        m_AppStatus.EnableEvents = Application.EnableEvents
        m_AppStatus.Calculation = Application.Calculation
        m_AppStatus.DisplayAlerts = Application.DisplayAlerts
        m_AppStatus.StatusBar = Application.StatusBar

        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        Application.DisplayAlerts = False
    End If

    m_AppStatusDepth = m_AppStatusDepth + 1
End Sub

Public Sub AppWiederherstellen()
    If m_AppStatusDepth <= 0 Then
        m_AppStatusDepth = 0
        Application.StatusBar = False
        Exit Sub
    End If

    m_AppStatusDepth = m_AppStatusDepth - 1
    If m_AppStatusDepth > 0 Then Exit Sub

    Application.ScreenUpdating = m_AppStatus.ScreenUpdating
    Application.EnableEvents = m_AppStatus.EnableEvents
    Application.Calculation = m_AppStatus.Calculation
    Application.DisplayAlerts = m_AppStatus.DisplayAlerts

    If IsEmpty(m_AppStatus.StatusBar) Then
        Application.StatusBar = False
    Else
        Application.StatusBar = m_AppStatus.StatusBar
    End If
End Sub

Public Function SekundenSeit(ByVal Startzeit As Date) As Double
    SekundenSeit = (Now - Startzeit) * 86400#
    If SekundenSeit < 0 Then SekundenSeit = 0
End Function

'===============================================================================
' FILESYSTEM / WORKSHEET
'===============================================================================

Public Function DateiExistiertSicher(ByVal Pfad As String) As Boolean
    On Error Resume Next
    DateiExistiertSicher = (LenB(Dir$(Pfad, vbNormal)) > 0)
    On Error GoTo 0
End Function

Public Function WorksheetHolen(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    On Error Resume Next
    Set WorksheetHolen = wb.Worksheets(wsName)
    On Error GoTo 0
End Function

Public Function WorksheetPflicht(ByVal wb As Workbook, ByVal wsName As String) As Worksheet
    Set WorksheetPflicht = WorksheetHolen(wb, wsName)
    If WorksheetPflicht Is Nothing Then
        Err.Raise 9001, "WorksheetPflicht", _
            "Pflicht-Worksheet '" & wsName & "' nicht gefunden."
    End If
End Function

Public Sub WorksheetLoeschen(ByVal wb As Workbook, ByVal wsName As String)
    Dim ws As Worksheet
    Dim errNr As Long
    Dim errText As String

    Set ws = WorksheetHolen(wb, wsName)
    If ws Is Nothing Then Exit Sub

    Dim vorherAlerts As Boolean
    vorherAlerts = Application.DisplayAlerts

    On Error GoTo EH
    Application.DisplayAlerts = False
    ws.Delete

Aufraeumen:
    Application.DisplayAlerts = vorherAlerts
    If errNr <> 0 Then
        Err.Raise errNr, "WorksheetLoeschen", errText
    End If
    Exit Sub

EH:
    errNr = Err.Number
    errText = Err.Description
    Resume Aufraeumen
End Sub

Public Function LetzteZeile(ByVal ws As Worksheet, ByVal Spalte As Long) As Long
    If Spalte <= 0 Then
        LetzteZeile = 1
        Exit Function
    End If

    LetzteZeile = ws.Cells(ws.Rows.Count, Spalte).End(xlUp).Row
    If LetzteZeile < 1 Then LetzteZeile = 1
End Function

Public Function LetzteSpalte(ByVal ws As Worksheet, Optional ByVal Zeile As Long = 1) As Long
    LetzteSpalte = ws.Cells(Zeile, ws.Columns.Count).End(xlToLeft).Column
    If LetzteSpalte < 1 Then LetzteSpalte = 1
End Function

'===============================================================================
' SEARCH / LOOKUP
'===============================================================================

Public Function FindeExakt(ByVal Suchbereich As Range, _
                           ByVal SuchText As String, _
                           Optional ByVal SuchReihenfolge As XlSearchOrder = xlByRows) As Range
    If Suchbereich Is Nothing Then Exit Function

    Dim startZelle As Range
    Set startZelle = Suchbereich.Cells(Suchbereich.Cells.Count)

    Set FindeExakt = Suchbereich.Find(What:=SuchText, _
                                      After:=startZelle, _
                                      LookIn:=xlValues, _
                                      LookAt:=xlWhole, _
                                      SearchOrder:=SuchReihenfolge, _
                                      SearchDirection:=xlNext, _
                                      MatchCase:=False, _
                                      SearchFormat:=False)
End Function

Public Function SpalteFinden(ByVal ws As Worksheet, ByVal Header As String, _
                              Optional ByVal LogModul As String = "SpalteFinden", _
                              Optional ByVal Pflicht As Boolean = True) As Long
    Dim gefunden As Range
    Set gefunden = FindeExakt(ws.UsedRange, Header, xlByRows)

    If gefunden Is Nothing Then
        SpalteFinden = -1
        If Pflicht Then
            Call Log(LVL_ERR, LogModul, "Pflicht-Spalte nicht gefunden", _
                "Worksheet: '" & ws.Name & "' | Header: '" & Header & "'")
        Else
            Call Log(LVL_WARN, LogModul, "Optionale Spalte nicht gefunden", _
                "Worksheet: '" & ws.Name & "' | Header: '" & Header & "'")
        End If
        Exit Function
    End If

    SpalteFinden = gefunden.Column
    Call Log(LVL_DEBUG, LogModul, "Spalte gefunden", _
        "'" & Header & "' -> Spalte " & gefunden.Column & " in " & ws.Name)
End Function

Public Function SpaltenMapErstellen(ByVal ws As Worksheet, _
                                    ParamArray Headers() As Variant) As Object
    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Dim headerWert As Variant
    For Each headerWert In Headers
        dict(CStr(headerWert)) = SpalteFinden(ws, CStr(headerWert), "SpaltenMap", False)
    Next headerWert

    Set SpaltenMapErstellen = dict
End Function

'===============================================================================
' REGEX
'===============================================================================

Public Function RegExErstellen(ByVal Muster As String, _
                               Optional ByVal IgnoreCase As Boolean = False, _
                               Optional ByVal GlobalSuche As Boolean = False) As Object
    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Pattern = Muster
    re.IgnoreCase = IgnoreCase
    re.Global = GlobalSuche
    Set RegExErstellen = re
End Function

Public Function ProbenRegExMuster(ByVal Praefix As String, _
                                  ByVal Jahr As Integer, _
                                  ByVal JahrAlt As Integer) As String
    ProbenRegExMuster = "^" & Praefix & _
        "(" & Format$(Jahr, "00") & "|" & Format$(JahrAlt, "00") & ")" & _
        "\d{5}$"
End Function

'===============================================================================
' PARAMETER-BLOECKE
'===============================================================================

Public Function ParameterKonfigMapLaden(ByVal wsParam As Worksheet) As Object
    Const MODUL As String = "ParameterKonfigMap"

    Dim dict As Object
    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    On Error GoTo EH

    Dim headerZelle As Range
    Set headerZelle = FindeExakt(wsParam.UsedRange, C_BEZEICHNUNG, xlByRows)
    If headerZelle Is Nothing Then
        Call Log(LVL_WARN, MODUL, "Konfigurations-Header nicht gefunden", C_BEZEICHNUNG)
        Set ParameterKonfigMapLaden = dict
        Exit Function
    End If

    Dim zeile As Long
    zeile = headerZelle.Row + 1

    Dim guard As Long
    Do While guard < LOOP_GUARD
        Dim bez As String
        bez = Trim$(CStr(wsParam.Cells(zeile, headerZelle.Column).Value))
        If Len(bez) = 0 Then Exit Do

        dict(bez) = Array( _
            CStr(wsParam.Cells(zeile, headerZelle.Column + 1).Value), _
            CStr(wsParam.Cells(zeile, headerZelle.Column + 2).Value))

        zeile = zeile + 1
        guard = guard + 1
    Loop

    If guard >= LOOP_GUARD Then
        Call Log(LVL_ERR, MODUL, "Endlosschleifen-Schutz beim Laden der Parameter-Konfiguration")
    End If

    Set ParameterKonfigMapLaden = dict
    Exit Function

EH:
    Call LogFehler(MODUL, "Fehler beim Laden der Parameter-Konfiguration", _
        Err.Number, Err.Description)
    Set ParameterKonfigMapLaden = dict
End Function

Public Function ParameterBlockLesen(ByVal wsParam As Worksheet, _
                                     ByVal BlockHeader As String, _
                                     ByVal BlockEnde As String, _
                                     ByVal ProbenNummer As String) As String()
    Const MODUL As String = "ParameterBlockLesen"

    Dim ergebnis() As String
    ergebnis = LeeresStringRaster()

    On Error GoTo EH

    Dim startZelle As Range
    Set startZelle = FindeExakt(wsParam.UsedRange, BlockHeader, xlByRows)
    If startZelle Is Nothing Then
        Call Log(LVL_ERR, MODUL, "Block-Header nicht gefunden", "Header: '" & BlockHeader & "'")
        ParameterBlockLesen = ergebnis
        Exit Function
    End If

    Dim parameterMap As Object
    Set parameterMap = ParameterKonfigMapLaden(wsParam)

    Dim offset As Long
    Dim anzahl As Long
    offset = 1

    Do While Trim$(CStr(startZelle.Offset(offset, 0).Value)) <> BlockEnde
        If offset > LOOP_GUARD Then
            Call Log(LVL_ERR, MODUL, "Endlos-Schleifen-Schutz ausgelöst", _
                "Block: '" & BlockHeader & "' | Ende-Marker: '" & BlockEnde & "'")
            ParameterBlockLesen = ergebnis
            Exit Function
        End If

        If Len(Trim$(CStr(startZelle.Offset(offset, 0).Value))) > 0 Then
            anzahl = anzahl + 1
        End If

        offset = offset + 1
    Loop

    If anzahl = 0 Then
        Call Log(LVL_WARN, MODUL, "Block leer", "Block: '" & BlockHeader & "'")
        ParameterBlockLesen = ergebnis
        Exit Function
    End If

    ReDim ergebnis(1 To 3, 1 To anzahl)

    Dim index As Long
    index = 0
    For offset = 1 To LOOP_GUARD
        Dim bezWert As String
        bezWert = Trim$(CStr(startZelle.Offset(offset, 0).Value))

        If bezWert = BlockEnde Then Exit For
        If Len(bezWert) = 0 Then GoTo NaechsterBlockEintrag

        index = index + 1
        If bezWert = PHI50_MATRIX Then
            If Len(ProbenNummer) > 0 Then
                ergebnis(1, index) = ProbenNummer
            Else
                ergebnis(1, index) = PHI50_MATRIX
            End If
        Else
            ergebnis(1, index) = bezWert
        End If

        If parameterMap.Exists(bezWert) Then
            Dim configWerte As Variant
            configWerte = parameterMap(bezWert)
            ergebnis(2, index) = CStr(configWerte(0))
            ergebnis(3, index) = CStr(configWerte(1))
        Else
            Call Log(LVL_WARN, MODUL, "Bezeichnung nicht in Konfiguration gefunden", bezWert)
        End If

NaechsterBlockEintrag:
    Next offset

    Call Log(LVL_DEBUG, MODUL, "Block gelesen", "'" & BlockHeader & "' -> " & index & " Einträge")
    ParameterBlockLesen = ergebnis
    Exit Function

EH:
    Call LogFehler(MODUL, "Fehler beim Lesen von Block '" & BlockHeader & "'", _
        Err.Number, Err.Description)
    ParameterBlockLesen = LeeresStringRaster()
End Function

Private Function LeeresStringRaster() As String()
    Dim arr() As String
    ReDim arr(1 To 3, 0 To 0)
    LeeresStringRaster = arr
End Function

'===============================================================================
' VALIDIERUNG
'===============================================================================

Public Function ValidiereSpalten(ByVal spaltenMap As Object, _
                                  ByVal Modul As String, _
                                  ParamArray PflichtSpalten() As Variant) As Boolean
    ValidiereSpalten = True

    Dim eintrag As Variant
    For Each eintrag In PflichtSpalten
        Dim spaltenName As String
        spaltenName = CStr(eintrag)

        If Not spaltenMap.Exists(spaltenName) Then
            Call Log(LVL_ERR, Modul, "Pflicht-Spalte fehlt im Mapping", spaltenName)
            ValidiereSpalten = False
        ElseIf spaltenMap(spaltenName) <= 0 Then
            Call Log(LVL_ERR, Modul, "Pflicht-Spalte hat ungültigen Index", _
                spaltenName & " -> " & spaltenMap(spaltenName))
            ValidiereSpalten = False
        End If
    Next eintrag
End Function

Public Function BegrenzeInteger(ByVal Wert As Long, _
                                ByVal Minimum As Integer, _
                                ByVal Maximum As Integer) As Integer
    If Wert < Minimum Then
        BegrenzeInteger = Minimum
    ElseIf Wert > Maximum Then
        BegrenzeInteger = Maximum
    Else
        BegrenzeInteger = CInt(Wert)
    End If
End Function

'===============================================================================
' STATUS / PROGRESS
'===============================================================================

Public Sub Fortschritt(ByVal Modul As String, ByVal Schritt As Long, _
                        ByVal Gesamt As Long, Optional ByVal Extra As String = "")
    Dim prozent As Integer
    If Gesamt > 0 Then
        prozent = CInt((Schritt / Gesamt) * 100)
    Else
        prozent = 100
    End If

    Application.StatusBar = "[" & Modul & "] " & prozent & "% (" & _
                            Schritt & "/" & Gesamt & ") " & Extra
End Sub

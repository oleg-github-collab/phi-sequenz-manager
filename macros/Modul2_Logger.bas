Attribute VB_Name = "Modul2_Logger"
'===============================================================================
' MODUL: Modul2_Logger
' ZWECK: Zentrales Logging in ein dediziertes Worksheet.
' VERSION: 2.1.0
'===============================================================================
Option Explicit

Private m_LogWS As Worksheet
Private m_SessionStart As Date
Private m_Zeile As Long
Private m_Aktiv As Boolean

Private Const LC_NR As Byte = 1
Private Const LC_ZEIT As Byte = 2
Private Const LC_ELAPSED As Byte = 3
Private Const LC_LEVEL As Byte = 4
Private Const LC_MODUL As Byte = 5
Private Const LC_MSG As Byte = 6
Private Const LC_DETAIL As Byte = 7

Public Sub LogInit()
    On Error GoTo EH

    m_Aktiv = False
    m_SessionStart = Now
    Set m_LogWS = WorksheetHolen(ThisWorkbook, WS_LOG)

    If m_LogWS Is Nothing Then
        Set m_LogWS = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        m_LogWS.Name = WS_LOG
        Call LogSheet_Formatieren
    ElseIf LenB(CStr(m_LogWS.Cells(1, LC_NR).Value)) = 0 Then
        Call LogSheet_Formatieren
    End If

    m_Zeile = LetzteZeile(m_LogWS, LC_NR) + 1
    If m_Zeile < 2 Then m_Zeile = 2

    m_Aktiv = True
    Call Log(LVL_SECTION, "Logger", _
        "SESSION START: " & Format$(Now, "yyyy-mm-dd hh:nn:ss"))
    Exit Sub

EH:
    Debug.Print "Logger-Init-Fehler: " & Err.Number & " - " & Err.Description
    Set m_LogWS = Nothing
    m_Aktiv = False
End Sub

Public Sub Log(ByVal Level As String, ByVal Modul As String, _
               ByVal Nachricht As String, Optional ByVal Details As String = "")
    If Not m_Aktiv Then
        Debug.Print "[" & Level & "] " & Modul & ": " & Nachricht & IIf(Details <> "", " | " & Details, "")
        Exit Sub
    End If

    On Error Resume Next

    Dim elapsed As String
    elapsed = Format$(SekundenSeit(m_SessionStart), "0.000") & "s"

    With m_LogWS
        .Cells(m_Zeile, LC_NR).Value = m_Zeile - 1
        .Cells(m_Zeile, LC_ZEIT).Value = Format$(Now, "hh:nn:ss")
        .Cells(m_Zeile, LC_ELAPSED).Value = elapsed
        .Cells(m_Zeile, LC_LEVEL).Value = Level
        .Cells(m_Zeile, LC_MODUL).Value = Modul
        .Cells(m_Zeile, LC_MSG).Value = Nachricht
        .Cells(m_Zeile, LC_DETAIL).Value = Details
        Call LogZeileFormatieren(m_LogWS, m_Zeile, Level)
    End With

    If Level <> LVL_DEBUG Then
        Application.StatusBar = "[" & Level & "] " & Modul & ": " & Nachricht
    End If

    m_Zeile = m_Zeile + 1
    On Error GoTo 0
End Sub

Public Sub LogFehler(ByVal Modul As String, ByVal Kontext As String, _
                     ByVal ErrNr As Long, ByVal ErrBeschr As String)
    Call Log(LVL_ERR, Modul, Kontext, "Err " & ErrNr & ": " & ErrBeschr)
End Sub

Public Sub LogAbschluss(ByVal Erfolg As Boolean)
    If Not m_Aktiv Then Exit Sub

    Dim dauer As String
    dauer = Format$(SekundenSeit(m_SessionStart), "0.00") & "s"

    If Erfolg Then
        Call Log(LVL_OK, "Logger", "Session erfolgreich beendet", "Gesamtdauer: " & dauer)
    Else
        Call Log(LVL_ERR, "Logger", "Session mit Fehler beendet", "Gesamtdauer: " & dauer)
    End If

    Call Log(LVL_SECTION, "Logger", String$(60, "="))

    On Error Resume Next
    m_LogWS.Columns.AutoFit
    m_LogWS.Columns(LC_MSG).ColumnWidth = 55
    m_LogWS.Columns(LC_DETAIL).ColumnWidth = 65
    m_LogWS.Range(m_LogWS.Cells(1, 1), m_LogWS.Cells(1, LC_DETAIL)).AutoFilter
    On Error GoTo 0

    m_Aktiv = False
    Application.StatusBar = False
End Sub

Private Sub LogSheet_Formatieren()
    With m_LogWS
        .Tab.Color = RGB(0, 80, 200)
        .Cells.Font.Name = "Consolas"
        .Cells.Font.Size = 9

        .Cells(1, LC_NR).Value = "#"
        .Cells(1, LC_ZEIT).Value = "Zeit"
        .Cells(1, LC_ELAPSED).Value = "Elapsed"
        .Cells(1, LC_LEVEL).Value = "Level"
        .Cells(1, LC_MODUL).Value = "Modul"
        .Cells(1, LC_MSG).Value = "Nachricht"
        .Cells(1, LC_DETAIL).Value = "Details"

        With .Range(.Cells(1, 1), .Cells(1, LC_DETAIL))
            .Font.Bold = True
            .Font.Size = 10
            .Interior.Color = RGB(15, 25, 50)
            .Font.Color = RGB(0, 200, 255)
        End With

        .Columns(LC_NR).ColumnWidth = 5
        .Columns(LC_ZEIT).ColumnWidth = 12
        .Columns(LC_ELAPSED).ColumnWidth = 10
        .Columns(LC_LEVEL).ColumnWidth = 11
        .Columns(LC_MODUL).ColumnWidth = 22
        .Columns(LC_MSG).ColumnWidth = 55
        .Columns(LC_DETAIL).ColumnWidth = 65
        .Rows(1).RowHeight = 22
    End With
End Sub

Private Sub LogZeileFormatieren(ByVal wsLog As Worksheet, ByVal Zeile As Long, ByVal Level As String)
    Dim gesamteZeile As Range
    Set gesamteZeile = wsLog.Range(wsLog.Cells(Zeile, 1), wsLog.Cells(Zeile, LC_DETAIL))

    gesamteZeile.Interior.Pattern = xlSolid
    gesamteZeile.Font.Bold = False
    gesamteZeile.Font.Color = RGB(0, 0, 0)

    Select Case Level
        Case LVL_ERR
            gesamteZeile.Interior.Color = RGB(255, 235, 235)
            wsLog.Cells(Zeile, LC_LEVEL).Interior.Color = RGB(220, 40, 40)
            wsLog.Cells(Zeile, LC_LEVEL).Font.Color = RGB(255, 255, 255)
            gesamteZeile.Font.Bold = True

        Case LVL_WARN
            gesamteZeile.Interior.Color = RGB(255, 247, 212)
            wsLog.Cells(Zeile, LC_LEVEL).Interior.Color = RGB(220, 160, 0)

        Case LVL_OK
            wsLog.Cells(Zeile, LC_LEVEL).Interior.Color = RGB(0, 180, 80)
            wsLog.Cells(Zeile, LC_LEVEL).Font.Color = RGB(255, 255, 255)

        Case LVL_DEBUG
            gesamteZeile.Font.Color = RGB(120, 120, 120)

        Case LVL_SECTION
            gesamteZeile.Interior.Color = RGB(20, 40, 80)
            gesamteZeile.Font.Color = RGB(100, 180, 255)
            gesamteZeile.Font.Bold = True
    End Select
End Sub

Attribute VB_Name = "Modul6_Hauptsteuerung"
Option Explicit

Private Const MODUL As String = "Hauptsteuerung"

Public Sub MakrosAusfuehren()
    Dim startZeit As Date
    Dim importErfolgreich As Boolean
    Dim sequenzErfolgreich As Boolean
    Dim dauer As String

    startZeit = Now
    Call LogInit
    Call Log(LVL_SECTION, MODUL, "PHI GC Sequenz-Manager")
    Call Log(LVL_INFO, MODUL, "Vollstaendiger Durchlauf gestartet")

    On Error GoTo EH
    Call AppOptimieren

    Call Log(LVL_SECTION, MODUL, "Phase 1: CSV-Import")
    importErfolgreich = ImportCSV()
    If Not importErfolgreich Then GoTo FehlerAbschluss

    Call Log(LVL_SECTION, MODUL, "Phase 2: Sequenz-Erzeugung")
    sequenzErfolgreich = SequenzSchreiben()
    If Not sequenzErfolgreich Then GoTo FehlerAbschluss

    Call AppWiederherstellen
    dauer = Format$(SekundenSeit(startZeit), "0.00") & "s"
    Call Log(LVL_OK, MODUL, "Vollstaendiger Durchlauf erfolgreich", "Gesamtdauer: " & dauer)
    Call LogAbschluss(True)

    MsgBox "Sequenz erfolgreich erstellt!" & vbCrLf & vbCrLf & _
           "Gesamtdauer: " & dauer & vbCrLf & _
           "Details im Log-Sheet '" & WS_LOG & "'.", _
           vbInformation, "PHI Sequenz-Manager - Erfolg"
    Exit Sub

FehlerAbschluss:
    Call AppWiederherstellen
    dauer = Format$(SekundenSeit(startZeit), "0.00") & "s"
    Call Log(LVL_ERR, MODUL, "Durchlauf fehlgeschlagen", "Dauer: " & dauer)
    Call LogAbschluss(False)

    MsgBox "Makro abgebrochen." & vbCrLf & _
           "Bitte Log-Sheet '" & WS_LOG & "' pruefen.", _
           vbCritical, "PHI Sequenz-Manager - Fehler"
    Exit Sub

EH:
    Call AppWiederherstellen
    Call LogFehler(MODUL, "Unbehandelter Fehler in MakrosAusfuehren", Err.Number, Err.Description)
    Call LogAbschluss(False)

    MsgBox "Unerwarteter Fehler!" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Details im Log-Sheet '" & WS_LOG & "'.", _
           vbCritical, "PHI Sequenz-Manager - Kritischer Fehler"
End Sub

Public Sub NurImportAusfuehren()
    Dim ok As Boolean

    Call LogInit
    Call Log(LVL_INFO, MODUL, "Nur-Import-Modus gestartet")

    On Error GoTo EH
    Call AppOptimieren
    ok = ImportCSV()
    Call AppWiederherstellen
    Call LogAbschluss(ok)

    If ok Then
        MsgBox "Import erfolgreich. Daten-Sheet '" & WS_DATEN & "' wurde erstellt.", _
               vbInformation, "Import - OK"
    Else
        MsgBox "Import fehlgeschlagen. Log pruefen.", vbCritical, "Import - Fehler"
    End If
    Exit Sub

EH:
    Call AppWiederherstellen
    Call LogFehler(MODUL, "Fehler in NurImportAusfuehren", Err.Number, Err.Description)
    Call LogAbschluss(False)
    MsgBox "Import fehlgeschlagen: " & Err.Description, vbCritical, "Import - Fehler"
End Sub

Public Sub NurSequenzAusfuehren()
    Dim wsDaten As Worksheet
    Dim ok As Boolean

    Set wsDaten = WorksheetHolen(ThisWorkbook, WS_DATEN)
    If wsDaten Is Nothing Then
        MsgBox "Daten-Sheet '" & WS_DATEN & "' nicht gefunden." & vbCrLf & _
               "Bitte zuerst den Import ausfuehren.", _
               vbExclamation, "Sequenz - Voraussetzung fehlt"
        Exit Sub
    End If

    Call LogInit
    Call Log(LVL_INFO, MODUL, "Nur-Sequenz-Modus gestartet")

    On Error GoTo EH
    Call AppOptimieren
    ok = SequenzSchreiben()
    Call AppWiederherstellen
    Call LogAbschluss(ok)

    If ok Then
        MsgBox "Sequenz erfolgreich erstellt.", vbInformation, "Sequenz - OK"
    Else
        MsgBox "Sequenz-Erzeugung fehlgeschlagen. Log pruefen.", vbCritical, "Sequenz - Fehler"
    End If
    Exit Sub

EH:
    Call AppWiederherstellen
    Call LogFehler(MODUL, "Fehler in NurSequenzAusfuehren", Err.Number, Err.Description)
    Call LogAbschluss(False)
    MsgBox "Sequenz fehlgeschlagen: " & Err.Description, vbCritical, "Sequenz - Fehler"
End Sub

Public Sub LogAnzeigen()
    Dim wsLog As Worksheet

    Set wsLog = WorksheetHolen(ThisWorkbook, WS_LOG)
    If wsLog Is Nothing Then
        MsgBox "Log-Sheet '" & WS_LOG & "' nicht gefunden." & vbCrLf & _
               "Bitte zuerst ein Makro ausfuehren.", _
               vbInformation, "Log"
        Exit Sub
    End If

    wsLog.Activate
End Sub

Public Sub LogLeeren()
    Dim antwort As VbMsgBoxResult

    antwort = MsgBox("Log-Sheet '" & WS_LOG & "' komplett leeren?" & vbCrLf & _
                     "Diese Aktion kann nicht rueckgaengig gemacht werden.", _
                     vbQuestion + vbYesNo, "Log leeren")
    If antwort <> vbYes Then Exit Sub

    On Error GoTo EH
    Call WorksheetLoeschen(ThisWorkbook, WS_LOG)
    MsgBox "Log-Sheet geloescht. Es wird beim naechsten Makro-Aufruf neu erstellt.", _
           vbInformation, "Log geleert"
    Exit Sub

EH:
    MsgBox "Log-Sheet konnte nicht geloescht werden: " & Err.Description, _
           vbCritical, "Log - Fehler"
End Sub

Public Sub TastenkombinationAktivieren()
    Application.OnKey "^+{RETURN}", MakroVollqualifiziert("MakrosAusfuehren")
    MsgBox "Tastenkombination aktiv:" & vbCrLf & _
           "Ctrl + Shift + Enter -> MakrosAusfuehren()", _
           vbInformation, "Tastenkombination"
End Sub

Public Sub TastenkombinationDeaktivieren()
    Application.OnKey "^+{RETURN}"
End Sub

Public Sub Selbsttest()
    Dim fehler As Integer
    Dim wsNames(0 To 2) As String
    Dim i As Integer

    Call LogInit
    Call Log(LVL_SECTION, MODUL, "Selbsttest")

    wsNames(0) = WS_NEEDLE1
    wsNames(1) = WS_PARAMETER
    wsNames(2) = WS_LOG

    For i = 0 To 2
        Dim ws As Worksheet
        Set ws = WorksheetHolen(ThisWorkbook, wsNames(i))
        If ws Is Nothing Then
            Call Log(LVL_ERR, "Selbsttest", "Worksheet fehlt", wsNames(i))
            fehler = fehler + 1
        Else
            Call Log(LVL_OK, "Selbsttest", "Worksheet vorhanden", wsNames(i))
        End If
    Next i

    If DateiExistiertSicher(PFAD_DAT) Then
        Call Log(LVL_OK, "Selbsttest", "Quelldatei vorhanden", PFAD_DAT)
    Else
        Call Log(LVL_WARN, "Selbsttest", "Quelldatei nicht gefunden", PFAD_DAT)
    End If

    Dim wsN As Worksheet
    Set wsN = WorksheetHolen(ThisWorkbook, WS_NEEDLE1)
    If Not wsN Is Nothing Then
        Dim pflichtSpalten(0 To 4) As String
        pflichtSpalten(0) = C_BEZEICHNUNG
        pflichtSpalten(1) = C_POSITION
        pflichtSpalten(2) = C_TYP
        pflichtSpalten(3) = C_EXT_VERD
        pflichtSpalten(4) = C_GEWICHT

        Dim j As Integer
        For j = 0 To 4
            Dim col As Long
            col = SpalteFinden(wsN, pflichtSpalten(j), "Selbsttest", False)
            If col <= 0 Then
                Call Log(LVL_WARN, "Selbsttest", "Spalte in Needle1 fehlt", pflichtSpalten(j))
            Else
                Call Log(LVL_OK, "Selbsttest", "Needle1-Spalte OK", pflichtSpalten(j) & " -> Spalte " & col)
            End If
        Next j
    End If

    Call LogAbschluss(fehler = 0)

    If fehler = 0 Then
        MsgBox "Selbsttest erfolgreich!" & vbCrLf & _
               "Alle kritischen Konfigurationen sind vorhanden." & vbCrLf & vbCrLf & _
               "Details im Log-Sheet '" & WS_LOG & "'.", _
               vbInformation, "Selbsttest - OK"
    Else
        MsgBox fehler & " kritische(r) Fehler gefunden!" & vbCrLf & _
               "Bitte Log-Sheet '" & WS_LOG & "' pruefen.", _
               vbCritical, "Selbsttest - Fehler"
    End If
End Sub

Private Function MakroVollqualifiziert(ByVal prozedur As String) As String
    MakroVollqualifiziert = "'" & ThisWorkbook.Name & "'!" & prozedur
End Function

Attribute VB_Name = "Modul5_SequenzSchreiben"
Option Explicit

Private Const MODUL As String = "SequenzSchreiben"

Private Type Needle1Spalten
    Position As Long
    Typ As Long
    Bezeichnung As Long
    Kommentar As Long
    Gewicht As Long
    ExtVerd As Long
End Type

Private Type DatenSpalten
    ProbenNr As Long
    Matrix As Long
    Analyse As Long
    Prio As Long
End Type

Public Function SequenzSchreiben() As Boolean
    Dim wb As Workbook
    Dim wsDaten As Worksheet
    Dim wsNeedle1 As Worksheet
    Dim wsParam As Worksheet
    Dim n1 As Needle1Spalten
    Dim d As DatenSpalten
    Dim parameterMap As Object
    Dim kontrollprobenListe As Collection
    Dim probenListe As Collection
    Dim arrZI1() As String
    Dim arrZI2() As String
    Dim arrES() As String
    Dim reKProbe As Object
    Dim rePProbe As Object
    Dim rePosition As Object
    Dim jahrAktuell As Integer
    Dim jahrAlt As Integer
    Dim phi50ProbenNr As String
    Dim intervallLaenge As Integer
    Dim intervallAnzahl As Integer
    Dim maxPos As Integer
    Dim vialType As String
    Dim naechsteZeile As Long

    SequenzSchreiben = False
    Call Log(LVL_SECTION, MODUL, "Sequenz-Erzeugung gestartet")

    On Error GoTo EH

    Set wb = ThisWorkbook
    Set wsDaten = WorksheetPflicht(wb, WS_DATEN)
    Set wsNeedle1 = WorksheetPflicht(wb, WS_NEEDLE1)
    Set wsParam = WorksheetPflicht(wb, WS_PARAMETER)

    n1.Position = SpalteFinden(wsNeedle1, C_POSITION, MODUL, True)
    n1.Typ = SpalteFinden(wsNeedle1, C_TYP, MODUL, True)
    n1.Bezeichnung = SpalteFinden(wsNeedle1, C_BEZEICHNUNG, MODUL, True)
    n1.Kommentar = SpalteFinden(wsNeedle1, C_KOMMENTAR, MODUL, False)
    n1.Gewicht = SpalteFinden(wsNeedle1, C_GEWICHT, MODUL, False)
    n1.ExtVerd = SpalteFinden(wsNeedle1, C_EXT_VERD, MODUL, False)

    d.ProbenNr = DATA_COL_PROBENNR
    d.Matrix = SpalteFinden(wsDaten, C_MATRIX, MODUL, True)
    d.Analyse = SpalteFinden(wsDaten, C_ANALYSE, MODUL, False)
    d.Prio = SpalteFinden(wsDaten, C_PRIO, MODUL, False)

    If n1.Position <= 0 Or n1.Typ <= 0 Or n1.Bezeichnung <= 0 Then
        Call Log(LVL_ERR, MODUL, "Kritische Needle1-Spalten fehlen")
        Exit Function
    End If

    If d.Matrix <= 0 Then
        Call Log(LVL_ERR, MODUL, "Pflicht-Spalte 'Matrix' in Daten fehlt")
        Exit Function
    End If

    Call Needle1Zuruecksetzen(wsNeedle1)
    naechsteZeile = 2
    Set parameterMap = ParameterKonfigMapLaden(wsParam)

    jahrAktuell = CInt(Right$(CStr(Year(Date)), 2))
    jahrAlt = (jahrAktuell + 99) Mod 100

    Set reKProbe = RegExErstellen(ProbenRegExMuster("K", jahrAktuell, jahrAlt))
    Set rePProbe = RegExErstellen(ProbenRegExMuster("P", jahrAktuell, jahrAlt))
    Set rePosition = RegExErstellen(REGEX_POSITION)

    If Not DatenSortieren(wsDaten, d) Then
        Call Log(LVL_ERR, MODUL, "Sortierung fehlgeschlagen")
        Exit Function
    End If

    phi50ProbenNr = PHI50ProbenNrErmitteln(wsDaten, d, reKProbe)
    If Len(phi50ProbenNr) = 0 Then
        Call Log(LVL_WARN, MODUL, "Keine phI-50 Kontrollprobe gefunden")
    Else
        Call Log(LVL_OK, MODUL, "phI-50 Kontrollprobe", phi50ProbenNr)
    End If

    Set kontrollprobenListe = KontrollprobenSammeln(wsDaten, d, reKProbe)
    Set probenListe = ProbenListeAufbauen(wsDaten, d, rePProbe, kontrollprobenListe)

    Call Log(LVL_INFO, MODUL, "Probenlisten aufgebaut", _
        "Kontrollproben: " & kontrollprobenListe.Count & " | Gesamt: " & probenListe.Count)

    intervallLaenge = IntervallLaengeLaden(wsParam)
    intervallAnzahl = IntervallAnzahlErmitteln(probenListe.Count, intervallLaenge)

    arrZI1 = ParameterBlockLesen(wsParam, P_ZW1, P_ZW1_ENDE, phi50ProbenNr)
    arrZI2 = ParameterBlockLesen(wsParam, P_ZW2, P_ZW2_ENDE, phi50ProbenNr)
    arrES = ParameterBlockLesen(wsParam, P_END_SEQ, P_END_ENDE, phi50ProbenNr)

    If Not InitialSequenzSchreiben(wsParam, wsNeedle1, n1, parameterMap, naechsteZeile, phi50ProbenNr) Then
        Call Log(LVL_ERR, MODUL, "Initial-Sequenz konnte nicht geschrieben werden")
        Exit Function
    End If

    maxPos = MaxPositionErmitteln(wsNeedle1, n1.Position, rePosition)
    vialType = ProbenTypLaden(wsParam)

    If Not MesssequenzSchreiben(wsNeedle1, n1, naechsteZeile, probenListe, _
        intervallLaenge, intervallAnzahl, maxPos, vialType, arrZI1, arrZI2, arrES) Then
        Call Log(LVL_ERR, MODUL, "Messsequenz-Erzeugung fehlgeschlagen")
        Exit Function
    End If

    Call KommentarSpalteSchreiben(wsNeedle1, wsDaten, wsParam, n1, d, reKProbe, rePProbe)

    If n1.ExtVerd > 0 And n1.Gewicht > 0 Then
        Call EinheitsspaltenFuellen(wsNeedle1, n1)
        Call Log(LVL_OK, MODUL, "Ext. Verd. und Gewicht auf 1 gesetzt")
    Else
        Call Log(LVL_WARN, MODUL, "Ext. Verd. / Gewicht Spalten nicht gefunden - uebersprungen")
    End If

    Call Log(LVL_OK, MODUL, "Sequenz erfolgreich erstellt", _
        "Gesamtzeilen in Needle1: " & (naechsteZeile - 2))

    SequenzSchreiben = True
    Exit Function

EH:
    Call LogFehler(MODUL, "Kritischer Fehler in SequenzSchreiben", Err.Number, Err.Description)
    MsgBox "Sequenz-Erzeugung fehlgeschlagen!" & vbCrLf & vbCrLf & _
           "Fehler " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Details im Log-Sheet '" & WS_LOG & "'.", _
           vbCritical, "Sequenz - Kritischer Fehler"
End Function

Private Function DatenSortieren(ByVal ws As Worksheet, ByRef d As DatenSpalten) As Boolean
    Const SORT_ORDER_HEADER As String = "Sortierung_Originalzeile"

    Dim laenge As Long
    Dim letzteSpalteDaten As Long
    Dim sortSpalte As Long
    Dim orderSpalte As Long
    Dim i As Long

    DatenSortieren = False
    On Error GoTo EH

    laenge = LetzteZeile(ws, d.ProbenNr)
    If laenge < 2 Then
        DatenSortieren = True
        Exit Function
    End If

    letzteSpalteDaten = LetzteSpalte(ws, 1)
    sortSpalte = letzteSpalteDaten + 1
    orderSpalte = sortSpalte + 1

    ws.Cells(1, sortSpalte).Value = C_SORT_KEY
    ws.Cells(1, orderSpalte).Value = SORT_ORDER_HEADER

    For i = 2 To laenge
        Dim key As Integer
        Dim prioNum As Integer
        Dim hatPrio As Boolean
        Dim matrixWert As String
        Dim prioText As String

        key = KEY_KEIN_PRIO
        prioNum = 0
        hatPrio = False
        matrixWert = LCase$(Trim$(CStr(ws.Cells(i, d.Matrix).Value)))

        If d.Prio > 0 Then
            prioText = Trim$(CStr(ws.Cells(i, d.Prio).Value))
            If Len(prioText) > 0 And IsNumeric(prioText) Then
                prioNum = CInt(Val(prioText))
                hatPrio = True
            End If
        End If

        Select Case True
            Case prioNum = 1
                key = KEY_PRIO1
            Case InStr(1, matrixWert, "w", vbTextCompare) > 0 And hatPrio And prioNum = 2
                key = KEY_WASSER_P2
            Case InStr(1, matrixWert, "w", vbTextCompare) > 0 And hatPrio And prioNum = 3
                key = KEY_WASSER_P3
            Case matrixWert = ASP_MATRIX And hatPrio And prioNum = 2
                key = KEY_ASP_P2
            Case matrixWert = ASP_MATRIX And hatPrio And prioNum = 3
                key = KEY_ASP_P3
            Case matrixWert <> ASP_MATRIX And InStr(1, matrixWert, "w", vbTextCompare) = 0 And hatPrio And prioNum = 2
                key = KEY_BODEN_P2
            Case matrixWert <> ASP_MATRIX And InStr(1, matrixWert, "w", vbTextCompare) = 0 And hatPrio And prioNum = 3
                key = KEY_BODEN_P3
        End Select

        ws.Cells(i, sortSpalte).Value = key
        ws.Cells(i, orderSpalte).Value = i

        If i Mod 50 = 0 Then
            Call Fortschritt("Sortierung", i, laenge - 1)
        End If
    Next i

    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add Key:=ws.Range(ws.Cells(2, sortSpalte), ws.Cells(laenge, sortSpalte)), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    ws.Sort.SortFields.Add Key:=ws.Range(ws.Cells(2, orderSpalte), ws.Cells(laenge, orderSpalte)), _
                           SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal

    With ws.Sort
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(laenge, orderSpalte))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    ws.Columns(orderSpalte).Delete
    ws.Columns(sortSpalte).Delete

    Call Log(LVL_DEBUG, "DatenSortieren", "Sortierung angewendet", (laenge - 1) & " Zeilen")
    DatenSortieren = True
    Exit Function

EH:
    On Error Resume Next
    If orderSpalte > 0 Then ws.Columns(orderSpalte).Delete
    If sortSpalte > 0 Then ws.Columns(sortSpalte).Delete
    On Error GoTo 0
    Call LogFehler("DatenSortieren", "Fehler bei Sortierung", Err.Number, Err.Description)
End Function

Private Function PHI50ProbenNrErmitteln(ByVal ws As Worksheet, ByRef d As DatenSpalten, ByVal reKProbe As Object) As String
    Dim laenge As Long
    Dim i As Long

    PHI50ProbenNrErmitteln = ""
    laenge = LetzteZeile(ws, d.ProbenNr)

    For i = 2 To laenge
        Dim probenNr As String
        Dim matrixWert As String

        probenNr = Trim$(CStr(ws.Cells(i, d.ProbenNr).Value))
        matrixWert = Trim$(CStr(ws.Cells(i, d.Matrix).Value))

        If reKProbe.Test(probenNr) Then
            If StrComp(matrixWert, PHI50_MATRIX, vbTextCompare) = 0 Then
                PHI50ProbenNrErmitteln = probenNr
                Exit Function
            End If
        End If
    Next i
End Function

Private Function KontrollprobenSammeln(ByVal ws As Worksheet, ByRef d As DatenSpalten, ByVal reKProbe As Object) As Collection
    Dim liste As New Collection
    Dim laenge As Long
    Dim i As Long

    laenge = LetzteZeile(ws, d.ProbenNr)

    For i = 2 To laenge
        Dim probenNr As String
        Dim matrixWert As String

        probenNr = Trim$(CStr(ws.Cells(i, d.ProbenNr).Value))
        matrixWert = Trim$(CStr(ws.Cells(i, d.Matrix).Value))

        If reKProbe.Test(probenNr) Then
            If StrComp(matrixWert, PHI50_MATRIX, vbTextCompare) <> 0 Then
                liste.Add probenNr
            End If
        End If
    Next i

    Set KontrollprobenSammeln = liste
End Function

Private Function ProbenListeAufbauen(ByVal ws As Worksheet, ByRef d As DatenSpalten, ByVal rePProbe As Object, ByVal kListe As Collection) As Collection
    Dim liste As New Collection
    Dim eintrag As Variant
    Dim laenge As Long
    Dim i As Long

    For Each eintrag In kListe
        liste.Add CStr(eintrag)
    Next eintrag

    laenge = LetzteZeile(ws, d.ProbenNr)
    For i = 2 To laenge
        Dim probenNr As String
        probenNr = Trim$(CStr(ws.Cells(i, d.ProbenNr).Value))
        If rePProbe.Test(probenNr) Then
            liste.Add probenNr
        End If
    Next i

    Set ProbenListeAufbauen = liste
End Function

Private Function IntervallLaengeLaden(ByVal wsParam As Worksheet) As Integer
    Dim zelle As Range
    Dim defaultWert As Integer
    Dim roherWert As Long
    Dim roherText As String

    defaultWert = (MIN_INTERVALL + MAX_INTERVALL) \ 2
    Set zelle = FindeExakt(wsParam.UsedRange, P_INTERVALL, xlByRows)

    If zelle Is Nothing Then
        Call Log(LVL_WARN, "IntervallLaengeLaden", _
            "Parameter 'Intervall-Laenge' nicht gefunden - Default wird verwendet", _
            CStr(defaultWert))
        IntervallLaengeLaden = defaultWert
        Exit Function
    End If

    roherText = Trim$(CStr(zelle.Offset(1, 0).Value))
    If Len(roherText) > 0 And IsNumeric(roherText) Then
        roherWert = CLng(Val(roherText))
    Else
        roherWert = defaultWert
    End If

    IntervallLaengeLaden = BegrenzeInteger(roherWert, MIN_INTERVALL, MAX_INTERVALL)

    If roherWert <> IntervallLaengeLaden Then
        Call Log(LVL_WARN, "IntervallLaengeLaden", _
            "Intervall-Laenge ausserhalb des zulaessigen Bereichs", _
            "Eingelesen: " & roherWert & " -> " & IntervallLaengeLaden)
    End If
End Function

Private Function IntervallAnzahlErmitteln(ByVal probenAnzahl As Long, ByVal intervallLaenge As Integer) As Integer
    If probenAnzahl <= 0 Then
        IntervallAnzahlErmitteln = 1
    Else
        IntervallAnzahlErmitteln = CInt((probenAnzahl + intervallLaenge - 1) \ intervallLaenge)
    End If
End Function

Private Function MaxPositionErmitteln(ByVal wsNeedle As Worksheet, ByVal posCol As Long, ByVal rePos As Object) As Integer
    Dim lz As Long
    Dim i As Long

    MaxPositionErmitteln = 0
    lz = LetzteZeile(wsNeedle, posCol)

    For i = 2 To lz
        Dim posWert As String
        Dim posZahl As Integer

        posWert = Trim$(CStr(wsNeedle.Cells(i, posCol).Value))
        If rePos.Test(posWert) Then
            posZahl = PositionsnummerAusText(posWert)
            If posZahl > MaxPositionErmitteln Then
                MaxPositionErmitteln = posZahl
            End If
        End If
    Next i

    Call Log(LVL_DEBUG, "MaxPositionErmitteln", "Hoechste Position nach Initial-Sequenz", "A" & MaxPositionErmitteln)
End Function

Private Function ProbenTypLaden(ByVal wsParam As Worksheet) As String
    Dim zelle As Range
    Dim wert As String

    Set zelle = FindeExakt(wsParam.UsedRange, P_PROBENTYP, xlByRows)
    If zelle Is Nothing Then
        ProbenTypLaden = DEFAULT_PROBENTYP
        Call Log(LVL_WARN, "ProbenTypLaden", "ProbenTyp nicht gefunden - Default wird verwendet", DEFAULT_PROBENTYP)
        Exit Function
    End If

    wert = Trim$(CStr(zelle.Offset(0, 1).Value))
    If Len(wert) = 0 Then
        ProbenTypLaden = DEFAULT_PROBENTYP
        Call Log(LVL_WARN, "ProbenTypLaden", "ProbenTyp leer - Default wird verwendet", DEFAULT_PROBENTYP)
    Else
        ProbenTypLaden = wert
    End If
End Function

Private Function InitialSequenzSchreiben(ByVal wsParam As Worksheet, ByVal wsNeedle As Worksheet, ByRef n1 As Needle1Spalten, _
                                         ByVal parameterMap As Object, ByRef naechsteZeile As Long, ByVal phi50Nr As String) As Boolean
    Dim startZelle As Range
    Dim offset As Long
    Dim guard As Long
    Dim counter As Long
    Dim endeGefunden As Boolean

    InitialSequenzSchreiben = False
    Set startZelle = FindeExakt(wsParam.UsedRange, P_INIT_SEQ, xlByRows)

    If startZelle Is Nothing Then
        Call Log(LVL_ERR, "InitialSequenzSchreiben", "'" & P_INIT_SEQ & "' nicht gefunden")
        Exit Function
    End If

    offset = 1
    Do While guard < LOOP_GUARD
        Dim bezWert As String
        Dim resolvedBez As String
        Dim typWert As String
        Dim posWert As String

        bezWert = Trim$(CStr(startZelle.Offset(offset, 0).Value))
        If bezWert = P_INIT_ENDE Then
            endeGefunden = True
            Exit Do
        End If

        If Len(bezWert) > 0 Then
            resolvedBez = BlockBezeichnungAufloesen(bezWert, phi50Nr)
            Call KonfigEintragLaden(parameterMap, bezWert, typWert, posWert)
            Call SchreibeNeedleZeile(wsNeedle, n1, naechsteZeile, resolvedBez, typWert, posWert)
            naechsteZeile = naechsteZeile + 1
            counter = counter + 1
        End If

        offset = offset + 1
        guard = guard + 1
    Loop

    If Not endeGefunden Then
        Call Log(LVL_ERR, "InitialSequenzSchreiben", "Ende-Marker fuer Initial-Sequenz nicht gefunden", P_INIT_ENDE)
        Exit Function
    End If

    Call Log(LVL_OK, "InitialSequenzSchreiben", "Initial-Sequenz geschrieben", counter & " Eintraege")
    InitialSequenzSchreiben = True
End Function

Private Function MesssequenzSchreiben(ByVal wsNeedle As Worksheet, ByRef n1 As Needle1Spalten, ByRef naechsteZeile As Long, _
                                      ByVal probenListe As Collection, ByVal intervallLaenge As Integer, ByVal intervallAnzahl As Integer, _
                                      ByVal maxPosNachInit As Integer, ByVal vialType As String, _
                                      ByRef arrZI1() As String, ByRef arrZI2() As String, ByRef arrES() As String) As Boolean
    Dim gesamtProben As Long
    Dim startPosition As Integer
    Dim aktuellePosition As Integer
    Dim probenCounter As Long
    Dim i As Integer
    Dim k As Integer

    MesssequenzSchreiben = False
    On Error GoTo EH

    gesamtProben = probenListe.Count
    startPosition = FreieStartPosition(maxPosNachInit)

    If startPosition <= 0 Then
        Call Log(LVL_ERR, "MesssequenzSchreiben", _
            "Keine freie Tellerposition fuer Messsequenz verfuegbar", _
            "MaxPos nach Initial-Sequenz: " & maxPosNachInit)
        Exit Function
    End If

    aktuellePosition = startPosition
    probenCounter = 1

    If gesamtProben = 0 Then
        Call Log(LVL_WARN, "MesssequenzSchreiben", "Keine Proben vorhanden - nur End-Sequenz wird geschrieben")
        Call BlockInNeedle1Schreiben(wsNeedle, n1, arrES, naechsteZeile)
        MesssequenzSchreiben = True
        Exit Function
    End If

    For i = 1 To intervallAnzahl
        For k = 1 To intervallLaenge
            If probenCounter > gesamtProben Then Exit For

            Dim probenNr As String
            probenNr = Trim$(CStr(probenListe(probenCounter)))

            Call SchreibeNeedleZeile(wsNeedle, n1, naechsteZeile, probenNr, vialType, "A" & aktuellePosition)

            naechsteZeile = naechsteZeile + 1
            probenCounter = probenCounter + 1
            aktuellePosition = NaechsteMessposition(aktuellePosition, startPosition)

            If (probenCounter - 1) Mod 10 = 0 Then
                Call Fortschritt("Sequenz", probenCounter - 1, gesamtProben, "Intervall " & i & "/" & intervallAnzahl)
            End If
        Next k

        If i = intervallAnzahl Then
            Call BlockInNeedle1Schreiben(wsNeedle, n1, arrES, naechsteZeile)
        ElseIf i Mod 2 = 1 Then
            Call BlockInNeedle1Schreiben(wsNeedle, n1, arrZI1, naechsteZeile)
        Else
            Call BlockInNeedle1Schreiben(wsNeedle, n1, arrZI2, naechsteZeile)
        End If
    Next i

    Call Log(LVL_OK, "MesssequenzSchreiben", "Sequenz-Schleife abgeschlossen", _
        "Verarbeitete Proben: " & (probenCounter - 1) & "/" & gesamtProben)
    MesssequenzSchreiben = True
    Exit Function

EH:
    Call LogFehler("MesssequenzSchreiben", "Fehler in Sequenz-Hauptschleife", Err.Number, Err.Description)
End Function

Private Sub BlockInNeedle1Schreiben(ByVal wsNeedle As Worksheet, ByRef n1 As Needle1Spalten, _
                                    ByRef arr() As String, ByRef naechsteZeile As Long)
    Dim anzahl As Long
    Dim index As Long

    anzahl = BlockLaenge(arr)
    If anzahl <= 0 Then Exit Sub

    For index = 1 To anzahl
        If Len(Trim$(arr(1, index))) > 0 Then
            Call SchreibeNeedleZeile(wsNeedle, n1, naechsteZeile, arr(1, index), arr(2, index), arr(3, index))
            naechsteZeile = naechsteZeile + 1
        End If
    Next index
End Sub

Private Sub KommentarSpalteSchreiben(ByVal wsNeedle As Worksheet, ByVal wsDaten As Worksheet, ByVal wsParam As Worksheet, _
                                     ByRef n1 As Needle1Spalten, ByRef d As DatenSpalten, _
                                     ByVal reKProbe As Object, ByVal rePProbe As Object)
    Dim kommentarMap As Object
    Dim kKommentarMap As Object
    Dim pKommentarQueues As Object
    Dim pKommentarIndex As Object
    Dim lzNeedle As Long
    Dim needleZeile As Long

    If n1.Kommentar <= 0 Then
        Call Log(LVL_WARN, "KommentarSpalteSchreiben", "Kommentar-Spalte nicht gefunden - uebersprungen")
        Exit Sub
    End If

    On Error GoTo EH

    Set kommentarMap = KommentarMapLaden(wsParam)
    Set pKommentarIndex = CreateObject("Scripting.Dictionary")
    pKommentarIndex.CompareMode = vbTextCompare

    Call KommentarDatenIndizesAufbauen(wsDaten, d, kommentarMap, reKProbe, rePProbe, kKommentarMap, pKommentarQueues)

    lzNeedle = LetzteZeile(wsNeedle, n1.Bezeichnung)
    For needleZeile = 2 To lzNeedle
        Dim bezWert As String
        bezWert = Trim$(CStr(wsNeedle.Cells(needleZeile, n1.Bezeichnung).Value))

        If Len(bezWert) = 0 Then GoTo NaechsteNeedleZeile

        If reKProbe.Test(bezWert) Then
            If kKommentarMap.Exists(bezWert) Then
                wsNeedle.Cells(needleZeile, n1.Kommentar).Value = kKommentarMap(bezWert)
            End If
        ElseIf rePProbe.Test(bezWert) Then
            If pKommentarQueues.Exists(bezWert) Then
                Dim index As Long
                Dim kommentarWert As String
                Dim queue As Collection

                index = 1
                If pKommentarIndex.Exists(bezWert) Then index = CLng(pKommentarIndex(bezWert))

                Set queue = pKommentarQueues(bezWert)
                If index <= queue.Count Then
                    kommentarWert = CStr(queue.Item(index))
                    If Len(kommentarWert) > 0 Then
                        wsNeedle.Cells(needleZeile, n1.Kommentar).Value = kommentarWert
                    End If
                    pKommentarIndex(bezWert) = index + 1
                End If
            End If
        ElseIf StrComp(bezWert, DRIFT_BEZ, vbTextCompare) = 0 Then
            If kommentarMap.Exists(DRIFT_BEZ) Then
                wsNeedle.Cells(needleZeile, n1.Kommentar).Value = kommentarMap(DRIFT_BEZ)
            End If
        End If

NaechsteNeedleZeile:
    Next needleZeile

    Call Log(LVL_OK, "KommentarSpalteSchreiben", "Kommentare eingetragen")
    Exit Sub

EH:
    Call LogFehler("KommentarSpalteSchreiben", "Fehler beim Kommentar-Schreiben", Err.Number, Err.Description)
End Sub

Private Sub EinheitsspaltenFuellen(ByVal wsNeedle As Worksheet, ByRef n1 As Needle1Spalten)
    Dim lz As Long
    lz = LetzteZeile(wsNeedle, n1.Bezeichnung)

    If lz < 2 Then Exit Sub

    wsNeedle.Range(wsNeedle.Cells(2, n1.ExtVerd), wsNeedle.Cells(lz, n1.ExtVerd)).Value = 1
    wsNeedle.Range(wsNeedle.Cells(2, n1.Gewicht), wsNeedle.Cells(lz, n1.Gewicht)).Value = 1
End Sub

Private Sub Needle1Zuruecksetzen(ByVal wsNeedle As Worksheet)
    Dim letzteZeileGesamt As Long
    Dim letzteSpalteGesamt As Long

    letzteZeileGesamt = LetzteBelegteZeileGesamt(wsNeedle)
    letzteSpalteGesamt = LetzteBelegteSpalteGesamt(wsNeedle)

    If letzteZeileGesamt >= 2 Then
        wsNeedle.Range(wsNeedle.Cells(2, 1), wsNeedle.Cells(letzteZeileGesamt, letzteSpalteGesamt)).ClearContents
        Call Log(LVL_INFO, MODUL, "Vorherige Needle1-Sequenz geloescht", "Zeilen: " & (letzteZeileGesamt - 1))
    End If
End Sub

Private Sub SchreibeNeedleZeile(ByVal wsNeedle As Worksheet, ByRef n1 As Needle1Spalten, ByVal zielZeile As Long, _
                                ByVal bezeichnung As String, Optional ByVal typWert As String = "", Optional ByVal positionsWert As String = "")
    wsNeedle.Cells(zielZeile, n1.Bezeichnung).Value = bezeichnung

    If n1.Typ > 0 Then
        wsNeedle.Cells(zielZeile, n1.Typ).Value = typWert
    End If

    If n1.Position > 0 Then
        wsNeedle.Cells(zielZeile, n1.Position).Value = positionsWert
    End If
End Sub

Private Function KommentarMapLaden(ByVal wsParam As Worksheet) As Object
    Dim dict As Object
    Dim matrixHeader As Range
    Dim kommentarHeader As Range
    Dim zeile As Long
    Dim guard As Long

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    Set matrixHeader = FindeExakt(wsParam.UsedRange, P_MATRIX_BEZ, xlByRows)
    Set kommentarHeader = FindeExakt(wsParam.UsedRange, P_KOMMENTAR, xlByRows)

    If matrixHeader Is Nothing Or kommentarHeader Is Nothing Then
        Call Log(LVL_WARN, "KommentarMapLaden", "Kommentar-Konfiguration nicht vollstaendig gefunden")
        Set KommentarMapLaden = dict
        Exit Function
    End If

    zeile = kommentarHeader.Row + 1
    Do While guard < LOOP_GUARD
        Dim kommentarWert As String
        Dim schluessel As String

        kommentarWert = Trim$(CStr(wsParam.Cells(zeile, kommentarHeader.Column).Value))
        If kommentarWert = P_KOMMENTAR_ENDE Then Exit Do

        schluessel = Trim$(CStr(wsParam.Cells(zeile, matrixHeader.Column).Value))
        If Len(schluessel) > 0 Then
            dict(schluessel) = kommentarWert
        End If

        zeile = zeile + 1
        guard = guard + 1
    Loop

    Call Log(LVL_DEBUG, "KommentarMapLaden", "Kommentar-Map geladen", dict.Count & " Eintraege")
    Set KommentarMapLaden = dict
End Function

Private Sub KommentarDatenIndizesAufbauen(ByVal wsDaten As Worksheet, ByRef d As DatenSpalten, _
                                          ByVal kommentarMap As Object, ByVal reKProbe As Object, ByVal rePProbe As Object, _
                                          ByRef kKommentarMap As Object, ByRef pKommentarQueues As Object)
    Dim lzDaten As Long
    Dim i As Long

    Set kKommentarMap = CreateObject("Scripting.Dictionary")
    kKommentarMap.CompareMode = vbTextCompare

    Set pKommentarQueues = CreateObject("Scripting.Dictionary")
    pKommentarQueues.CompareMode = vbTextCompare

    lzDaten = LetzteZeile(wsDaten, d.ProbenNr)
    For i = 2 To lzDaten
        Dim probenNr As String
        probenNr = Trim$(CStr(wsDaten.Cells(i, d.ProbenNr).Value))
        If Len(probenNr) = 0 Then GoTo NaechsteDatenzeile

        If reKProbe.Test(probenNr) Then
            Dim matrixWert As String
            matrixWert = Trim$(CStr(wsDaten.Cells(i, d.Matrix).Value))
            If kommentarMap.Exists(matrixWert) Then
                kKommentarMap(probenNr) = kommentarMap(matrixWert)
            End If
        ElseIf rePProbe.Test(probenNr) And d.Analyse > 0 Then
            Dim analyseWert As String
            Dim kommentarWert As String

            analyseWert = Trim$(CStr(wsDaten.Cells(i, d.Analyse).Value))
            kommentarWert = ""
            If kommentarMap.Exists(analyseWert) Then
                kommentarWert = kommentarMap(analyseWert)
            End If

            Call KommentarQueueHinzufuegen(pKommentarQueues, probenNr, kommentarWert)
        End If

NaechsteDatenzeile:
    Next i
End Sub

Private Sub KommentarQueueHinzufuegen(ByVal queueMap As Object, ByVal schluessel As String, ByVal kommentarWert As String)
    Dim queue As Collection

    If queueMap.Exists(schluessel) Then
        Set queue = queueMap(schluessel)
    Else
        Set queue = New Collection
        queueMap.Add schluessel, queue
    End If

    queue.Add kommentarWert
End Sub

Private Sub KonfigEintragLaden(ByVal parameterMap As Object, ByVal bezeichnung As String, ByRef typWert As String, ByRef positionsWert As String)
    typWert = ""
    positionsWert = ""

    If parameterMap Is Nothing Then Exit Sub
    If Not parameterMap.Exists(bezeichnung) Then
        Call Log(LVL_WARN, "KonfigEintragLaden", "Bezeichnung nicht in Konfiguration gefunden", bezeichnung)
        Exit Sub
    End If

    Dim configWerte As Variant
    configWerte = parameterMap(bezeichnung)
    typWert = CStr(configWerte(0))
    positionsWert = CStr(configWerte(1))
End Sub

Private Function BlockBezeichnungAufloesen(ByVal bezWert As String, ByVal phi50Nr As String) As String
    If StrComp(bezWert, PHI50_MATRIX, vbTextCompare) = 0 Then
        If Len(phi50Nr) > 0 Then
            BlockBezeichnungAufloesen = phi50Nr
        Else
            BlockBezeichnungAufloesen = PHI50_MATRIX
        End If
    Else
        BlockBezeichnungAufloesen = bezWert
    End If
End Function

Private Function FreieStartPosition(ByVal maxPosNachInit As Integer) As Integer
    If maxPosNachInit < 0 Then maxPosNachInit = 0

    If maxPosNachInit >= MAX_TELLER_POS Then
        FreieStartPosition = 0
    Else
        FreieStartPosition = maxPosNachInit + 1
    End If
End Function

Private Function NaechsteMessposition(ByVal aktuellePosition As Integer, ByVal startPosition As Integer) As Integer
    If aktuellePosition >= MAX_TELLER_POS Then
        NaechsteMessposition = startPosition
    Else
        NaechsteMessposition = aktuellePosition + 1
    End If
End Function

Private Function PositionsnummerAusText(ByVal positionsText As String) As Integer
    If Len(positionsText) < 2 Then Exit Function
    If Not IsNumeric(Mid$(positionsText, 2)) Then Exit Function
    PositionsnummerAusText = CInt(Val(Mid$(positionsText, 2)))
End Function

Private Function BlockLaenge(ByRef arr() As String) As Long
    On Error GoTo Leer
    BlockLaenge = UBound(arr, 2)
    If BlockLaenge < 1 Then BlockLaenge = 0
    Exit Function

Leer:
    BlockLaenge = 0
End Function

Private Function LetzteBelegteZeileGesamt(ByVal ws As Worksheet) As Long
    Dim letzteZelle As Range

    Set letzteZelle = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)

    If letzteZelle Is Nothing Then
        LetzteBelegteZeileGesamt = 1
    Else
        LetzteBelegteZeileGesamt = letzteZelle.Row
    End If
End Function

Private Function LetzteBelegteSpalteGesamt(ByVal ws As Worksheet) As Long
    Dim letzteZelle As Range

    Set letzteZelle = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, LookAt:=xlPart, _
                                    SearchOrder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False, SearchFormat:=False)

    If letzteZelle Is Nothing Then
        LetzteBelegteSpalteGesamt = 1
    Else
        LetzteBelegteSpalteGesamt = letzteZelle.Column
    End If
End Function

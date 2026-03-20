Attribute VB_Name = "Modul1_Konstanten"
Option Explicit

Public Const PFAD_DAT           As String = "C:\dialims\messsequenzPHI.dat"

Public Const WS_DATEN           As String = "Daten"
Public Const WS_NEEDLE1         As String = "Needle1"
Public Const WS_PARAMETER       As String = "Parameter"
Public Const WS_LOG             As String = "_Execution_Log"
Public Const WS_DATEN_TEMP      As String = "_TMP_Daten_Import"

Public Const C_BEZEICHNUNG      As String = "Bezeichnung"
Public Const C_POSITION         As String = "Pos."
Public Const C_POSITION_PARAM   As String = "Position auf Messteller"
Public Const C_TYP              As String = "Typ"
Public Const C_KOMMENTAR        As String = "Kommentar"
Public Const C_MATRIX           As String = "Matrix"
Public Const C_ANALYSE          As String = "Analyse"
Public Const C_PROBE            As String = "Probe"
Public Const C_PRIO             As String = "Auftrag Priorität"
Public Const C_EXT_VERD         As String = "Ext. Verd."
Public Const C_GEWICHT          As String = "Gewicht"
Public Const C_SORT_KEY         As String = "SortKey"

Public Const P_INIT_SEQ         As String = "Initial-Sequenz Reihenfolge"
Public Const P_INIT_ENDE        As String = "Ende Initial-Sequenz"
Public Const P_ZW1              As String = "Zwischen-Sequenz 1 Reihenfolge"
Public Const P_ZW1_ENDE         As String = "Ende Zwischen-Sequenz 1"
Public Const P_ZW2              As String = "Zwischen-Sequenz 2 Reihenfolge"
Public Const P_ZW2_ENDE         As String = "Ende Zwischen-Sequenz 2"
Public Const P_END_SEQ          As String = "End-Sequenz Reihenfolge"
Public Const P_END_ENDE         As String = "Ende End-Sequenz"
Public Const P_INTERVALL        As String = "Intervall-Länge"
Public Const P_PROBENTYP        As String = "Probentyp"
Public Const P_KOMMENTAR        As String = "Kommentar-Spalte"
Public Const P_KOMMENTAR_ENDE   As String = "Ende Kommentar-Spalte"
Public Const P_MATRIX_BEZ       As String = "Matrix/Bezeichnung"

Public Const PHI50_MATRIX       As String = "phI-50"
Public Const DRIFT_BEZ          As String = "Drift"
Public Const ASP_MATRIX         As String = "asp"
Public Const DEFAULT_PROBENTYP  As String = "U"

Public Const MAX_TELLER_POS     As Integer = 50
Public Const MIN_INTERVALL      As Integer = 10
Public Const MAX_INTERVALL      As Integer = 20
Public Const LOOP_GUARD         As Long = 50000
Public Const REGEX_POSITION     As String = "^A\d+$"
Public Const DATA_COL_PROBENNR  As Integer = 2

Public Const KEY_PRIO1          As Integer = 1
Public Const KEY_WASSER_P2      As Integer = 2
Public Const KEY_WASSER_P3      As Integer = 3
Public Const KEY_ASP_P2         As Integer = 4
Public Const KEY_ASP_P3         As Integer = 5
Public Const KEY_BODEN_P2       As Integer = 6
Public Const KEY_BODEN_P3       As Integer = 7
Public Const KEY_KEIN_PRIO      As Integer = 9

Public Const LVL_INFO           As String = "INFO"
Public Const LVL_OK             As String = "OK"
Public Const LVL_WARN           As String = "WARNUNG"
Public Const LVL_ERR            As String = "FEHLER"
Public Const LVL_DEBUG          As String = "DEBUG"
Public Const LVL_SECTION        As String = "--------"

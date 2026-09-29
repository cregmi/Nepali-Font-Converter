Attribute VB_Name = "PreetiUnicode_Excel"
Option Explicit

'===============================================================================
' PreetiUnicode_Excel.bas - Unicode (Devanagari) <-> Preeti font converter for Excel
'
' Author: Chandan Regmi
' Note: Compiled and debugged with assistance from AI.
' 
' Rebuilt on a validated longest-match tokenizer (Unicode->Preeti) and a
' char-map + reordering-rules pipeline (Preeti->Unicode), cross-checked against
' multiple independent open-source Preeti mapping tables and verified word-for-word
' against real ground-truth Preeti text before being converted to VBA.
'===============================================================================

' ==========================================
' MODULE-LEVEL CONSTANTS AND CACHED LOOKUP TABLES
' ==========================================
Private VIRAMA As String
Private RA_CHAR As String
Private CONSONANT_SET As String
Private MATRA_OR_BINDU_SET As String

Private Const PREETI_FONT As String = "Preeti"
Private Const UNICODE_FONT As String = "Kokila"  ' any installed Unicode Devanagari font, e.g. "Mangal" or "Kalimati"

Private mMapsReady As Boolean
Private mU2P_Len4 As Object  ' Scripting.Dictionary: 4-char Unicode key -> Preeti bytes
Private mU2P_Len3 As Object
Private mU2P_Len2 As Object
Private mU2P_Len1 As Object
Private mP2U_Char As Object  ' Scripting.Dictionary: single Preeti char -> Unicode string


' ==========================================
' EXCEL ENTRY POINTS
' ==========================================
' Run on the current cell selection. Only plain-text cells are converted; numbers,
' empty cells and formulas are left untouched. A cell whose text does not change
' (for example a bare English word or a number stored as text) is skipped and keeps
' its font.
'
' Worksheet functions (usable in formulas, e.g. =UnicodeToPreeti(A1)) are the two
' Public functions UnicodeToPreeti and PreetiToUnicode further down in this module.
Public Sub ConvertSelectionToPreeti()
    ConvertSelectionCells True
End Sub

Public Sub ConvertSelectionToUnicode()
    ConvertSelectionCells False
End Sub

Private Sub ConvertSelectionCells(ByVal toPreeti As Boolean)
    Dim target As Range
    Dim cell As Range
    Dim fontName As String
    Dim oldText As String
    Dim newText As String
    Dim msg As String
    Dim counter As Long

    If TypeName(Application.Selection) <> "Range" Then
        MsgBox "Select one or more cells first.", vbExclamation, "Preeti Converter"
        Exit Sub
    End If

    Set target = TextConstantCells(Application.Selection)
    If target Is Nothing Then
        msg = "The selection contains no plain-text cells to convert."
        msg = msg & vbCrLf & "Formula results are not changed; use =UnicodeToPreeti(A1) or =PreetiToUnicode(A1) for those."
        MsgBox msg, vbInformation, "Preeti Converter"
        Exit Sub
    End If

    If toPreeti Then
        fontName = PREETI_FONT
        Application.StatusBar = "Converting to Preeti..."
    Else
        fontName = UNICODE_FONT
        Application.StatusBar = "Converting to Unicode..."
    End If
    Application.ScreenUpdating = False

    On Error GoTo Fail
    For Each cell In target.Cells
        oldText = CStr(cell.Value)
        If toPreeti Then
            newText = UnicodeToPreeti(oldText)
        Else
            newText = PreetiToUnicode(oldText)
        End If
        If newText <> oldText Then
            WriteText cell, newText
            cell.Font.Name = fontName
        End If
        counter = counter + 1
        If counter Mod 25 = 0 Then DoEvents
    Next cell

Done:
    Application.StatusBar = False
    Application.ScreenUpdating = True
    Exit Sub

Fail:
    msg = "Conversion stopped: " & Err.Description
    msg = msg & vbCrLf & "If the sheet is protected, unprotect it and try again."
    MsgBox msg, vbExclamation, "Preeti Converter"
    Resume Done
End Sub

' Returns just the plain-text (non-formula) cells of the selection, or Nothing.
' SpecialCells is used for multi-cell selections because it skips blanks, numbers and
' formulas and never walks a whole column; a single cell is checked directly because
' SpecialCells would otherwise search the entire sheet.
Private Function TextConstantCells(ByVal sel As Range) As Range
    Dim v As Variant
    On Error Resume Next
    If sel.CountLarge = 1 Then
        v = sel.Cells(1, 1).Value
        If Not sel.Cells(1, 1).HasFormula Then
            If VarType(v) = vbString Then
                If Len(v) > 0 Then Set TextConstantCells = sel.Cells(1, 1)
            End If
        End If
    Else
        Set TextConstantCells = sel.SpecialCells(xlCellTypeConstants, xlTextValues)
    End If
    On Error GoTo 0
End Function

' Writes converted text into a cell as text. Excel would otherwise treat a value that
' begins with = + - or @ as a formula. In Preeti an opening bracket is stored as a hyphen,
' so such values do occur. A leading apostrophe stores them as text; cells already
' formatted as Text do not need it (and would show it literally).
Private Sub WriteText(ByVal cell As Range, ByVal s As String)
    Dim lead As String
    If Len(s) > 0 Then lead = Left(s, 1)
    If cell.NumberFormat <> "@" And (lead = "=" Or lead = "+" Or lead = "-" Or lead = "@") Then
        cell.Value = "'" & s
    Else
        cell.Value = s
    End If
End Sub

' ==========================================
' CHARACTER CLASSIFICATION HELPERS
' ==========================================
Private Function IsConsonant(ByVal ch As String) As Boolean
    IsConsonant = (Len(ch) = 1) And (InStr(1, CONSONANT_SET, ch, vbBinaryCompare) > 0)
End Function

Private Function IsMatraOrBindu(ByVal ch As String) As Boolean
    IsMatraOrBindu = (Len(ch) = 1) And (InStr(1, MATRA_OR_BINDU_SET, ch, vbBinaryCompare) > 0)
End Function

' ==========================================
' UNICODE PRE-PROCESSING PASSES
' ==========================================

' Moves a Reph (R + virama at the head of a cluster) to after the consonant
' cluster + matras it belongs to. Skips a R+virama that is itself preceded by
' a virama, since that means it is a medial subjoined-RA inside an existing
' conjunct chain (e.g. clusters like tra), not a genuine Reph.
Private Function RephShiftRight(ByVal txt As String) As String
    Dim outStr As String, cluster As String, prevChar As String
    Dim i As Long, j As Long, n As Long
    Dim handled As Boolean

    n = Len(txt)
    i = 1
    Do While i <= n
        handled = False
        If i < n Then
            If Mid(txt, i, 2) = RA_CHAR & VIRAMA Then
                If i = 1 Then
                    prevChar = ""
                Else
                    prevChar = Mid(txt, i - 1, 1)
                End If
                If prevChar <> VIRAMA Then
                    j = i + 2
                    cluster = ""
                    Do While j <= n
                        If Not IsConsonant(Mid(txt, j, 1)) Then Exit Do
                        cluster = cluster & Mid(txt, j, 1)
                        j = j + 1
                        If j <= n Then
                            If Mid(txt, j, 1) = VIRAMA Then
                                cluster = cluster & VIRAMA
                                j = j + 1
                            Else
                                Exit Do
                            End If
                        Else
                            Exit Do
                        End If
                    Loop
                    Do While j <= n
                        If Not IsMatraOrBindu(Mid(txt, j, 1)) Then Exit Do
                        cluster = cluster & Mid(txt, j, 1)
                        j = j + 1
                    Loop
                    outStr = outStr & cluster & RA_CHAR & VIRAMA
                    i = j
                    handled = True
                End If
            End If
        End If
        If Not handled Then
            outStr = outStr & Mid(txt, i, 1)
            i = i + 1
        End If
    Loop
    RephShiftRight = outStr
End Function

' Moves a Raswa (short i-matra) to just before the consonant (and its optional
' half-form chain) that it phonetically belongs to.
Private Function RaswaShiftLeft(ByVal txt As String) As String
    Dim outStr As String, ch As String, bc As String
    Dim i As Long, k As Long, n As Long
    Dim RASWA As String
    RASWA = ChrW(&H93F)

    n = Len(txt)
    outStr = ""
    For i = 1 To n
        ch = Mid(txt, i, 1)
        If ch = RASWA Then
            k = Len(outStr)
            Do While k > 0
                bc = Mid(outStr, k, 1)
                If IsConsonant(bc) Then
                    k = k - 1
                    If k > 0 Then
                        If Mid(outStr, k, 1) = VIRAMA Then
                            k = k - 1
                        Else
                            Exit Do
                        End If
                    End If
                Else
                    Exit Do
                End If
            Loop
            outStr = Left(outStr, k) & ch & Mid(outStr, k + 1)
        Else
            outStr = outStr & ch
        End If
    Next i
    RaswaShiftLeft = outStr
End Function

' ==========================================
' LOOKUP TABLE INITIALIZATION (built once, cached for the VBA session)
' ==========================================
Private Sub EnsureMapsReady()
    If mMapsReady Then Exit Sub

    VIRAMA = ChrW(&H94D)
    RA_CHAR = ChrW(&H930)
    CONSONANT_SET = ChrW(&H915) & ChrW(&H916) & ChrW(&H917) & ChrW(&H918) & ChrW(&H919) & ChrW(&H91A) & ChrW(&H91B) & ChrW(&H91C) & ChrW(&H91D) & ChrW(&H91E) & ChrW(&H91F) & ChrW(&H920) & ChrW(&H921) & ChrW(&H922) & ChrW(&H923) & ChrW(&H924) & ChrW(&H925) & ChrW(&H926) & ChrW(&H927) & ChrW(&H928) & ChrW(&H92A) & ChrW(&H92B) & ChrW(&H92C) & ChrW(&H92D) & ChrW(&H92E) & ChrW(&H92F) & ChrW(&H930) & ChrW(&H932) & ChrW(&H935) & ChrW(&H936) & ChrW(&H937) & ChrW(&H938) & ChrW(&H939) & ChrW(&H921) & ChrW(&H93C) & ChrW(&H922) & ChrW(&H93C)
    MATRA_OR_BINDU_SET = ChrW(&H93E) & ChrW(&H93F) & ChrW(&H940) & ChrW(&H941) & ChrW(&H942) & ChrW(&H943) & ChrW(&H947) & ChrW(&H948) & ChrW(&H94B) & ChrW(&H94C) & ChrW(&H949) & ChrW(&H902) & ChrW(&H903) & ChrW(&H901)

    Set mU2P_Len4 = CreateObject("Scripting.Dictionary")
    Set mU2P_Len3 = CreateObject("Scripting.Dictionary")
    Set mU2P_Len2 = CreateObject("Scripting.Dictionary")
    Set mU2P_Len1 = CreateObject("Scripting.Dictionary")
    Set mP2U_Char = CreateObject("Scripting.Dictionary")

    ' --- Unicode -> Preeti: 4-character keys (4 entries) ---
    mU2P_Len4.Add ChrW(&H915) & ChrW(&H94D) & ChrW(&H937) & ChrW(&H94D), "I"
    mU2P_Len4.Add ChrW(&H91C) & ChrW(&H94D) & ChrW(&H91E) & ChrW(&H94D), ChrW(&HA1)
    ' NOTE: a doubled conjunct's half-form (TA+virama+TA+virama, and the same
    ' for DA) intentionally has no entry here -- e.g. the middle of "mahattva".
    ' Leaving it out lets the tokenizer fall through to the 3-char base
    ' ligature followed by the ordinary trailing-virama character, which is
    ' how Preeti actually represents it; there is no separate dedicated glyph.

    ' --- Unicode -> Preeti: 3-character keys (25 entries) ---
    mU2P_Len3.Add ChrW(&H915) & ChrW(&H94D) & ChrW(&H924), "Qm"
    mU2P_Len3.Add ChrW(&H915) & ChrW(&H94D) & ChrW(&H937), "If"
    mU2P_Len3.Add ChrW(&H919) & ChrW(&H94D) & ChrW(&H915), ChrW(&HCD)
    mU2P_Len3.Add ChrW(&H919) & ChrW(&H94D) & ChrW(&H916), ChrW(&HCE)
    mU2P_Len3.Add ChrW(&H919) & ChrW(&H94D) & ChrW(&H917), ChrW(&HCB)
    mU2P_Len3.Add ChrW(&H919) & ChrW(&H94D) & ChrW(&H918), ChrW(&H2039)
    mU2P_Len3.Add ChrW(&H919) & ChrW(&H94D) & ChrW(&H922), ChrW(&HB0)
    mU2P_Len3.Add ChrW(&H91C) & ChrW(&H94D) & ChrW(&H91E), "1"
    mU2P_Len3.Add ChrW(&H91F) & ChrW(&H94D) & ChrW(&H91F), ChrW(&HA7)
    mU2P_Len3.Add ChrW(&H91F) & ChrW(&H94D) & ChrW(&H920), ChrW(&HDD)
    mU2P_Len3.Add ChrW(&H920) & ChrW(&H94D) & ChrW(&H920), ChrW(&HB6)
    mU2P_Len3.Add ChrW(&H921) & ChrW(&H94D) & ChrW(&H921), ChrW(&H2022)
    mU2P_Len3.Add ChrW(&H924) & ChrW(&H94D) & ChrW(&H924), "Q"
    mU2P_Len3.Add ChrW(&H924) & ChrW(&H94D) & ChrW(&H930), "q"
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H918), ChrW(&HA2)
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H926), "2"
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H927), "4"
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H92E), ChrW(&HDF)
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H92F), "B"
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H930), ChrW(&H203A)
    mU2P_Len3.Add ChrW(&H926) & ChrW(&H94D) & ChrW(&H935), ChrW(&HE5)
    mU2P_Len3.Add ChrW(&H927) & ChrW(&H94D) & ChrW(&H930), ChrW(&H201E)
    mU2P_Len3.Add ChrW(&H928) & ChrW(&H94D) & ChrW(&H928), ChrW(&HCC)
    mU2P_Len3.Add ChrW(&H930) & ChrW(&H94D) & ChrW(&H200D), ChrW(&HA5)
    mU2P_Len3.Add ChrW(&H936) & ChrW(&H94D) & ChrW(&H930), ">"

    ' --- Unicode -> Preeti: 2-character keys (32 entries) ---
    mU2P_Len2.Add ChrW(&H915) & ChrW(&H94D), "S"
    mU2P_Len2.Add ChrW(&H916) & ChrW(&H94D), "V"
    mU2P_Len2.Add ChrW(&H917) & ChrW(&H94D), "U"
    mU2P_Len2.Add ChrW(&H918) & ChrW(&H94D), ChrW(&HA3)
    mU2P_Len2.Add ChrW(&H91A) & ChrW(&H94D), "R"
    mU2P_Len2.Add ChrW(&H91C) & ChrW(&H94D), "H"
    mU2P_Len2.Add ChrW(&H91D) & ChrW(&H94D), ChrW(&H2030)
    mU2P_Len2.Add ChrW(&H91E) & ChrW(&H94D), "~"
    mU2P_Len2.Add ChrW(&H921) & ChrW(&H93C), "?"
    mU2P_Len2.Add ChrW(&H922) & ChrW(&H93C), "?"
    mU2P_Len2.Add ChrW(&H923) & ChrW(&H94D), "0"
    mU2P_Len2.Add ChrW(&H924) & ChrW(&H94D), "T"
    mU2P_Len2.Add ChrW(&H925) & ChrW(&H94D), "Y"
    mU2P_Len2.Add ChrW(&H927) & ChrW(&H94D), "W"
    mU2P_Len2.Add ChrW(&H928) & ChrW(&H94D), "G"
    mU2P_Len2.Add ChrW(&H92A) & ChrW(&H94D), "K"
    mU2P_Len2.Add ChrW(&H92B) & ChrW(&H94D), ChrW(&H2C6)
    mU2P_Len2.Add ChrW(&H92C) & ChrW(&H94D), "A"
    mU2P_Len2.Add ChrW(&H92D) & ChrW(&H94D), "E"
    mU2P_Len2.Add ChrW(&H92E) & ChrW(&H94D), "D"
    mU2P_Len2.Add ChrW(&H930) & ChrW(&H941), "?"
    mU2P_Len2.Add ChrW(&H930) & ChrW(&H942), ChrW(&HBF)
    mU2P_Len2.Add ChrW(&H930) & ChrW(&H94D), "{"
    mU2P_Len2.Add ChrW(&H932) & ChrW(&H94D), "N"
    mU2P_Len2.Add ChrW(&H935) & ChrW(&H94D), "J"
    mU2P_Len2.Add ChrW(&H936) & ChrW(&H94D), "Z"
    mU2P_Len2.Add ChrW(&H937) & ChrW(&H94D), "i"
    mU2P_Len2.Add ChrW(&H938) & ChrW(&H94D), ":"
    mU2P_Len2.Add ChrW(&H939) & ChrW(&H943), ChrW(&HC5)
    mU2P_Len2.Add ChrW(&H939) & ChrW(&H94D), "X"
    mU2P_Len2.Add ChrW(&H94D) & ChrW(&H92F), ChrW(&HD8)
    mU2P_Len2.Add ChrW(&H94D) & ChrW(&H930), "|"

    ' --- Unicode -> Preeti: 1-character keys (93 entries) ---
    mU2P_Len1.Add "!", ChrW(&HDB)
    mU2P_Len1.Add "%", ChrW(&HDC)
    mU2P_Len1.Add "(", "-"
    mU2P_Len1.Add ")", "_"
    mU2P_Len1.Add "+", ChrW(&HB1)
    mU2P_Len1.Add ",", ","
    mU2P_Len1.Add ".", "="
    mU2P_Len1.Add "/", ChrW(&HF7)
    mU2P_Len1.Add ";", ChrW(&HD9)
    mU2P_Len1.Add "=", ChrW(&HD6)
    mU2P_Len1.Add "?", "<"
    mU2P_Len1.Add ChrW(&HA8), ChrW(&HD2)
    mU2P_Len1.Add ChrW(&HD7), ChrW(&HD7)
    mU2P_Len1.Add ChrW(&H901), "F"
    mU2P_Len1.Add ChrW(&H902), "+"
    mU2P_Len1.Add ChrW(&H903), "M"
    mU2P_Len1.Add ChrW(&H905), "c"
    mU2P_Len1.Add ChrW(&H906), "cf"
    mU2P_Len1.Add ChrW(&H907), "O"
    mU2P_Len1.Add ChrW(&H908), "O{"
    mU2P_Len1.Add ChrW(&H909), "p"
    mU2P_Len1.Add ChrW(&H90A), "pm"
    mU2P_Len1.Add ChrW(&H90B), "C"
    mU2P_Len1.Add ChrW(&H90F), "P"
    mU2P_Len1.Add ChrW(&H910), "P]"
    mU2P_Len1.Add ChrW(&H913), "cf]"
    mU2P_Len1.Add ChrW(&H914), "cf}"
    mU2P_Len1.Add ChrW(&H915), "s"
    mU2P_Len1.Add ChrW(&H916), "v"
    mU2P_Len1.Add ChrW(&H917), "u"
    mU2P_Len1.Add ChrW(&H918), "3"
    mU2P_Len1.Add ChrW(&H919), ChrW(&HAA)
    mU2P_Len1.Add ChrW(&H91A), "r"
    mU2P_Len1.Add ChrW(&H91B), "5"
    mU2P_Len1.Add ChrW(&H91C), "h"
    mU2P_Len1.Add ChrW(&H91D), "em"
    mU2P_Len1.Add ChrW(&H91E), "`"
    mU2P_Len1.Add ChrW(&H91F), "6"
    mU2P_Len1.Add ChrW(&H920), "7"
    mU2P_Len1.Add ChrW(&H921), "8"
    mU2P_Len1.Add ChrW(&H922), "9"
    mU2P_Len1.Add ChrW(&H923), "0f"
    mU2P_Len1.Add ChrW(&H924), "t"
    mU2P_Len1.Add ChrW(&H925), "y"
    mU2P_Len1.Add ChrW(&H926), "b"
    mU2P_Len1.Add ChrW(&H927), "w"
    mU2P_Len1.Add ChrW(&H928), "g"
    mU2P_Len1.Add ChrW(&H92A), "k"
    mU2P_Len1.Add ChrW(&H92B), "km"
    mU2P_Len1.Add ChrW(&H92C), "a"
    mU2P_Len1.Add ChrW(&H92D), "e"
    mU2P_Len1.Add ChrW(&H92E), "d"
    mU2P_Len1.Add ChrW(&H92F), "o"
    mU2P_Len1.Add ChrW(&H930), "/"
    mU2P_Len1.Add ChrW(&H932), "n"
    mU2P_Len1.Add ChrW(&H935), "j"
    mU2P_Len1.Add ChrW(&H936), "z"
    mU2P_Len1.Add ChrW(&H937), "if"
    mU2P_Len1.Add ChrW(&H938), ";"
    mU2P_Len1.Add ChrW(&H939), "x"
    mU2P_Len1.Add ChrW(&H93D), ChrW(&H2DC)
    mU2P_Len1.Add ChrW(&H93E), "f"
    mU2P_Len1.Add ChrW(&H93F), "l"
    mU2P_Len1.Add ChrW(&H940), "L"
    mU2P_Len1.Add ChrW(&H941), "'"
    mU2P_Len1.Add ChrW(&H942), ChrW(&H22)
    mU2P_Len1.Add ChrW(&H943), "["
    mU2P_Len1.Add ChrW(&H945), ChrW(&H2018)
    mU2P_Len1.Add ChrW(&H947), "]"
    mU2P_Len1.Add ChrW(&H948), "}"
    mU2P_Len1.Add ChrW(&H949), "f" & ChrW(&H2018)
    mU2P_Len1.Add ChrW(&H94B), "f]"
    mU2P_Len1.Add ChrW(&H94C), "f}"
    mU2P_Len1.Add ChrW(&H94D), "\"
    mU2P_Len1.Add ChrW(&H950), ChrW(&HE7)
    mU2P_Len1.Add ChrW(&H960), "C["
    mU2P_Len1.Add ChrW(&H961), "N" & ChrW(&H2DC)
    mU2P_Len1.Add ChrW(&H964), "."
    mU2P_Len1.Add ChrW(&H965), ".."
    mU2P_Len1.Add ChrW(&H966), ")"
    mU2P_Len1.Add ChrW(&H967), "!"
    mU2P_Len1.Add ChrW(&H968), "@"
    mU2P_Len1.Add ChrW(&H969), "#"
    mU2P_Len1.Add ChrW(&H96A), "$"
    mU2P_Len1.Add ChrW(&H96B), "%"
    mU2P_Len1.Add ChrW(&H96C), "^"
    mU2P_Len1.Add ChrW(&H96D), "&"
    mU2P_Len1.Add ChrW(&H96E), "*"
    mU2P_Len1.Add ChrW(&H96F), "("
    mU2P_Len1.Add ChrW(&H2018), ChrW(&H2026)
    mU2P_Len1.Add ChrW(&H2019), ChrW(&HDA)
    mU2P_Len1.Add ChrW(&H201C), ChrW(&HE6)
    mU2P_Len1.Add ChrW(&H201D), ChrW(&HC6)

    ' --- Preeti -> Unicode: single-character map (135 entries) ---
    mP2U_Char.Add "!", ChrW(&H967)
    mP2U_Char.Add ChrW(&H22), ChrW(&H942)
    mP2U_Char.Add "#", ChrW(&H969)
    mP2U_Char.Add "$", ChrW(&H96A)
    mP2U_Char.Add "%", ChrW(&H96B)
    mP2U_Char.Add "&", ChrW(&H96D)
    mP2U_Char.Add "'", ChrW(&H941)
    mP2U_Char.Add "(", ChrW(&H96F)
    mP2U_Char.Add ")", ChrW(&H966)
    mP2U_Char.Add "*", ChrW(&H96E)
    mP2U_Char.Add "+", ChrW(&H902)
    mP2U_Char.Add ",", ","
    mP2U_Char.Add "-", "("
    mP2U_Char.Add ".", ChrW(&H964)
    mP2U_Char.Add "/", ChrW(&H930)
    mP2U_Char.Add "0", ChrW(&H923) & ChrW(&H94D)
    mP2U_Char.Add "1", ChrW(&H91C) & ChrW(&H94D) & ChrW(&H91E)
    mP2U_Char.Add "2", ChrW(&H926) & ChrW(&H94D) & ChrW(&H926)
    mP2U_Char.Add "3", ChrW(&H918)
    mP2U_Char.Add "4", ChrW(&H926) & ChrW(&H94D) & ChrW(&H927)
    mP2U_Char.Add "5", ChrW(&H91B)
    mP2U_Char.Add "6", ChrW(&H91F)
    mP2U_Char.Add "7", ChrW(&H920)
    mP2U_Char.Add "8", ChrW(&H921)
    mP2U_Char.Add "9", ChrW(&H922)
    mP2U_Char.Add ":", ChrW(&H938) & ChrW(&H94D)
    mP2U_Char.Add ";", ChrW(&H938)
    mP2U_Char.Add "<", "?"
    mP2U_Char.Add "=", "."
    mP2U_Char.Add ">", ChrW(&H936) & ChrW(&H94D) & ChrW(&H930)
    mP2U_Char.Add "?", ChrW(&H930) & ChrW(&H941)
    mP2U_Char.Add "@", ChrW(&H968)
    mP2U_Char.Add "A", ChrW(&H92C) & ChrW(&H94D)
    mP2U_Char.Add "B", ChrW(&H926) & ChrW(&H94D) & ChrW(&H92F)
    mP2U_Char.Add "C", ChrW(&H90B)
    mP2U_Char.Add "D", ChrW(&H92E) & ChrW(&H94D)
    mP2U_Char.Add "E", ChrW(&H92D) & ChrW(&H94D)
    mP2U_Char.Add "F", ChrW(&H901)
    mP2U_Char.Add "G", ChrW(&H928) & ChrW(&H94D)
    mP2U_Char.Add "H", ChrW(&H91C) & ChrW(&H94D)
    mP2U_Char.Add "I", ChrW(&H915) & ChrW(&H94D) & ChrW(&H937) & ChrW(&H94D)
    mP2U_Char.Add "J", ChrW(&H935) & ChrW(&H94D)
    mP2U_Char.Add "K", ChrW(&H92A) & ChrW(&H94D)
    mP2U_Char.Add "L", ChrW(&H940)
    mP2U_Char.Add "M", ChrW(&H903)
    mP2U_Char.Add "N", ChrW(&H932) & ChrW(&H94D)
    mP2U_Char.Add "O", ChrW(&H907)
    mP2U_Char.Add "P", ChrW(&H90F)
    mP2U_Char.Add "Q", ChrW(&H924) & ChrW(&H94D) & ChrW(&H924)
    mP2U_Char.Add "R", ChrW(&H91A) & ChrW(&H94D)
    mP2U_Char.Add "S", ChrW(&H915) & ChrW(&H94D)
    mP2U_Char.Add "T", ChrW(&H924) & ChrW(&H94D)
    mP2U_Char.Add "U", ChrW(&H917) & ChrW(&H94D)
    mP2U_Char.Add "V", ChrW(&H916) & ChrW(&H94D)
    mP2U_Char.Add "W", ChrW(&H927) & ChrW(&H94D)
    mP2U_Char.Add "X", ChrW(&H939) & ChrW(&H94D)
    mP2U_Char.Add "Y", ChrW(&H925) & ChrW(&H94D)
    mP2U_Char.Add "Z", ChrW(&H936) & ChrW(&H94D)
    mP2U_Char.Add "[", ChrW(&H943)
    mP2U_Char.Add "\", ChrW(&H94D)
    mP2U_Char.Add "]", ChrW(&H947)
    mP2U_Char.Add "^", ChrW(&H96C)
    mP2U_Char.Add "_", ")"
    mP2U_Char.Add "`", ChrW(&H91E)
    mP2U_Char.Add "a", ChrW(&H92C)
    mP2U_Char.Add "b", ChrW(&H926)
    mP2U_Char.Add "c", ChrW(&H905)
    mP2U_Char.Add "d", ChrW(&H92E)
    mP2U_Char.Add "e", ChrW(&H92D)
    mP2U_Char.Add "f", ChrW(&H93E)
    mP2U_Char.Add "g", ChrW(&H928)
    mP2U_Char.Add "h", ChrW(&H91C)
    mP2U_Char.Add "i", ChrW(&H937) & ChrW(&H94D)
    mP2U_Char.Add "j", ChrW(&H935)
    mP2U_Char.Add "k", ChrW(&H92A)
    mP2U_Char.Add "l", ChrW(&H93F)
    mP2U_Char.Add "n", ChrW(&H932)
    mP2U_Char.Add "o", ChrW(&H92F)
    mP2U_Char.Add "p", ChrW(&H909)
    mP2U_Char.Add "q", ChrW(&H924) & ChrW(&H94D) & ChrW(&H930)
    mP2U_Char.Add "r", ChrW(&H91A)
    mP2U_Char.Add "s", ChrW(&H915)
    mP2U_Char.Add "t", ChrW(&H924)
    mP2U_Char.Add "u", ChrW(&H917)
    mP2U_Char.Add "v", ChrW(&H916)
    mP2U_Char.Add "w", ChrW(&H927)
    mP2U_Char.Add "x", ChrW(&H939)
    mP2U_Char.Add "y", ChrW(&H925)
    mP2U_Char.Add "z", ChrW(&H936)
    mP2U_Char.Add "|", ChrW(&H94D) & ChrW(&H930)
    mP2U_Char.Add "}", ChrW(&H948)
    mP2U_Char.Add "~", ChrW(&H91E) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&HA1), ChrW(&H91C) & ChrW(&H94D) & ChrW(&H91E) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&HA2), ChrW(&H926) & ChrW(&H94D) & ChrW(&H918)
    mP2U_Char.Add ChrW(&HA3), ChrW(&H918) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&HA4), ChrW(&H91D) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&HA5), ChrW(&H930) & ChrW(&H94D) & ChrW(&H200D)
    mP2U_Char.Add ChrW(&HA7), ChrW(&H91F) & ChrW(&H94D) & ChrW(&H91F)
    mP2U_Char.Add ChrW(&HA9), ChrW(&H930)
    mP2U_Char.Add ChrW(&HAA), ChrW(&H919)
    mP2U_Char.Add ChrW(&HAB), ChrW(&H94D) & ChrW(&H930)
    mP2U_Char.Add ChrW(&HB0), ChrW(&H919) & ChrW(&H94D) & ChrW(&H922)
    mP2U_Char.Add ChrW(&HB1), "+"
    mP2U_Char.Add ChrW(&HB4), ChrW(&H91D)
    mP2U_Char.Add ChrW(&HB6), ChrW(&H920) & ChrW(&H94D) & ChrW(&H920)
    mP2U_Char.Add ChrW(&HBF), ChrW(&H930) & ChrW(&H942)
    mP2U_Char.Add ChrW(&HC5), ChrW(&H939) & ChrW(&H943)
    mP2U_Char.Add ChrW(&HC6), ChrW(&H201D)
    mP2U_Char.Add ChrW(&HCB), ChrW(&H919) & ChrW(&H94D) & ChrW(&H917)
    mP2U_Char.Add ChrW(&HCC), ChrW(&H928) & ChrW(&H94D) & ChrW(&H928)
    mP2U_Char.Add ChrW(&HCD), ChrW(&H919) & ChrW(&H94D) & ChrW(&H915)
    mP2U_Char.Add ChrW(&HCE), ChrW(&H919) & ChrW(&H94D) & ChrW(&H916)
    mP2U_Char.Add ChrW(&HD2), ChrW(&HA8)
    mP2U_Char.Add ChrW(&HD6), "="
    mP2U_Char.Add ChrW(&HD7), ChrW(&HD7)
    mP2U_Char.Add ChrW(&HD8), ChrW(&H94D) & ChrW(&H92F)
    mP2U_Char.Add ChrW(&HD9), ";"
    mP2U_Char.Add ChrW(&HDA), ChrW(&H2019)
    mP2U_Char.Add ChrW(&HDB), "!"
    mP2U_Char.Add ChrW(&HDC), "%"
    mP2U_Char.Add ChrW(&HDD), ChrW(&H91F) & ChrW(&H94D) & ChrW(&H920)
    mP2U_Char.Add ChrW(&HDF), ChrW(&H926) & ChrW(&H94D) & ChrW(&H92E)
    mP2U_Char.Add ChrW(&HE5), ChrW(&H926) & ChrW(&H94D) & ChrW(&H935)
    mP2U_Char.Add ChrW(&HE6), ChrW(&H201C)
    mP2U_Char.Add ChrW(&HE7), ChrW(&H950)
    mP2U_Char.Add ChrW(&HF7), "/"
    mP2U_Char.Add ChrW(&H2C6), ChrW(&H92B) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&H2DC), ChrW(&H93D)
    mP2U_Char.Add ChrW(&H2018), ChrW(&H945)
    mP2U_Char.Add ChrW(&H201E), ChrW(&H927) & ChrW(&H94D) & ChrW(&H930)
    mP2U_Char.Add ChrW(&H2022), ChrW(&H921) & ChrW(&H94D) & ChrW(&H921)
    mP2U_Char.Add ChrW(&H2026), ChrW(&H2018)
    mP2U_Char.Add ChrW(&H2030), ChrW(&H91D) & ChrW(&H94D)
    mP2U_Char.Add ChrW(&H2039), ChrW(&H919) & ChrW(&H94D) & ChrW(&H918)
    mP2U_Char.Add ChrW(&H203A), ChrW(&H926) & ChrW(&H94D) & ChrW(&H930)

    mMapsReady = True
End Sub

' ==========================================
' LOOKUP HELPERS ACROSS THE FOUR LENGTH-KEYED DICTIONARIES
' ==========================================
Private Function DictHasKey(ByVal s As String) As Boolean
    Select Case Len(s)
        Case 4: DictHasKey = mU2P_Len4.Exists(s)
        Case 3: DictHasKey = mU2P_Len3.Exists(s)
        Case 2: DictHasKey = mU2P_Len2.Exists(s)
        Case 1: DictHasKey = mU2P_Len1.Exists(s)
        Case Else: DictHasKey = False
    End Select
End Function

Private Function LookupU2P(ByVal s As String) As String
    Select Case Len(s)
        Case 4: LookupU2P = mU2P_Len4(s)
        Case 3: LookupU2P = mU2P_Len3(s)
        Case 2: LookupU2P = mU2P_Len2(s)
        Case 1: LookupU2P = mU2P_Len1(s)
        Case Else: LookupU2P = s
    End Select
End Function

' Longest-match tokenizer: at each position, tries a 4-char key, then 3, then
' 2, then 1, and emits the mapped Preeti bytes for whichever matches. A single
' unmapped character (space, digit-that-has-no-entry, stray symbol, etc.) is
' passed through unchanged.
Private Function TokenizeUnicodeToPreeti(ByVal txt As String) As String
    Dim outStr As String, matched As String, candidate As String, baseForm As String
    Dim i As Long, n As Long, matchLen As Long

    Call EnsureMapsReady
    n = Len(txt)
    i = 1
    outStr = ""
    Do While i <= n
        matchLen = 0

        If i + 3 <= n Then
            candidate = Mid(txt, i, 4)
            If mU2P_Len4.Exists(candidate) Then matched = candidate: matchLen = 4
        End If
        If matchLen = 0 And i + 2 <= n Then
            candidate = Mid(txt, i, 3)
            If mU2P_Len3.Exists(candidate) Then matched = candidate: matchLen = 3
        End If
        If matchLen = 0 And i + 1 <= n Then
            candidate = Mid(txt, i, 2)
            If mU2P_Len2.Exists(candidate) Then matched = candidate: matchLen = 2
        End If
        If matchLen = 0 Then
            candidate = Mid(txt, i, 1)
            If mU2P_Len1.Exists(candidate) Then matched = candidate: matchLen = 1
        End If

        If matchLen = 0 Then
            outStr = outStr & Mid(txt, i, 1)
            i = i + 1
        Else
            ' prefer the base consonant before a bare RA, to avoid a spurious
            ' extra virama glyph on generic (non-ligature) consonant+ra conjuncts
            If Right(matched, 1) = VIRAMA Then
                If Mid(txt, i + matchLen, 1) = RA_CHAR Then
                    baseForm = Left(matched, Len(matched) - 1)
                    If DictHasKey(baseForm) Then
                        matched = baseForm
                        matchLen = Len(matched)
                    End If
                End If
            End If
            outStr = outStr & LookupU2P(matched)
            i = i + matchLen
        End If
    Loop
    TokenizeUnicodeToPreeti = outStr
End Function

' ==========================================
' CORE CONVERSION ALGORITHMS (public entry points)
' ==========================================
Public Function UnicodeToPreeti(ByVal txt As String) As String
    If Len(txt) = 0 Then Exit Function
    Call EnsureMapsReady
    txt = RephShiftRight(txt)
    txt = RaswaShiftLeft(txt)
    UnicodeToPreeti = TokenizeUnicodeToPreeti(txt)
End Function

Public Function PreetiToUnicode(ByVal txt As String) As String
    Dim outStr As String, ch As String
    Dim i As Long, n As Long

    If Len(txt) = 0 Then Exit Function
    Call EnsureMapsReady
    n = Len(txt)
    outStr = ""
    For i = 1 To n
        ch = Mid(txt, i, 1)
        If mP2U_Char.Exists(ch) Then
            outStr = outStr & mP2U_Char(ch)
        Else
            outStr = outStr & ch
        End If
    Next i

    PreetiToUnicode = ApplyPostRules(outStr)
End Function

' ==========================================
' PREETI->UNICODE REORDERING/CLEANUP RULES
' Applied in order after the character map, to fix up matra reordering,
' Reph placement, and a handful of compound-vowel and duplicate-mark cases.
' ==========================================
Private Function ApplyPostRules(ByVal txt As String) As String
    Static re As Object
    If re Is Nothing Then
        Set re = CreateObject("VBScript.RegExp")
        re.Global = True
        re.IgnoreCase = False
        re.MultiLine = False
    End If

    ' Rule 1
    re.Pattern = ChrW(&H90B) & ChrW(&H943)
    txt = re.Replace(txt, ChrW(&H960))
    ' Rule 2
    re.Pattern = ChrW(&H932) & ChrW(&H94D) & ChrW(&H93D)
    txt = re.Replace(txt, ChrW(&H961))
    ' Rule 3
    re.Pattern = ChrW(&H94D) & ChrW(&H93E)
    txt = re.Replace(txt, "")
    ' Rule 4
    re.Pattern = "(" & ChrW(&H924) & ChrW(&H94D) & ChrW(&H930) & "|" & ChrW(&H924) & ChrW(&H94D) & ChrW(&H924) & ")([^" & ChrW(&H909) & ChrW(&H92D) & ChrW(&H92A) & "]+?)m"
    txt = re.Replace(txt, "$1m$2")
    ' Rule 5
    re.Pattern = ChrW(&H924) & ChrW(&H94D) & ChrW(&H930) & "m"
    txt = re.Replace(txt, ChrW(&H915) & ChrW(&H94D) & ChrW(&H930))
    ' Rule 6
    re.Pattern = ChrW(&H924) & ChrW(&H94D) & ChrW(&H924) & "m"
    txt = re.Replace(txt, ChrW(&H915) & ChrW(&H94D) & ChrW(&H924))
    ' Rule 7
    re.Pattern = "([^" & ChrW(&H909) & ChrW(&H92D) & ChrW(&H92A) & "]+?)m"
    txt = re.Replace(txt, "m$1")
    ' Rule 8
    re.Pattern = ChrW(&H909) & "m"
    txt = re.Replace(txt, ChrW(&H90A))
    ' Rule 9
    re.Pattern = ChrW(&H92D) & "m"
    txt = re.Replace(txt, ChrW(&H91D))
    ' Rule 10
    re.Pattern = ChrW(&H92A) & "m"
    txt = re.Replace(txt, ChrW(&H92B))
    ' Rule 11
    re.Pattern = ChrW(&H907) & "{"
    txt = re.Replace(txt, ChrW(&H908))
    ' Rule 12
    re.Pattern = ChrW(&H93F) & "((." & ChrW(&H94D) & ")*[^" & ChrW(&H94D) & "])"
    txt = re.Replace(txt, "$1" & ChrW(&H93F))
    ' Rule 13
    re.Pattern = "(.[" & ChrW(&H93E) & ChrW(&H93F) & ChrW(&H940) & ChrW(&H941) & ChrW(&H942) & ChrW(&H943) & ChrW(&H947) & ChrW(&H948) & ChrW(&H94B) & ChrW(&H94C) & ChrW(&H902) & ChrW(&H903) & ChrW(&H901) & "]*?){"
    txt = re.Replace(txt, "{$1")
    ' Rule 14
    re.Pattern = "((." & ChrW(&H94D) & ")*){"
    txt = re.Replace(txt, "{$1")
    ' Rule 15
    re.Pattern = "{"
    txt = re.Replace(txt, ChrW(&H930) & ChrW(&H94D))
    ' Rule 16
    re.Pattern = "([" & ChrW(&H93E) & ChrW(&H940) & ChrW(&H941) & ChrW(&H942) & ChrW(&H943) & ChrW(&H947) & ChrW(&H948) & ChrW(&H94B) & ChrW(&H94C) & ChrW(&H902) & ChrW(&H903) & ChrW(&H901) & "]+?)(" & ChrW(&H94D) & "(." & ChrW(&H94D) & ")*[^" & ChrW(&H94D) & "])"
    txt = re.Replace(txt, "$2$1")
    ' Rule 17
    re.Pattern = ChrW(&H94D) & "([" & ChrW(&H93E) & ChrW(&H940) & ChrW(&H941) & ChrW(&H942) & ChrW(&H943) & ChrW(&H947) & ChrW(&H948) & ChrW(&H94B) & ChrW(&H94C) & ChrW(&H902) & ChrW(&H903) & ChrW(&H901) & "]+?)((." & ChrW(&H94D) & ")*[^" & ChrW(&H94D) & "])"
    txt = re.Replace(txt, ChrW(&H94D) & "$2$1")
    ' Rule 18
    re.Pattern = "([" & ChrW(&H902) & ChrW(&H901) & "])([" & ChrW(&H93E) & ChrW(&H93F) & ChrW(&H940) & ChrW(&H941) & ChrW(&H942) & ChrW(&H943) & ChrW(&H947) & ChrW(&H948) & ChrW(&H94B) & ChrW(&H94C) & ChrW(&H903) & "]*)"
    txt = re.Replace(txt, "$2$1")
    ' Rule 19
    re.Pattern = ChrW(&H901) & ChrW(&H901)
    txt = re.Replace(txt, ChrW(&H901))
    ' Rule 20
    re.Pattern = ChrW(&H902) & ChrW(&H902)
    txt = re.Replace(txt, ChrW(&H902))
    ' Rule 21
    re.Pattern = ChrW(&H947) & ChrW(&H947)
    txt = re.Replace(txt, ChrW(&H947))
    ' Rule 22
    re.Pattern = ChrW(&H948) & ChrW(&H948)
    txt = re.Replace(txt, ChrW(&H948))
    ' Rule 23
    re.Pattern = ChrW(&H941) & ChrW(&H941)
    txt = re.Replace(txt, ChrW(&H941))
    ' Rule 24
    re.Pattern = ChrW(&H942) & ChrW(&H942)
    txt = re.Replace(txt, ChrW(&H942))
    ' Rule 25
    re.Pattern = "^" & ChrW(&H903)
    txt = re.Replace(txt, ":")
    ' Rule 26
    re.Pattern = ChrW(&H91F) & ChrW(&H943)
    txt = re.Replace(txt, ChrW(&H91F) & ChrW(&H94D) & ChrW(&H91F))
    ' Rule 27
    re.Pattern = ChrW(&H947) & ChrW(&H93E)
    txt = re.Replace(txt, ChrW(&H93E) & ChrW(&H947))
    ' Rule 28
    re.Pattern = ChrW(&H948) & ChrW(&H93E)
    txt = re.Replace(txt, ChrW(&H93E) & ChrW(&H948))
    ' Rule 29
    re.Pattern = ChrW(&H905) & ChrW(&H93E) & ChrW(&H947)
    txt = re.Replace(txt, ChrW(&H913))
    ' Rule 30
    re.Pattern = ChrW(&H905) & ChrW(&H93E) & ChrW(&H948)
    txt = re.Replace(txt, ChrW(&H914))
    ' Rule 31
    re.Pattern = ChrW(&H905) & ChrW(&H93E)
    txt = re.Replace(txt, ChrW(&H906))
    ' Rule 32
    re.Pattern = ChrW(&H90F) & ChrW(&H947)
    txt = re.Replace(txt, ChrW(&H910))
    ' Rule 33
    re.Pattern = ChrW(&H93E) & ChrW(&H947)
    txt = re.Replace(txt, ChrW(&H94B))
    ' Rule 34
    re.Pattern = ChrW(&H93E) & ChrW(&H948)
    txt = re.Replace(txt, ChrW(&H94C))

    ApplyPostRules = txt
End Function

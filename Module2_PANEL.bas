Attribute VB_Name = "Module2_PANEL"
Option Explicit

' ==============================================================
' PANEL - pouzivatelske rozhranie ovladacieho panela a obsluha tlacidiel
' Rozlozenie z kroku 3: B5 = zdroj (PLAN/CSS), F5 = datum exportu
' ==============================================================

' ==============================================================
' Spolocne pomocne funkcie
' ==============================================================

' Aktualny vyber zdroja, normalizovany. Cokolvek okrem "CSS" -> "PLAN"
Public Function SourceMode() As String
    Dim s As String
    s = UCase$(Trim$(CStr(Worksheets("PANEL").Range(PANEL_SOURCE_CELL).Value)))
    If s <> "CSS" Then s = "PLAN"
    SourceMode = s
End Function

' Obnovi kazdu tabulku na harku STOCK napojenu na dopyt (query). Obycajne
' tabulky sa preskocia, nie je to kriticka chyba (review #2)
Public Sub RefreshStock()
    Dim lo As ListObject
    For Each lo In Worksheets("STOCK").ListObjects
        On Error Resume Next
        lo.QueryTable.Refresh BackgroundQuery:=False
        On Error GoTo 0
    Next lo
End Sub

' Obnovi dotaz Packaging (velkosti balenia z harku Packaging v subore PLAN).
' RefreshStock chyby ignoruje; tu sa po obnoveni overi cas nacitania
' (stlpec LoadedAt), aby sa neaktualne udaje nepouzili bez upozornenia.
' Vrati sa True, ak moze pokracovat: udaje su cerstve, alebo pouzivatel
' potvrdil pouzitie poslednych nacitanych udajov.
Public Function RefreshPackaging() As Boolean

    Dim lo As ListObject
    Dim stampCol As Long
    Dim loadedAt As Variant
    Dim ageMin As Double
    Dim problem As String
    
    RefreshPackaging = False
    
    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(PACKAGING_SHEET).ListObjects(1)
    On Error GoTo 0
    
    If lo Is Nothing Then
        MsgBox "Hárok " & PACKAGING_SHEET & " neobsahuje tabuľku s veľkosťami balení." & vbCrLf & _
                "Vytvor dotaz Packaging (Údaje > Získať údaje).", vbCritical
        Exit Function
    End If
    
    ' obnovenie - chyba sa zachyti, nie ignoruje
    SyncPackagingPath                         ' cesta dotazu podla PATH_PLAN
    On Error Resume Next
    lo.QueryTable.Refresh BackgroundQuery:=False
    If Err.Number <> 0 Then problem = Err.Number & ": " & Err.Description
    On Error GoTo 0
    
    ' aj bez chyby over, ze udaje su naozaj cerstve
    If problem = "" Then
        On Error Resume Next
        stampCol = lo.ListColumns(PKG_STAMP_HEADER).Index
        On Error GoTo 0

        If stampCol = 0 Then
            problem = "chýba stĺpec " & PKG_STAMP_HEADER
        ElseIf lo.DataBodyRange Is Nothing Then
            problem = "tabuľka je prázdna"
        Else
            loadedAt = lo.DataBodyRange.Cells(1, stampCol).Value
            If IsError(loadedAt) Then
                problem = "neplatný čas načítania"
            ElseIf Not (IsDate(loadedAt) Or IsNumeric(loadedAt)) Then
                problem = "neplatný čas načítania"
            Else
                ageMin = (CDbl(Now) - CDbl(loadedAt)) * 1440
                If ageMin > PKG_MAX_AGE_MIN Then
                    problem = "údaje sú staré " & Format(ageMin, "0") & " min"
                End If
            End If
        End If
    End If

    If problem = "" Then
        RefreshPackaging = True
        Exit Function
    End If

    RefreshPackaging = (MsgBox( _
        "Obnovenie veľkostí balení (hárok " & PACKAGING_SHEET & ") sa nepodarilo:" & vbCrLf & _
        problem & vbCrLf & vbCrLf & _
        "Skontroluj, či je dostupný súbor PLAN." & vbCrLf & _
        "Použiť posledné načítané veľkosti balení?", _
        vbYesNo + vbExclamation, "Veľkosti balení") = vbYes)

End Function

Private Function Stamp() As String
    Stamp = Format(Now(), "DD.MM.YYYY HH:NN")
End Function

' ==============================================================
' TLACIDLA
' ==============================================================

Sub Btn_GenerateList()

    Dim src As String
    Dim dateText As String

    src = SourceMode()
    dateText = Trim$(CStr(Worksheets("PANEL").Range(PANEL_DATE_CELL).Value))

    If dateText = "" Then
        MsgBox "Vyber deň exportu (PANEL F5).", vbExclamation
        Exit Sub
    End If

    If src = "CSS" Then
        If Not CSSHasData() Then
            MsgBox "CSS hárok je prázdny." & vbCrLf & _
                   "Najprv stlač 'Aktualizácia zdroja'.", vbExclamation
            Exit Sub
        End If
        If Not CSSDateLoaded(dateText) Then
            MsgBox "Dátum " & dateText & " nie je v načítanom CSS súbore." & vbCrLf & _
                   "Spusti 'Aktualizácia zdroja' a vyber deň znova.", vbExclamation
            Exit Sub
        End If
    Else
        If SlovakDayName(ExportDate()) = "" Then
            MsgBox "Dátum " & dateText & " padne na víkend." & vbCrLf & _
                   "Pre PLAN vyber pracovný deň.", vbExclamation
            Exit Sub
        End If
    End If

    Call Vytvor_zoznam

End Sub

Sub Btn_RefreshStock()
    RefreshStock
    Worksheets("PANEL").Range("B15").Value = "Sklad: " & Stamp()
    MsgBox "Sklad aktualizovaný.", vbInformation
End Sub

' Jedno tlacidlo na obnovenie - rozhoduje podla vybraneho zdroja
Sub Btn_RefreshSource()

    Dim src As String
    Dim ok As Boolean

    src = SourceMode()

    If src = "CSS" Then
        RefreshCSS                       ' pri zlyhani zobrazi vlastnu hlasku
        ok = CSSHasData()
        If ok Then Worksheets("PANEL").Range("B18").Value = "CSS: " & Stamp()
    Else
        ok = RefreshPlan()               ' Boolean vid. krok 4
        If ok Then Call BuildExportDateList
    End If

    If ok Then
        Worksheets("PANEL").Range("B16").Value = "Zdroj " & src & ": " & Stamp()
    End If

End Sub

' ==============================================================
' DEFAULTS - volane z Workbook_Open a na konci BuildPanel
' ==============================================================
Sub SetDefaultExport()

    Dim wsPanel As Worksheet
    Set wsPanel = Worksheets("PANEL")

    If UCase$(Trim$(CStr(wsPanel.Range(PANEL_SOURCE_CELL).Value))) <> "CSS" Then
        wsPanel.Range(PANEL_SOURCE_CELL).Value = "PLAN"
    End If

    ' znovu zostavi zoznam v F5 (datumy CSS, ak su nacitane, inak najblizsie
    ' pracovne dni) a ponecha v F5 platny vyber
    Call BuildExportDateList

End Sub

' ==============================================================
' BUILD PANEL
' ==============================================================
Sub BuildPanel()

    Dim ws As Worksheet
    Set ws = Worksheets("PANEL")

    Application.ScreenUpdating = False

    ' ---- clear ----
    ws.Cells.Clear
    ws.Cells.Interior.ColorIndex = xlNone

    Dim shp As Shape
    For Each shp In ws.Shapes
        shp.Delete
    Next shp

    ws.Tab.Color = RGB(241, 245, 249)

    ' ---- sirky stlpcov ----
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B").ColumnWidth = 18
    ws.Columns("C").ColumnWidth = 18
    ws.Columns("D").ColumnWidth = 18
    ws.Columns("E").ColumnWidth = 2
    ws.Columns("F").ColumnWidth = 18
    ws.Columns("G").ColumnWidth = 18
    ws.Columns("H").ColumnWidth = 2

    ' ---- vysky riadkov (index 1..18) ----
    Dim rh As Variant
    rh = Array(0, 8, 54, 12, 22, 34, 10, 40, 6, 40, 10, 8, 8, 10, 22, 22, 22, 22, 22)
    Dim i As Long
    For i = 1 To 18
        ws.Rows(i).RowHeight = rh(i)
    Next i

    ' ---- nadpis ----
    With ws.Range("B2:G2")
        .Merge
        .Value = "Picking list - kontrolný panel"
        .Font.Size = 22
        .Font.Color = RGB(30, 41, 59)
        .VerticalAlignment = xlVAlignCenter
    End With

    ' ---- popisky ----
    LabelCell ws.Range("B4:D4"), "Zdroj dát"
    LabelCell ws.Range("F4:G4"), "Deň exportu"

    ' ---- rozbalovaci zoznam zdroja  B5:D5 ----
    SelectorCell ws.Range("B5:D5")
    With ws.Range("B5").Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="PLAN,CSS"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    ' ---- rozbalovaci zoznam datumu  F5:G5  (validaciu pridava BuildExportDateList) ----
    SelectorCell ws.Range("F5:G5")

    ' ---- hlavne tlacidla ----
    PanelButton ws, "BtnRefreshStock", "Aktualizácia skladu", "Btn_RefreshStock", _
        ws.Range("B7:D7"), ws.Rows(7).Height - 4
    PanelButton ws, "BtnRefreshSource", "Aktualizácia zdroja", "Btn_RefreshSource", _
        ws.Range("B9:D9"), ws.Rows(9).Height - 4

    ' Vytvorit zoznam - vyrazne tlacidlo, cez F7:G9
    Dim btnGL As Shape
    Set btnGL = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        ws.Range("F7").Left, ws.Range("F7").Top, ws.Range("F7:G7").Width, _
        ws.Rows(7).Height + ws.Rows(8).Height + ws.Rows(9).Height - 4)
    With btnGL
        .Name = "BtnGenerateList"
        .Fill.ForeColor.RGB = RGB(37, 99, 235)
        .Line.Visible = msoFalse
        .TextFrame2.TextRange.Text = "Vytvor zoznam"
        .TextFrame2.TextRange.Font.Size = 18
        .TextFrame2.TextRange.Font.Bold = msoTrue
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .onAction = "Btn_GenerateList"
        .Adjustments.Item(1) = 0.05
    End With

    ' ---- blok stavu ----
    LabelCell ws.Range("B14:D14"), "Stav"
    StatusCell ws.Range("B15:G15"), "Sklad: —"
    StatusCell ws.Range("B16:G16"), "Zdroj: —"
    StatusCell ws.Range("B17:G17"), "Zoznam: —"
    StatusCell ws.Range("B18:G18"), "CSS: —"

    ' ---- defaults ----
    Call SetDefaultExport

    Application.ScreenUpdating = True
    MsgBox "Panel vytvorený.", vbInformation

End Sub

' ==============================================================
' BuildPanel pomocne funkcie
' ==============================================================

Private Sub LabelCell(rng As Range, txt As String)
    With rng
        .Merge
        .Value = txt
        .Font.Size = 13
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub StatusCell(rng As Range, txt As String)
    With rng
        .Merge
        .Value = txt
        .Font.Size = 13
        .Font.Color = RGB(100, 116, 139)
        .VerticalAlignment = xlVAlignCenter
    End With
End Sub

Private Sub SelectorCell(rng As Range)
    With rng
        .Merge
        .Font.Size = 18
        .Font.Bold = False
        .Font.Color = RGB(29, 78, 216)
        .Interior.Color = RGB(219, 234, 254)
        .HorizontalAlignment = xlHAlignLeft
        .VerticalAlignment = xlVAlignCenter
        .IndentLevel = 1
        .BorderAround LineStyle:=xlContinuous, Weight:=xlHairline, Color:=RGB(191, 219, 254)
    End With
End Sub

Private Sub PanelButton(ws As Worksheet, nm As String, caption As String, _
                        onAction As String, spanRange As Range, h As Double)
    Dim b As Shape
    Set b = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
        spanRange.Left, spanRange.Top, spanRange.Width, h)
    With b
        .Name = nm
        .Fill.ForeColor.RGB = RGB(248, 250, 252)
        .Line.ForeColor.RGB = RGB(203, 213, 225)
        .Line.Weight = 0.5
        .TextFrame2.TextRange.Text = caption
        .TextFrame2.TextRange.Font.Size = 14
        .TextFrame2.TextRange.Font.Fill.ForeColor.RGB = RGB(51, 65, 85)
        .TextFrame2.VerticalAnchor = msoAnchorMiddle
        .TextFrame2.TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .onAction = onAction
        .Adjustments.Item(1) = 0.05
    End With
End Sub


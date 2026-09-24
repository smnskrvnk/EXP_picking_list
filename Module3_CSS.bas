Attribute VB_Name = "Module3_CSS"
Option Explicit

' ==============================================================
' SPRACOVANIE ZDROJOV - CSS a PLAN
' Nacita "Ship Sched" zo suboru CSS do lokalneho harku CSS,
' zostavi spolocny rozbalovaci zoznam datumu exportu (PANEL F5) a poskytne
' jednotne nacitavace poloziek pouzivane vo Vytvor_zoznam.
' ==============================================================

Public Const CSS_SHEET          As String = "CSS"          ' harok v tomto zosite
Public Const CSS_SRC_SHEET      As String = "Ship Sched"   ' harok vo vnutri suboru CSS
Public Const CSS_HDR_DATE_ROW   As Long = 20               ' riadok so skutocnymi datumami
Public Const CSS_HDR_NAME_ROW   As Long = 21               ' riadok s hlavickami tabulky
Public Const CSS_FIRST_DATA_ROW As Long = 22

' bunky harku PANEL (Krok3: presunute z doacasnych B20/F20)
Public Const PANEL_SOURCE_CELL    As String = "B5"           ' PLAN / CSS
Public Const PANEL_DATE_CELL      As String = "F5"           ' zvoleny datum exportu

' pomocne bunky harku PACK_SIZE
Public Const PLAN_PATH_ROW      As Long = 1                ' X1 = cesta k suboru PLAN
Public Const CSS_PATH_ROW       As Long = 2                ' X2 = cesta k suboru CSS
Public Const CSS_PATH_COL       As Long = 24               ' stlpec X
Public Const CSS_LIST_COL       As Long = 26               ' stlpec Z = zoznam datumov pre validaciu

' Kolko vypocitanych pracovnych dni ponuknut, ak CSS este nie je nacitane
Public Const EXPORT_FALLBACK_DAYS As Long = 15

' ==============================================================
' MAIN - obnovi lokalny harok CSS zo suboru CSS
' ==============================================================
Sub RefreshCSS()

    Dim cssPath As String
    Dim wbCSS As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long
    Dim itemCol As Long
    Dim prevEvents As Boolean
    Dim errNum As Long
    Dim errDesc As String
    Dim usedLast As Long

    cssPath = Trim(CStr(ThisWorkbook.Worksheets("PACK_SIZE").Cells(CSS_PATH_ROW, CSS_PATH_COL).Value))

    If cssPath = "" Then
        MsgBox "Cesta k CSS súboru je prázdna." & vbCrLf & _
               "Zadaj cestu do hárku PACK_SIZE, bunka X2.", vbExclamation
        Exit Sub
    End If

    Set wsDest = GetOrCreateSheet(CSS_SHEET)

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False          ' aby sa nespustil vlastny Workbook_Open suboru CSS
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail
    Set wbCSS = Workbooks.Open(Filename:=cssPath, ReadOnly:=True, UpdateLinks:=0)
    Application.DisplayAlerts = True

    On Error Resume Next
    Set wsSrc = wbCSS.Worksheets(CSS_SRC_SHEET)
    On Error GoTo CleanFail

    If wsSrc Is Nothing Then
        wbCSS.Close SaveChanges:=False
        MsgBox "Hárok '" & CSS_SRC_SHEET & "' sa v súbore CSS nenašiel." & vbCrLf & vbCrLf & _
   "Skontrolujte, prosím:" & vbCrLf & _
   "- či ste vybrali správny súbor CSS," & vbCrLf & _
   "- či hárok nebol premenovaný alebo odstránený." & vbCrLf & vbCrLf & _
   "Hárok musí mať presne názov '" & CSS_SRC_SHEET & "'." & vbCrLf & vbCrLf & _
   "Súbor: " & wbCSS.Name, _
   vbCritical, "Chýba hárok v súbore CSS"
        GoTo Restore
    End If

    ' posledny riadok s datami urcuje stlpec Acme Item
    itemCol = FindHeaderCol(wsSrc, CSS_HDR_NAME_ROW, "Acme Item")
    If itemCol = 0 Then
        wbCSS.Close SaveChanges:=False
        MsgBox "V hárku '" & CSS_SRC_SHEET & "' sa v riadku " & CSS_HDR_NAME_ROW & _
       " nenašiel stĺpec s hlavičkou 'Acme Item'." & vbCrLf & vbCrLf & _
       "Skontrolujte, prosím:" & vbCrLf & _
       "- či nad hlavičkou neboli vložené alebo odstránené riadky," & vbCrLf & _
       "- či stĺpec nebol premenovaný alebo odstránený," & vbCrLf & _
       "- či ide o správny export CSS (formát sa mohol zmeniť)." & vbCrLf & vbCrLf & _
       "Hlavička musí byť v riadku " & CSS_HDR_NAME_ROW & " a mať presne názov 'Acme Item'.", _
       vbCritical, "Chýba hlavička v súbore CSS"
        GoTo Restore
    End If

    lastRow = wsSrc.Cells(wsSrc.Rows.count, itemCol).End(xlUp).Row

    usedLast = wsSrc.UsedRange.Row + wsSrc.UsedRange.Rows.count - 1
    If usedLast > lastRow Then lastRow = usedLast

    If lastRow < CSS_FIRST_DATA_ROW Then
        wbCSS.Close SaveChanges:=False
        MsgBox "V hárku '" & CSS_SRC_SHEET & "' sa nenašli žiadne riadky s dátami." & vbCrLf & vbCrLf & _
       "Skontrolujte, prosím:" & vbCrLf & _
       "- či ste vybrali správny a aktuálny súbor CSS," & vbCrLf & _
       "- či sú dáta vložené pod riadkom s hlavičkou," & vbCrLf & _
       "- či export nebol vygenerovaný prázdny (napr. nesprávny filter alebo obdobie)," & vbCrLf & _
       "- či na hárku nie je zapnutý filter, ktorý skrýva riadky.", _
       vbExclamation, "Súbor CSS neobsahuje dáta"
        GoTo Restore
    End If

    ' iba hodnoty, priame priradenie - odolne voci filtrom, skrytym riadkom a zlucenym bunkam
    wsDest.Cells.Clear
    wsDest.Range(wsDest.Cells(CSS_HDR_DATE_ROW, 1), wsDest.Cells(lastRow, 100)).Value = _
        wsSrc.Range(wsSrc.Cells(CSS_HDR_DATE_ROW, 1), wsSrc.Cells(lastRow, 100)).Value

    wbCSS.Close SaveChanges:=False

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True

    Call BuildExportDateList

    Exit Sub

Restore:
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

CleanFail:
    errNum = Err.Number
    errDesc = Err.Description

    On Error Resume Next
    If Not wbCSS Is Nothing Then wbCSS.Close SaveChanges:=False
    Application.EnableEvents = prevEvents
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    On Error GoTo 0

    MsgBox "Obnovenie dát z CSS zlyhalo." & vbCrLf & vbCrLf & _
       "Chyba " & errNum & ": " & errDesc & vbCrLf & vbCrLf & _
       "Skontrolujte, prosím:" & vbCrLf & _
       "- či je súbor CSS dostupný (sieťový disk, správna cesta)," & vbCrLf & _
       "- či súbor nemá otvorený iný používateľ alebo nie je uzamknutý," & vbCrLf & _
       "- či súbor nie je poškodený alebo v nesprávnom formáte." & vbCrLf & vbCrLf & _
       "Ak problém pretrváva, pošlite snímku tohto hlásenia správcovi makra.", _
       vbCritical, "Chyba pri obnovení CSS"
End Sub

' ==============================================================
' Zostavi spolocny rozbalovaci zoznam datumu exportu (PANEL F5)
' - ak lokalny harok CSS obsahuje stlpce s datumami FIX, pouziju sa tie
' - inak sa pouzije najblizsich EXPORT_FALLBACK_DAYS pracovnych dni
' Zapise zoznam do stlpca Z harku PACK_SIZE, znovu vytvori pomenovany rozsah
' CSS_DATES a validaciu v F5, a zachova aktualny vyber, ak je platny.
' (Nahradza povodnu funkciu BuildCSSDateList)
' ==============================================================
Public Sub BuildExportDateList()

    Dim wsCSS As Worksheet
    Dim wsPack As Worksheet
    Dim wsPanel As Worksheet
    Dim cols() As Long
    Dim dts() As Date
    Dim listDates() As Date
    Dim n As Long
    Dim i As Long
    Dim keep As String
    Dim colLetter As String

    Set wsCSS = ThisWorkbook.Worksheets(CSS_SHEET)
    Set wsPack = ThisWorkbook.Worksheets("PACK_SIZE")
    Set wsPanel = ThisWorkbook.Worksheets("PANEL")

    n = CSSFixDateCols(wsCSS, cols, dts)

    If n > 0 Then
        ReDim listDates(1 To n)
        For i = 1 To n
            listDates(i) = dts(i)
        Next i
    Else
        n = NextWorkingDays(EXPORT_FALLBACK_DAYS, listDates)
    End If

    keep = Trim(CStr(wsPanel.Range(PANEL_DATE_CELL).Value))

    wsPack.Range(wsPack.Cells(1, CSS_LIST_COL), wsPack.Cells(200, CSS_LIST_COL)).ClearContents
    For i = 1 To n
        wsPack.Cells(i, CSS_LIST_COL).Value = Format(listDates(i), "DD.MM.YYYY")
    Next i

    colLetter = Split(wsPack.Cells(1, CSS_LIST_COL).Address, "$")(1)

    On Error Resume Next
    ThisWorkbook.Names("CSS_DATES").Delete
    On Error GoTo 0
    ThisWorkbook.Names.Add Name:="CSS_DATES", _
        RefersTo:="='PACK_SIZE'!$" & colLetter & "$1:$" & colLetter & "$" & n

    With wsPanel.Range(PANEL_DATE_CELL).Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Formula1:="=CSS_DATES"
        .IgnoreBlank = False
        .InCellDropdown = True
    End With

    ' zachova predchadzajuci vyber, ak stale existuje, inak vezme prvy datum
    If keep <> "" And _
       Not IsError(Application.Match(keep, wsPack.Range(wsPack.Cells(1, CSS_LIST_COL), _
                                                        wsPack.Cells(n, CSS_LIST_COL)), 0)) Then
        wsPanel.Range(PANEL_DATE_CELL).Value = keep
    Else
        wsPanel.Range(PANEL_DATE_CELL).Value = wsPack.Cells(1, CSS_LIST_COL).Value
    End If

End Sub

' ==============================================================
' Next N working days (Mon-Fri) starting from tomorrow.
' Vyplni out() 1..N, vrati N
' ==============================================================
Public Function NextWorkingDays(ByVal count As Long, ByRef out() As Date) As Long

    Dim d As Date
    Dim k As Long

    ReDim out(1 To count)
    d = Date

    Do While k < count
        d = d + 1
        Select Case Weekday(d, vbMonday)
            Case 1 To 5
                k = k + 1
                out(k) = d
        End Select
    Loop

    NextWorkingDays = count

End Function

' ==============================================================
' Slovensky nazov dna pre dany datum. Vikend -> "" (prazdne)
' ==============================================================
Public Function SlovakDayName(ByVal d As Date) As String

    Select Case Weekday(d, vbMonday)
        Case 1: SlovakDayName = "Pondelok"
        Case 2: SlovakDayName = "Utorok"
        Case 3: SlovakDayName = "Streda"
        Case 4: SlovakDayName = "Štvrtok"
        Case 5: SlovakDayName = "Piatok"
        Case Else: SlovakDayName = ""
    End Select

End Function

' ==============================================================
' PANEL F5 ako skutocny Date. Vrati 0, ak je F5 prazdne/nie je datum.
' Explicitne parsuje "DD.MM.YYYY", takze to nezavisi od
' regionalneho nastavenia pocitaca
' ==============================================================
Public Function ExportDate() As Date

    Dim v As Variant
    Dim s As String
    Dim p() As String

    v = ThisWorkbook.Worksheets("PANEL").Range(PANEL_DATE_CELL).Value

    If IsNumeric(v) Then
        If CDbl(v) >= 30000 And CDbl(v) <= 80000 Then ExportDate = CDate(CDbl(v))
        Exit Function
    End If

    s = Trim(CStr(v))
    If s = "" Then Exit Function

    p = Split(s, ".")
    If UBound(p) = 2 Then
        If IsNumeric(p(0)) And IsNumeric(p(1)) And IsNumeric(p(2)) Then
            On Error Resume Next
            ExportDate = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
            On Error GoTo 0
            Exit Function
        End If
    End If

    If IsDate(v) Then ExportDate = CDate(v)

End Function

' ==============================================================
' TRUE, ak lokaly harok CSS aktualne obsahuje stlpce s datumami FIX
' ==============================================================
Public Function CSSHasData() As Boolean

    Dim ws As Worksheet
    Dim cols() As Long
    Dim dts() As Date

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(CSS_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function

    CSSHasData = (CSSFixDateCols(ws, cols, dts) > 0)

End Function

' ==============================================================
' TRUE, ak dateText ("DD.MM.YYYY") je jeden z nacitanych stlpcov CSS
' ==============================================================
Public Function CSSDateLoaded(ByVal dateText As String) As Boolean

    Dim ws As Worksheet
    Dim cols() As Long
    Dim dts() As Date
    Dim n As Long
    Dim i As Long

    Set ws = ThisWorkbook.Worksheets(CSS_SHEET)
    n = CSSFixDateCols(ws, cols, dts)

    For i = 1 To n
        If Format(dts(i), "DD.MM.YYYY") = dateText Then
            CSSDateLoaded = True
            Exit Function
        End If
    Next i

End Function

' ==============================================================
' Najde stlpce dni FIX
' Kotva = "Ship_Sch_Backlog" v riadku hlavicky, potom sa postupuje doprava,
' kym riadok 20 stale obsahuje skutocny datum.
' Vrati pocet; naplni cols() a dts()
' ==============================================================
Public Function CSSFixDateCols(ws As Worksheet, ByRef cols() As Long, ByRef dts() As Date) As Long

    Dim startCol As Long
    Dim c As Long
    Dim n As Long
    Dim v As Variant

    CSSFixDateCols = 0

    startCol = FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Ship_Sch_Backlog")
    If startCol = 0 Then Exit Function

    ReDim cols(1 To 100)
    ReDim dts(1 To 100)
    n = 0

    For c = startCol + 1 To startCol + 100
        v = ws.Cells(CSS_HDR_DATE_ROW, c).Value

        If IsError(v) Then Exit For

        If IsDate(v) Then
            n = n + 1
            cols(n) = c
            dts(n) = CDate(v)
        ElseIf IsNumeric(v) Then
            ' vlozenie iba hodnot meni datumy na seriove cisla
            If CDbl(v) >= 30000 And CDbl(v) <= 80000 Then
                n = n + 1
                cols(n) = c
                dts(n) = CDate(CDbl(v))
            Else
                Exit For
            End If
        Else
            Exit For
        End If
    Next c

    If n = 0 Then Exit Function

    ReDim Preserve cols(1 To n)
    ReDim Preserve dts(1 To n)
    CSSFixDateCols = n

End Function

' ==============================================================
' Najde stlpec podla textu hlavicky v danom riadku
' ==============================================================
Public Function FindHeaderCol(ws As Worksheet, headerRow As Long, headerText As String) As Long

    Dim c As Long
    Dim lastCol As Long
    Dim v As Variant

    FindHeaderCol = 0

    lastCol = ws.Cells(headerRow, ws.Columns.count).End(xlToLeft).Column
    If lastCol > 200 Then lastCol = 200

    For c = 1 To lastCol
        v = ws.Cells(headerRow, c).Value
        If Not IsError(v) Then
            If StrComp(Trim(CStr(v)), headerText, vbTextCompare) = 0 Then
                FindHeaderCol = c
                Exit Function
            End If
        End If
    Next c

End Function

' ==============================================================
' Vytvori lokalny harok CSS, ak este neexistuje
' ==============================================================
Private Function GetOrCreateSheet(sheetName As String) As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.count))
        ws.Name = sheetName
    End If

    Set GetOrCreateSheet = ws

End Function

' ==============================================================
' Nacita harok PLAN do rovnakeho jednotneho zoznamu poloziek
' items(n,1)=Delivery note  (n,2)=Acme PN    (n,3)=Customer PN   (n,4)=qty
' Logika presunuta bez zmeny z povodnej nacitavacej slucky vo Vytvor_zoznam
' ==============================================================
Public Function LoadItemsPLAN(ByRef items() As Variant) As Long

    Dim wsPlan As Worksheet
    Dim planColPN As Long
    Dim planColCustPN As Long
    Dim planColQty As Long
    Dim planColNote As Long
    Dim planLastCol As Long
    Dim lastRowPlan As Long
    Dim c As Long
    Dim i As Long
    Dim cnt As Long
    Dim deliveryNote As String
    Dim tempNote As String
    Dim planPN As String
    Dim customerPN As String
    Dim requiredQty As Long

    LoadItemsPLAN = 0

    Set wsPlan = ThisWorkbook.Worksheets("PLAN")

    ' ---- najst stlpce podla nazvu hlavicky v riadku 3 ----
    planLastCol = wsPlan.Cells(3, wsPlan.Columns.count).End(xlToLeft).Column

    For c = 1 To planLastCol
        Select Case Trim(CStr(wsPlan.Cells(3, c).Text))
            Case "Acme PN": planColPN = c
            Case "Zakaznik PN": planColCustPN = c
            Case "KS": planColQty = c
            Case "Delivery note": planColNote = c
        End Select
    Next c

    If planColPN = 0 Or planColCustPN = 0 Or planColQty = 0 Or planColNote = 0 Then
        MsgBox "V hárku PLAN sa nenašiel jeden alebo viac povinných stĺpcov." & vbCrLf & vbCrLf & _
       "Skontrolujte, prosím:" & vbCrLf & _
       "- či hlavičky stĺpcov neboli premenované alebo odstránené," & vbCrLf & _
       "- či nad hlavičkou neboli vložené alebo odstránené riadky," & vbCrLf & _
       "- či je vložený správny a aktuálny export PLAN." & vbCrLf & vbCrLf & _
       "Názvy hlavičiek musia presne zodpovedať očakávaným názvom (pozor na medzery na konci).", _
       vbCritical, "Chýbajú stĺpce v hárku PLAN"
        Exit Function
    End If

    lastRowPlan = wsPlan.Cells(wsPlan.Rows.count, planColPN).End(xlUp).Row
    If lastRowPlan < 4 Then Exit Function

    ReDim items(1 To lastRowPlan, 1 To 4)
    cnt = 0
    deliveryNote = ""

    For i = 4 To lastRowPlan

        ' dodaci list sa cita PRED kontrolou PN, takze riadky iba s hlavickou
        ' stale aktualizuju prenasanu hodnotu
        If Not IsError(wsPlan.Cells(i, planColNote).Value) Then
            tempNote = Trim(CStr(wsPlan.Cells(i, planColNote).Value))
            If Left(tempNote, 4) = "1000" Or Left(tempNote, 4) = "2000" Then
                deliveryNote = tempNote
            End If
        End If

        If Not IsError(wsPlan.Cells(i, planColPN).Value) Then
            planPN = Trim(UCase(CStr(wsPlan.Cells(i, planColPN).Value)))
        Else
            planPN = ""
        End If

        If planPN = "" Then GoTo NextPlanRow

        ' iba dodacie listy 1000 / 2000
        If Left(deliveryNote, 4) <> "1000" And Left(deliveryNote, 4) <> "2000" Then
            GoTo NextPlanRow
        End If

        If Not IsError(wsPlan.Cells(i, planColCustPN).Value) Then
            customerPN = Trim(CStr(wsPlan.Cells(i, planColCustPN).Value))
        Else
            customerPN = ""
        End If

        If IsNumeric(wsPlan.Cells(i, planColQty).Value) Then
            requiredQty = CLng(wsPlan.Cells(i, planColQty).Value)
        Else
            requiredQty = 0
        End If

        If requiredQty <= 0 Then GoTo NextPlanRow

        cnt = cnt + 1
        items(cnt, 1) = deliveryNote
        items(cnt, 2) = planPN
        items(cnt, 3) = customerPN
        items(cnt, 4) = requiredQty

NextPlanRow:
    Next i

    LoadItemsPLAN = cnt

End Function

' ==============================================================
' Nacita riadky zvoleneho dna z harku CSS
' items(n,1)=Customer  (n,2)=Acme PN  (n,3)=Customer PN  (n,4)=qty
' ==============================================================
Public Function LoadItemsCSS(ByVal dateText As String, ByRef items() As Variant) As Long

    Dim ws As Worksheet
    Dim cols() As Long
    Dim dts() As Date
    Dim n As Long
    Dim i As Long
    Dim qtyCol As Long
    Dim custCol As Long
    Dim itemCol As Long
    Dim cpnCol As Long
    Dim lastRow As Long
    Dim r As Long
    Dim cnt As Long
    Dim v As Variant
    Dim q As Double
    Dim lastCust As String

    LoadItemsCSS = 0

    ' normalizuje cokolvek, co volajuci odovzdal, na DD.MM.YYYY (review #4)
    If IsDate(dateText) Then dateText = Format(CDate(dateText), "DD.MM.YYYY")

    Set ws = ThisWorkbook.Worksheets(CSS_SHEET)

    n = CSSFixDateCols(ws, cols, dts)
    If n = 0 Then
        MsgBox "V hárku CSS sa nenašli žiadne stĺpce s dátumami FIX." & vbCrLf & vbCrLf & _
       "Najprv spustite obnovenie dát tlačidlom 'Aktualizácia zdroja' a potom to skúste znova." & vbCrLf & vbCrLf & _
       "Ak sa hlásenie zobrazí aj po obnovení, skontrolujte, prosím:" & vbCrLf & _
       "- či obnovenie prebehlo bez chyby," & vbCrLf & _
       "- či zdrojový súbor CSS obsahuje stĺpce s dátumami FIX (správny a aktuálny export).", _
       vbExclamation, "Chýbajú dátumy FIX"
        Exit Function
    End If

    qtyCol = 0
    For i = 1 To n
        If Format(dts(i), "DD.MM.YYYY") = dateText Then
            qtyCol = cols(i)
            Exit For
        End If
    Next i

    If qtyCol = 0 Then
        MsgBox "Dátum " & dateText & " nie je v načítanom súbore CSS." & vbCrLf & _
               "Spusti 'Aktualizácia zdroja' a vyber deň znova.", vbExclamation
        Exit Function
    End If

    custCol = FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Customer")
    itemCol = FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Acme Item")
    cpnCol = FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Customer Item2")

    If custCol = 0 Or itemCol = 0 Or cpnCol = 0 Then
        MsgBox "Na hárku CSS chýba hlavička " & _
               "(Customer / Acme Item / Customer Item2).", vbCritical
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.count, itemCol).End(xlUp).Row

    ReDim items(1 To lastRow, 1 To 4)
    cnt = 0
    lastCust = ""

    For r = CSS_FIRST_DATA_ROW To lastRow

        ' prenasa zakaznika cez prazdne bunky
        v = ws.Cells(r, custCol).Value
        If Not IsError(v) Then
            If Trim(CStr(v)) <> "" Then lastCust = Trim(CStr(v))
        End If

        v = ws.Cells(r, qtyCol).Value
        If IsError(v) Then GoTo NextRow
        If Not IsNumeric(v) Then GoTo NextRow

        q = CDbl(v)
        If q <= 0 Then GoTo NextRow

        v = ws.Cells(r, itemCol).Value
        If IsError(v) Then GoTo NextRow
        If Trim(CStr(v)) = "" Then GoTo NextRow

        cnt = cnt + 1
        items(cnt, 1) = lastCust
        items(cnt, 2) = Trim(UCase(CStr(v)))

        v = ws.Cells(r, cpnCol).Value
        If IsError(v) Then
            items(cnt, 3) = ""
        Else
            items(cnt, 3) = Trim(CStr(v))
        End If

        items(cnt, 4) = CLng(q)

NextRow:
    Next r

    LoadItemsCSS = cnt

End Function


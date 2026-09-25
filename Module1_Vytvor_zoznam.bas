Attribute VB_Name = "Module1_Vytvor_zoznam"
Option Explicit

' ==============================================================
' Vytvor_zoznam - vytvori vychystavaci zoznam na harku TVOJ ZOZNAM
'
' Postup:
' 1. Obnovi data skladu (STOCK) a nacita polozky zo zvoleneho zdroja
'    (PLAN alebo CSS). Stary zoznam sa maze az potom, takze pri
'    predcasnom ukonceni ostane predchadzajuci zoznam nezmeneny.
' 2. Pre kazdu polozku najde PACK SIZE (harok PACKAGING); ak chyba,
'    polozka sa preskoci (s hlaskou). Potom spocita dostupne kusy,
'    podla lokacie z harka STOCK (C = stav, G = lokacia,
'    H = cislo dielu, R = mnozstvo).
'     - preskakuju sa lokacie LOC_EXCL_1, LOC_EXCL_2, LOC_EXCL_3, LOC_EXCL_4, LOC_EXCL_5
'       a riadky so stavom PICKED
' 3. Pravidlo celych baleni: ak je pozadovane mnozstvo >= PACK SIZE,
'    kazdy riadok skladu (davka) sa pocita len po celych nasobkoch
'    PACK SIZE, pocita sa cele dostupne mnozstvo.
' 4. Pridelenie: najskor sa hlada lokacia s presne pozadovanym
'    mnozstvom; ak taka nie je, berie sa od najvacsej lokacie a
'    pokraucje sa dalsimi, kym sa mnozstvo nepokryje.
' 5. Nedostatok sa zapise cervenym ako "-N" = pocet chybajucich
'    baleni (zaokruhlene nahor). Ak nie je ziadny pouzitelny sklad,
'    chybajucej casti (napr. treba 32 ks, je 16 ks a PACK SIZE 16
'    -> jedna lokacia + "-1").
' 6. Vysledok sa rozlozi do 3 stlpcov; hlavicka skupiny (dodaci list
'    alebo zakaznik) sa drzi pokope s prvou polozkou.
' ==============================================================

Sub Vytvor_zoznam()

    ' ==============================
    ' VARIABLES
    ' ==============================

    Dim wsStock As Worksheet
    Dim packDict As Object
    Dim wsDest As Worksheet
    Dim prevSU As Boolean

    Dim lastRowStock As Long
    Dim stockData As Variant

    Dim planPN As String
    Dim customerPN As String
    Dim requiredQty As Long
    Dim stillNeeded As Long
    Dim takeQty As Long

    Dim stockPN As String
    Dim location As String
    Dim packQty As Long
    Dim packSize As Long
    Dim rawPack As String

    Dim totalPicked As Long

    Dim groupHeader As String
    Dim lastGroup As String

    Dim allBlocks() As Variant
    Dim blockCount As Long
    Dim blockStarts() As Long
    Dim blockLengths() As Long
    Dim blockTypes() As String
    Dim numBlocks As Long

    ' --- vyber zdroja ---
    Dim srcMode As String
    Dim srcDate As String
    Dim items() As Variant
    Dim itemCount As Long
    Dim n As Long

    Dim key As Variant
    Dim tempArray()
    Dim k As Long
    Dim x As Long, y As Long
    Dim tempLoc As Variant
    Dim tempQty As Variant
    Dim b As Long
    Dim r As Long
    Dim j As Long

    Dim missingPacks As Long
    Dim missingQty As Long

    Dim allocationDict As Object
    Dim locationTotals As Object

    Dim totalRows As Long
    Dim targetRows As Long
    Dim colStarts(1 To 3) As Long
    Dim currentCol As Long
    Dim currentRow As Long
    Dim rowsInCurrentCol As Long
    Dim startIdx As Long
    Dim writeRow As Long
    Dim writeCol As Long
    Dim nextBlockLen As Long
    Dim cellVal As Variant

    Dim exactFound As Boolean
    Dim requiresFullPack As Boolean

    ' ==============================
    ' NAJSKOR OBNOVIT DATA SKLADU
    ' ==============================

    prevSU = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo CleanFail

    RefreshStock

    ' velkosti balenia z dotazu Packaging; pri zlyhani obnovenia sa pouzivatela opyta
    If Not RefreshPackaging() Then GoTo CleanExit
    Set packDict = LoadPackSizes()
    If packDict Is Nothing Then GoTo CleanExit

    ' ==============================
    ' NACITAT POLOZKY Z VYBRANEHO ZDROJA
    ' vykona sa PRED zmazanim stareho zoznamu, takze predcasne ukoncenie
    ' ponecha predchadzajuci zoznam nezmeneny
    ' ==============================

    srcMode = UCase(Trim(CStr(Worksheets("PANEL").Range(PANEL_SOURCE_CELL).Value)))
    If srcMode = "" Then srcMode = "PLAN"

    If srcMode = "CSS" Then
        If ExportDate() = 0 Then
            MsgBox "Nie je vybraný dátum exportu. Vyber dátum na hárku PANEL.", vbExclamation
            GoTo CleanExit
        End If
        srcDate = Format(ExportDate(), "DD.MM.YYYY")
        itemCount = LoadItemsCSS(srcDate, items)
    Else
        itemCount = LoadItemsPLAN(items)
    End If

    If itemCount = 0 Then
        MsgBox "Pre vybraný zdroj (" & srcMode & ") nie sú žiadne položky na vychystanie." & vbCrLf & _
       "Predchádzajúci zoznam zostal nezmenený." & vbCrLf & vbCrLf & _
       "Skontrolujte, prosím:" & vbCrLf & _
       "- či je vybraný správny zdroj," & vbCrLf & _
       "- či sú zdrojové dáta aktuálne (načítané/obnovené)," & vbCrLf & _
       "- či položky nemajú nulové množstvo alebo už neboli vychystané.", _
       vbExclamation, "Žiadne položky na vychystanie"
        GoTo CleanExit
    End If

    ' ==============================
    ' PRIPRAVA HARKOV
    ' ==============================

    Set wsStock = Worksheets("STOCK")

    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("TVOJ ZOZNAM").Delete
    Application.DisplayAlerts = True
    On Error GoTo CleanFail

    Set wsDest = Worksheets.Add
    wsDest.Name = "TVOJ ZOZNAM"
    wsDest.Activate
    ActiveWindow.DisplayGridlines = False

    lastRowStock = wsStock.Cells(wsStock.Rows.count, "H").End(xlUp).Row

    ' nacita stlpce harku STOCK, ktore pouzivame, do pamate naraz - A:R pokryva C, G, H, R
    If lastRowStock >= 2 Then stockData = wsStock.Range("A1:R" & lastRowStock).Value

    blockCount = 0
    numBlocks = 0
    ReDim allBlocks(1 To 1000, 1 To 2)
    ReDim blockStarts(1 To 1000)
    ReDim blockLengths(1 To 1000)
    ReDim blockTypes(1 To 1000)

    lastGroup = ""

    ' ==============================
    ' LOOP CEZ POLOZKY
    ' ==============================

    For n = 1 To itemCount

        groupHeader = CStr(items(n, 1))
        planPN = CStr(items(n, 2))
        customerPN = CStr(items(n, 3))
        requiredQty = CLng(items(n, 4))

        ' ==============================
        ' NAJST PACK_SIZE
        ' ==============================

        packSize = 0
        If packDict.Exists(planPN) Then packSize = packDict(planPN)

        If packSize = 0 Then
            MsgBox "Chýba PACK SIZE pre položku: " & planPN & vbCrLf & _
                   "(skontroluj hárok Packaging v súbore PLAN)"
            GoTo NextItem
        End If
        
        requiresFullPack = (requiredQty >= packSize)

        totalPicked = 0

        Set allocationDict = CreateObject("Scripting.Dictionary")
        Set locationTotals = CreateObject("Scripting.Dictionary")

        ' ==============================
        ' SUCET KUSOV PODLA LOKACIE
        ' ==============================

        For j = 2 To lastRowStock

            If Not IsError(stockData(j, 8)) Then
                stockPN = Trim(UCase(CStr(stockData(j, 8))))
            Else
                stockPN = ""
            End If

            If stockPN = planPN Then

                rawPack = Trim(CStr(stockData(j, 18)))
                rawPack = Replace(rawPack, ".", ",")
                If IsNumeric(rawPack) Then
                    packQty = CLng(CDbl(rawPack))
                Else
                    packQty = 0
                End If

                If requiresFullPack Then
                    ' objednavka potrebuje jedno alebo viac celych baleni - davka mensia
                    ' ako jedno balenie nemoze pokryt ani jedno z nich; ponecha sa v sklade
                    packQty = Int(packQty / packSize) * packSize
                End If

                If IsError(stockData(j, 7)) Then GoTo SkipLocation
                location = stockData(j, 7)

                If UCase(location) = "LOC_EXCL_1" Then GoTo SkipLocation
                If UCase(location) = "LOC_EXCL_2" Then GoTo SkipLocation
                If UCase(location) = "LOC_EXCL_3" Then GoTo SkipLocation
                If UCase(location) = "LOC_EXCL_4" Then GoTo SkipLocation
                If UCase(location) = "LOC_EXCL_5" Then GoTo SkipLocation
                If UCase(Trim(CStr(stockData(j, 3)))) = "PICKED" Then GoTo SkipLocation

                If packQty > 0 Then
                    If Not locationTotals.Exists(location) Then
                        locationTotals.Add location, packQty
                    Else
                        locationTotals(location) = locationTotals(location) + packQty
                    End If
                End If

            End If

SkipLocation:
        Next j

        If locationTotals.count = 0 Then
            missingPacks = Application.WorksheetFunction.RoundUp(requiredQty / packSize, 0)
            allocationDict.Add "-" & missingPacks, ""
            GoTo StoreBlock
        End If

        ' ==============================
        ' PREVOD NA POLE A ZORADENIE
        ' ==============================

        ReDim tempArray(1 To locationTotals.count, 1 To 2)
        k = 1
        For Each key In locationTotals.Keys
            tempArray(k, 1) = key
            tempArray(k, 2) = locationTotals(key)
            k = k + 1
        Next key

        For x = 1 To UBound(tempArray) - 1
            For y = x + 1 To UBound(tempArray)
                If tempArray(x, 2) < tempArray(y, 2) Then
                    tempLoc = tempArray(x, 1)
                    tempQty = tempArray(x, 2)
                    tempArray(x, 1) = tempArray(y, 1)
                    tempArray(x, 2) = tempArray(y, 2)
                    tempArray(y, 1) = tempLoc
                    tempArray(y, 2) = tempQty
                End If
            Next y
        Next x

        ' ==============================
        ' ALLOCATE
        ' ==============================

        ' Prvy prechod - hladanie presnej zhody
        exactFound = False
        For x = 1 To UBound(tempArray)
            If tempArray(x, 2) = requiredQty Then
                allocationDict.Add tempArray(x, 1), tempArray(x, 2)
                totalPicked = requiredQty
                exactFound = True
                Exit For
            End If
        Next x

        ' Druhy prechod - ak nie je presna zhoda, berie sa najskor z najvacsej lokacie
        If Not exactFound Then
            For x = 1 To UBound(tempArray)
                If totalPicked >= requiredQty Then Exit For
                location = tempArray(x, 1)
                packQty = tempArray(x, 2)
                If packQty <= 0 Then GoTo NextLocation
                stillNeeded = requiredQty - totalPicked
                takeQty = Application.WorksheetFunction.Min(packQty, stillNeeded)
                allocationDict.Add location, takeQty
                totalPicked = totalPicked + takeQty
NextLocation:
            Next x
        End If

    If totalPicked < requiredQty Then
        missingQty = requiredQty - totalPicked
        missingPacks = Application.WorksheetFunction.RoundUp(missingQty / packSize, 0)
        allocationDict.Add "-" & missingPacks, ""
    End If

        ' ==============================
        ' ULOZIT HLAVICKY PRI ZMENE
        ' ==============================

StoreBlock:
        If allocationDict.count > 0 Then

            ' Hlavicka skupiny - dodaci list (PLAN) alebo zakaznik (CSS)
            If groupHeader <> lastGroup Then
                numBlocks = numBlocks + 1
                If numBlocks > UBound(blockStarts) Then
                    ReDim Preserve blockStarts(1 To numBlocks + 100)
                    ReDim Preserve blockLengths(1 To numBlocks + 100)
                    ReDim Preserve blockTypes(1 To numBlocks + 100)
                End If
                blockStarts(numBlocks) = blockCount + 1
                blockTypes(numBlocks) = "NOTE"

                blockCount = blockCount + 1
                If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
                allBlocks(blockCount, 1) = "-- " & groupHeader & " --"
                allBlocks(blockCount, 2) = ""

                blockCount = blockCount + 1
                If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
                allBlocks(blockCount, 1) = ""
                allBlocks(blockCount, 2) = ""

                blockLengths(numBlocks) = blockCount - blockStarts(numBlocks) + 1
                lastGroup = groupHeader
            End If

            ' Blok polozky
            numBlocks = numBlocks + 1
            If numBlocks > UBound(blockStarts) Then
                ReDim Preserve blockStarts(1 To numBlocks + 100)
                ReDim Preserve blockLengths(1 To numBlocks + 100)
                ReDim Preserve blockTypes(1 To numBlocks + 100)
            End If
            blockStarts(numBlocks) = blockCount + 1
            blockTypes(numBlocks) = "ITEM"

            blockCount = blockCount + 1
            If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
            allBlocks(blockCount, 1) = customerPN
            allBlocks(blockCount, 2) = ""

            For Each key In allocationDict.Keys
                blockCount = blockCount + 1
                If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
                allBlocks(blockCount, 1) = key
                If allocationDict(key) <> "" Then
                    allBlocks(blockCount, 2) = allocationDict(key)
                Else
                    allBlocks(blockCount, 2) = ""
                End If
            Next key

            blockCount = blockCount + 1
            If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
            allBlocks(blockCount, 1) = ""
            allBlocks(blockCount, 2) = ""

            blockCount = blockCount + 1
            If blockCount > UBound(allBlocks) Then ReDim Preserve allBlocks(1 To blockCount + 100, 1 To 2)
            allBlocks(blockCount, 1) = ""
            allBlocks(blockCount, 2) = ""

            blockLengths(numBlocks) = blockCount - blockStarts(numBlocks) + 1

        End If

NextItem:
    Next n

    ' ==============================
    ' ROZDELENIE DO 3 STLPCOV
    ' ==============================

    totalRows = blockCount
    targetRows = Application.WorksheetFunction.RoundUp(totalRows / 3, 0)

    colStarts(1) = 1
    colStarts(2) = 4
    colStarts(3) = 7

    currentCol = 1
    currentRow = 1
    rowsInCurrentCol = 0

    For b = 1 To numBlocks

        nextBlockLen = 0
        If b < numBlocks Then nextBlockLen = blockLengths(b + 1)

        If blockTypes(b) = "NOTE" Then
            If rowsInCurrentCol > 0 And rowsInCurrentCol + blockLengths(b) + nextBlockLen > targetRows And currentCol < 3 Then
                currentCol = currentCol + 1
                currentRow = 1
                rowsInCurrentCol = 0
            End If
        Else
            If rowsInCurrentCol > 0 And rowsInCurrentCol + blockLengths(b) > targetRows And currentCol < 3 Then
                currentCol = currentCol + 1
                currentRow = 1
                rowsInCurrentCol = 0
            End If
        End If

        startIdx = blockStarts(b)

        For r = 0 To blockLengths(b) - 1

            writeRow = currentRow + r
            writeCol = colStarts(currentCol)
            cellVal = allBlocks(startIdx + r, 1)

            wsDest.Cells(writeRow, writeCol).Value = cellVal
            wsDest.Cells(writeRow, writeCol + 1).Value = allBlocks(startIdx + r, 2)
            wsDest.Cells(writeRow, writeCol).HorizontalAlignment = xlHAlignCenter
            wsDest.Cells(writeRow, writeCol + 1).HorizontalAlignment = xlHAlignCenter

            If blockTypes(b) = "NOTE" And r = 0 Then
                wsDest.Cells(writeRow, writeCol).Font.Bold = True
                wsDest.Cells(writeRow, writeCol).Font.Size = 11
                wsDest.Cells(writeRow, writeCol).Font.Italic = True
            End If

            If blockTypes(b) = "ITEM" And r = 0 Then
                wsDest.Cells(writeRow, writeCol).Font.Bold = True
                wsDest.Cells(writeRow, writeCol).Font.Size = 12
                wsDest.Cells(writeRow, writeCol).Font.Underline = xlUnderlineStyleSingle
            End If

            If Left(CStr(allBlocks(startIdx + r, 1)), 1) = "-" Then
                wsDest.Cells(writeRow, writeCol).Font.Color = RGB(255, 0, 0)
            End If

        Next r

        currentRow = currentRow + blockLengths(b)
        rowsInCurrentCol = rowsInCurrentCol + blockLengths(b)

    Next b

    ' ==============================
    ' FINALNE FORMATOVANIE
    ' ==============================

    wsDest.Columns("A").ColumnWidth = 20
    wsDest.Columns("B").ColumnWidth = 8
    wsDest.Columns("C").ColumnWidth = 2
    wsDest.Columns("D").ColumnWidth = 20
    wsDest.Columns("E").ColumnWidth = 8
    wsDest.Columns("F").ColumnWidth = 2
    wsDest.Columns("G").ColumnWidth = 20
    wsDest.Columns("H").ColumnWidth = 8

    With wsDest.PageSetup
        .Orientation = xlLandscape
        .PrintArea = "$A:$H"
        .FitToPagesWide = 1
        .FitToPagesTall = False
        .Zoom = False
    End With

    If srcMode = "CSS" Then
        Worksheets("PANEL").Cells(17, 2).Value = "Zoznam: " & Now() & "  (CSS " & srcDate & ")"
    Else
        Worksheets("PANEL").Cells(17, 2).Value = "Zoznam: " & Now() & "  (PLAN)"
    End If

    MsgBox "Zoznam úspešne vytvorený" & vbCrLf & _
           "Zdroj: " & srcMode & "  počet položiek: " & itemCount, vbInformation

CleanExit:
    Application.ScreenUpdating = prevSU
    Exit Sub

CleanFail:
    Application.ScreenUpdating = prevSU
    MsgBox "Vytvor zoznam - chyba " & Err.Number & ": " & Err.Description, vbCritical
End Sub

' ==============================================================
' Nacita harok zvoleneho dna z externeho suboru PLAN do
' lokalneho harku PLAN. Slovensky nazov dna sa odvodzuje z
' datumu exportu (PANEL F5), nie z bunky.
' Vrati sa True len ak je uspesny
' ==============================================================
Public Function RefreshPlan() As Boolean

    Dim planPath As String
    Dim dayName As String
    Dim chosen As Date
    Dim wbPlan As Workbook
    Dim wsSrc As Worksheet
    Dim wsDest As Worksheet
    Dim src As Range
    Dim prevEvents As Boolean
    Dim errNum As Long
    Dim errDesc As String

    RefreshPlan = False

    chosen = ExportDate()
    If chosen = 0 Then
        MsgBox "Vyber deň exportu (PANEL F5).", vbExclamation
        Exit Function
    End If

    dayName = SlovakDayName(chosen)
    If dayName = "" Then
        MsgBox "Dátum " & Format(chosen, "DD.MM.YYYY") & " padne na víkend.", vbExclamation
        Exit Function
    End If

    planPath = ResolvePath(PATH_PLAN_NAME)
    If planPath = "" Then
        MsgBox "Cesta k PLAN súboru je prázdna." & vbCrLf & _
               "Zadaj cestu do hárku NASTAVENIA.", vbExclamation
        Exit Function
    End If

    Set wsDest = ThisWorkbook.Worksheets("PLAN")

    prevEvents = Application.EnableEvents
    Application.EnableEvents = False
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    On Error GoTo CleanFail
    Set wbPlan = Workbooks.Open(Filename:=planPath, ReadOnly:=True, UpdateLinks:=0)

    On Error Resume Next
    Set wsSrc = wbPlan.Worksheets(dayName)
    On Error GoTo CleanFail

    If wsSrc Is Nothing Then
        wbPlan.Close SaveChanges:=False
        GoTo Restore_Warn
    End If

    ' priradenie hodnot - odolne voci filtru ponechanemu na zdrojovom harku
    Set src = wsSrc.UsedRange
    wsDest.Cells.Clear
    wsDest.Range(src.Address).Value = src.Value

    wbPlan.Close SaveChanges:=False

    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True

    ' Krok 4: subor PLAN ma en 5 harkov dni, ktore sa opakuju kazdy tyzden -
    ' overi sa, ci prave nacitany harok je naozaj datovany na zvoleny den.
    ' Pri zamietnuti sa harok PLAN vymaze, aby nasledujuci Vytvor_zoznam nenasiel nic
    If VerifyPlanDate(wsDest, dayName, chosen) Then
        RefreshPlan = True
    Else
        wsDest.Cells.Clear
        ThisWorkbook.Worksheets("PANEL").Range("B16").Value = "Zdroj: PLAN zamietnutý " & Format(Now(), "DD.MM.YYYY HH:NN")
        RefreshPlan = False
    End If
    Exit Function

Restore_Warn:
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    MsgBox "PLAN súbor nemá hárok '" & dayName & "'.", vbCritical
    Exit Function

CleanFail:
    errNum = Err.Number
    errDesc = Err.Description
    On Error Resume Next
    If Not wbPlan Is Nothing Then wbPlan.Close SaveChanges:=False
    Application.EnableEvents = prevEvents
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error GoTo 0
    MsgBox "Aktualizácia PLAN zlyhala: " & errNum & " - " & errDesc, vbCritical

End Function

' ==============================================================
' Pomocne funkcie ku kroku 4 - overenie, ci je nacitany harok PLAN pre zvoleny den
' ==============================================================

' Datum "Vyvoz:", ktory harok planu nesie v bunke C1 (seriove cislo datumu).
' Ak sa nenajde, prehlada sa lavy horny blok pre popisok "Vyvoz".
' Vrati 0, ak sa nenajde ziadny hodnoverny datum.
Public Function PlanSheetDate(ws As Worksheet) As Date

    Dim v As Variant
    Dim r As Long
    Dim c As Long

    v = ws.Range("C1").Value
    If PlausibleDate(v) Then
        PlanSheetDate = CoerceDate(v)
        Exit Function
    End If

    For r = 1 To 3
        For c = 1 To 6
            If InStr(1, CStr(ws.Cells(r, c).Value), "Vývoz", vbTextCompare) > 0 Then
                v = ws.Cells(r, c + 1).Value
                If PlausibleDate(v) Then PlanSheetDate = CoerceDate(v)
                Exit Function
            End If
        Next c
    Next r

End Function

Private Function PlausibleDate(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsDate(v) Then
        PlausibleDate = True
    ElseIf IsNumeric(v) Then
        PlausibleDate = (CDbl(v) >= 30000 And CDbl(v) <= 80000)
    End If
End Function

Private Function CoerceDate(ByVal v As Variant) As Date
    If IsDate(v) Then
        CoerceDate = CDate(v)
    ElseIf IsNumeric(v) Then
        CoerceDate = CDate(CDbl(v))
    End If
End Function

' Vrati sa True na pokracovanie (datumy sa zhoduju, alebo sa pouzivatel
' rozhodol plan pouzit aj tak), False, ak pouzivatel zamietol
' neaktualny / neoveritelny plan.
Private Function VerifyPlanDate(ws As Worksheet, ByVal dayName As String, ByVal chosen As Date) As Boolean

    Dim sheetDate As Date
    sheetDate = PlanSheetDate(ws)

    If sheetDate = 0 Then
        VerifyPlanDate = (MsgBox( _
            "Hárok '" & dayName & "' nemá dátum vývozu (C1)." & vbCrLf & _
            "Nedá sa overiť, či je plán pre " & Format(chosen, "DD.MM.YYYY") & " hotový." & vbCrLf & vbCrLf & _
            "Pokračovať?", vbYesNo + vbExclamation, "Kontrola dátumu PLAN") = vbYes)

    ElseIf Int(sheetDate) <> Int(chosen) Then
        VerifyPlanDate = (MsgBox( _
            "POZOR - nesúlad dátumov:" & vbCrLf & vbCrLf & _
            "   hárok '" & dayName & "' je datovaný   " & Format(sheetDate, "DD.MM.YYYY") & vbCrLf & _
            "   zvolený deň exportu   " & Format(chosen, "DD.MM.YYYY") & vbCrLf & vbCrLf & _
            "Plán pre tvoj deň zrejme ešte nie je pripravený" & vbCrLf & _
            "a použil sa starší plán z hárku '" & dayName & "'." & vbCrLf & vbCrLf & _
            "Použiť tento plán aj tak?", vbYesNo + vbExclamation, "Kontrola dátumu PLAN") = vbYes)

    Else
        VerifyPlanDate = True
    End If

End Function




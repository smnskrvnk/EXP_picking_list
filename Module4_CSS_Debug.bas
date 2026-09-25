Attribute VB_Name = "Module4_CSS_Debug"
Option Explicit

' ==============================================================
' TEST HELPER - Immediate window
' ==============================================================
Sub TestLoadCSS()

    Dim items() As Variant
    Dim n As Long
    Dim i As Long
    Dim dateText As String
    Dim cust As String
    Dim total As Double

    dateText = Trim(CStr(ThisWorkbook.Worksheets("PANEL").Range(PANEL_DATE_CELL).Value))
    If dateText = "" Then
        dateText = Trim(CStr(ThisWorkbook.Worksheets("PACK_SIZE").Cells(1, CSS_LIST_COL).Value))
    End If

    Debug.Print "--- CSS items for " & dateText & " ---"

    n = LoadItemsCSS(dateText, items)
    Debug.Print "rows: " & n

    cust = Chr(1)
    For i = 1 To n
        If CStr(items(i, 1)) <> cust Then
            cust = CStr(items(i, 1))
            Debug.Print "  [" & cust & "]"
        End If
        Debug.Print "    " & items(i, 2) & "   " & items(i, 3) & "   qty " & items(i, 4)
        total = total + CDbl(items(i, 4))
    Next i

    Debug.Print "total pcs: " & total

End Sub

' ==============================================================
' TESTOVACIA POMÔCKA - spusti po RefreshCSS, skontroluj okno Immediate
' ==============================================================
Sub TestCSSDates()

    Dim cols() As Long
    Dim dts() As Date
    Dim n As Long
    Dim i As Long
    Dim ws As Worksheet

    Set ws = ThisWorkbook.Worksheets(CSS_SHEET)

    Debug.Print "Acme Item col : " & FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Acme Item")
    Debug.Print "Customer Item2 col: " & FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Customer Item2")
    Debug.Print "Customer col     : " & FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Customer")
    Debug.Print "Backlog anchor   : " & FindHeaderCol(ws, CSS_HDR_NAME_ROW, "Ship_Sch_Backlog")

    n = CSSFixDateCols(ws, cols, dts)
    Debug.Print "FIX date columns : " & n

    For i = 1 To n
        Debug.Print "  " & Format(dts(i), "DD.MM.YYYY") & "  ->  col " & cols(i) & _
                    " (" & Split(ws.Cells(1, cols(i)).Address, "$")(1) & ")"
    Next i

End Sub

Sub DumpCSSHeaders()

    Dim ws As Worksheet
    Dim c As Long
    Dim v20 As Variant
    Dim v21 As Variant
    Dim s20 As String
    Dim s21 As String

    Set ws = ThisWorkbook.Worksheets(CSS_SHEET)

    For c = 1 To 100
        v20 = ws.Cells(CSS_HDR_DATE_ROW, c).Value
        v21 = ws.Cells(CSS_HDR_NAME_ROW, c).Value

        s20 = "": s21 = ""
        If Not IsError(v20) Then s20 = Left(CStr(v20), 24)
        If Not IsError(v21) Then s21 = Left(CStr(v21), 40)

        If s20 <> "" Or s21 <> "" Then
            Debug.Print Split(ws.Cells(1, c).Address, "$")(1) & _
                        " | r20=" & s20 & _
                        " | r21=" & s21
        End If
    Next c

End Sub

' ==============================================================
' TESTOVACIA POMOCKA - porovna rucnu tabulku PACK_SIZE (stlpce B:C)
' s harkom PACKAGING (dotaz Packaging). Vystup je v okne Immediate.
' Najprv spusti RefreshPackaging. Nic v zosite nemeni.
' ==============================================================
Sub ComparePackSizes()

    Dim wsPack As Worksheet
    Dim lo As ListObject
    Dim pkg As Variant
    Dim man As Variant
    Dim dPkg As Object
    Dim dMan As Object
    Dim i As Long
    Dim lastRow As Long
    Dim k As String
    Dim mVal As Double
    Dim pVal As Double
    Dim key As Variant
    Dim nSame As Long
    Dim nDiff As Long
    Dim nOnlyMan As Long
    Dim nOnlyPkg As Long
    Dim nDupPkg As Long
    Dim nDupMan As Long
    Dim sDiff As String
    Dim sOnlyMan As String
    Dim sOnlyPkg As String
    Dim sDupPkg As String
    Dim sDupMan As String

    Set wsPack = ThisWorkbook.Worksheets("PACK_SIZE")

    On Error Resume Next
    Set lo = ThisWorkbook.Worksheets(PACKAGING_SHEET).ListObjects(1)
    On Error GoTo 0
    If lo Is Nothing Then
        Debug.Print "No table on sheet " & PACKAGING_SHEET
        Exit Sub
    End If
    If lo.DataBodyRange Is Nothing Then
        Debug.Print "PACKAGING table is empty"
        Exit Sub
    End If

    Set dPkg = CreateObject("Scripting.Dictionary")
    Set dMan = CreateObject("Scripting.Dictionary")

    ' --- PACKAGING: stlpec 2 = cislo dielu, stlpec 3 = velkost balenia ---
    pkg = lo.DataBodyRange.Value
    For i = 1 To UBound(pkg, 1)
        If Not IsError(pkg(i, 2)) Then
            k = UCase(Trim(CStr(pkg(i, 2))))
            If k <> "" Then
                pVal = ToNum(pkg(i, 3))
                If dPkg.Exists(k) Then
                    If dPkg(k) <> pVal Then
                        nDupPkg = nDupPkg + 1
                        AddLine sDupPkg, nDupPkg, k & "  " & dPkg(k) & " vs " & pVal
                    End If
                Else
                    dPkg.Add k, pVal
                End If
            End If
        End If
    Next i

    ' --- rucna tabulka PACK_SIZE: stlpec B = cislo dielu, C = velkost balenia ---
    lastRow = wsPack.Cells(wsPack.Rows.count, "B").End(xlUp).Row
    If lastRow < 2 Then
        Debug.Print "PACK_SIZE has no data rows"
        Exit Sub
    End If

    man = wsPack.Range("B2:C" & lastRow).Value
    For i = 1 To UBound(man, 1)
        If Not IsError(man(i, 1)) Then
            k = UCase(Trim(CStr(man(i, 1))))
            If k <> "" Then
                mVal = ToNum(man(i, 2))

                If dMan.Exists(k) Then
                    If dMan(k) <> mVal Then
                        nDupMan = nDupMan + 1
                        AddLine sDupMan, nDupMan, k & "  " & dMan(k) & " vs " & mVal
                    End If
                Else
                    dMan.Add k, mVal
                End If

                If Not dPkg.Exists(k) Then
                    nOnlyMan = nOnlyMan + 1
                    AddLine sOnlyMan, nOnlyMan, k & "  manual=" & mVal
                ElseIf dPkg(k) <> mVal Then
                    nDiff = nDiff + 1
                    AddLine sDiff, nDiff, k & "  manual=" & mVal & "  packaging=" & dPkg(k)
                Else
                    nSame = nSame + 1
                End If
            End If
        End If
    Next i

    ' --- polozky, ktore su len v PACKAGING ---
    For Each key In dPkg.Keys
        If Not dMan.Exists(CStr(key)) Then
            nOnlyPkg = nOnlyPkg + 1
            AddLine sOnlyPkg, nOnlyPkg, key & "  packaging=" & dPkg(key)
        End If
    Next key

    Debug.Print "=== PACK_SIZE (manual) vs PACKAGING ==="
    Debug.Print "manual rows compared    : " & (nSame + nDiff + nOnlyMan)
    Debug.Print "same size               : " & nSame
    Debug.Print "DIFFERENT size          : " & nDiff
    Debug.Print "only in manual          : " & nOnlyMan
    Debug.Print "only in PACKAGING       : " & nOnlyPkg
    Debug.Print "conflicting duplicates  : manual " & nDupMan & ", packaging " & nDupPkg

    PrintList "DIFFERENT size (max 25)", sDiff
    PrintList "Only in manual (max 25)", sOnlyMan
    PrintList "Only in PACKAGING (max 25)", sOnlyPkg
    PrintList "Conflicting duplicates in manual (max 25)", sDupMan
    PrintList "Conflicting duplicates in PACKAGING (max 25)", sDupPkg

End Sub

Private Sub AddLine(ByRef lst As String, ByVal cnt As Long, ByVal txt As String)
    If cnt <= 25 Then lst = lst & "  " & txt & vbCrLf
End Sub

Private Sub PrintList(ByVal title As String, ByVal lst As String)
    If lst <> "" Then Debug.Print vbCrLf & title & ":" & vbCrLf & lst
End Sub

Private Function ToNum(ByVal v As Variant) As Double
    If IsError(v) Then
        ToNum = -1
    ElseIf IsNumeric(v) Then
        ToNum = CDbl(v)
    Else
        ToNum = -1
    End If
End Function

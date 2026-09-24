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




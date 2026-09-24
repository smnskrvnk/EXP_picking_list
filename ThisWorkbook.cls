VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Tento_zošit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
    ' Krok 3: predvolene nastaví zdroj na PLAN a (znovu) zostaví zoznam dátumov v F5.
    SetDefaultExport
End Sub


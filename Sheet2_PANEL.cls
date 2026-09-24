VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

' Hárok PANEL (kódový názov Sheet2)
' Krok 3: B5 obsahuje zdroj (PLAN / CSS). Keď sa zmení, znovu sa zostaví
' rozbaľovací zoznam dátumu exportu v F5. F5 obsluhu nepotrebuje - je to vstup.

Private Sub Worksheet_Change(ByVal Target As Range)

    If Application.Intersect(Target, Me.Range("B5")) Is Nothing Then Exit Sub

    Application.EnableEvents = False
    On Error Resume Next
    BuildExportDateList
    On Error GoTo 0
    Application.EnableEvents = True

End Sub


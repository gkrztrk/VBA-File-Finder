VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub Adres_Bul()
Dim a, b, c, d As Integer
a = 2000
b = 3
a = ThisWorkbook.Worksheets("kontrol").Cells(a, 5).End(xlDown).Row + 1
Do


If ThisWorkbook.Worksheets("kontrol").Cells(a, 5) = 0 Then

ThisWorkbook.Worksheets("Sheet1 (2)").Cells(3, 3).Value = ThisWorkbook.Worksheets("kontrol").Cells(a, 3).Value


    If ThisWorkbook.Worksheets("Sheet1 (2)").Cells(6, 2).Value <> "" Then
    
    ThisWorkbook.Worksheets("kontrol").Cells(a, 5) = ThisWorkbook.Worksheets("Sheet1 (2)").Cells(6, 2).Value
    
    Else
    
    ThisWorkbook.Worksheets("kontrol").Cells(a, 5) = "DOSYA BULUNAMADI"
    
    
    End If
    
    d = 300
    c = d Mod 101
    'If c = 0 Then
    
    'ThisWorkbook.Save
    
    'End If
    
    
End If
d = d + 1
a = a + 1
DoEvents

Loop While ThisWorkbook.Worksheets("kontrol").Cells(a, 2).Value <> 0

ThisWorkbook.Save

End Sub

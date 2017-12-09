Attribute VB_Name = "test"
Sub test()
  Dim wkb As Workbook, wks As Worksheet
     
     Dim i As Integer, j As Integer, k As Integer                    'i IS order amount
     
     Dim TXTS As String, TXTF As String
     Dim ts As Range, tf As Range, tm As Range, ms As Range, mf As Range
     Application.ScreenUpdating = False
     
     
     Set TWS1 = ThisWorkbook.Worksheets("order detail")                 'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")                   'target worksheet
'   Set TWS3 = ThisWorkbook.Worksheets("bank detail collect report")   'target worksheet
    Set TWS4 = ThisWorkbook.Worksheets("shipping mark")         'target worksheet
    Set TWS5 = ThisWorkbook.Worksheets("collect information")   'target worksheet
    Set TWS6 = ThisWorkbook.Worksheets("checkdata")             'target worksheet
    
    ActiveCell.NumberFormat = "гд #,##0.00"
    MsgBox (ActiveCell.Font.Name)
    
End Sub
Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function

Function autosum1(rng As Range) As String
Dim i As Integer, j As Integer, R As Integer, C As Integer

For i = 1 To 10000
    If IsNumeric(rng.Offset(0 - i, 0)) Then
    Else: Exit For
    End If
    Next
rng.Value = "=SUM(R[-" & i - 1 & "]C:R[-1]C)"
 
 
End Function



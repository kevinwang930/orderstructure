Attribute VB_Name = "banktoorder"
Sub banktoorder()
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
    
     
     Dim orderrowstart(70) As Long
     Dim orderrowfinish(70) As Long
     Dim orderno(70) As String
     Dim orderqty(70) As Long
     Dim orderamount(70) As Double
     Dim orderctn(70) As Long
     Dim ordergweight(70) As Long
     Dim ordernweight(70) As Long
     Dim modelstart(70) As Long
     Dim modelfinish(70) As Long
     Dim ordercbm(70) As Double
     Dim suppliername(70) As String
     
     Dim test As String
     

     
     TXTS = "YW1117"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    Set tf = TWS1.Range("A1")
    
    Dim bankheader As Integer
    bankheader = 7
    TWS2.Cells.Font.Size = 16
    TWS2.Rows(5).Font.Size = 18
    
    f = Dir(ThisWorkbook.path & "\information*")
     If f <> "" Then
        Set wkb = Workbooks.Open(ThisWorkbook.path & "\" & f)
        Set wks1 = wkb.Worksheets("bank detail")
     End If
    For i = 1 To 70
     
            Set ts = finddown(TWS1.UsedRange, TXTS, ts)                             'find order start
            If ts Is Nothing Then
                Exit For
            Else
                Set tf = finddown(TWS1.UsedRange, TXTF, ts)                         'find order finish
                If tf Is Nothing Then
                    
                    MsgBox ("ORDERROWS have start, but do not have finish")
                    Exit For
                Else
                    Set ms = finddown(TWS1.UsedRange, "Article No", ts)
                    modelstart(i) = ms.Row + 1
                    modelfinish(i) = tf.Row - 1
                    orderrowstart(i) = ts.Row - 1
                    orderrowfinish(i) = tf.Row
                    orderno(i) = ts.Value
                    orderqty(i) = Application.Sum(TWS1.Range("H" & orderrowstart(i) & ":H" & orderrowfinish(i) - 1))
                    orderamount(i) = Round(Application.Sum(TWS1.Range("J" & orderrowstart(i) & ":J" & orderrowfinish(i) - 1)), 2)
                    orderctn(i) = Application.Sum(TWS1.Range("G" & orderrowstart(i) & ":G" & orderrowfinish(i) - 1))
                    ordergweight(i) = Application.Sum(TWS1.Range("p" & orderrowstart(i) & ":p" & orderrowfinish(i) - 1))
                    ordernweight(i) = Application.Sum(TWS1.Range("q" & orderrowstart(i) & ":q" & orderrowfinish(i) - 1))
                    ordercbm(i) = Application.Sum(TWS1.Range("N" & modelstart(i) & ":N" & modelfinish(i)))
                    suppliername(i) = TWS1.Range("A" & orderrowstart(i)).Value
                   
                End If
            End If
            
    'cut chinese supplier name in suppliername(i) and then find it in wks1
    
    Dim ps As Integer, pf As Integer, middlename As String, BP As Range
    
    For j = 1 To Len(suppliername(i))
        If Mid(suppliername(i), j, 1) Like "[Ò»-ý›]" Then
                    ps = j
                    Exit For
        End If
            
    Next
    
    pf = ps
    For j = ps To Len(suppliername(i))
        If Mid(suppliername(i), j, 1) Like "[Ò»-ý›]" Then
            pf = pf + 1
        Else
            Exit For
        End If
    Next
     
    
    middlename = Mid(suppliername(i), ps, pf - ps)
    
    Set BP = TWS2.UsedRange.Find(middlename)
    If BP Is Nothing Then
    Else
        TWS1.Range("A" & orderrowstart(i)).Value = BP.Value                             'find and set suppliername and show in order detail
        TWS1.Range("A" & orderrowstart(i)).Font.Size = 20
        TWS1.Range("A" & orderrowstart(i)).WrapText = False
        
    End If

        
            
    Next
    wkb.Close
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



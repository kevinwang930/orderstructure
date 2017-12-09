Attribute VB_Name = "checkdata"
Sub generatedata()
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
     
     
     

     
     TXTS = "YW0817"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    Set tf = TWS1.Range("A1")
    For i = 0 To 70
     
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
                    
                    TWS6.Range("A" & i + 1).Value = orderno(i)              'output checkdata value
                    TWS6.Range("b" & i + 1).Value = orderqty(i)
                    TWS6.Range("c" & i + 1).Value = orderamount(i)
                    TWS6.Range("d" & i + 1).Value = orderctn(i)
                    TWS6.Range("e" & i + 1).Value = ordergweight(i)
                    TWS6.Range("f" & i + 1).Value = ordernweight(i)
                    
                    TWS6.Range("b" & i + 1).Interior.ColorIndex = 0
                    TWS6.Range("c" & i + 1).Interior.ColorIndex = 0
                    TWS6.Range("d" & i + 1).Interior.ColorIndex = 0
                    TWS6.Range("e" & i + 1).Interior.ColorIndex = 0
                    TWS6.Range("f" & i + 1).Interior.ColorIndex = 0
                    
                    If orderqty(i) <> TWS1.Range("H" & orderrowfinish(i)) Then
                       TWS6.Range("b" & i + 1).Interior.ColorIndex = 37
                    
                    ElseIf orderamount(i) <> TWS1.Range("C" & orderrowfinish(i)) Then
  
                       TWS6.Range("c" & i + 1).Interior.ColorIndex = 37
                    
                    ElseIf orderctn(i) <> TWS1.Range("k" & orderrowfinish(i)) Then

                       TWS6.Range("d" & i + 1).Interior.ColorIndex = 37
                  
                    ElseIf ordergweight(i) <> TWS1.Range("s" & orderrowfinish(i)) Then
                       TWS6.Range("e" & i + 1).Interior.ColorIndex = 37
                       
                    ElseIf ordernweight(i) <> TWS1.Range("U" & orderrowfinish(i)) Then
                       TWS6.Range("f" & i + 1).Interior.ColorIndex = 37
                    End If
                    
                       
                End If
            End If
            
            For j = modelstart(i) To modelfinish(i)
                If TWS1.Range("B" & j).Value = "" Then
                    TWS1.Range("B" & j).Value = TWS1.Range("A" & j).Value
                End If
            Next
    Next
End Sub
Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function


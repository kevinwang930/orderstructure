Attribute VB_Name = "shippingspecialnamechange"
Sub shippingsepcialnamechange()

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
  
    
    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    
     Dim i As Integer
     Dim TXTS As String, TXTF As String
     Dim ts As Range, tf As Range
     
     
     
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
     Dim orderrequirement(70) As String
     Dim shippingnamechange(70) As String
     
     
     

     
     TXTS = "YW1117"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    Set tf = TWS1.Range("A1")
    Set wkb1 = Workbooks.Add
    TWS1.Copy wkb1.Worksheets("sheet1")
    Set wks1 = wkb1.Worksheets("order detail")
    
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
                    orderrequirement(i) = TWS1.Range("C" & ts.Row + 1).Value
                    orderqty(i) = Application.Sum(TWS1.Range("H" & orderrowstart(i) & ":H" & orderrowfinish(i) - 1))
                    orderamount(i) = Round(Application.Sum(TWS1.Range("J" & orderrowstart(i) & ":J" & orderrowfinish(i) - 1)), 2)
                    orderctn(i) = Application.Sum(TWS1.Range("G" & orderrowstart(i) & ":G" & orderrowfinish(i) - 1))
                    ordergweight(i) = Application.Sum(TWS1.Range("p" & orderrowstart(i) & ":p" & orderrowfinish(i) - 1))
                    ordernweight(i) = Application.Sum(TWS1.Range("q" & orderrowstart(i) & ":q" & orderrowfinish(i) - 1))

                    shippingnamechange(i) = ""
                    
                    For j = modelstart(i) To modelfinish(i)
                    waterbottle = ""
                    sunglass = ""
                    lunchbox = ""
                    TWS1.Range("u" & j).Value = TWS1.Range("u" & j).Value
                    If TWS1.Range("E" & j).Value = "Ë®±­" Then
                       TWS1.Range("E" & j).Value = "ÀñÆ·ºÐ"
                       TWS1.Range("C" & j).Value = "gift box"
                       waterbottle = "water bottle description gift box"
                    ElseIf TWS1.Range("E" & j).Value = "Ì«Ñô¾µ" Then
                       TWS1.Range("E" & j).Value = "ÀñÆ·"
                       TWS1.Range("C" & j).Value = "gift set"
                       sunglass = "sunglass description gift"
                    ElseIf TWS1.Range("E" & j).Value = "²ÍºÐ" Then
                       TWS1.Range("E" & j).Value = "ÀñÆ·ºÐ"
                       TWS1.Range("C" & j).Value = "gift box"
                       lunchbox = "lunchbox description gift box"
                    End If
                    
                    Next
                
                End If
            End If
 'orderrequirement(i) = orderrequirement(i) & waterbottle & sunglass & lunchbox
 'wks1.Range("C" & ts.Row + 1).Value = orderrequirement(i)
 'TWS1.Range("C" & ts.Row + 1).Value = orderrequirement(i)

Next
End Sub

  Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function

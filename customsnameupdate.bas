Attribute VB_Name = "customsnameupdate"
Public TWS1 As Worksheet, TWS2 As Worksheet, TWS3 As Worksheet, TWS4 As Worksheet, TWS5 As Worksheet, TWS6 As Worksheet


Sub dataupdate()


Dim wkb As Workbook, wks As Worksheet
     
     Dim i As Integer, j As Integer, k As Integer                    'i IS order amount
     
     Dim TXTS As String, TXTF As String
     Dim ts As Range, tf As Range, tm As Range, ms As Range, mf As Range
     Application.ScreenUpdating = False
     
     
   Set TWS1 = ThisWorkbook.Worksheets("order detail")                 'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")                   'target worksheet
   'Set TWS3 = ThisWorkbook.Worksheets("bank detail collect report")   'target worksheet
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
     Dim orderstatus(70) As String
     
     
     
    TWS6.Cells.ClearContents
    TWS6.Cells.ClearFormats
    TWS6.Cells.ClearOutline
    
    
    
     
     TXTS = "YW1117"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    Set tf = TWS1.Range("A1")
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
                    suppliername(i) = TWS1.Range("A" & orderrowstart(i)).Value
                    orderrowfinish(i) = tf.Row
                    orderno(i) = ts.Value
                    orderstatus(i) = TWS1.Range("C" & ts.Row + 1).Value


                    
                End If
            End If
            
 
            
                If InStr(orderstatus(i), "water bottle") <> 0 Then
                    For j = modelstart(i) To modelfinish(i)
                        If TWS1.Range("c" & j).Value = "gift box" Then
                            TWS1.Range("c" & j).Value = "water bottle"
                            TWS1.Range("d" & j).Value = "gift box"
                            TWS1.Range("f" & j).Value = "水杯"
                            TWS1.Range("g" & j).Value = "礼品盒"
                        End If
                    Next
                
                
                ElseIf InStr(orderstatus(i), "lunchbox") <> 0 Then
                    For j = modelstart(i) To modelfinish(i)
                        If TWS1.Range("c" & j).Value = "gift box" Then
                            TWS1.Range("c" & j).Value = "lunch box"
                            TWS1.Range("d" & j).Value = "gift box"
                            TWS1.Range("f" & j).Value = "餐盒"
                            TWS1.Range("g" & j).Value = "礼品盒"
                        End If
                    Next
                 ElseIf InStr(orderstatus(i), "sunglss") <> 0 Then
                    For j = modelstart(i) To modelfinish(i)
                        If TWS1.Range("c" & j).Value = "gift set" Then
                            TWS1.Range("c" & j).Value = "sunglass"
                            TWS1.Range("d" & j).Value = "gift set"
                            TWS1.Range("f" & j).Value = "太阳镜"
                            TWS1.Range("g" & j).Value = "礼品"
                        End If
                    Next
                Else
                     For j = modelstart(i) To modelfinish(i)
                        If InStr(TWS1.Range("c" & j).Value, "CRYSTAL") <> 0 Then
                            TWS1.Range("c" & j).Value = "HANDI CRAFT"
                            TWS1.Range("d" & j).Value = "HANDI CRAFT"
                            TWS1.Range("f" & j).Value = "工艺品"
                            TWS1.Range("g" & j).Value = "工艺品"
                        Else
                            TWS1.Range("d" & j).Value = TWS1.Range("c" & j).Value
                            TWS1.Range("g" & j).Value = TWS1.Range("f" & j).Value
                        End If
                    Next
                End If
           
                
    Next
End Sub
Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function


Attribute VB_Name = "dataupdate"

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
                    orderqty(i) = Application.Sum(TWS1.Range("H" & orderrowstart(i) & ":H" & orderrowfinish(i) - 1))
                    orderamount(i) = Round(Application.Sum(TWS1.Range("J" & orderrowstart(i) & ":J" & modelfinish(i))), 2)
                    orderctn(i) = Application.Sum(TWS1.Range("G" & orderrowstart(i) & ":G" & orderrowfinish(i) - 1))
                    ordergweight(i) = Application.Sum(TWS1.Range("p" & orderrowstart(i) & ":p" & orderrowfinish(i) - 1))
                    ordernweight(i) = Application.Sum(TWS1.Range("q" & orderrowstart(i) & ":q" & orderrowfinish(i) - 1))
                    ordercbm(i) = Application.Sum(TWS1.Range("N" & modelstart(i) & ":N" & modelfinish(i)))
                    orderstatus(i) = TWS1.Range("C" & ms.Row).Value


                    TWS1.Range("A" & modelstart(i) & ":v" & modelfinish(i)).Font.Name = "Times New Roman"                'set order records font and size
                    TWS1.Range("A" & modelstart(i) & ":v" & modelfinish(i)).Font.Size = 14
                    
                    TWS1.Range("H" & orderrowfinish(i)).Value = "=sum(H" & modelstart(i) & ":H" & modelfinish(i) & ")"        'sum quantity
                    TWS1.Range("H" & orderrowfinish(i)).NumberFormat = "0"
                    TWS1.Range("C" & orderrowfinish(i)).Value = "=sum(J" & modelstart(i) & ":J" & modelfinish(i) & ")"        'sum money amount
                    TWS1.Range("C" & orderrowfinish(i)).NumberFormat = "�� #,##0.00"
                    TWS1.Range("E" & orderrowfinish(i)).NumberFormat = "�� #,##0.00"
                    TWS1.Range("k" & orderrowfinish(i)).Value = "=sum(G" & modelstart(i) & ":G" & modelfinish(i) & ")"        'sum cartoon amount
                    TWS1.Range("k" & orderrowfinish(i)).NumberFormat = "0 CT\N"
                    
                End If
            End If
            
            
            
                TWS1.Range("U" & modelstart(i)).ClearFormats                                        'set first carton container No
                TWS1.Range("U" & modelstart(i)).NumberFormat = "0"
                TWS1.Range("U" & modelstart(i)).Font.Size = 16
                TWS1.Range("U" & modelstart(i)).Font.Name = "Times New Roman"
                TWS1.Range("U" & modelstart(i)).HorizontalAlignment = xlCenter
                TWS1.Range("U" & modelstart(i)).VerticalAlignment = xlCenter
                TWS1.Range("U" & modelstart(i)).Borders.LineStyle = xlContinuous
            If TWS1.Range("G" & modelstart(i)).Value = 1 Then
                TWS1.Range("U" & modelstart(i)).Value = 1
            ElseIf TWS1.Range("G" & modelstart(i)).Value > 1 Then
                TWS1.Range("U" & modelstart(i)).Value = "=" & """1~""" & "&" & "G" & modelstart(i)
            Else: TWS1.Range("U" & modelstart(i)).Value = ""
            End If
            
                
            For j = modelstart(i) + 1 To modelfinish(i)
                TWS1.Range("U" & j).ClearFormats
                TWS1.Range("U" & j).NumberFormat = "0"
                TWS1.Range("U" & j).Font.Size = 16
                TWS1.Range("U" & j).Font.Name = "Times New Roman"
                TWS1.Range("U" & j).HorizontalAlignment = xlCenter
                TWS1.Range("U" & j).VerticalAlignment = xlCenter
                TWS1.Range("U" & j).Borders.LineStyle = xlContinuous
                
                
    
                If TWS1.Range("G" & j) = 1 Then
                TWS1.Range("U" & j) = "=SUM(G" & modelstart(i) & ":" & "G" & j & ")"
                
                ElseIf TWS1.Range("G" & j) > 1 Then
                TWS1.Range("U" & j).Value = "=SUM(G" & modelstart(i) & ":G" & j - 1 & ",1)" & "&""~""&" & "SUM(G" & modelstart(i) & ":G" & j & ")"
                '&\""-\""" & "sum(g" & modelstart(I) & ":g" & J & ")"
                ElseIf TWS1.Range("G" & j).Value = 0 And TWS1.Range("G" & j).MergeCells Then
                TWS1.Range("U" & j).Value = TWS1.Range("u" & j - 1)
                End If
                
                If TWS1.Range("s" & j).Value = "" And TWS1.Range("s" & j - 1).Value <> "" Then                                      'update material
                    TWS1.Range("s" & j).Value = TWS1.Range("S" & j - 1).Value
                    
                End If
                
            Next
            
            For j = modelstart(i) To modelfinish(i)
                If TWS1.Range("k" & j) = "" And TWS1.Range("L" & j) = "" And TWS1.Range("M" & j) = "" Then
                ElseIf TWS1.Range("N" & j) = 0 And TWS1.Range("L" & j) = "" And TWS1.Range("M" & j) = "" Then
                TWS1.Range("N" & j).Value = "=K" & j & "*G" & j                                                                     'update every model volume formula
                End If
                
                If TWS1.Range("G" & j).Value = 0 And TWS1.Range("G" & j).MergeCells Then                                            'update quantity formula
                    TWS1.Range("h" & j).Value = TWS1.Range("G" & j).MergeArea.Cells(1, 1).Value * TWS1.Range("f" & j).Value
                ElseIf TWS1.Range("G" & j).Value = 0 And TWS1.Range("v" & j).Value > 0 Then
                    TWS1.Range("h" & j).Value = TWS1.Range("F" & j).Value                                                           'update checkdate
                    TWS6.Range("a" & TWS6.UsedRange.Rows.count + 1).Value = orderno(i)
                    TWS6.Range("B" & TWS6.UsedRange.Rows.count + 1).Value = "Single term ctn 0 pack with other"
                    TWS6.Range("c" & TWS6.UsedRange.Rows.count + 1).Value = TWS1.Range("a" & j).Value
                End If
                
                
                TWS1.Range("H" & j).Interior.ColorIndex = 0
                If TWS1.Range("H" & j).Value = "" Then
                    TWS1.Range("H" & j).Value = "=F" & j & "*" & "G" & j                                                            'update every model quantity formula and check the result
                ElseIf TWS1.Range("H" & j).Value <> TWS1.Range("F" & j).Value * TWS1.Range("G" & j).Value Then
                    TWS1.Range("H" & j).Interior.ColorIndex = 35
                End If
                
                
                If TWS1.Range("O" & j).Value = 0 And TWS1.Range("p" & j).Value = 0 Then                                           'update net weight formula
                    TWS1.Range("Q" & j).NumberFormat = "0"
                    TWS1.Range("Q" & j).Value = 0
                ElseIf TWS1.Range("O" & j).Value <> "" Then
                    TWS1.Range("Q" & j).NumberFormat = "0"
                    TWS1.Range("Q" & j).Value = "=(O" & j & "-1)*" & "G" & j
                ElseIf TWS1.Range("O" & j).Value = 0 And TWS1.Range("p" & j).Value <> 0 Then
                    TWS1.Range("Q" & j).NumberFormat = "0"
                    TWS1.Range("Q" & j).Value = "=p" & j & "-" & "G" & j
                End If
                
                
                
                
               'add invoice model No automatically
                If TWS1.Range("B" & j).Value = "" Then
                    TWS1.Range("B" & j).NumberFormat = TWS1.Range("A" & j).NumberFormat
                    TWS1.Range("B" & j) = TWS1.Range("A" & j)
                End If
               
               
                
            Next
            
            
                If InStr(orderstatus(i), "water bottle") <> 0 Then
                    For j = modelstart(i) To modelfinish(i)
                        If TWS1.Range("c" & j).Value = "gift box" Then
                            TWS1.Range("c" & j).Value = "water bottle"
                            TWS1.Range("d" & j).Value = "gift box"
                            TWS1.Range("f" & j).Value = "ˮ��"
                            TWS1.Range("g" & j).Value = "��Ʒ��"
                        End If
                    Next
                End If
                
                If InStr(orderstatus(i), "lunch box") <> 0 Then
                    For j = modelstart(i) To modelfinish(i)
                        If TWS1.Range("c" & j).Value = "gift box" Then
                            TWS1.Range("c" & j).Value = "lunch box"
                            TWS1.Range("d" & j).Value = "gift box"
                            TWS1.Range("f" & j).Value = "�ͺ�"
                            TWS1.Range("g" & j).Value = "��Ʒ��"
                        End If
                    Next
                End If
            
                            
                        
            
                
                
           If TWS1.Range("s" & orderrowfinish(i)) > 0 And TWS1.Range("U" & orderrowfinish(i)).Value <= 0 And TWS1.Range("k" & orderrowfinish(i)).Value > 0 Then
            TWS1.Range("U" & orderrowfinish(i)).Value = "=s" & orderrowfinish(i) & "-k" & orderrowfinish(i)
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




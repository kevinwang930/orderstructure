Attribute VB_Name = "datainitialization"
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
     Dim ordercbm(70) As Double
     Dim suppliername(70) As String
     
     Dim test As String
     Dim totalpurchaseamount As Double, totalpurchasectn As Long
     totalpurchaseamount = 0
     totalpurchasectn = 0
     

     
     TXTS = "YW1117"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    
    Set tf = TWS1.Range("A1")
    Dim bankheader As Integer
    bankheader = 7
    TWS2.Cells.Font.Size = 16
    TWS2.Cells.Font.Name = "Calibri"
    TWS2.Rows(5).Font.Size = 22
    For i = 1 To 70
     
            Set ts = finddown(TWS1.UsedRange, TXTS, ts)                             'find order start
            If ts Is Nothing Then
            
                TWS2.Range("F" & i + bankheader).Value = "Purchase total"
                Call autosum1(TWS2.Range("G" & i + bankheader))
                Call autosum1(TWS2.Range("h" & i + bankheader))
                Call autosum1(TWS2.Range("i" & i + bankheader))
                Call autosum1(TWS2.Range("J" & i + bankheader))
                Call autosum1(TWS2.Range("M" & i + bankheader))
                Call autosum1(TWS2.Range("N" & i + bankheader))
                Call autosum1(TWS2.Range("O" & i + bankheader))
                TWS2.Range("G" & i + bankheader).NumberFormat = "гд #,##0.00"
                TWS2.Range("G" & i + bankheader).Font.Size = 22
                TWS2.Range("h" & i + bankheader).NumberFormat = "гд #,##0.00"
                TWS2.Range("h" & i + bankheader).Font.Size = 22
                TWS2.Range("I" & i + bankheader).NumberFormat = "гд #,##0.00"
                TWS2.Range("I" & i + bankheader).Font.Size = 20
                TWS2.Range("J" & i + bankheader).NumberFormat = "0"
                TWS2.Range("M" & i + bankheader).NumberFormat = "0"
                TWS2.Range("N" & i + bankheader).NumberFormat = "0.00 C\B\M"
                TWS2.Range("O" & i + bankheader).NumberFormat = "0.00 K\G"
                
                Exit For
            Else
                If i < 10 Then
                    ts.Value = "YW1117-ST0" & i                                     'set supplier code
                Else: ts.Value = "YW1117-ST" & i
                End If
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
                   
                    
                    
                    
                    For j = modelstart(i) To modelfinish(i)             'set recap cell number format
                        TWS1.Range("I" & j).NumberFormat = "гд #,##0.00"
                        TWS1.Range("j" & j).NumberFormat = "гд #,##0.00"
                        TWS1.Range("G" & j).NumberFormat = "0 ct\n"               'set every ctn format
                        TWS1.Range("F" & j).NumberFormat = "0"
                        TWS1.Range("H" & j).NumberFormat = "0"
                        TWS1.Range("J" & j).Value = "=H" & j & "*" & "I" & j       'set every model amount formula
                        TWS1.Range("N" & j).Value = "=K" & j & "*" & "L" & j & "*M" & j & "*G" & j & "*0.000001"     'set every model volume formula
                        TWS1.Range("N" & j).NumberFormat = "0.000"
                        TWS1.Range("P" & j).Value = "=O" & j & "*" & "G" & j       'set every model gross weight formula
                        TWS1.Range("Q" & j).Value = "=(O" & j & "-1)*" & "G" & j    'set every model total net weight formula
                        
                         If TWS1.Range("H" & j).Value = "" Then
                            TWS1.Range("H" & j).Value = "=F" & j & "*" & "G" & j      'set every model quantity formula
                         End If
                    Next
                    
                    TWS1.Range("H" & orderrowfinish(i)).Value = "=sum(H" & modelstart(i) & ":H" & modelfinish(i) & ")"        'sum quantity
                    TWS1.Range("H" & orderrowfinish(i)).NumberFormat = "0"
                    TWS1.Range("C" & orderrowfinish(i)).Value = "=sum(J" & modelstart(i) & ":J" & modelfinish(i) & ")"        'sum money amount
                    TWS1.Range("C" & orderrowfinish(i)).NumberFormat = "гд #,##0.00"
                    TWS1.Range("E" & orderrowfinish(i)).NumberFormat = "гд #,##0.00"
                    TWS1.Range("k" & orderrowfinish(i)).Value = "=sum(G" & modelstart(i) & ":G" & modelfinish(i) & ")"        'sum cartoon amount
                    TWS1.Range("k" & orderrowfinish(i)).NumberFormat = "0 CT\N"
                    TWS1.Range("s" & orderrowfinish(i)).Value = "=sum(P" & modelstart(i) & ":P" & modelfinish(i) & ")"        'sum gross weight
                    TWS1.Range("s" & orderrowfinish(i)).NumberFormat = "0.0 k\g"
                    TWS1.Range("U" & orderrowfinish(i)).Value = "=sum(Q" & modelstart(i) & ":Q" & modelfinish(i) & ")"        'sum net weight
                    TWS1.Range("U" & orderrowfinish(i)).NumberFormat = "0.0 k\g"
                    
                    TWS1.Range("I" & orderrowfinish(i)).Value = "=sum(N" & modelstart(i) & ":N" & modelfinish(i) & ")"        'sum CBM
                    TWS1.Range("I" & orderrowfinish(i)).NumberFormat = "0.0 C\B\M"
                    
                    TWS2.Range("A" & i + bankheader).Value = orderno(i)                                                              'bank detail information generate
                    TWS2.Range("A" & i + bankheader).Font.Size = 22
                    TWS2.Range("B" & i + bankheader).Value = suppliername(i)
                    TWS2.Range("E" & i + bankheader).Font.Size = 20
                    TWS2.Range("F" & i + bankheader).Font.Size = 20
                    TWS2.Range("G" & i + bankheader) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("C" & orderrowfinish(i)).Address   'amount
                    TWS2.Range("G" & i + bankheader).NumberFormat = "гд #,##0.00"
                    TWS2.Range("G" & i + bankheader).Font.Size = 22
                    TWS2.Range("H" & i + bankheader) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("E" & orderrowfinish(i)).Address   'deposit
                    TWS2.Range("H" & i + bankheader).NumberFormat = "гд #,##0.00"
                    TWS2.Range("H" & i + bankheader).Font.Size = 22
                    
                    TWS2.Range("I" & i + 7) = "=G" & i + 7 & "-H" & i + 7                                                       'BALANCE
                    TWS2.Range("I" & i + 7).NumberFormat = "гд #,##0.00"
                    TWS2.Range("I" & i + bankheader).Font.Size = 20
                    
                    
                    TWS2.Range("J" & i + 7) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("H" & orderrowfinish(i)).Address   'qty
                    TWS2.Range("J" & i + 7).NumberFormat = "0"
                    TWS2.Range("m" & i + 7) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("K" & orderrowfinish(i)).Address   'carton
                    TWS2.Range("M" & i + 7).NumberFormat = "0 ct\n"
                    TWS2.Range("N" & i + 7) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("I" & orderrowfinish(i)).Address   'CBM
                    TWS2.Range("N" & i + 7).NumberFormat = "0.00 c\b\m"
                    TWS2.Range("O" & i + 7) = "=" & "'" & TWS1.Name & "'" & "!" & TWS1.Range("S" & orderrowfinish(i)).Address   'GROSS WEIGHT
                    TWS2.Range("O" & i + 7).NumberFormat = "0.0 k\g"
                   ' MsgBox (TWS1.Range("c" & orderrowfinish(I)).NumberFormat & "blank" & TWS1.Range("c" & orderrowfinish(I)).NumberFormatLocal)
                End If
            End If
            
          
            
            For j = modelstart(i) + 1 To modelfinish(i)                    'set every model some value
            
            
            If TWS1.Range("C" & j).Value = "" Then
                If TWS1.Range("C" & j - 1).Value <> "" Then
                    TWS1.Range("C" & j).Value = TWS1.Range("C" & j - 1).Value
                End If
            End If
            
            If TWS1.Range("E" & j).Value = "" Then
                If TWS1.Range("E" & j - 1).Value <> "" Then
                    TWS1.Range("E" & j).Value = TWS1.Range("E" & j - 1).Value
                End If
            End If
            
            If TWS1.Range("F" & j).Value = "" Then
                If TWS1.Range("F" & j - 1).Value <> "" Then
                    TWS1.Range("F" & j).Value = TWS1.Range("F" & j - 1).Value
                End If
            End If
            
            If TWS1.Range("G" & j).Value = "" Then
                If TWS1.Range("G" & j - 1).Value <> "" Then
                    TWS1.Range("G" & j).Value = TWS1.Range("G" & j - 1).Value
                End If
            End If
            
            If TWS1.Range("I" & j).Value = "" Then
                If TWS1.Range("I" & j - 1).Value <> "" Then
                    TWS1.Range("I" & j).Value = TWS1.Range("I" & j - 1).Value
                End If
            End If
            
            Next
    totalpurchaseamount = totalpurchaseamount + orderamount(i)
    totalpurchasectn = totalpurchasectn + orderctn(i)
    Next
    TWS1.Range("J679").NumberFormat = "гд #,##0.00"
    TWS1.Range("J680").NumberFormat = "0 ct\n"
    TWS1.Range("J679:J680").Font.Size = 18
    TWS1.Range("J679:J680").Font.Bold = True
    
    
    TWS1.Range("J679").Value = totalpurchaseamount
    TWS1.Range("J680").Value = totalpurchasectn
    
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


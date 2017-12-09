Attribute VB_Name = "clearcontent"
Sub clearcontent()
Attribute clearcontent.VB_ProcData.VB_Invoke_Func = " \n14"
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
                    orderrowfinish(i) = tf.Row
                End If
                
            End If
            
            For j = modelstart(i) To modelfinish(i)                         'clear model content
                If IsNull(TWS1.Rows(j).MergeCells) Then                     'jump merge cells
                    'TWS1.Rows(j).MergeArea.ClearContents
                Else: TWS1.Rows(j).ClearContents
                TWS1.Range("A" & j & ":V" & j).Interior.Color = 16777215
                End If
            
            Next
        TWS1.Rows(orderrowstart(i)).ClearContents
        If TWS1.Range("A" & ts.Row + 1).Value = TWS1.Range("A10").Value Then
            TWS1.Range("A" & ts.Row + 1).EntireRow.Insert
            TWS1.Range("A" & ts.Row + 1).Value = "order requirement:"
        Else: TWS1.Range("A" & ts.Row + 1).Value = "order requirement:"
        End If
        TWS1.Range("A" & ts.Row - 1 & ":V" & ts.Row + 2).Interior.Color = 16777215
        
        
        Next
            
            
                    
End Sub


Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function

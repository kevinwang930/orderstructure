Attribute VB_Name = "depositreport"
''wait to show infomration in order details


Sub reportdeposit()
'creat 2 separate files, order details and bank details.
'create folder to put the 2 file.
'remove the formulas of the container No. remove the reference outside of worksheet.

Application.ScreenUpdating = False
Application.DisplayAlerts = False

    Dim wkb As Workbook, wks1 As Worksheet, wks2 As Worksheet
    Dim TWS1 As Worksheet, TWS2 As Worksheet, TWS3 As Worksheet
    Dim filename As String, ANAME As String, directory As String, projectname As String, Planinformation As String
    
    Dim fso As Object, otname As Object
    Set fso = CreateObject("scripting.filesystemobject")
    
    projectname = "ST1117"
    directory = fso.GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder.path
    directory = directory & "\Market order\" & projectname & "\YW\recap"         'set file location
    
    
    
    
    
    Dim d As String
    d = Format(Date, "yyyy-mm-dd")
    filename = fso.GETBASENAME(ThisWorkbook.FullName)                           'set name
    Planinformation = "Bank detail for TT deposit"
    Call createfolder(directory)
    directory = directory & "\" & Planinformation & " " & d     'set new folder name
    
    Call createfolder(directory)
    
    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    
    'Call getsequence(TWS1, TWS2, "YW1117", "Container")
    
    'create order detail report file
    'Call createorderdetail(directory, filename, Planinformation, d)
    Call createbankdetail(directory, filename, Planinformation, d)
    
    
    
End Sub


Sub createfolder(sname As String)

    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
 'check if target folder exist
    If fso.FolderExists(sname) Then
    
    Else
            fso.createfolder (sname)
            
    End If
End Sub
Sub createorderdetail(directory As String, filename As String, Planinformation As String, d As String)

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
                    orderqty(i) = Application.Sum(TWS1.Range("H" & orderrowstart(i) & ":H" & orderrowfinish(i) - 1))
                    orderamount(i) = Round(Application.Sum(TWS1.Range("J" & orderrowstart(i) & ":J" & orderrowfinish(i) - 1)), 2)
                    orderctn(i) = Application.Sum(TWS1.Range("G" & orderrowstart(i) & ":G" & orderrowfinish(i) - 1))
                    ordergweight(i) = Application.Sum(TWS1.Range("p" & orderrowstart(i) & ":p" & orderrowfinish(i) - 1))
                    ordernweight(i) = Application.Sum(TWS1.Range("q" & orderrowstart(i) & ":q" & orderrowfinish(i) - 1))
    

                    For j = modelstart(i) To modelfinish(i)
                    wks1.Range("u" & j).Value = wks1.Range("u" & j).Value
                    Next
                End If
            End If
        
Next

    
    wks1.Range("P3").Value = d
    wks1.Range("D3").Value = Planinformation
    wkb1.SaveAs directory & "\" & filename & " " & TWS1.Name & " " & Planinformation & " " & d & ".xls", FileFormat:=56
    wkb1.Close False
End Sub

Sub createbankdetail(directory As String, filename As String, Planinformation As String, d As String)
    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    

    
    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    
    
    

    Set wkb1 = Workbooks.Add
    TWS2.Copy wkb1.Worksheets("sheet1")
    Set wks1 = wkb1.Worksheets("bank detail")
    wks1.Columns("A:H").Value = TWS2.Columns("A:H").Value
    wks1.Columns("J:J").Value = TWS2.Columns("J:J").Value
    wks1.Columns("L:N").Value = TWS2.Columns("L:N").Value
    
    
    Dim header As Integer, i As Integer
    header = 7
    For i = header + 1 To wks1.UsedRange.Rows.count
    
    If InStr(wks1.Range("A" & i).Value, "YW1117") = 0 Then          'delete rows not bank information or total ethier, delete 0 deposit
        If InStr(wks1.Range("F" & i).Value, "total") = 0 Then
            wks1.Rows(i).Delete
            i = i - 1
        Else
            wks1.Range("G" & i).Value = "=SUM(G" & header + 1 & ":G" & i - 1 & ")"
            wks1.Range("H" & i).Value = "=SUM(H" & header + 1 & ":H" & i - 1 & ")"
            wks1.Range("I" & i).Value = "=SUM(I" & header + 1 & ":I" & i - 1 & ")"
            wks1.Range("J" & i).Value = "=SUM(J" & header + 1 & ":J" & i - 1 & ")"
            wks1.Range("M" & i).Value = "=SUM(M" & header + 1 & ":M" & i - 1 & ")"
            wks1.Range("N" & i).Value = "=SUM(N" & header + 1 & ":N" & i - 1 & ")"
            wks1.Range("O" & i).Value = "=SUM(O" & header + 1 & ":O" & i - 1 & ")"
        End If
        ElseIf wks1.Range("H" & i).Value = 0 Then
            wks1.Rows(i).Delete
            i = i - 1
    End If
    If i >= wks1.UsedRange.Rows.count Then
        Exit For
    End If
    
    Next
    
    
    
    wks1.Range("O3").Value = d
    wks1.Range("D3").Value = Planinformation
    
    wkb1.SaveAs directory & "\" & filename & " " & Planinformation & " " & d & ".xls", FileFormat:=56
    wkb1.Close False
End Sub


Sub getsequence(TWS1 As Worksheet, TWS2 As Worksheet, suppliercode As String, container As String)
'from A1,if cell start with YW0817 COMPARE WITH ORDER DETAIL SHEET

    Dim i As Integer
    Dim add1 As Range, add2 As Range, adde As Range, adds2 As Range, adds1 As Range, addc1 As Range, addc2 As Range
    
    
    
    

   
    
    Set add1 = TWS1.Range("A1")
    Set adds1 = TWS1.Range("A1")
    Set adds2 = TWS2.Range("A1")
    Set add2 = TWS2.Range("A1")
    For i = 1 To TWS2.UsedRange.Rows.count                        ' loop for supplier code sequence
    
    
    
        Set adds2 = finddown(TWS2.Range("A:A"), suppliercode, adds2)
        Set adds1 = finddown(TWS1.Range("A:A"), suppliercode, adds1)
        If adds2 Is Nothing Then
            Exit For
        ElseIf adds1 Is Nothing Then
            Exit For
        
        ElseIf adds1.Value <> adds2.Value Then                     'find different order sequence, then adjust by cut and insert one by one
            Call insertorder(adds1, adds2.Value, "Total")
            Set adds1 = TWS1.Range("A:A").Find(adds2.Value)        ' to make TWS1 has same order sequence with TWS2.
        End If
    
    Next
    
    Set addc1 = TWS1.Range("A1")
    Set addc2 = TWS2.Range("A1")
    
     For i = 1 To TWS2.UsedRange.Rows.count                        ' loop for container sequence
    
    
    
        Set addc2 = finddown(TWS2.Range("A:A"), container, addc2)
        If addc2 Is Nothing Then                                    'if container does not exist in bank detail
            Exit For
        Else
            Set adds2 = finddown(TWS2.Range("A:A"), suppliercode, addc2)
            If adds2 Is Nothing Then                                'check if container is the last row
                Exit For
            Else
                 Set addc1 = finddown(TWS1.Range("A:A"), container, addc1)
                  If addc1 Is Nothing Then                         'if container does not exist in order detail then insert addc1 to right position
                    Set adds1 = TWS1.Range("A:A").Find(adds2.Value)
                    TWS2.Rows(addc2.Row).Copy
                    TWS1.Rows(adds1.Row - 1).Insert
                  Else                                             'check container position is same in TWS1 and TWS2.
                    Set adds1 = finddown(TWS1.Range("A:A"), suppliercode, addc1)
                    If adds1 <> adds2 Then                         'if container not same in order detail, cut container insert to right position
                        Set adds1 = TWS1.Range("A:A").Find(adds2.Value)
                        TWS1.Rows(addc1.Row).Cut
                        TWS1.Rows(adds1.Row - 1).Insert
                    End If
                  End If
            End If
        End If
        
    Next
   
End Sub

Sub insertorder(ip As Range, starting As String, finishing As String)

    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    
    Set adds = TWS1.Range("A:A").Find(starting, ip)
    Set addf = TWS1.Range("A:A").Find(finishing, adds)  'find the order range waiting for cut and insert
    Dim sr As Integer, fr As Integer, k As Integer
    
    sr = adds.Row
    fr = addf.Row
    
    k = ip.Row
    
    TWS1.Rows(sr - 1 & ":" & fr).Cut
    
    TWS1.Rows(k - 1).Insert shift:=xlDown
    

    
    
End Sub

Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function



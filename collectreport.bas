Attribute VB_Name = "collectreport"
    Public TWS1 As Worksheet, TWS2 As Worksheet, TWS3 As Worksheet, TWS4 As Worksheet, TWS5 As Worksheet
 
   
   
Sub collectreport()                     'report the collect situation
'creat 2 separate files, order details and bank details.
'create folder to put the 2 file.
'remove the formulas of the container No. remove the reference outside of worksheet.

Application.ScreenUpdating = False
Application.DisplayAlerts = False

    
    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("bank detail collect report")   'target worksheet
    Set TWS4 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    Set TWS5 = ThisWorkbook.Worksheets("collect information")   'target worksheet
    
    Dim tname As String, ANAME As String
    Dim d As String
    d = Format(Date, "yyyy-mm-dd")
    
    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
    tname = fso.GETBASENAME(ThisWorkbook.path & "\" & ThisWorkbook.Name) 'get current file name
    ANAME = "collect goods" & " " & d                        'define name information
    
    
                       
    Call createfolder(ThisWorkbook.path & "\RECAPTEST")
    Call createfolder(ThisWorkbook.path & "\RECAPTEST" & "\" & ANAME)
  
    Call getsequence(TWS1, TWS2, "YW0817", "Container")              'check and create order sequence in original date.
    
    
   
    Call createreport(tname, ANAME, d)
   
    
   
End Sub
Sub createreport(tname As String, ANAME As String, d As String)


    Dim ra As Range
    
    
    Set wkb1 = Workbooks.Add                              'copy to new worksheet, avoid damage to original data
    TWS1.Copy wkb1.Worksheets("sheet1")
    Set wks1 = wkb1.Worksheets("order detail")
    TWS2.Copy wkb1.Worksheets("sheet1")
    Set wks2 = wkb1.Worksheets("bank detail")
    TWS3.Copy wkb1.Worksheets("sheet1")
    Set wks3 = wkb1.Worksheets("bank detail collect report")
    
    
    wks3.Columns("A:H").Value = wks3.Columns("A:H").Value  'set wks3 report format
    wks3.Columns("L:N").Value = wks3.Columns("L:N").Value
    wks3.Range("M" & 3).Value = d
    
    
    Set ra = TWS3.Range("A1")
    
    For i = 1 To 10
        Set ra = finddown(TWS3.UsedRange, "total =", ra)
        If ra Is Nothing Then
            Exit For
        Else
            TWS3.Rows(ra.Row).Copy wks3.Rows(ra.Row)
        End If
    Next
    
    wks2.Columns("A:H").Value = wks2.Columns("A:H").Value  'set wks2 report format
    wks2.Columns("L:N").Value = wks2.Columns("L:N").Value
    wks2.Range("M3").Value = d
    
    Set ra = TWS2.Range("A1")
    
    For i = 1 To 10
        Set ra = finddown(TWS2.UsedRange, "total =", ra)
        If ra Is Nothing Then
            Exit For
        Else
            TWS2.Rows(ra.Row).Copy wks2.Rows(ra.Row)
        End If
    Next
    
     wks1.Columns("U").NumberFormat = "@"                   'set wks1 report format
     wks1.Columns("u").Value = wks1.Columns("u").Value
     wks1.Range("p" & 3).Value = d
     wks1.Columns("ac:ad").Delete
     
     Set ra = wks3.UsedRange.Find("left orders")
     If ra Is Nothing Then                                  'check if orders left and create left order report
     Else
        Set wks4 = wkb1.Add
        wks4.Name = "left order"
        Set wks5 = wkb1.Add
        wks5.Name = "left order bank detail"
        wks3.Rows(ra.Row + 1 & ":" & wks5.UsedRange.Rows.count).Cut
        wks5.Rows("a").Insert
        wks3.Rows("1:7").Copy
        wks5.Rows("a").Insert
        
        Set ra1 = wks3.Range(ra.Row + 1, 1)
        Set ra1 = wks1.UsedRange.Find(ra1.Value)
        
        wks1.Rows(ra1.Row - 1 & ":" & wks4.UsedRange.Rows.count).Cut
        wks4.Rows("a").Insert
        wks1.Rows("1:7").Copy
        wks5.Rows("a").Insert
    End If
    
        
   wkb1.SaveAs ThisWorkbook.path & "\RECAPTEST" & "\" & ANAME & "\" & tname & " " & ANAME & ".xls", FileFormat:=56
    wkb1.Close False
    
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
                    If adds1 Is Nothing Then                          'if container not same in order detail, cut container insert to right position
                        Set adds1 = TWS1.Range("A:A").Find(adds2.Value)
                        TWS1.Rows(addc1.Row).Cut
                        TWS1.Rows(adds1.Row - 1).Insert
                    ElseIf adds1 <> adds2 Then                         'if container not same in order detail, cut container insert to right position
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

    Dim wks1 As Worksheet
    
    Set wks1 = ip.Parent
    
    Set adds = wks1.Range("A:A").Find(starting, ip)
    Set addf = wks1.Range("A:A").Find(finishing, adds)
    Dim i As Integer, j As Integer, k As Integer
    
    Ir = adds.Row
    jr = addf.Row
    ipr = ip.Row
    
    k = ip.Row
    
    wks1.Rows(Ir - 1 & ":" & jr).Cut
    
    wks1.Rows(k - 1).Insert shift:=xlDown
 
    
End Sub

Function finddown(rng As Range, target As String, after1 As Range) As Range


Set finddown = rng.Find(target, after1)

If finddown Is Nothing Then
    
ElseIf finddown.Row <= after1.Row Then
    Set finddown = rng.Find("I don't know how to express nothing")
End If

    
End Function


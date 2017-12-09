Attribute VB_Name = "inform_orderstart"
 Sub EXSPLIT()
     Dim wkb As Workbook, wks As Worksheet
     
     Dim i As Integer, k As Integer, j As Integer                     'i IS order amount
     
     Dim TXTS As String, TXTF As String
     Dim ts As Range, tf As Range, tm As Range, BD As Range
     
     Dim filename As String, ANAME As String, directory As String, projectname As String, Planinformation As String
     Application.ScreenUpdating = False
     
     
     Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
     Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
     Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
     
     Dim orderrowstart(70) As Integer
     Dim orderrowfinish(70) As Integer
     Dim orderno(70) As String
     Dim bankdetail(70) As String
     Dim suppliername(70) As String
     Dim deliverydate As String
     
     deliverydate = "12月10日左右"                                                      'set deliverydate
     
     
    Dim fso As Object, otname As Object
    Set fso = CreateObject("scripting.filesystemobject")
     
    projectname = "ST1117"
    directory = fso.GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder.path
    directory = directory & "\Market order\" & projectname & "\YW\packing listtest"         'set file location
     
     Call createfolder(directory)
     
     TXTS = "YW1117"
     TXTF = "Total Amount"
     
    Set ts = TWS1.Range("A1")
    Set tf = TWS1.Range("A1")
    Set BD = TWS2.Range("A1")
    
    For i = 1 To 70
     
            Set ts = finddown(TWS1.UsedRange, TXTS, ts)
            If ts Is Nothing Then
                MsgBox ("do not find order details" & TXTS)
                Exit For
            Else
                Set tf = finddown(TWS1.UsedRange, TXTF, ts)
                Set BD = finddown(TWS2.UsedRange, ts.Value, BD)
                If tf Is Nothing Then
                    
                    MsgBox ("ORDERROWS have start, but do not have finish")
                    Exit For
                ElseIf BD Is Nothing Then
                    MsgBox ("do not find bank details" & ts.Value)
                    Exit For
                
                Else
                    
                    orderrowstart(i) = ts.Row - 1
                    orderrowfinish(i) = tf.Row
                    orderno(i) = ts.Value
                    bankdetail(i) = BD.Row
                    suppliername(i) = TWS1.Range("A" & orderrowstart(i)).Value
                    
                     Dim ps As Integer, pf As Integer, middlename As String, BP As Range
    
    For j = 1 To Len(suppliername(i))
        If Mid(suppliername(i), j, 1) Like "[一-龥]" Then
                    ps = j
                    Exit For
        End If
            
    Next
    
    pf = ps
    For j = ps To Len(suppliername(i))
        If Mid(suppliername(i), j, 1) Like "[一-龥]" Then
            pf = pf + 1
        Else
            Exit For
        End If
    Next
     
    
    middlename = Mid(suppliername(i), ps, pf - ps)                                                       'supplier name
                    
                    
                    
                    Set TCR = TWS1.Range(orderrowstart(i) & ":" & orderrowfinish(i))                    'target copy range
                    Set TCRC = TWS1.Range("U" & orderrowstart(i) & ":" & "W" & orderrowfinish(i) - 1)   'set container no range
                    Set wkb = Workbooks.Add
                    TWS3.Copy wkb.Worksheets("sheet1")
                    Set wks = wkb.Worksheets("shipping mark")
                    wks.Range("C17").Value = orderno(i)                                             'set supplier code
                    wks.Range("H17").Value = wks.Range("C17").Value
                    wks.Range("C17").Font.Size = 20
                    wks.Range("H17").Font.Size = 20
            
            
                    wks.Range("H2").Value = middlename                                              'set supplier name
                    wks.Range("H2").Font.Size = 20
                    wks.Range("H5").Value = deliverydate                                            'set deliverydate
                    wks.Range("H4").Value = TWS2.Range("L" & bankdetail(i)).Value                   'set payment term
                    wks.Range("H4").Font.Size = 20
                    wks.Range("H5").Font.Size = 20
                    TWS2.Rows(bankdetail(i)).Copy                                                   'copy bank detail
                    wks.Range("A7").Insert , shift:=xlDown
                    wks.Rows(7).Value = TWS2.Rows(bankdetail(i)).Value
                    
                    MIDDLE = orderrowfinish(i) - orderrowstart(i)
                    MIDDLE = MIDDLE + 8
                    wks.Rows("8:" & MIDDLE).Insert , shift:=xlDown
                    TCR.Copy
                    wks.Range("A8").PasteSpecial xlPasteFormats
                    TWS1.Range("A" & orderrowstart(i) & ":" & "T" & orderrowfinish(i)).Copy
                    wks.Range("A8").PasteSpecial xlPasteAll
                    TWS1.Range("U" & orderrowstart(i) & ":" & "V" & orderrowfinish(i)).Copy
                    wks.Range("U8").PasteSpecial xlPasteValues
                    TWS1.Rows(orderrowfinish(i)).Copy                                               'copy original total formulas
                    wks.Paste Destination:=wks.Range("A" & MIDDLE)
                    
         
                    
                    wks.PageSetup.FitToPagesWide = 1
                    Application.DisplayAlerts = False
            
                    wkb.SaveAs directory & "\" & "箱唛发你银行账号请核对体积重量材质品牌请回传 " & orderno(i) & " " & middlename & ".xls", FileFormat:=56
            
                    wkb.Close False
                End If
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

Sub createfolder(sname As String)

    Dim fso As Object
    Set fso = CreateObject("scripting.filesystemobject")
 'check if target folder exist
    If fso.FolderExists(sname) Then
    
    Else
            fso.createfolder (sname)
            
    End If
End Sub


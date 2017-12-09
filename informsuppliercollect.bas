Attribute VB_Name = "informsuppliercollect"

Public wkb As Workbook, wks1 As Worksheet, wks2 As Worksheet
Public TWS1 As Worksheet, TWS2 As Worksheet, TWS3 As Worksheet, TWS4 As Worksheet, TWS5 As Worksheet
Public collecttime(100) As Date, collectdaterange(100) As Range, deposit(100) As Double, orderqty As Integer
Public paymentterm(100) As String, suppliercode(100) As String
Public suppliername(100) As String, supplierchinesename(100) As String






Sub informsuppliercollect()


Application.ScreenUpdating = False
Application.DisplayAlerts = False

    
    
    
    Dim tname As String, projectname As String, directory As String
    
    
    
    Dim fso As Object, otname As Object
    Set fso = CreateObject("scripting.filesystemobject")
     
    projectname = "ST1117"
    directory = fso.GetFile(ThisWorkbook.FullName).ParentFolder.ParentFolder.path
    directory = directory & "\Market order\" & projectname & "\YW\inform supplier collect date"         'set file location
    
    
    Call createfolder(directory)
    
    
    Set TWS1 = ThisWorkbook.Worksheets("order detail")   'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")   'target worksheet
    Set TWS3 = ThisWorkbook.Worksheets("shipping mark")   'target worksheet
    'Set TWS4 = ThisWorkbook.Worksheets("bank detail collecting")   'target worksheet
    Set TWS5 = ThisWorkbook.Worksheets("collect information")   'target worksheet
    
    Call ordermodule
    Call savepicture(TWS2, TWS5, directory)
    
    
    
    
End Sub

Sub ordermodule()

Dim suppliercode1 As Range, suppliercode2 As Range
 Dim ps As Integer, pf As Integer, middlename As String, BP As Range


    Set suppliercode1 = TWS2.Range("A1")
    Set suppliercode2 = TWS2.Range("A1")
    
    
    
    Dim k As Integer
    Dim collecttimeclock As Date
    collecttimeclock = CDate("2017-12-12 10:00")
   
    
    'collecttimeclock = Format(collecttimeclock, "YY-MM-DD HH")
    orderqty = 0
    
        
    
    Dim intevel As Integer
    intevel = 0         'set how many suppiers collect at same time
    
    For i = 1 To 100
    
        Set suppliercode2 = TWS2.Range("A:A").Find("YW1117", suppliercode1)                  'find target bank detail records
        
        
        
        
        If suppliercode2.Row < suppliercode1.Row Then
            MsgBox ("finish" & suppliercode1.Row)
            Exit For
    
        Else
            Set suppliercode1 = suppliercode2
        End If
        
       
        k = suppliercode1.Row
        
        Set collectdaterange(i) = TWS2.Range("k" & k)
        deposit(i) = TWS2.Range("H" & k).Value
        orderqty = orderqty + 1
        paymentterm(i) = TWS2.Range("L" & k).Value
        suppliercode(i) = TWS2.Range("A" & k).Value
        suppliername(i) = TWS2.Range("b" & k).Value
        
        
          'cut chinese supplier name in suppliername(i) and then find it in wks1
    
   
    
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
     
    
    supplierchinesename(i) = Mid(suppliername(i), ps, pf - ps)
        
        
    Next
    
    For i = 1 To orderqty
    If InStr(paymentterm(i), "当天转账") <> 0 Then
        intevel = intevel + 1
        If intevel <= 4 Then
            collecttime(i) = collecttimeclock
            collectdaterange(i).Value = collecttime(i)
        Else
            collecttimeclock = DateAdd("n", 60, collecttimeclock)
            
            If Hour(collecttimeclock) = 12 Then
                collecttimeclock = DateAdd("n", 30, collecttimeclock)
            ElseIf Hour(collecttimeclock) = 17 Then
                collecttimeclock = DateAdd("h", 16, collecttimeclock)
                collecttimeclock = DateAdd("n", 30, collecttimeclock)
                
            End If
        
            collecttime(i) = collecttimeclock
            collectdaterange(i).Value = collecttime(i)
            intevel = 1
        End If
    End If
    Next
    
     For i = 1 To orderqty
    
    If InStr(paymentterm(i), "第二天转账") <> 0 Then
        intevel = intevel + 1
        If intevel <= 4 Then
            collecttime(i) = collecttimeclock
            collectdaterange(i).Value = Format(collecttime(i), "YYYY-MM-DD-HH:nn")
        Else
            collecttimeclock = DateAdd("n", 60, collecttimeclock)
            
            If Hour(collecttimeclock) = 12 Then
                collecttimeclock = DateAdd("n", 30, collecttimeclock)
            ElseIf Hour(collecttimeclock) = 17 Then
                collecttimeclock = DateAdd("h", 16, collecttimeclock)
                collecttimeclock = DateAdd("n", 30, collecttimeclock)
                
            End If
        
            collecttime(i) = collecttimeclock
            collectdaterange(i).Value = Format(collecttime(i), "YYYY-MM-DD-HH:nn")
            intevel = 1
        End If
    End If
    Next


End Sub
    
    




Sub savepicture(Sourcedatasheet As Worksheet, Sourcepicturesheet As Worksheet, directory As String)


    Dim rng As Range, Fn As String, time As String
    
     
    
    
    
    
    Set rng = Sourcepicturesheet.Range("b7:g16")
    
    For i = 1 To orderqty
        Sourcepicturesheet.Range("b7").Value = supplierchinesename(i) & "您好"
        Sourcepicturesheet.Range("D9").Value = Format(collecttime(i), "YYYY年MM月DD日HH时") & "左右送到"
        Application.CutCopyMode = False
        rng.CopyPicture xlScreen
        

    
        Set ocht1 = Worksheets.Add
        Set ocht2 = ocht1.Shapes.AddChart(Width:=rng.Width, Height:=rng.Height).Chart
    
    
   'Set ocht = Charts.Add
   ' ocht.ChartArea.Clear
    
        ocht2.Parent.Select
        ocht2.Paste
        
        With ocht2
   ' .ChartArea.Width = Rng.Width
   ' .ChartArea.Height = Rng.Height
    
  '          .Paste
            .Export filename:=directory & "\" & "送货确认" & Format(collecttime(i), "YYYY年MM月DD日HH时") & "左右送到" & suppliercode(i) & " " & supplierchinesename(i) & ".jpg", Filtername:="JPG"
        End With
        ocht1.Delete
    Next
    
    
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

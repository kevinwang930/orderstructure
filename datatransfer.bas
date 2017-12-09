Attribute VB_Name = "datatransfer"
Sub ADOFromExcelToAccess()
' exports data from the active worksheet to a table in an Access database
' this procedure must be edited before use
Dim cn As ADODB.Connection, rs As ADODB.Recordset, R As Long
Dim strcn As String
Dim directory As String


 
    Set TWS1 = ThisWorkbook.Worksheets("order detail")                 'target worksheet
    Set TWS2 = ThisWorkbook.Worksheets("bank detail")                   'target worksheet
'   Set TWS3 = ThisWorkbook.Worksheets("bank detail collect report")   'target worksheet
    Set TWS4 = ThisWorkbook.Worksheets("shipping mark")         'target worksheet
    Set TWS5 = ThisWorkbook.Worksheets("collect information")   'target worksheet
    Set TWS6 = ThisWorkbook.Worksheets("checkdata")             'target worksheet
    
    Dim fso As Object, otname As Object
    Set fso = CreateObject("scripting.filesystemobject")
    
    directory = fso.GetFile(ThisWorkbook.FullName).ParentFolder.path
    
' connect to the Access database
Set cn = New ADODB.Connection
strcn = directory & "\abc.accdb;"
With cn
.Provider = "Microsoft.ACE.OLEDB.12.0"
.ConnectionString = "Data source=" & strcn
.Open
End With
' open a recordset
Set rs = New ADODB.Recordset
rs.ActiveConnection = cn


 Dim sSQL As String
    sSQL = "CREATE TABLE BridgerSubstitute (" & _
        "Auto_Increment COUNTER CONSTRAINT PrimaryKey PRIMARY KEY, " & _
        "First_Name varchar(255), " & "Middle_Name varchar(255), " & "Last_Name varchar(255), " & _
        "Entity_Type varchar(255), " & "Address_1 varchar(255), " & "City_1 varchar(255), " & _
        "State_1 varchar(255), " & "Zip_Code_1 varchar(255), " & "Country_1 varchar(255), " & _
        "Address_2 varchar(255), " & "City_2 varchar(255), " & "State_2 varchar(255), " & _
        "Zip_Code_2 varchar(255), " & "Country_2 varchar(255), " & "Aliases varchar(255), " & _
        "Alternate_Spellings varchar(255), " & "Additional_Information varchar(255))"
    cn.Execute sSQL

' all records in a table
R = 11                           'the start row in the worksheet
Do While Len(Range("A" & R).Formula) > 0
' repeat until first empty cell in column A
With rs
.AddNew                         'create a new record
' add values to each field in the record
.Fields("FieldName1") = TWS1.Range("A" & R).Value
.Fields("FieldName2") = TWS1.Range("B" & R).Value
.Fields("FieldNameN") = TWS1.Range("C" & R).Value
' add more fields if necessary¡­
.Update 'stores the new record
End With
R = R + 1           ' next row
Loop
rs.Close
Set rs = Nothing
cn.Close
Set cn = Nothing
End Sub


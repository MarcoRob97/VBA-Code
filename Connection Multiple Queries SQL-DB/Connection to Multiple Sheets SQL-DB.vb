
Sub main()

    Call unhide
    Call ExecuteSQLQueries
    Call hide
    ThisWorkbook.RefreshAll
    Sheets("Executive Summary").Select
    
    
End Sub


Sub ExecuteSQLQueries()

    ' Define the necessary variables
    Dim conn As Object ' ADODB.Connection
    Dim rs As Object ' ADODB.Recordset
    Dim strConn As String
    Dim strSQL() As String
    Dim i As Integer
    Dim fldCount As Integer
    Dim ws() As Worksheet
    Dim pivotTableRange() As Range
    
    ' Set the connection string
    strConn = "Provider='sqloledb';Data Source='saturnv2';Initial Catalog='DataMart';Integrated Security='SSPI';"
    
    ' Set the SQL queries
    ReDim strSQL(1 To 5)
    strSQL(1) = "select * from table1;"
    strSQL(2) = "select * from table2;"
    strSQL(3) = "select * from table3;"
    strSQL(4) = "select * from table4;"
    strSQL(5) = "select * from table5;"
  

    ' Set the worksheets to paste the results
    ReDim ws(1 To 5)
    Set ws(1) = ThisWorkbook.Sheets("Sheet1") ' Replace "Sheet1" with the desired sheet name for query 1
    Set ws(2) = ThisWorkbook.Sheets("Sheet2") ' Replace "Sheet2" with the desired sheet name for query 2
    Set ws(3) = ThisWorkbook.Sheets("Sheet3") ' Replace "Sheet3" with the desired sheet name for query 3
    Set ws(4) = ThisWorkbook.Sheets("Sheet4") ' Replace "Sheet4" with the desired sheet name for query 4
    Set ws(5) = ThisWorkbook.Sheets("Sheet5") ' Replace "Sheet5" with the desired sheet name for query 5
    
    ' Set the pivot table ranges for each query
    ReDim pivotTableRange(1 To 5)
    Set pivotTableRange(1) = ws(1).Range("A1").CurrentRegion ' Replace "A1" with the top-left cell of your pivot table for query 1
    Set pivotTableRange(2) = ws(2).Range("A1").CurrentRegion
    Set pivotTableRange(3) = ws(3).Range("A1").CurrentRegion
    Set pivotTableRange(4) = ws(4).Range("A1").CurrentRegion
    Set pivotTableRange(5) = ws(5).Range("A1").CurrentRegion
    
    
    ' Initialize the connection and recordset objects
    Set conn = CreateObject("ADODB.Connection")
    Set rs = CreateObject("ADODB.Recordset")
    
    ' Open the connection
    conn.Open strConn
    
    ' Loop through each query
    For i = 1 To 5
        ' Execute the SQL query
        rs.Open strSQL(i), conn
        
        ' Clear existing data within the pivot table range
        pivotTableRange(i).Offset(1, 0).Resize(pivotTableRange(i).Rows.Count - 1, pivotTableRange(i).Columns.Count).ClearContents
        
        ' Paste the field names in the first row
        fldCount = rs.Fields.Count
        For j = 1 To fldCount
            ws(i).Cells(pivotTableRange(i).Row, pivotTableRange(i).Column + j - 1).Value = rs.Fields(j - 1).Name
        Next j
        
        ' Paste the query result starting from the second row within the pivot table range
        ws(i).Cells(pivotTableRange(i).Row + 1, pivotTableRange(i).Column).CopyFromRecordset rs
        
        ' Close the recordset
        rs.Close
    Next i
    
    ' Close the connection
    conn.Close
    
    ' Cleanup
    Set rs = Nothing
    Set conn = Nothing
    
    
End Sub

Sub unhide():
     

    Sheets("Sheet1").Visible = True

    Sheets("Sheet2").Visible = True

    Sheets("Sheet3").Visible = True
    
    Sheets("Sheet4").Visible = True
    
    Sheets("Sheet5").Visible = True
    
  
End Sub

Sub hide():
     

    Sheets("Sheet1").Visible = False
    
    Sheets("Sheet2").Visible = False

    Sheets("Sheet3").Visible = False
    
    Sheets("Sheet4").Visible = False
    
    Sheets("Sheet5").Visible = False
    
  
End Sub
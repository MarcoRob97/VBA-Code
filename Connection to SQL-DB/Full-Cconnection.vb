Sub re_hourly_act()
      
      Call clean ' First is cleaning the table an inserting fresh data
      Call MyConnectionToDB ' Finally do the conection and paste it into the sheet 
      ThisWorkbook.RefreshAll

End Sub


Sub clean()
' This code open the unhide the sheet where i have the pivot table that feeds the dashboard
    Sheets("sheet 1").Select
    Sheets("sheet 2").Visible = True
    Range("sheet 2").Select
    Selection.ClearContents
    Sheets("sheet 2").Select
    ActiveWindow.SelectedSheets.Visible = False
' at the end hide again the sheet in that way user don't notice how the dashboard is feed it
    
End Sub


Sub MyConnectionToDB()
    ' Define variables
    Dim rs As ADODB.Recordset
    Dim cn As ADODB.Connection
    Dim strConn As String
    Dim strSQL As String
    Dim intPublisherCount As Integer
    Dim strCountry As String
    Dim strMessage As String

    ' Specify your database connection details
    Dim dbServer As String   ' Database server name or IP address
    Dim dbCatalog As String  ' Database name
    Dim dbUser As String     ' Database username
    Dim dbPassword As String ' Database password

    ' Set your database connection details here
    dbServer = "YourServerName"     ' Replace with your database server
    dbCatalog = "YourDatabaseName"  ' Replace with your database name
    dbUser = "YourUsername"        ' Replace with your database username
    dbPassword = "YourPassword"    ' Replace with your database password

    ' Create a new ADODB Connection
    Set cn = New ADODB.Connection

    ' Construct the connection string with placeholders for username and password
    strConn = "Provider='sqloledb';Data Source='" & dbServer & "';" & _
              "Initial Catalog='" & dbCatalog & "';User ID='" & dbUser & "';" & _
              "Password='" & dbPassword & "';"

    ' Open the database connection
    cn.Open strConn

    ' Define your SQL query
    strSQL = "SELECT * FROM table1"

    ' Create a new ADODB Recordset
    Set rs = New ADODB.Recordset

    ' Execute the SQL query and open the Recordset
    rs.Open strSQL, cn, adOpenStatic, adLockReadOnly

    ' Check if the Recordset is not empty
    If Not rs.EOF Then
        ' Copy data from the Recordset to the worksheet "Data" starting from cell A2
        Sheets("Data").Range("A2").CopyFromRecordset rs
    Else
        ' Handle the case when the Recordset is empty
        ' (Add your code or message here if needed)
    End If

    ' Close the Recordset and database connection
    rs.Close
    cn.Close

    ' Clean up and release resources
    Set rs = Nothing
    Set cn = Nothing

    
End Sub

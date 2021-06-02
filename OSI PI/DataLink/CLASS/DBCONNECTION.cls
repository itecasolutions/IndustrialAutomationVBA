VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DBCONNECTION"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Covington  Library
'Written 29OCT2020 by Nicholas Stom
'Class is used for connection to DATABASE

Option Explicit
' These are properties
    Public dataSource As String
    Public password As String
    Public userName As String
    Public catalog As String
    Public command As String
    Public dataSet As Variant
    Public dbA As ADODB.Connection
    Public rsA As ADODB.Recordset

Sub OpenConnection()
    Set dbA = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    dbA.ConnectionString = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & userName & ";Initial Catalog=" & catalog & ";Data Source=" & dataSource & ";"
    
    dbA.Open
    
End Sub


Sub Execute()
    Dim csSql As String
    csSql = command
    Set rsA = dbA.Execute(csSql)
    dataSet = rsA.GetRows
    
End Sub

Sub CloseConnection()
    dbA.Close
    
End Sub

Sub GetData()
 'Set Up Connection To DB
    Dim dbA As ADODB.Connection
    Dim rsA As ADODB.Recordset
    Set dbA = New ADODB.Connection
    Set rsA = New ADODB.Recordset
    dbA.ConnectionString = "Provider=SQLOLEDB.1;Password=" & password & ";Persist Security Info=True;User ID=" & userName & ";Initial Catalog=" & catalog & ";Data Source=" & dataSource & ";"
    
    dbA.Open
    Dim csSql As String
    csSql = command
    Set rsA = dbA.Execute(csSql)
    dataSet = rsA.GetRows
    Debug.Print "This method is depreciated"
End Sub

Attribute VB_Name = "Module1"
Public gAdoConn As ADODB.Connection


Public Function SQLSafe(strValue As String) As String
    Dim strTemp1 As String
    
    strTemp1 = Replace(strValue, "'", "''")
    
    SQLSafe = strTemp1
End Function

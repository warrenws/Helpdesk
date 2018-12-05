'Created by Matthew Hull 5/1/14

'This script will remove expired sessions

On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strCurrentFolder = objFSO.GetParentFolderName(strCurrentFolder)
strDatabase = strCurrentFolder & "\Database\helpdesk.mdb"

'Connect to the database
Set objConnection = CreateObject("ADODB.Connection")

'Attempt to connect using the Jet engine
strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDatabase & ";"
objConnection.Open strConnection

'If using the Jet engine failed try using the Access engine
If Err Then
   strConnection = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & strDatabase & ";"
   Err.Clear
   objConnection.Open strConnection
End If

'Set the theme
strSQL = "DELETE FROM Sessions WHERE Date() >= ExpirationDate"
objConnection.Execute(strSQL)


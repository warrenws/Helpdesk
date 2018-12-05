'Created by Matthew Hull 10/12/12

'This script will correct the techs times in the database

On Error Resume Next

Set objFSO = CreateObject("Scripting.FileSystemObject")
strCurrentFolder = objFSO.GetAbsolutePathName(".")
strCurrentFolder = objFSO.GetParentFolderName(strCurrentFolder)
strCurrentFolder = objFSO.GetParentFolderName(strCurrentFolder)
strDatabase = strCurrentFolder & "\helpdesk.mdb"

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

strSQL = "UPDATE Log SET TaskTime=null"
objConnection.Execute(strSQL)

strSQL = "SELECT DISTINCT Ticket" & vbCRLF
strSQL = strSQL & "FROM Log" & vbCRLF
strSQL = strSQL & "ORDER BY Ticket"
Set objTicketList = objConnection.Execute(strSQL)

If Not objTicketList.EOF Then
   Do Until objTicketList.EOF
   
      strSQL = "SELECT ID,NewValue,UpdateDate,UpdateTime" & vbCRLF
      strSQL = strSQL & "FROM Log" & vbCRLF
      strSQL = strSQL & "WHERE Ticket=" & objTicketList(0) & " And " 
      strSQL = strSQL & "(Type='Assigned' Or Type='Auto Assigned' Or Type='New Ticket' Or Type='Tech Reassigned' Or NewValue='Complete' or Type='Tech Changed')" & vbCRLF
      strSQL = strSQL & "ORDER BY ID"
      Set objTicket = objConnection.Execute(strSQL)
   
      intCurrentID = ""
      strCurrentNewTech = ""
      datCurrentDateTime = ""
      intPreviousID = ""
      strPreviousNewTech = ""
      datPreviousDateTime = ""
      
      If Not objTicket.EOF Then
         Do Until objTicket.EOF
            datTimeSpent = ""
            intCurrentID = objTicket(0)
            strCurrentNewTech = objTicket(1)
            datCurrentDateTime = objTicket(2) & " " & objTicket(3)
            
            If (strCurrentNewTech <> strPreviousNewTech) Then
               If datPreviousDateTime <> "" Then
                  datTimeSpent = DateDiff("n",datPreviousDateTime,datCurrentDateTime)
                  strSQL = "UPDATE Log" & vbCRLF
                  strSQL = strSQL & "SET TaskTime=" & datTimeSpent & vbCRLF
                  strSQL = strSQL & "WHERE ID=" & intPreviousID
                  objConnection.Execute(strSQL)
               End If
            End If
            
            intPreviousID = intCurrentID
            strPreviousNewTech = strCurrentNewTech
            datPreviousDateTime = datCurrentDateTime
            
            objTicket.MoveNext
         Loop
      End If
   
      objTicketList.MoveNext
   Loop
End If

MsgBox "Tech times have been fixed",vbOkOnly,"Fixed"
'Created by Matthew Hull 3/21/12

'This script will send an email to each tech about unassigned tickets.

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

strSQL = "SELECT AdminURL, BBC, SendMailFrom, SMTPPickupFolder" & vbCRLF
strSQL = strSQL & "FROM Settings"
Set objSettings = objConnection.Execute(strSQL)
strURL = objSettings(0)
strBCC = objSettings(1)
strFrom = objSettings(2)
strSMTPPickupFolder = objSettings(3)

strSQL = "SELECT Tech, EMail FROM Tech WHERE Active"
Set objTechList = objConnection.Execute(strSQL)

strSQL = "SELECT ID, DisplayName, SubmitDate, SubmitTime, Problem, Notes, Location" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "Where Tech=''"

Set objUnassignedTickets = objConnection.Execute(strSQL)

If NOT objUnassignedTickets.EOF Then
   Do Until objTechList.EOF
      SendEMail   
      objTechList.MoveNext
   Loop
End If

Sub SendEMail
   
   Const cdoSendUsingPickup = 1

   'Create the objects required to send the mail.
   Set objMessage = CreateObject("CDO.Message")
   Set objConf = objMessage.Configuration
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = strSMTPPickupFolder
      .Update
   End With
   
   strMessage = objTechList(0) & "," & vbCRLF & vbCRLF
   strMessage = strMessage & "This is an automated email from the Help Desk.  "
   strMessage = strMessage & "Below you will find a list of help desk tickets that aren't assigned to anyone.  "
   strMessage = strMessage & "This message has been sent to all techs in the help desk to ensure no tasks are lost.  "
   strMessage = strMessage & "If any of the tickets belong to you please assign them to yourself." & vbCRLF & vbCRLF
   strMessage = strMessage & "**************************************************************************************" & vbCRLF 
   
   Do Until objUnassignedTickets.EOF
      strMessage = strMessage & "Ticket #" & objUnassignedTickets(0) & " - " & objUnassignedTickets(1) & vbCRLF
      strMessage = strMessage & "   - Location: " & objUnassignedTickets(6) & vbCRLF
      strMessage = strMessage & "   - Submitted: " & objUnassignedTickets(2) & " - " & objUnassignedTickets(3) & vbCRLF
      strMessage = strMessage & "   - Problem: " & objUnassignedTickets(4) & vbCRLF 
      If objUnassignedTickets(5) <> "" Then
         strMessage = strMessage & "   - Notes: " & objUnassignedTickets(5) & vbCRLF
      End If
      strMessage = strMessage & "   - " & strURL & "/modify.asp?ID=" & objUnassignedTickets(0) & vbCRLF
      strMessage = strMessage & "**************************************************************************************" & vbCRLF 
      objUnassignedTickets.MoveNext
   Loop
   
   objUnassignedTickets.MoveFirst
   
   strMessage = strMessage & vbCRLF & "Thank you"  & vbCRLF & vbCRLF & "Please do not respond to this message..."

   With objMessage
      .To = "mhull@wswheboces.org" 'objTechList(1)
      .From = strFrom
      .Subject = "Help Desk - Unassigned Tickets"
      .TextBody = strMessage
      If strBCC <> "" Then
         .BCC = strBCC
      End If
      .Send
   End With
   
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub
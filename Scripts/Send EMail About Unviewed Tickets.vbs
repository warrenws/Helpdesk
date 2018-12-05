'Created by Matthew Hull 12/19/11

'This script will send an email to each tech that has unviewed tickets.

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

strSQL = "SELECT Tech, EMail FROM Tech"
Set objTechList = objConnection.Execute(strSQL)

Do Until objTechList.EOF
   strSQL = "SELECT ID, DisplayName, SubmitDate, SubmitTime, Problem, Notes" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "Where Tech='" & objTechList(0) & "' AND TicketViewed=False AND Status<>'Complete'"
   
   Set objUnviewedTickets = objConnection.Execute(strSQL)
   If NOT objUnviewedTickets.EOF Then
      SendEMail
   End If
   
   objTechList.MoveNext
Loop

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
   strMessage = strMessage & "Below you will find a list of help desk tickets that are assigned to you and not marked as viewed.  "
   strMessage = strMessage & "Please use the link to view the tickets." & vbCRLF & vbCRLF
   strMessage = strMessage & "**************************************************************************************" & vbCRLF
   
   Do Until objUnviewedTickets.EOF
      strMessage = strMessage & "Ticket #" & objUnviewedTickets(0) & " - " & objUnviewedTickets(1) & vbCRLF
      strMessage = strMessage & "   - Submitted: " & objUnviewedTickets(2) & " - " & objUnviewedTickets(3) & vbCRLF
      strMessage = strMessage & "   - Problem: " & objUnviewedTickets(4) & vbCRLF 
      If objUnviewedTickets(5) <> "" Then
         strMessage = strMessage & "   - Notes: " & objUnviewedTickets(5) & vbCRLF
      End If
      strMessage = strMessage & "   - " & strURL & "/modify.asp?ID=" & objUnviewedTickets(0) & vbCRLF
      strMessage = strMessage & "**************************************************************************************" & vbCRLF 
      objUnviewedTickets.MoveNext
   Loop
   
   strMessage = strMessage & vbCRLF & "Thank you"  & vbCRLF & vbCRLF & "Please do not respond to this message..."

   With objMessage
      .To = objTechList(1)
      .From = strFrom
      .Subject = "Help Desk - List of Unviewed Tickets"
      .TextBody = strMessage
      If strBCC <> "" Then
         .BCC = strBCC
      End If
      .Send
   End With
   
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub
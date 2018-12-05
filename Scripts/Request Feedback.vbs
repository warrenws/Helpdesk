'Created by Matthew Hull 1/13/12

'This script will send an email to each person who had a ticket closed
'the day before.  It will ask them for feedback

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
strURL = Left(objSettings(0),Len(objSettings(0))-5) & "feedback.asp?"
strBCC = objSettings(1)
strFrom = objSettings(2)
strSMTPPickupFolder = objSettings(3)

strSQL = "SELECT FeedBack FROM Counters WHERE ID = 1"
Set objCounter = objConnection.Execute(strSQL)
intCounter = objCounter(0)

strSQL = "SELECT ID, Email, Problem, DisplayName, Tech" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "WHERE LastUpdatedDate=Date()-1 AND Status='Complete'"

Set objFeedback = objConnection.Execute(strSQL)

If Not objFeedback.EOF Then
   Do Until objFeedback.EOF
      If UCase(objFeedback(3)) <> UCase(objFeedback(4)) Then
         intCounter = intCounter + 1
         SendEMail
      End If
      objFeedback.MoveNext
   Loop
End If

strSQL = "UPDATE Counters" & vbCRLF
strSQL = strSQL & "SET Feedback=" & intCounter & vbCRLF
strSQL = strSQL & "WHERE ID=1"
objConnection.Execute(strSQL)

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
   strMessage = "<html><body style=""font-family:Verdana,Arial,Helvetica,sans-serif;font-size:10pt;"">"
   strMessage = strMessage & "<p>Your answer to this one quick question will help us improve our service. "
   strMessage = strMessage & "How satisfied are you with the performance of our team regarding your recently closed "
   strMessage = strMessage & "help desk ticket #" & objFeedback(0) & "?  Click the link below to submit your feedback: </p>"
   strMessage = strMessage & "<ul>" & vbCRLF
   strMessage = strMessage & "<li><a href=""" & strURL & "rating=5&ticket=" & objFeedback(0) & """>Very satisfied</a></li>" & vbCRLF
   strMessage = strMessage & "<li><a href=""" & strURL & "rating=4&ticket=" & objFeedback(0) & """>Satisfied</a></li>" & vbCRLF
   strMessage = strMessage & "<li><a href=""" & strURL & "rating=3&ticket=" & objFeedback(0) & """>Neutral</a></li>" & vbCRLF
   strMessage = strMessage & "<li><a href=""" & strURL & "rating=2&ticket=" & objFeedback(0) & """>Dissatisfied</a></li>" & vbCRLF
   strMessage = strMessage & "<li><a href=""" & strURL & "rating=1&ticket=" & objFeedback(0) & """>Very dissatisfied</a></li>" & vbCRLF
   strMessage = strMessage & "</ul>"
   strMessage = strMessage & vbCRLF & "<p><b>Original Problem</b>: " & objFeedback(2) & "</p>"
   strMessage = strMessage & vbCRLF & "<p>Thank you for your time.</p>"
   strMessage = strMessage & vbCRLF & "<p>Please do not respond to this email...</p>"
   strMessage = strMessage & "</body></html>"

   With objMessage
      .To = objFeedback(1)
      .From = strFrom
      .Subject = "How did we do?"
      .HTMLBody = strMessage
      If strBCC <> "" Then
         .BCC = strBCC
      End If
      .Send
   End With
   
   Set objMessage = Nothing
   Set objConf = Nothing
   
End Sub
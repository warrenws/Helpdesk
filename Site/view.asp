<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 12/7/04
'Last Updated 6/16/14

'This page will list all the calls for the user.

Option Explicit

On Error Resume Next

Dim strSQL, strTitle, strSQLLocation, strUser, objNetwork, strFilter, objTrackedTickets
Dim objRecordSet, strIcon, intIconSpan, strMessage, intRecordCount, bolShowLogout
Dim strSymbol, strLinkBar, strNotes, strProblem, bolComplete, strStatusMessage
Dim strTrackedTickets, strRequestedUpdate, strUserAgent

'Redirect the user the SSL version if required
If Application("ForceSSL") Then
   If Request.ServerVariables("SERVER_PORT")=80 Then
      If Request.ServerVariables("QUERY_STRING") = "" Then
         Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL")
      Else
         Response.Redirect "https://" & Request.ServerVariables("SERVER_NAME") & Request.ServerVariables("URL") & "?" & Request.ServerVariables("QUERY_STRING")
      End If
   End If
End If

'If the database and the website are not the same version then let them know
If Application("VersionError") Then
   VersionProblem
End If

'Kick them out if they try to sneak in
If Not Application("UserCanViewCallStatus") Then
   Response.Redirect("index.asp")
End If
   
'Get the users logon name
Set objNetwork = CreateObject("WSCRIPT.Network")   
strUser = objNetwork.UserName

'Check and see if anonymous access is enabled
If LCase(Left(strUser,4)) = "iusr" Then
   strUser = GetUser
   bolShowLogout = True
Else
   bolShowLogout = False
End If
   
If Request.Form("Log Out") = "Log Out" Then
   Response.Redirect "login.asp?action=logout"
End If 
   
'Send you to the home page if home is clicked on the iPhone version
If Request.Form("home") = "Home" or Request.Form("back") = "Back" Then
   Response.Redirect "index.asp"
End If
   
'Get the Filter from the URL and build the link bar
strFilter = LCase(Request.QueryString("Filter"))
If strFilter = "" Then
   strFilter = LCase(Request.Form("Filter"))
End If

Select Case strFilter
   Case "closed", "closed tickets"
      strSymbol = "="
      bolComplete = True
   Case "open", "open tickets"
      strSymbol = "<>"
      bolComplete = False
   Case Else
      strSymbol = "<>"
      bolComplete = False
End Select
      
'Check and see if a button was hit
Select Case Request.Form("cmdSubmit")
   Case "Request Update"
      RequestUpdate
      strMessage = "Update requested..."
   Case "Cancel Update Request"
      CancelUpdateRequest
      strMessage = "Requested update cancelled..."
   Case "Track Ticket"
      TrackTicket
      strMessage = "Enabled ticket tracking..."
   Case "Stop Tracking"
      DontTrackTicket
      strMessage = "Disabled ticket tracking..."
   Case "Close Ticket"
      CloseTicket
      strMessage = "Ticket closed..."
End Select
      
'Get the list of tickets they are tracking or have requested an update
strSQL = "SELECT Ticket,Type FROM Tracking WHERE TrackedBy='" & strUser & "'"
Set objTrackedTickets = Application("Connection").Execute(strSQL)

'Build the strings that contain the list of tracked tickets and tickets where they requested updates
strTrackedTickets = ";"
strRequestedUpdate = ";"
If Not objTrackedTickets.EOF Then
   Do Until objTrackedTickets.EOF
      Select Case objTrackedTickets(1)
         Case "Request"
            strRequestedUpdate = strRequestedUpdate & objTrackedTickets(0) & ";"
         Case "Track"
            strTrackedTickets = strTrackedTickets & objTrackedTickets(0) & ";"
      End Select
      objTrackedTickets.MoveNext
   Loop
End If
      
'Get the tickets from the database
strSQL = "SELECT ID,Name,Location,Status,Category,Tech,SubmitDate,SubmitTime,Problem,Notes,LastUpdatedDate,LastUpdatedTime,Custom1,Custom2,DisplayName,EMail" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "Where Name='" & strUser & "' And Status" & strSymbol & "'Complete'"
strSQL = strSQL & "ORDER BY ID DESC;"
Set objRecordSet = Application("Connection").Execute(strSQL)
      
'Count the number of returned records.  If none are returned a message will be
'displayed to the user.
intRecordCount = 0
If Not objRecordSet.EOF Then
   Do Until objRecordSet.EOF
      intRecordCount = intRecordCount + 1
      objRecordSet.MoveNext
   Loop
   objRecordSet.MoveFirst
End If

If intRecordCount = 0 Then
   strMessage = "No Tickets Found"
End If       
      
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
      
If IsMobile Then
   MobileVersion
Else
   FullVersion
End If

%>

<%
Sub FullVersion %>

      <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
      "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
      <html>
      
      <head>
         <title>Help Desk</title>
         <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
         <link rel="stylesheet" type="text/css" href="themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
         <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadusericon.png" />
         <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
         <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 9") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.99, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 7") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.78, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 5") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.47, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"Nexus 6") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.53, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
   <% If InStr(strUserAgent,"GT-N5110") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.77, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
         <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
      </head>
      
      <body>
      <div class="header">
         <%=Application("SchoolName")%> Help Desk (<%=intRecordCount%>)
      </div>

      <div class="version">
         Version <%=Application("Version")%>
      </div>

      <hr class="usertopbar"/>
      <div class="usertopbar">
         <ul class="topbar">
            <li class="topbar"><a href="index.asp">Home</a><font class="separator"> | </font></li>
         <% If bolComplete Then %>
            <li class="topbar"><a href="view.asp?filter=open">Open Tickets</a><font class="separator"> | </font></li>
            <li class="topbar">Closed Tickets</li>
         <% Else %>
            <li class="topbar">Open Tickets<font class="separator"> | </font></li>
            <li class="topbar"><a href="view.asp?filter=closed">Closed Tickets</a></li>  
         <% End If %>
         <% If bolShowLogout Then %>
            <font class="separator"> | </font></li>
            <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
         <% Else %>
            </li>
         <% End If %>
         </ul>
      </div>
      <hr class="userbottombar"/>
      <div align="center">
      <table border="0" width="750" cellspacing="0" cellpadding="0">
      <% If strMessage <> "" Then %>
         <tr><td><font class="information"><%=strMessage%></font></td></tr>
         <tr><td><hr /></td></tr>
      <% End If%>
   <% Do  Until objRecordSet.EOF
         
         'Change a carriage return to a <br> so it will display properly in HTML.
         strNotes = HideText(objRecordSet(9))
         If strNotes <> "" Then
            strNotes = Replace(strNotes,vbCRLF,"<br />")
         End If
         If objRecordSet(8) <> "" Then
            strProblem = FixURLs(objRecordSet(8),1)
            strProblem = Replace(strProblem,vbCRLF,"<br />")
         End If
         
         'If the notes field contains data then change the icon to the page with the N on it
         If objRecordSet(9) <> "" Then
            strIcon = "nedit"
             
            'Set the number of rows that the icon section should span.  This will change if the call'
            'is closed
            If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then
               intIconSpan = 7 'Has notes and is complete
            Else
               intIconSpan = 7 'Has notes and is not complete
            End If
         Else
            strIcon = "edit"
             
            'Set the number of rows that the icon section should span.  This will change if the call'
            'is closed
            If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then
               intIconSpan = 4 'No notes and is complete
            Else
               intIconSpan = 4 'No notes and is not complete
            End If
         End If 
           
         'Drop one more row down in a custom variable is used.
         If Not Application("UseCustom1") And Not Application("UseCustom2") Then 
            intIconSpan = intIconSpan - 1
         End If 
         %>
         <form method="POST" action="view.asp">
         <input type="hidden" name="ID" value="<%=objRecordSet(0)%>">

         <tr><td>
         <table width="100%">
      	<tr>
      		<td rowspan="<%=intIconSpan%>" width="10%" valign="top">
               <center>
                  <img border="0" src="themes/<%=Application("Theme")%>/images/<%=strIcon%>.gif" width="25" height="32" />
                  <br /><%=objRecordSet(0)%>
               </center>
            </td>
            <td><b>Problem</b>:</td>
         </tr>
         <tr><td><%=strProblem%></td></tr>

<%       'Display the notes if they are in the database
         If objRecordSet(9) <> "" Then %>
            <tr><td>&nbsp;</td></tr>
            <tr><td><b>Notes</b>:</td></tr>
            <tr><td><%=FixURLs(strNotes,1)%></td></tr>
<%       End If %>
         
<%       'Display the date completed if the ticket is closed
         If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then %>
   	      <tr><td><hr /></td></tr>
            <tr><td>Closed: <%=objRecordSet(10)%> - <%=objRecordSet(11)%></td></tr>
<%       Else %>
            <tr><td><hr /></td></tr>
            <tr><td>
               <table width="100%">
                  <tr>
                     <td>Submitted: <%=objRecordSet(6)%> - <%=objRecordSet(7)%></td>
                     
               <% If Application("ShowUserButtons") Then %>     
                     <td align="right">

                  <% If InStr(strRequestedUpdate,";" & objRecordSet(0) & ";") Then %>
                        <input type="submit" value="Cancel Update Request" name="cmdSubmit">
                  <% Else %>
                        <input type="submit" value="Request Update" name="cmdSubmit">
                  <% End If %>

                  <% If InStr(strTrackedTickets,";" & objRecordSet(0) & ";") Then %>
                        <input type="submit" value="Stop Tracking" name="cmdSubmit">
                  <% Else %>
                        <input type="submit" value="Track Ticket" name="cmdSubmit">
                  <% End If %>
                        <input type="submit" value="Close Ticket" name="cmdSubmit">
                     </td>
               <% Else %>
                     <td>&nbsp;</td>
               <% End If %>
                  </tr>
               </table>
            </td></tr>
<%       End If %>
         
   	   <tr>
   	   	<td colspan="4"><hr></td>
   	   </tr>
         </form>
         </table>
         </td></tr>
<%       objRecordSet.MoveNext
      Loop %>
      </table>   
   </div>
   </body>   
   </html>
<%
End Sub %>

<%
Sub MobileVersion %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      
      <title>Help Desk</title>
      <meta http-equiv="Content-Type" content="text/html; charset=utf-8" />
      <link rel="stylesheet" type="text/css" href="themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadusericon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />      
      <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
   <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then %>
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %>  

   </head>
   
   <body>  
      <center><b><%=Application("SchoolName")%> Help Desk</b></center>
      <center>
      <table align="center">
         <tr><td><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               <div align="center">
                  <input type="submit" value="Home" name="home">
                  <input type="submit" value="Open Tickets" name="filter"> 
                  <input type="submit" value="Closed Tickets" name="filter">
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table>
      <table><tr><td width="<%=Application("MobileSiteWidth")%>">
      <table align="center">
      <% If strMessage <> "" Then %>
         <tr><td><font class="information"><%=strMessage%></font></td></tr>
         <tr><td><hr /></td></tr>
      <% End If%>
   <% Do  Until objRecordSet.EOF
         
         'Change a carriage return to a <br> so it will display properly in HTML.
         strNotes = HideText(objRecordSet(9))
         If strNotes <> "" Then
            strNotes = Replace(strNotes,vbCRLF,"<br />")
         End If
         If objRecordSet(8) <> "" Then
            strProblem = FixURLs(objRecordSet(8),1)
            strProblem = Replace(strProblem,vbCRLF,"<br />")
         End If
         %>
         <form method="POST" action="view.asp">
         <input type="hidden" name="ID" value="<%=objRecordSet(0)%>" />

         <tr><td>
         <table width="100%">
            <tr><td align="center"><b>Ticket #<%=objRecordSet(0)%></b></td></tr>
            <tr><td><b>Problem</b>:</td></tr>
            <tr><td><%=strProblem%></td></tr> 
            
<%       'Display the notes if they are in the database
         If objRecordSet(9) <> "" Then %>
            <tr><td>&nbsp;</td></tr>
            <tr><td><b>Notes</b>:</td></tr>
            <tr><td><%=Trim(FixURLs(strNotes,1))%></td></tr>
<%       End If %>

<%       'Display the date completed if the call is closed
         If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then %>
   	      <tr><td>&nbsp;</td></tr>
            <tr><td><b>Closed</b>:</td></tr>
            <tr><td><%=objRecordSet(10)%> - <%=objRecordSet(11)%></td></tr>
<%       Else %>
            <tr><td>&nbsp;</td></tr>
            <tr><td><b>Submitted</b>:</td></tr>
            <tr><td><%=objRecordSet(6)%> - <%=objRecordSet(7)%></td></tr>
            
         <% If Application("ShowUserButtons") Then %>      
               <tr><td align="center">
            <% If InStr(strRequestedUpdate,";" & objRecordSet(0) & ";") Then %>
                  <input type="submit" value="Cancel Update Request" name="cmdSubmit">
            <% Else %>
                  <input type="submit" value="Request Update" name="cmdSubmit">
            <% End If %>

            <% If InStr(strTrackedTickets,";" & objRecordSet(0) & ";") Then %>
                  <input type="submit" value="Stop Tracking" name="cmdSubmit">
            <% Else %>
                  <input type="submit" value="Track Ticket" name="cmdSubmit">
            <% End If %>
                  <input type="submit" value="Close Ticket" name="cmdSubmit">
               </td></tr>
         <% End If %> 
         
      <% End If %>
            <tr><td><hr /></td></tr>
         </table>
         </td></tr>
         </form>
   <%    objRecordSet.MoveNext
      Loop %>
      </table>
   </body>
   </html>
<%
End Sub %>

<%
Function IsMobile

   'It's not mobile if the user is requesting the full site
   Select Case LCase(Request.QueryString("Site"))
      Case "full"
         IsMobile = False
         Response.Cookies("SiteVersion") = "Full"
         Response.Cookies("SiteVersion").Expires = Date() + 14
         Exit Function
      Case "mobile"
         IsMobile = True
         Response.Cookies("SiteVersion") = "Mobile"
         Response.Cookies("SiteVersion").Expires = Date() + 14
         Exit Function
   End Select
   
   'Choose the site based on the cookie
   Select Case LCase(Request.Cookies("SiteVersion"))
      Case "full"
         IsMobile = False
         Exit Function
      Case "mobile"
         IsMobile = True
         Exit Function
   End Select
   
   'Check the user agent for signs they are on a mobile browser
   If InStr(strUserAgent,"iPhone") Then
      IsMobile = True
   ElseIf Instr(strUserAgent,"Android") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Windows Phone") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"BlackBerry") Then
      IsMobile = True
   ElseIf InStr(strUserAgent,"Nintendo") Then
      IsMobile = True 
   ElseIf InStr(strUserAgent,"PlayStation Vita") Then
      IsMobile = True
   Else
      IsMobile = False
   End If 

End Function 
%>

<%
Function IsTablet
   If InStr(strUserAgent,"Nexus 7") Then
      IsTablet = True
   ElseIf InStr(strUserAgent,"iPad") Then
      IsTablet = True
   Else
      IsTablet = False
   End If
End Function
%>


<%
Sub RequestUpdate

   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strProblem, strNotes, strStatus, strEMail
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objMessageText
   Dim strSubject, strName, strTechEMail, objTicket, arrAddresses
   
   'Get the ID for the ticket
   intID = Request.Form("ID")
   
   'Get the data about the ticket from the database
   strSQL = "SELECT DisplayName,Location,Status,Tech,Problem,Notes,EMail,Custom1,Custom2" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "Where ID=" & intID
   Set objTicket = Application("Connection").Execute(strSQL)

   'Assign the ticket's data to variables
   strName = objTicket(0)
   strLocation = objTicket(1)
   strStatus = objTicket(2)
   strTech = objTicket(3)
   strProblem = objTicket(4)
   strNotes = objTicket(5)
   strUserEMail = objTicket(6)
   strCustom1 = objTicket(7)
   strCustom2 = objTicket(8)
   strCurrentUser = GetFirstandLastName(strUser)
  
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   'Get the message from the database
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Request for Update'"
   Set objMessageText = Application("Connection").Execute(strSQL)
   
   'Merge the data into the message
   strMessage = objMessageText(1)
   strMessage = Replace(strMessage,"#TICKET#",intID)
   strMessage = Replace(strMessage,"#USER#",strName)
   strMessage = Replace(strMessage,"#LOCATION#",strLocation)
   strMessage = Replace(strMessage,"#STATUS#",strStatus)
   strMessage = Replace(strMessage,"#TECH#",strTech)
   strMessage = Replace(strMessage,"#PROBLEM#",strProblem)
   If Not IsNull(strNotes) Then
      strMessage = Replace(strMessage,"#NOTES#",strNotes)
   Else
      strMessage = Replace(strMessage,"#NOTES#","")
   End If
   strMessage = Replace(strMessage,"#USEREMAIL#",strUserEMail)
   strMessage = Replace(strMessage,"#CUSTOM1#",strCustom1)
   strMessage = Replace(strMessage,"#CUSTOM2#",strCustom2)
   strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   'Set the subject and merge data
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   'Get the email address from the tech who is assigned the ticket
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
   Set objTechSet = Application("Connection").Execute(strSQL)
   
   'If a tech is found send them the email, if not send it to all admins
   If objTechSet.EOF Then
      strTechEMail = Application("AdminEMail")
   Else
      strTechEMail = objTechSet(0)
   End If
   
   'If this is a list of emails then split the emails
   arrAddresses = Split(strTechEMail,";")
   
   'Send the email to each address in the array
   For Each strEMail in arrAddresses
      If strTechEMail <> "" Then 
         With objMessage
            .To = strTechEMail
            .From = strUserEMail 
            '.CC = strUserEMail
            .Subject = strSubject
            .TextBody = strMessage
            If Application("BCC") <> "" Then
               .BCC = Application("BCC")
            End If
            .Send
         End With
      End If
   Next
   
   'Build the SQL string that will write to the database who is requesting the update.
   strSQL = "INSERT INTO Tracking (Ticket,Type,TrackedBy)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Request','" & strUser & "')" & vbCRLF
   Application("Connection").Execute(strSQL)
   
   'Build the SQL string that will update the log saying who requested the update
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Update Requested','" & strUser & "','" & strTech & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   'Close objects
   Set objConf = Nothing
   Set objMessage = Nothing

End Sub%> 

<%
Sub CancelUpdateRequest

   Dim intID, strSQL

   intID = Request.Form("ID")

   'Build the SQL string that will remove from the database who is requesting the update.
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE (Ticket=" & intID & " And TrackedBy='" & strUser & "' And Type='Request')"
   Application("Connection").Execute(strSQL)   
   
   'Build the SQL string that will update the log saying the update request has been answered
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Cancelled Update Request','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
End Sub
%>

<%
Sub TrackTicket  

   Dim intID, strSQL

   intID = Request.Form("ID")
  
   'Build the SQL string that will write to the database who is requesting the update.
   strSQL = "INSERT INTO Tracking (Ticket,Type,TrackedBy)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Track','" & strUser & "')" & vbCRLF

   Application("Connection").Execute(strSQL)
  
   'Update the log
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Ticket Tracked','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)

End Sub
%> 

<%
Sub DontTrackTicket

   Dim intID, strSQL

   intID = Request.Form("ID")
  
   'Build the SQL string that will remove from the database who is requesting the update.
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE (Ticket=" & intID & " And TrackedBy='" & strUser & "' And Type='Track')"
   Application("Connection").Execute(strSQL)
  
   'Update the log
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Ticket Not Tracked','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)

End Sub
%> 

<%
Sub CloseTicket
   
   Dim intID, strSQL, objTicket, strNotes, strCategory, strOpenTime
   Dim strDate, strTime, strSubmitDate, strSubmitTime
   
   'Get the current date and time
   strDate = Date()
   strTime = Time()
   
   'Get the ID of the ticket we need to close
   intID = Request.Form("ID")
   
   'Get the info on the ticket
   strSQL = "SELECT Notes,Category,Tech,SubmitDate,SubmitTime,Status "
   strSQL = strSQL & "FROM Main WHERE ID=" & intID
   Set objTicket = Application("Connection").Execute(strSQL)
   
   'Update the notes to write back the datebase.
   If objTicket(0) <> "" Then
      strNotes = objTicket(0) & vbCRLF & vbCRLF & "Ticket closed by user."
   Else
      strNotes = "Ticket closed by user."
   End If
   
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Notes Updated','" & strUser & "','" & "" & "','" & "" & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   'Update the category if no category is set
   If IsNull(objTicket(1)) or objTicket(1) = " " Then
      strCategory = "Closed By User"
      
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'Category Changed','" & strUser & "',' ','" & strCategory & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)    
      
   Else
      strCategory = objTicket(1)
   End If
      
   'Stop all tracking of this ticket
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE Ticket=" & intID
   Application("Connection").Execute(strSQL)
   
   'Get the amount of time the ticket was open
   strSubmitDate = objTicket(3)
   strSubmitTime = objTicket(4)
   strOpenTime = DateDiff("n",strSubmitDate,strDate)
   strOpenTime = strOpenTime + DateDiff("n",strSubmitTime,strTime)
   
   'Update the log
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Status Changed','" & strUser & "','" & objTicket(5) & "','Complete','" & strDate & "','" & strTime & "');"
   Application("Connection").Execute(strSQL)
   
   'Close the ticket
   strSQL = "UPDATE Main SET TicketViewed=True,Status='Complete',Notes='" & Replace(strNotes,"'","''")
   strSQL = strSQL & "',Category='" & strCategory & "',LastUpdatedDate=#" & strDate & "#,"
   strSQL = strSQL & "LastUpdatedTime=#" & strTime & "#,OpenTime='" & strOpenTime
   strSQL = strSQL & "' WHERE ID=" & intID
   Application("Connection").Execute(strSQL)
   
   'Let the tech know the ticket was closed by the user.
   If objTicket(2) <> "" or IsNull(objTicket(2)) Then
      UpdateTech
   End If
   
End Sub
%>

<%
Sub UpdateTech
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strProblem, strNotes, strStatus, strEMail
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objMessageText
   Dim strSubject, strName, strTechEMail, objTicket, arrAddresses
   
   'Get the ID for the ticket
   intID = Request.Form("ID")
   
   'Get the data about the ticket from the database
   strSQL = "SELECT DisplayName,Location,Status,Tech,Problem,Notes,EMail,Custom1,Custom2" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "Where ID=" & intID
   Set objTicket = Application("Connection").Execute(strSQL)

   'Assign the ticket's data to variables
   strName = objTicket(0)
   strLocation = objTicket(1)
   strStatus = objTicket(2)
   strTech = objTicket(3)
   strProblem = objTicket(4)
   strNotes = objTicket(5)
   strUserEMail = objTicket(6)
   strCustom1 = objTicket(7)
   strCustom2 = objTicket(8)
   strCurrentUser = GetFirstandLastName(strUser)
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strCurrentUser = "Heat Help Desk"
      Case "TPERKINS"
         strCurrentUser = "Tech Services"
      Case Else
         strCurrentUser = GetFirstandLastName(strUser)
   End Select

   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration

   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   'Get the Tech's email address
   strSQL = "Select Tech.EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where ((Tech.Tech)=""" & strTech & """);"
   Set objTechSet = Application("Connection").Execute(strSQL)
   If objTechSet.EOF Then
      strTechEmail = ""
      Exit Sub
   Else
      strTechEmail = objTechSet(0)
   End If

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Ticket Closed By User'"
   Set objMessageText = Application("Connection").Execute(strSQL)
   
   strMessage = objMessageText(1)
   strMessage = Replace(strMessage,"#TICKET#",intID)
   strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
   strMessage = Replace(strMessage,"#USER#",strName)
   strMessage = Replace(strMessage,"#TECH#",strTech)
   strMessage = Replace(strMessage,"#STATUS#",strStatus)
   strMessage = Replace(strMessage,"#USEREMAIL#",strUserEMail)
   strMessage = Replace(strMessage,"#LOCATION#",strLocation)
   strMessage = Replace(strMessage,"#CUSTOM1#",strCustom1)
   strMessage = Replace(strMessage,"#CUSTOM2#",strCustom2)
   strMessage = Replace(strMessage,"#PROBLEM#",strProblem)
   If Not IsNull(strNotes) Then
      strMessage = Replace(strMessage,"#NOTES#",strNotes)
   Else
      strMessage = Replace(strMessage,"#NOTES#","")
   End If
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   If strTechEmail <> "" Then
      With objMessage
         .To = strTechEmail
         .From = Application("SendFromEMail") 
         .Subject = strSubject
         .TextBody = strMessage
         If Application("BCC") <> "" Then
            .BCC = Application("BCC")
         End If
         .Send
      End With
   End If
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub
%> 

<%
Function FixURLs(strMessage, intPosition)

   Dim intStartURL, intEndURL, strURL, strNewMessage
   
   If InStr(intPosition,LCase(strMessage),"http") Then
      intStartURL = InStr(intPosition,LCase(strMessage),"http")
      If InStr(InStr(intPosition,LCase(strMessage),"http"),LCase(strMessage)," ") < InStr(InStr(intPosition,LCase(strMessage),"http"),LCase(strMessage),vbCRLF) Then
         intEndURL = InStr(InStr(intPosition,LCase(strMessage),"http"),LCase(strMessage)," ")
      Else
         intEndURL = InStr(InStr(intPosition,LCase(strMessage),"http"),LCase(strMessage),vbCRLF)
      End If
      If intEndURL <> 0 Then
         strURL = Mid(strMessage,intStartURL,intEndURL - intStartURL)
      Else
         strURL = Mid(strMessage,intStartURL)
      End If
   End If

   'If the link has spaces in it then something went wrong, don't turn it into a link
   If InStr(strURL," ") > 0 Then
      strNewMessage = strMessage
   Else
      strNewMessage = Replace(strMessage,strURL,"<a href=""" & Replace(strURL,vbCRLF,"") & """>Link</a>")
   End If
   
   'Remove iframes
   strNewMessage = Replace(strNewMessage,"<iframe","")
   strNewMessage = Replace(strNewMessage,"</iframe>","")
   
   If intEndURL <> 0 Then
      If InStr(intEndURL,LCase(strNewMessage),"http") Then
         strNewMessage = FixURLs(strNewMessage,intEndURL)
      End If
   End If
   
   FixURLs = strNewMessage
   
End Function

Function HideText(strMessage)
   If InStr(strMessage,"----" & vbCRLF) Then
      strMessage = Left(strMessage,(InStr(strMessage,vbCRLF & "----" & vbCRLF))-1)
   End If
   HideText = strMessage
End Function
%>

<%
Function GetFirstandLastName(strUserName)

   On Error Resume Next

   Dim objConnection, objCommand, objRootDSE, objRecordSet,strDNSDomain

   If Application("UseAD") Then
      'Create objects required to connect to AD
      Set objConnection = CreateObject("ADODB.Connection")
      Set objCommand = CreateObject("ADODB.Command")
      Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

      'Create a connection to AD
      objConnection.Provider = "ADSDSOObject"

      objConnection.Open "Active Directory Provider", Application("ADUsername"), Application("ADPassword")
      objCommand.ActiveConnection = objConnection
      strDNSDomain = objRootDSE.Get("DefaultNamingContext")
      objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); GivenName,SN,name ;subtree"

      'Initiate the LDAP query and return results to a RecordSet object.
      Set objRecordSet = objCommand.Execute

      If NOT objRecordSet.EOF Then
         If objRecordSet(0) = "" Then
            GetFirstandLastName = strUserName
         Else
            GetFirstandLastName = objRecordSet(0) & " " & objRecordSet(1)
         End If
      Else
         GetFirstandLastName= strUserName
      End If

   Else
      GetFirstandLastName= strUserName
   End If

End Function
%>

<%
Function HideText(strMessage)
   If InStr(strMessage,"----" & vbCRLF) Then
      strMessage = Left(strMessage,(InStr(strMessage,"----" & vbCRLF))-1)
   End If
   HideText = strMessage
End Function
%>

<%
Function BuildReturnLink(bolIncludeID)

   Dim strLinkID, strLinkLocation, strLinkStatus, strLinkTech, strLinkCategory, strLinkUser, strLinkFilter, strLinkProblem
   Dim strLinkNotes, strLinkEMail, strLinkSort, strLinkDays, strLinkBack, strLinkViewed

   'Build the return link
   If bolIncludeID Then
      strLinkID = Request.QueryString("ID")
   End If
   strLinkLocation = Request.QueryString("Location")
   strLinkStatus = Request.QueryString("Status")
   strLinkTech = Request.QueryString("Tech")
   strLinkCategory = Request.QueryString("Category")
   strLinkUser = Request.QueryString("User")
   strLinkFilter = Request.QueryString("Filter")
   strLinkProblem = Request.QueryString("Problem")
   strLinkNotes = Request.QueryString("Notes")
   strLinkEMail = Request.QueryString("EMail")
   strLinkSort = Request.QueryString("Sort")
   strLinkDays = Request.QueryString("Days")
   strLinkBack = Request.QueryString("Back")
   strLinkViewed = Request.QueryString("Viewed")

   If strLinkID <> "" Then
      BuildReturnLink = BuildReturnLink & "&ID=" & Replace(strLinkID," ","%20")
   End If   
   If strLinkLocation <> "" Then
      BuildReturnLink = BuildReturnLink & "&Location=" & Replace(strLinkLocation," ","%20")
   End If
   If strLinkStatus <> "" Then
      BuildReturnLink = BuildReturnLink & "&Status=" & Replace(strLinkStatus," ","%20")
   End If
   If strLinkTech <> "" Then
      BuildReturnLink = BuildReturnLink & "&Tech=" & Replace(strLinkTech," ","%20")
   End If
   If strLinkCategory <> "" Then
      BuildReturnLink = BuildReturnLink & "&Category=" & Replace(strLinkCategory," ","%20")
   End If
   If strLinkUser <> "" Then
      BuildReturnLink = BuildReturnLink & "&User=" & Replace(strLinkUser," ","%20")
   End If
   If strLinkFilter <> "" Then
      BuildReturnLink = BuildReturnLink & "&Filter=" & Replace(strLinkFilter," ","%20")
   End If
   If strLinkProblem <> "" Then
      BuildReturnLink = BuildReturnLink & "&Problem=" & Replace(strLinkProblem," ","%20")
   End If
   If strLinkNotes <> "" Then
      BuildReturnLink = BuildReturnLink & "&Notes=" & Replace(strLinkNotes," ","%20")
   End If
   If strLinkEMail <> "" Then
      BuildReturnLink = BuildReturnLink & "&EMail=" & Replace(strLinkEMail," ","%20")
   End If
   If strLinkSort <> "" Then
      BuildReturnLink = BuildReturnLink & "&Sort=" & Replace(strLinkSort," ","%20")
   End If
   If strLinkBack <> "" Then
      BuildReturnLink = BuildReturnLink & "&Back=Yes"
   End If
   If strLinkDays <> "" Then
      BuildReturnLink = BuildReturnLink & "&Days=" & strLinkDays
   End If
   If strLinkViewed <> "" Then
      BuildReturnLink = BuildReturnLink & "&Viewed=" & strLinkViewed
   End If
   
   If BuildReturnLink <> "" Then
      BuildReturnLink = "?" & Right(BuildReturnLink,(Len(BuildReturnLink)-1))
   End If

End Function
%>

<%
Function GetUser

   Const USERNAME = 1

   Dim strUserAgent, strSessionID, objSessionLookup
   
   'Get some needed data
   strSessionID = Request.Cookies("SessionID")
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Send them to the logon screen if they don't have a Session ID
   If strSessionID = "" Then
      SendToLogonScreen

   'Get the username from the database
   Else
   
      strSQL = "SELECT ID,UserName,SessionID,IPAddress,UserAgent,ExpirationDate FROM Sessions "
      strSQL = strSQL & "WHERE UserAgent='" & Left(Replace(strUserAgent,"'","''"),250) & "' And SessionID='" & Replace(strSessionID,"'","''") & "'"
      strSQL = strSQL & " And ExpirationDate > Date()"
      Set objSessionLookup = Application("Connection").Execute(strSQL)
      
      'If a session isn't found kick them out
      If objSessionLookup.EOF Then
         SendToLogonScreen
      Else
         GetUser = objSessionLookup(USERNAME)
      End If
   End If  
   
End Function
%>

<%
Sub SendToLogonScreen

   Dim strReturnLink, strSourcePage
      
   'Build the return link before sending them away.
   strReturnLink = BuildReturnLink(True)
   strSourcePage = Request.ServerVariables("SCRIPT_NAME")
   strSourcePage = Right(strSourcePage,Len(strSourcePage) - InStrRev(strSourcePage,"/"))
   If strReturnLink = "" Then
      strReturnLink = "?SourcePage=" & strSourcePage
   Else
      strReturnLink = strReturnLink & "&SourcePage=" & strSourcePage
   End If
   
   Response.Redirect("login.asp" & strReturnLink)
   
End Sub 
%>

<%Sub VersionProblem %>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
      <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
      <meta name="viewport" content="width=device-width" />
   </head>
   <body>
      <center><b>Web and Database versions don't match</b></center>
   </body>
   </html>

<%
   Response.End

End Sub%>
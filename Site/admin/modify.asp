<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/13/04
'Last Updated 6/16/14

'This page will display a form the the help desk administrator allowing them to update the
'status of a help desk ticket.  It contains both the main form and the successfully submit
'page.  Once the user submits the page all the fields are checked, if they are correct then
'the success page is displayed and the database is updated

Option Explicit

On Error Resume Next

Dim intID, strUserTemp, strEMailTemp, strLocation, strCategory, strProblem, strNotes, bolUpdated
Dim strStatus, strTech, strSQL, objRecordSet, strCustom1Temp, strCustom2Temp, bolUserUpdated
Dim bolCallClosed, strTechEmail, bolUpdateRequested, bolTicketReOpened, strReturnLink, bolLog
Dim strLinkID, strLinkLocation, strLinkStatus, strLinkTech, strLinkCategory, strLinkUser, strLinkFilter
Dim strLinkProblem, strLinkNotes, strLinkEMail, strLinkSort, strLinkBack, Upload, strCMD, objLog
Dim strShowLog, bolShowLogButton, objNetwork, bolTrackTicket, bolTrackTicketOff, bolMissingTech
Dim strLinkViewed, objNameCheckSet, bolTechUpdated, bolKeepData, strLinkDays, strUserName, bolShowLogout
Dim intTest, strType, intCount, strDate, strTime, strDays, objUpdateRequest, bolCancelledUpdateRequest
Dim strMinutes, strHours, strTimeActive, objCategorySet, objTechSet, objStatusSet, strAttachTech
Dim objLocationSet, objFSO, objFolder, objFile, strUserAgent, strAttachmentFolder, bolMobileVersion
Dim objTracking, strTaskTime, bolTicketViewed, strRole, bolTracking, bolRequest, strEMail, strUser
Dim intZoom

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

'Find the current users role
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)
bolMobileVersion = objNameCheckSet(4)

'See if the user has the rights to visit this page
If objNameCheckSet(2) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      AccessDenied
   Else
      AccessGranted
   End If
Else
   AccessDenied
End If

Sub AccessGranted

   Dim strDBStatus

   bolTechUpdated = False
   bolUserUpdated = False
   bolUpdated = False
   bolCallClosed = False
   bolLog = False
   bolKeepData = False

   Set Upload = New FreeASPUpload
   Upload.Save(Application("FileLocation"))

   'Build the return link
   strReturnLink = BuildReturnLink(False)
   strLinkBack = Request.QueryString("Back")
   strLinkID = Request.QueryString("ID")

   intID = request.querystring("ID")
   strUserTemp = Upload.Form("Name")
   strEMailTemp = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1Temp = Upload.Form("Custom1")
   strCustom2Temp = Upload.Form("Custom2")
   
   If strStatus = "Auto Assigned" or strStatus = "New Assignment" Then
      strStatus = "In Progress"
   End If

   strShowLog = Request.QueryString("ShowLog")
   If strShowLog = "Yes" Then
      bolLog = True
   End If
   
   'Get the log items for this ticket
   strSQL = "SELECT * FROM Log WHERE Ticket=" & intID & " ORDER BY ID"
   Set objLog = Application("Connection").Execute(strSQL)

   If Not objLog.EOF Then
      bolShowLogButton = True
   Else
      bolShowLogButton = False
   End If
  
   strSQL = "Select Status" & vbCRLF
   strSQL = strSQL & "From Main" & vbCRLF
   strSQL = strSQL & "Where ID=" & intID
         
   Set objRecordSet = Application("Connection").Execute(strSQL)
   If Not objRecordSet.BOF or Not objRecordSet.EOF Then
      strDBStatus = objRecordSet(0)
   Else
      strDBStatus = ""
   End If
   
   'Check and see if we are sending and email to someone.
   If Upload.Form("cmdEMail") <> "" Then
      strEMail = Upload.Form("SendEMail")
      If IsEmailValid(strEMail) Then
         SendTicket
      End If
   End If

   'Make sure all the fields were filled out
   If (strUserTemp = "" Or strEMailTemp = "" Or strLocation = "" Or strCategory = " " Or strStatus = "New Assignment" Or strStatus = "Auto Assigned" Or strTech = "" Or (strStatus = "Complete" And strDBStatus = "Complete")) And (Upload.Form("cmdSubmit") <> "") Then
      Select Case Upload.Form("cmdSubmit")
         Case "Save"
            strCMD = "Save"
            Call Main()
         Case "Open Ticket"
       
            'Update the log
            strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
            strSQL = strSQL & "VALUES (" & intID & ",'Ticket Reopened','" & strUser & "','" & Date() & "','" & Time() & "');"
            Application("Connection").Execute(strSQL)
            
            strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,NewValue,UpdateDate,UpdateTime)"
            strSQL = strSQL & "VALUES (" & intID & ",'Tech Reassigned','" & strUser & "','" & strTech & "','" & Date() & "','" & Time() & "');"
            Application("Connection").Execute(strSQL)

            bolTicketReOpened = True
            Call Submit
         Case "Update Tech"    
            bolKeepData = True
            bolTechUpdated = True
            UpdateTech
            Call Main
         Case "Update User"
            bolKeepData = True
            bolUserUpdated = True
            UpdateUser()
            Call Main
         Case "Request Update"
            strTech = Upload.Form("Tech")
            If strTech <> "" Then
               bolUpdateRequested = True
               Call RequestUpdate
            Else
               bolMissingTech = True
            End If
            Call Main
         Case "Cancel Update Request"
            bolCancelledUpdateRequest = True
            CancelUpdateRequest
            Call Main
         Case "Track Ticket"
            bolTrackTicket = True
            Call TrackTicket()
            Call Main
         Case "Stop Tracking"
            bolTrackTicketOff = True
            Call DontTrackTicket()
            Call Main
         Case "Show Log"
            bolLog = True
            bolKeepData = True
            Call Main
         Case "Hide Log"
            bolLog = False
            bolKeepData = True
            Call Main
         Case "User History"
            Response.Redirect "view.asp?User=" & strUserTemp
         Case Else
            strCMD = "Save"
            Call Main()
      End Select
   Else
      Select Case Upload.Form("cmdSubmit")
         Case  "Save"
            strCMD = "Save"
            Call Submit()
         Case "Update Tech"
            bolKeepData = True
            bolTechUpdated = True
            Call Main
            UpdateTech
         Case "Update User"
            bolKeepData = True
            bolUserUpdated = True
            Call Main
            UpdateUser()
         Case "Request Update"
            strTech = Upload.Form("Tech")
            bolUpdateRequested = True
            Call RequestUpdate
            Call Submit
         Case "Cancel Update Request"
            bolCancelledUpdateRequest = True
            CancelUpdateRequest
            Call Main
         Case "Track Ticket"
            bolTrackTicket = True
            Call TrackTicket()
            Call Main
         Case "Stop Tracking"
            bolTrackTicketOff = True
            Call DontTrackTicket()
            Call Main
         Case "Show Log"
            bolLog = True
            bolKeepData = True
            Call Main
         Case "Hide Log"
            bolLog = False
            bolKeepData = True
            Call Main
         Case "User History"
            Response.Redirect "view.asp?User=" & strUserTemp
         Case Else
            Call Main()
      End Select
   End If
End Sub

Sub Main()
   
   'This is what the user first sees when they come to this page or if they entered an invalid
   'value on the page.  It will get the data from the database and display it to the user.  If
   'this page has already been submitted and there were errors it will show you what fields 
   'have errors.
   
   Dim strDisplayName
   
   On Error Resume Next
   
   Const ID = 0
   Const Name = 1
   Const Location = 2
   Const EMail = 3
   Const Problem = 4
   Const SubmitDate = 5
   Const SubmitTime = 6
   Const Notes = 7
   Const Status = 8
   Const Category = 9
   Const Tech = 10
   Const LastUpdatedDate = 11
   Const LastUpdatedTime = 12
   Const Custom1 = 13
   Const Custom2 = 14
   Const TicketViewed = 15
   
   intID = request.querystring("ID")
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
   
   'Set the zoom level
   If Request.Cookies("ZoomLevel") = "ZoomIn" Then
      If InStr(strUserAgent,"Silk") Then
         intZoom = 1.4
      Else
         intZoom = 1.9
      End If
   End If
   
   'Verify that the intID is a number.  If not then the user enter a non numeric value in for 
   'a ticket number.  If that is the case then set intID to 0 so it will kick out as an error
   'to the user.
   intTest = CInt(intID)
   strType = TypeName(intTest)
   If UCase(strType) <> "INTEGER" Then
      intID = "0"
   End If
   
   strShowLog = Request.QueryString("ShowLog")
   If strShowLog = "Yes" Then
      bolLog = True
   End If
   
   'Get the log items for this ticket
   strSQL = "SELECT * FROM Log WHERE Ticket=" & intID & " ORDER BY ID"
   Set objLog = Application("Connection").Execute(strSQL)

   'Build the SQL string that will get the data for the requested ticket
   strSQL = "SELECT ID,DisplayName,Location,Email,Problem,SubmitDate,SubmitTime,Notes,Status,Category,Tech,LastUpdatedDate,LastUpdatedTime,Custom1,Custom2,TicketViewed" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   
   'Execute the SQL string and assign the results to a Record Set
   Set objRecordSet = Application("Connection").Execute(strSQL)
   
   If Not objRecordSet.EOF Then
      bolTicketViewed = objRecordSet(TicketViewed) 
   End If
   
   'Get the tech's email address if they were sent an update.
   If bolTechUpdated Then
      'Get the tech's email address
      strSQL = "Select Tech.EMail" & vbCRLF
      strSQL = strSQL & "From Tech" & vbCRLF
      strSQL = strSQL & "Where ((Tech.Tech)=""" & strTech & """);"

      Set objTechSet = Application("Connection").Execute(strSQL)
      strTechEmail = objTechSet(0)
   End If
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strDisplayName = "Heat Help Desk"
      Case "TPERKINS"
         strDisplayName = "Tech Services"
      Case Else
         strDisplayName = GetFirstandLastName(strUser)
   End Select
   
   If Not objRecordSet.EOF Then
      If strDisplayName = objRecordSet(Tech) And Not bolTicketViewed Then
         strSQL = "UPDATE Main SET TicketViewed=True WHERE ID=" & intID
         Application("Connection").Execute(strSQL)
         bolTicketViewed = True
         
         strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
         strSQL = strSQL & "VALUES (" & intID & ",'Ticket Viewed','" & strUser & "','" & Date() & "','" & Time() & "');"
         Application("Connection").Execute(strSQL)
      End If
   End If
      
   'Determin if the call is closed or has never been modified.  If so then get the date and time
   'from the system.  Otherwise get the last update date and time from the database.  This will
   'be used to calculate how log a call has been open.
   If (objRecordSet(Status) <> "Complete") Or (objRecordSet(LastUpdatedDate) = "6/16/1978") Then
      strDate = Date
      strTime = Time
   Else
      strDate = objRecordSet(LastUpdatedDate)
      strTime = objRecordSet(LastUpdatedTime)
   End If            
   
   'Calculate how long a call has been open
   strDays = DateDiff("d",objRecordSet("SubmitDate"),strDate)
   strMinutes = DateDiff("n",objRecordSet("SubmitTime"),strTime)
   strHours = (strMinutes / 60)
   strMinutes = strMinutes Mod 60
   If Sgn(strHours) = -1 Then
      strHours = (24 + strHours)
      strDays = strDays - 1
   End If
   If Sgn(strMinutes) = -1 Then
      strMinutes = 60 + strMinutes
   End If
   strTimeActive = strDays & "d " & Int(strHours) & "h " & strMinutes & "m" 
   
   'Verify that at least one ticket was returned.  There should never be more then one.
   'If no tickets are returned then an invalid ticket number was given.
   intCount = 0
   Do  Until objRecordSet.EOF
      intCount = intCount + 1
      objRecordSet.MoveNext
   Loop   
   objRecordSet.MoveFirst
   
   'If only one ticket is returned then display it's information
   If intCount = 1 Then
   
      'Build the SQL string and execute it to populate the category pulldown list
      strSQL = "Select Category.Category" & vbCRLF
      strSQL = strSQL & "From Category" & vbCRLF
      strSQL = strSQL & "Where (((Category.Active)=Yes))" & vbCRLF
      strSQL = strSQL & "Order By Category.Category;"
      Set objCategorySet = Application("Connection").Execute(strSQL)
      
      'Build the SQL string and execute it to populate the tech pulldown list
      strSQL = "Select Tech" & vbCRLF
      strSQL = strSQL & "From Tech" & vbCRLF
      strSQL = strSQL & "Where Active=Yes And UserLevel<>'Data Viewer'" & vbCRLF
      strSQL = strSQL & "Order By Tech.Tech;"
      Set objTechSet = Application("Connection").Execute(strSQL)
      
      'Build the SQL string and execute it to populate the status pulldown list
      strSQL = "Select Status.Status" & vbCRLF
      strSQL = strSQL & "From Status" & vbCRLF
      Set objStatusSet = Application("Connection").Execute(strSQL)
      
      'Build the SQL string and execute it to populate the location pulldown list
      strSQL = "Select Location.Location" & vbCRLF
      strSQL = strSQL & "From Location" & vbCRLF
      strSQL = strSQL & "Where (((Location.Active)=Yes))" & vbCRLF
      strSQL = strSQL & "Order By Location.Location;"
      Set objLocationSet = Application("Connection").Execute(strSQL)
      
      strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
      
      'Find out if the current user is either tracking the ticket or has requested an update
      strSQL = "SELECT Type FROM Tracking WHERE Ticket=" & intID & " And TrackedBy='" & strUser & "'"
      Set objTracking = Application("Connection").Execute(strSQL)
      bolTracking = False
      bolRequest = False
      Do Until objTracking.EOF
         Select Case objTracking(0)
            Case "Track"
               bolTracking = True
            Case "Request"
               bolRequest = True
         End Select
         objTracking.MoveNext
      Loop
      
      'See if anyone has requested an updated on this ticket
      strSQL = "SELECT TrackedBy FROM Tracking WHERE Ticket=" & intID & " And Type='Request'"
      Set objUpdateRequest = Application("Connection").Execute(strSQL)

      If IsMobile Then
         MobileVersion
      ElseIf IsWatch Then
         WatchVersion
      Else
         MainVersion
      End If   
      
   Else%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <head>
      <title>HDL - Admin - <%=intID%></title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
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
      Invalid Ticket Number entered.  <a href="index.asp">Go Back</a>
   </body>
<% End If
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
      Case "watch"
         IsMobile = False
         Response.Cookies("SiteVersion") = "Watch"
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
      Case "watch"
         IsMobile = False
         Exit Function
   End Select
   
   'It's not mobile if the mobile version is turned off.
   If Not bolMobileVersion Then
      IsMobile = False
      Exit Function
   End If

   'It's not mobile if it's an Android watch
   If Instr(strUserAgent,"Android") > 0 And InStr(strUserAgent,"Watch") > 0 Then
      IsMobile = False
      Exit Function
   End If

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
Function IsWatch

   'Check and see if it's an Android Watch
   If Instr(strUserAgent,"Android") > 0 And InStr(strUserAgent,"Watch") > 0 Then
      IsWatch = True
   Else
   
      'Choose the site based on the cookie
      Select Case LCase(Request.Cookies("SiteVersion"))
         Case "watch"
            IsWatch = True
         Case Else
            IsWatch = False
      End Select
   
   End If

End Function 
%>

<%Sub MainVersion 

   Const ID = 0
   Const Name = 1
   Const Location = 2
   Const EMail = 3
   Const Problem = 4
   Const SubmitDate = 5
   Const SubmitTime = 6
   Const Notes = 7
   Const Status = 8
   Const Category = 9
   Const Tech = 10
   Const LastUpdatedDate = 11
   Const LastUpdatedTime = 12
   Const Custom1 = 13
   Const Custom2 = 14
   Const TicketViewed = 15
   
   Dim intInputSize

   If InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 29
   Else
      intInputSize = 36
   End If
   
   If strStatus = "Auto Assigned" or strStatus = "New Assignment" Then
      strStatus = "In Progress"
   End If
   
%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <head>
      <title>HDL - Admin - <%=intID%></title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
         <link rel="stylesheet" type="text/css" href="../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
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
<% If strReturnLink = "" Then %>
      <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>">
<% Else %>
      <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>&<%=Right(strReturnLink,Len(strReturnLink)-1)%>">
<% End If %>
      <div align="center">
   	<table border="0" width="750" cellspacing="0" cellpadding="0">
   		<tr>
   			<td width="204" valign="bottom">
   			<p style="margin-top: 0; margin-bottom: 0">
   			Open for <%=strTimeActive%></td>
   			<td width="311" valign="bottom">
   			<p style="margin-top: 0; margin-bottom: 0" align="center">
         <% If objRecordSet(Status) = "Complete" Then %>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/closed.gif" width="20" height="20">
            <% Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/closed.gif" width="20" height="20">
            <% End If %>
         <% End If %>            
<%       If bolTicketViewed Then %>   
<%          If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/viewed.gif" alt="Viewed by Tech" width="20" height="20">
<%          Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/viewed.gif" alt="Viewed by Tech" width="20" height="20">
<%          End If %>
<%       Else %>
<%          If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/notviewed.gif" alt="Not Viewed by Tech" width="20" height="20">
<%          Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/notviewed.gif" alt="Not Viewed by Tech" width="20" height="20">
<%          End If %>            
<%       End If %>

            Ticket #<%=intID%></b><font face="Arial">
         		</td>
   			<td width="235" valign="bottom">
   			<p style="margin-top: 0; margin-bottom: 0" align="right">
            Version <%=Application("Version")%>
   			</td>
   		</tr>
         <tr>
            <td colspan="3"><center>
<%          If strLinkBack = "Yes" Then %>
               <a href="view.asp<%=strReturnLink%>#<%=strLinkID%>">Back</a> |
<%          End If %>
               <a href="index.asp">Home</a> |
               <a href="view.asp?Filter=AllOpenTickets">Open Tickets</a> | 
            <% If strRole <> "Data Viewer" Then %>   
               <a href="view.asp?Filter=MyOpenTickets">Your Tickets</a> | 
            <% End If %> 
            <% If Application("UseTaskList") And objNameCheckSet(5) <> "Deny" Then %>
               <a class="linkbar" href="tasklist.asp">Tasks</a> | 
            <% End If %>
            <% If Application("UseStats") Then %>
               <a href="stats.asp">Stats</a> | 
            <% End If %>
            <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
               <a class="linkbar" href="docs.asp">Docs</a> | 
            <% End If %>
               <a href="settings.asp">Settings</a>
            <% If objNameCheckSet(1) = "Administrator" Then %>
               | <a href="setup.asp">Admin Mode</a>
            <% End If %> 
            <% If bolShowLogout Then %>
               | <a href="login.asp?action=logout">Log Out</a>
            <% End If %>
               </center>
            </td>
         </tr>
   		<tr>
   			<td width="750" colspan="3">
   			<p style="margin-top: 0; margin-bottom: 0">

   <%    'Display any error messages if the form is missing anything
         If (strCMD = "Save") And (strUserTemp = "" Or strEMailTemp = "" Or strLocation = "" Or strCategory = " " Or strTech = "") Then %>
            <font class="missing">Please fill out highlighted fields...</font>        
<%          bolUpdated = False
         End If
         If (strCMD = "Save") And strStatus = "New Assignment" Then %>
            <font class="missing">Status cannot be "New Assignment"</font>
<%          bolUpdated = False
         End If 
         If (strCMD = "Save") And strStatus = "Auto Assigned" Then %>
            <font class="missing">Status cannot be "Auto Assigned"</font>
<%          bolUpdated = False
         End If 
         If bolUserUpdated Then %>
            <font class="information">EMail Sent to <%=objRecordSet(EMail)%></font>
<%          bolUpdated = False
         End If
         If bolTechUpdated Then %>
            <font class="information">EMail Sent to <%=strTechEmail%></font>
<%          bolUpdated = False
         End If 
         If strCMD = "Save" And bolUpdated And Not bolCallClosed And Not bolTicketReOpened Then 
            If strTechEmail <> "" Then %>
               <font class="information">Ticket Updated - EMail Sent to <%=strTechEmail%></font>
<%          Else %>
               <font class="information">Ticket Updated</font>
<%          End If
         End If
         If strCMD = "Save"  And bolUpdated And bolCallClosed Then %>
            <font class="information">Ticket Closed - EMail Sent to <%=objRecordSet(EMail)%></font></b>
<%       End If  
         If bolUpdateRequested Then %>
            <font class="information">EMail Sent to <%=strTechEMail%></font>
<%       End If
         If bolMissingTech Then %>
            <font class="missing">No tech assigned...<%=strTechEMail%></font>
<%       End If 
         If bolTicketReOpened Then %>
            <font class="information">Ticket Reopened</font>
<%       End If 
         If bolTrackTicket Then %>
            <font class="information">You are now tracking this ticket</font>
<%       End If
         If bolTrackTicketOff Then %>
            <font class="information">You are no longer tracking this ticket</font>
<%       End If
         If bolCancelledUpdateRequest Then %>
            <font class="information">Request for update cancelled</font>
<%       End If %>
<%       If Upload.Form("cmdEMail") <> "" Then %>
<%          If IsEmailValid(strEMail) Then %>
               <font class="information">EMail Sent to <%=strEMail%></font>
<%          Else %>
               <font class="missing"><%=strEMail%> is an invalid address</font>
<%          End If %>
<%       End If %>
         
         &nbsp;</td>	 
   		</tr>
   		<tr>
   			<td colspan="3">
   			<table class="showborders" width="100%" cellspacing="0" cellpadding="0" id="table2">
   				<tr>

   <%             'Highlight the User label if it was blank when the form was submitted
                  If (strCMD = "Save") And (strUserTemp = "") Then %>
                     <td class="showborders" width="9%"><font class="missing">Name:</font></td>
   <%             Else %>
                     <td class="showborders" width="9%">Name:&nbsp;</td>
   <%             End If %>
   
   					<td class="showborders" width="45%">
   					

   <%             'Highlight the User label if it was blank when the form was submitted
                  If ((strCMD = "Save") And strUserTemp <> objRecordSet(Name)) or bolKeepData Then%>
                     <input type="text" name="Name" size="<%=intInputSize%>" value="<%=strUserTemp%>"></td>
   <%             Else%>
   					   <input type="text" name="Name" size="<%=intInputSize%>" value="<%=objRecordSet(Name)%>"></td>
   <%             End If%>
   					<td class="showborders" width="9%">Location:</td>
   					<td class="showborders" width="36%">
   					<select size="1" name="Location">

   <%             'If the user is visiting this page for the first time the default item in the location
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strLocation <> objRecordSet(Location)) or bolKeepData Then%>
                     <option value="<%=strLocation%>"><%=strLocation%></option>
   <%                Do Until objLocationSet.EOF
                        If strLocation = "" Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   ElseIf Trim(Ucase(strLocation)) <> Trim(Ucase(objLocationSet(0))) Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   End If
                        objLocationSet.MoveNext
                     Loop
                  Else%>
   					   <option value="<%=objRecordSet(Location)%>"><%=objRecordSet(Location)%></option>
   <%                Do Until objLocationSet.EOF
                        If objRecordSet(Location) = "" Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Location))) <> Trim(Ucase(objLocationSet(0))) Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   End If
                        objLocationSet.MoveNext
                     Loop
                  End If%>
                  
   					</select></td>
   				</tr>
   				<tr>

   <%             'Highlight the E-Mail label if it was blank when the form was submitted
                  If (strCMD = "Save") And (strEmailTemp = "") Then %>
                     <td class="showborders" width="9%"><font class="missing">E-Mail:</font></td>
   <%             Else %>
                     <td class="showborders" width="9%">EMail:</td>
   <%             End If %>
   
   					<td class="showborders" width="45%">

   <%             'If this is the first time the user is visiting the form the value will be pulled from
                  'the database.  If there was an error when the form is submitted the value that was in
                  'the box will be displayed.
                  If ((strCMD = "Save") And strEMailTemp <> objRecordSet(EMail)) or bolKeepData Then%>
                     <input type="text" name="EMail" size="<%=intInputSize%>" value="<%=strEMailTemp%>"></td>
   <%             Else%>
   					   <input type="text" name="EMail" size="<%=intInputSize%>" value="<%=objRecordSet(EMail)%>"></td>
   <%             End If%>

   <%             'Highlight the Category label if it was blank when the form was submitted
                  If (strCMD = "Save") And (strCategory = " ") Then %>
                     <td class="showborders" width="9%"><font class="missing">Category:</font></td>
   <%             Else %>
                     <td class="showborders" width="9%">Category:</td>
   <%             End If %>
   
   					<td class="showborders" width="36%">
   					<select size="1" name="Category">

   <%             'If the user is visiting this page for the first time the default item in the category
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strCategory <> objRecordSet(Category)) or bolKeepData Then%>
                     <option value="<%=strCategory%>"><%=strCategory%></option>
   <%                Do Until objCategorySet.EOF
                        If strCategory = "" Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   ElseIf Trim(Ucase(strCategory)) <> Trim(Ucase(objCategorySet(0))) Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   End If
                        objCategorySet.MoveNext
                     Loop
                  Else%>
                     <option value="<%=objRecordSet(Category)%>"><%=objRecordSet(Category)%></option>
   <%                Do Until objCategorySet.EOF
                        If objRecordSet(Category) = "" Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Category))) <> Trim(Ucase(objCategorySet(0))) Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   End If
                        objCategorySet.MoveNext
                     Loop
                  End If %>
   
   					</select></td>
   				</tr>
   				<tr>

   <%             'Highlight the Tech label if it was blank when the form was submitted
                  If (strCMD = "Save") And (strTech = "") Then %>
                     <td class="showborders" width="9%"><font class="missing">Tech:</font></td>
   <%             Else %>
                     <td class="showborders" width="9%">Tech:</td>
   <%             End If %>
   
   					<td class="showborders" >
   					<select size="1" name="Tech">
   					
   <%             'If the user is visiting this page for the first time the default item in the tech
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strTech <> objRecordSet(Tech)) or bolKeepData Then%>
   					   <option value="<%=strTech%>"><%=strTech%></option>
   <%                Do Until objTechSet.EOF
                        If strTech = "" Then%>
                           <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
   <%                   ElseIf Trim(Ucase(strTech)) <> Trim(Ucase(objTechSet(0))) Then%>
                           <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
   <%                   End If
                        objTechSet.MoveNext
                     Loop
                  Else%>
   					   <option value="<%=objRecordSet(Tech)%>"><%=objRecordSet(Tech)%></option>
   <%                Do Until objTechSet.EOF
                        If objRecordSet(Tech) = "" Then%>
                           <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Tech))) <> Trim(Ucase(objTechSet(0))) Then%>
                           <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
   <%                   End If
                        objTechSet.MoveNext
                     Loop
                  End If%>

   					</select></td>

   <%             'Highlight the Status label if it was blank when the form was submitted
                  If (strCMD = "Save") And (strStatus = "New Assignment" or strStatus = "Auto Assigned") Then%>
                     <td class="showborders" ><font class="missing">Status:</font></td>
   <%             Else%>
                     <td class="showborders" >Status:</td>
   <%             End If%>
   
   					<td class="showborders">
   					<select size="1" name="Status">

   <%             'If the user is visiting this page for the first time the default item in the Status
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strStatus <> objRecordSet(Status)) or bolKeepData Then%>
   					   <option value="<%=strStatus%>"><%=strStatus%></option>
   <%                Do Until objStatusSet.EOF
                        If strStatus = "" Then%>
                           <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   Else%>
                           <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   End If
                        objStatusSet.MoveNext
                     Loop%>
   
   <%             Else%>
   					   <option value="<%=objRecordSet(Status)%>"><%=objRecordSet(Status)%></option>
   <%                Do Until objStatusSet.EOF
                        If objRecordSet(Status) = "" Then
   %>                      <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Status))) <> Trim(Ucase(objStatusSet(0))) Then
   %>                      <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   End If
                        objStatusSet.MoveNext
                     Loop   
                  End If%>
                  
   					</select></td>
   				</tr>
	<% 		If Application("UseCustom1") or Application("UseCustom2") Then%>	
               <tr>
	<%          If Application("UseCustom1") Then%>   
                  <td class="showborders" width="9%"><%=Application("Custom1Text")%>:</td>
                  <td class="showborders" ><%=objRecordSet(Custom1)%> <input type="hidden" name="Custom1" value="<%=objRecordSet(Custom1)%>"></td>
	<%			   End If %>	
   <% 			If Application("UseCustom2") Then %>
                  <td class="showborders" width="9%"><%=Application("Custom2Text")%>:</td>
                  <td class="showborders" ><%=objRecordSet(Custom2)%><input type="hidden" name="Custom2" value="<%=objRecordSet(Custom2)%>"></td>
   <%          Else %>
					<td class="showborders" width="9%">&nbsp;</td>
					<td class="showborders" >&nbsp;</td>
   <%			   End If %>
               </tr>   				
   <%       End If %>
				
   				<tr>
   <%                Set objFSO = CreateObject("Scripting.FileSystemObject")
                     If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
                        strAttachmentFolder = Application("FileLocation") & "\" & intID
                        strAttachTech = ""
                     End If  
                     If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
                        strAttachmentFolder = Application("FileLocation") & "\" & intID & "-Tech"
                        strAttachTech = "-Tech"
                     End If
                     If objFSO.FolderExists(strAttachmentFolder) Then %>
                        <td class="showborders" colspan="2">Problem:</td>
                        <td class="showborders">Attachment:&nbsp;</td>
                        <td class="showborders">
   <%                   Set objFolder = objFSO.GetFolder(strAttachmentFolder)
                        For Each objFile in objFolder.Files
                           If UCase(objFile.Name) <> "THUMBS.DB" Then
                              Response.Write "<a href=""download.asp?folder=" & intID & strAttachTech & "&file=" & objFile.Name & """>Click Here</a>&nbsp;"
                           End If
                        Next
                     Else %>
                       
                  <%If (inStr(strUserAgent,"iPad") = False And inStr(strUserAgent,"iPhone") = False) And objRecordSet(Status) <> "Complete" Then
                     If InStr(strUserAgent,"Chrome") or InStr(strUserAgent,"Safari") Then %>
                        <td class="showborders" colspan="2">Problem:</td>
                        <td class="showborders">Attachment:&nbsp;</td>
                        <td class="showborders">
                        <input class="fileuploadchrome" type="file" name="Attachment" size="20">
                     <%Else%>
                        <td class="showborders" colspan="2">Problem:</td>
                        <td class="showborders">Attachment:&nbsp;</td>
                        <td class="showborders">
                        <input class="fileupload" type="file" name="Attachment" size="20">
                     <%End If
                  Else %>
                     <td class="showborders" colspan="4">Problem:</td>
                  <%End If%>
                        
                        
   <%                End If %>
                  </td>
   				</tr>
   				<tr>
   					<td class="showborders" colspan="4">
                  <%=Replace(FixURLs(objRecordSet(Problem),1),vbCRLF,"<br />")%>
                  <input type="hidden" name="Problem" value="<%=Replace(objRecordSet(Problem),"""","")%>">
               </tr>
   				<tr>
   					<td class="showborders" colspan="4">Notes:</td>
   				</tr>
   				<tr>
   					<td class="showborders" colspan="4">

   <%             'If the user is visiting this page for the first time the default item in the Notes
                  'text box will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If (strCMD = "Save") or bolKeepData Then%>
                     <textarea rows="8" name="Notes" cols="90" style="width: 99%;"><%=strNotes%></textarea></td>
   <%             Else%>
                     <textarea rows="8" name="Notes" cols="90" style="width: 99%;"><%=objRecordSet(Notes)%></textarea></td>
   <%             End If%>
   
   				</tr>
   				<tr>
   					<td class="showborders" colspan="4">
   					<table border="0" width="100%" cellspacing="0" cellpadding="0" id="table3">
   						<tr>
   							<td colspan="2" class="showborders">Submitted on <%=objRecordSet(SubmitDate)%> at 
   							<%=objRecordSet(SubmitTime)%> 

   <%                   'Check the last updated date in the database.  Display it if it is different then the default
                        If objRecordSet(LastUpdatedDate) = "6/16/1978" Then%>
                           - Never updated.
   <%                   Else%>
                           - Updated <%=objRecordSet(LastUpdatedDate)%> at 
   							<%=objRecordSet(LastUpdatedTime)%></td>
   <%                   End If %>
                     </tr>
                     <tr>
                  <% If strReturnLink = "" Then %>
                     <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>">
                  <% Else %>
                     <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>&<%=Right(strReturnLink,Len(strReturnLink)-1)%>">
                  <% End If %>
                        <td class="showborders" colspan="3">
                           Send this ticket to an additional email address 
   <%                   If Upload.Form("cmdEMail") = "" Then %>
                           <input type="text" name="SendEMail" size="30">
   <%                   Else %>
                           <input type="text" name="SendEMail" value="<%=Upload.Form("SendEMail")%>" size="30">
   <%                   End If %>
                           <input type="submit" value="EMail Ticket" name="cmdEMail">
                        </td>
                     <form>
                     </tr>
                     <tr>
                     <td valign="top">
   <%                If bolShowLogButton Then      
                        If bolLog Then %>
                           <input type="submit" value="Hide Log" name="cmdSubmit"><input type="submit" value="User History" name="cmdSubmit"></td>
   <%                   Else %>
                           <input type="submit" value="Show Log" name="cmdSubmit"><input type="submit" value="User History" name="cmdSubmit"></td>
   <%                   End If 
                     End If%>   
   							<td>
							<div align="right">
     
<%             If objRecordSet(Status) <> "Complete" Then %>
<%                If bolTracking Then%>                     
                     <input type="submit" value="Stop Tracking" name="cmdSubmit">
<%                Else %>                        
                     <input type="submit" value="Track Ticket" name="cmdSubmit">
<%                End If %>
<%                If bolRequest Then%>
                     <input type="submit" value="Cancel Update Request" name="cmdSubmit">
<%                Else %>
<%                   If objRecordSet(Tech) <> "" Then %>                     
                        <input type="submit" value="Request Update" name="cmdSubmit">
<%                   Else %>   
                        <input type="submit" disabled="disabled" value="Request Update" name="cmdSubmit">
<%                   End If %>   
<%                End If %>
<%                If strRole <> "Data Viewer" Then %> 
<%                   If objRecordSet(Tech) <> "" Then %>
                        <input type="submit" value="Update Tech" name="cmdSubmit">
<%                   Else %>
                        <input type="submit" disabled="disabled" value="Update Tech" name="cmdSubmit">
<%                   End If %>
                     <input type="submit" value="Update User" name="cmdSubmit">
<%                End If %>
<%             End If %>
                  <input type="button" value="Print" onClick="window.open('print.asp?ID=<%=intID%>','newwin');">
<%             If strRole <> "Data Viewer" Then %>
<%                If objRecordSet(Status) = "Complete" Then %>	             
                     <input type="submit" value="Open Ticket" name="cmdSubmit">
<%                Else %>
                     <input type="submit" value="Save" name="cmdSubmit">
<%                End If
               End If %>
   					&nbsp;</div></td>
   						</tr>
   					</table>
   					</td>
   				</tr>
   <%       If bolLog Then %>       
               <tr>
                  <td colspan="4">
                     Activity Log for Ticket #<%=intID%>
                     <ul>
   <%          Do Until objLog.EOF 
                  Select Case objLog(2)
                     Case "Location Changed"%>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the location from <%=objLog(4)%> to <%=objLog(5)%>.</li>
   <%                Case "EMail Changed" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the users email from <%=objLog(4)%> to <%=objLog(5)%>.</li>
   <%                Case "Category Changed" 
                        If objLog(4) = " " Then %>
                           <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                           the category to <%=objLog(5)%>.</li>
   <%                   Else %>
                           <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                           the category from <%=objLog(4)%> to <%=objLog(5)%>.</li>
   <%                   End If                        
                     Case "Tech Changed" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the assigned tech from <%=objLog(4)%> to <%=objLog(5)%>. 
   <%                   If objLog(8) <> "" Then 
   
                           'Calculate how long a task was assigned to the last t
                           strDays = Int(objLog(8)/1440)
                           strHours = Int((objLog(8)-strDays*1440)/60)
                           strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                           strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                           <ul><li>The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>.</li></ul>
   <%                   End If %>
                        </li>
   <%                Case "Status Changed" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the status from <%=objLog(4)%> to <%=objLog(5)%>.</li>
   <%                Case "Notes Updated" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> updated
                        the notes.</li>
   <%                Case "User Notified" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> sent the user an update.</li>
   <%                Case "Ticket EMailed" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> emailed this ticket to <%=objLog(5)%>.</li>
   <%                Case "Tech Notified" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> sent <%=objLog(5)%> an update.</li>
   <%                Case "Update Requested" %>
                     <% If objLog(4) = "" Then %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> requested an update.
                        </li>
                     <% Else %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> requested an update from
                        <%=objLog(4)%>.
                        </li>
                     <% End If %>
   <%                Case "Cancelled Update Request" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> cancelled their request for update.</li>
   <%                Case "Request Update Complete" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> updated the ticket after a request for update.</li>
   <%                Case "Auto Assigned" %>
                        <li>Ticket Auto Assigned to <%=objLog(5)%>. 
   <%                   If objLog(8) <> "" Then 
   
                           'Calculate how long a task was assigned to the lastt
                           strDays = Int(objLog(8)/1440)
                           strHours = Int((objLog(8)-strDays*1440)/60)
                           strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                           strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                           <ul><li>The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>.</li></ul>
   <%                   End If %>
                        </li>
   <%                Case "New Ticket" %>
                        <li>Ticket Entered. </li>
   <%                Case "Ticket Reopened" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> reopened the ticket.</li>
   <%                Case "Tech Reassigned" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> assigned the reopened ticket to <%=objLog(5)%>.
   <%                   If objLog(8) <> "" Then 
   
                           'Calculate how long a task was assigned to the last t
                           strDays = Int(objLog(8)/1440)
                           strHours = Int((objLog(8)-strDays*1440)/60)
                           strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                           strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                           <ul><li>The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>.</li></ul>
   <%                   End If %>      
                        </li>
   <%                Case "Assigned" %>
                        <li>On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> assigned the ticket to <%=objLog(5)%>.</li>
   <%                Case "Ticket Tracked" %>
                        <li>
                           On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> started tracking the ticket.
                        </li>
   <%                Case "Ticket Not Tracked" %>
                        <li>
                           On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> stopped tracking the ticket.
                        </li>
   <%                Case "Ticket Viewed" %>
                        <li>
                           On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> viewed the ticket.
                        </li>
   <%             End Select
                  objLog.MoveNext
               Loop %>
                     </ul>   
                  </td>
               </tr>
   <%       End If%>
   			</table>
   			</td>
   		</tr>
   	</table>
   	</div>
   		
   		<p>&nbsp;</p>
   	</form>
   </body>
<%End Sub%>

<%Sub MobileVersion 

   Const ID = 0
   Const Name = 1
   Const Location = 2
   Const EMail = 3
   Const Problem = 4
   Const SubmitDate = 5
   Const SubmitTime = 6
   Const Notes = 7
   Const Status = 8
   Const Category = 9
   Const Tech = 10
   Const LastUpdatedDate = 11
   Const LastUpdatedTime = 12
   Const Custom1 = 13
   Const Custom2 = 14
   Const TicketViewed = 15
   
   Dim intInputSize

   If InStr(strUserAgent,"Nexus 7") And Request.Cookies("ZoomLevel") <> "ZoomIn" Then
      intInputSize = 50
   ElseIf InStr(strUserAgent,"iPhone") Then
      intInputSize = 31
   ElseIf InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 22
   Else
      intInputSize = 25
   End If   
   
   If strStatus = "Auto Assigned" or strStatus = "New Assignment" Then
      strStatus = "In Progress"
   End If
   
%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin - <%=intID%></title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>   
      <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
      <link rel="stylesheet" type="text/css" href="../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
	   <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>" />
   <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then %>
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %>
      <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
   </head>
   <body>
      <center>
         <% If objRecordSet(Status) = "Complete" Then %>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/closed.gif" width="15" height="15">
            <% Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/closed.gif" width="15" height="15">
            <% End If %>
         <% End If %>            
<%       If bolTicketViewed Then %>   
<%          If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/viewed.gif" alt="Viewed by Tech" width="15" height="15">
<%          Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/viewed.gif" alt="Viewed by Tech" width="15" height="15">
<%          End If %>
<%       Else %>
<%          If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <img border="0" src="../themes/<%=Application("Theme")%>/images/notviewed.gif" alt="Not Viewed by Tech" width="15" height="15">
<%          Else %>
               <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/notviewed.gif" alt="Not Viewed by Tech" width="15" height="15">
<%          End If %>            
<%       End If %>
      <b>Ticket #<%=intID%></b> - <%=strTimeActive%>
      </center>
      <center>
      <table align="center">
   <% If strReturnLink = "" Then %>
         <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>">
   <% Else %>
         <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>&<%=Right(strReturnLink,Len(strReturnLink)-1)%>">
   <% End If %>
      
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         
         <tr>
            <td colspan="2" width="<%=Application("MobileSiteWidth")%>">
               <div align="center">
   <%          If strLinkBack = "Yes" Then %>
                  <input type="button" value="<" onclick="window.location.href='view.asp<%=strReturnLink%>#<%=strLinkID%>'">  
   <%          End If%>
                  <input type="button" value="Home" onclick="window.location.href='index.asp'">
                  <input type="button" value="Open Tickets" onclick="window.location.href='view.asp?Filter=AllOpenTickets'">  
            <% If strRole <> "Data Viewer" Then %>   
                  <input type="button" value="Your Tickets" onclick="window.location.href='view.asp?Filter=MyOpenTickets'">  
            <% End If %>
 
               </div>
            </td>
         </tr>
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
      </table>
      <table align="center">   
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>">
   <%    'Display any error messages if the form is missing anything
         If (strCMD = "Save") And (strUserTemp = "" Or strEMailTemp = "" Or strLocation = "" Or strCategory = " " Or strTech = "") Then %>
            <font class="missing">Please fill out highlighted fields...</font>        
<%          bolUpdated = False
         End If
         If (strCMD = "Save") And strStatus = "New Assignment" Then %>
            <font class="missing">Status cannot be "New Assignment"</font>
<%          bolUpdated = False
         End If 
         If (strCMD = "Save") And strStatus = "Auto Assigned" Then %>
            <font class="missing">Status cannot be "Auto Assigned"</font>
<%          bolUpdated = False
         End If 
         If bolUserUpdated Then %>
            <font class="information">EMail Sent to <%=objRecordSet(EMail)%></font>
<%          bolUpdated = False
         End If
         If bolTechUpdated Then %>
            <font class="information">EMail Sent to <%=strTechEmail%></font>
<%          bolUpdated = False
         End If 
         If strCMD = "Save" And bolUpdated And Not bolCallClosed And Not bolTicketReOpened Then 
            If strTechEmail <> "" Then %>
               <font class="information">Ticket Updated - EMail Sent to <%=strTechEmail%></font>
<%          Else %>
               <font class="information">Ticket Updated</font>
<%          End If
         End If
         If strCMD = "Save"  And bolUpdated And bolCallClosed Then %>
            <font class="information">Ticket Closed - EMail Sent to <%=objRecordSet(EMail)%></font></b>
<%       End If
         If bolUpdateRequested Then %>
            <font class="information">EMail Sent to <%=strTechEMail%></font>
<%       End If
         If bolMissingTech Then %>
            <font class="missing">No tech assigned...<%=strTechEMail%></font>
<%       End If 
         If bolTicketReOpened Then %>
            <font class="information">Ticket Reopened</font>
<%       End If 
         If bolTrackTicket Then %>
            <font class="information">You are now tracking this ticket</font>
<%       End If
         If bolTrackTicketOff Then %>
            <font class="information">You are no longer tracking this ticket</font>
<%       End If            
         If bolCancelledUpdateRequest Then %>
            <font class="information">Request for update cancelled</font>
<%       End If %>

<%       If Upload.Form("cmdEMail") <> "" Then %>
<%          If IsEmailValid(strEMail) Then %>
               <font class="information">EMail Sent to <%=strEMail%></font>
<%          Else %>
               <font class="missing"><%=strEMail%> is an invalid address</font>
<%          End If %>
<%       End If %>

         </td></tr>
      </table>
   <% Do  Until objRecordSet.EOF %> 
      <table align="center"><tr><td width="<%=Application("MobileSiteWidth")%>">
      <table align="center">
         <tr><td class="showborders">Name: </td>
            <td class="showborders">
               <%=objRecordSet(1)%>               
               <input type="hidden" name="Name" value="<%=Replace(objRecordSet(1),"""","")%>">
               <input type="hidden" name="EMail" value="<%=Replace(objRecordSet(3),"""","")%>">
               <input type="hidden" name="Problem" value="<%=Replace(objRecordSet(4),"""","")%>">
            <% If Application("UseCustom1") Then %>   
                  <input type="hidden" name="Custom1" value="<%=Replace(objRecordSet(13),"""","")%>">
            <% End If %>
            <% If Application("UseCustom2") Then %> 
               <input type="hidden" name="Custom2" value="<%=Replace(objRecordSet(14),"""","")%>">
            <% End If %>
            </td>
         </tr>
   <%    If Application("UseCustom1") Then %> 
            <tr><td class="showborders"><%=Application("Custom1Text")%>: </td><td class="showborders"><%=objRecordSet(13)%></td></tr>
   <%    End If
         If Application("UseCustom2") Then %>			
            <tr><td class="showborders"><%=Application("Custom2Text")%>: </td><td class="showborders"><%=objRecordSet(14)%></td></tr>
   <%    End If %>
         <tr>
<%          'Highlight the Tech label if it was blank when the form was submitted
            If (strCMD = "Save") And (strTech = "") Then %>
               <td class="showborders"><font class="missing">Tech:</font></td>
<%          Else %>
               <td class="showborders">Tech:</td>
<%          End If %>
            <td class="showborders">
              <select size="1" name="Tech">
<%            If ((strCMD = "Save") And strTech <> objRecordSet(Tech)) or bolKeepData Then%>
                  <option value="<%=strTech%>"><%=strTech%></option>
<%                Do Until objTechSet.EOF
                     If strTech = "" Then%>
                        <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
<%                   ElseIf Trim(Ucase(strTech)) <> Trim(Ucase(objTechSet(0))) Then%>
                        <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
<%                   End If
                     objTechSet.MoveNext
                  Loop
               Else%>
                  <option value="<%=objRecordSet(Tech)%>"><%=objRecordSet(Tech)%></option>
<%                Do Until objTechSet.EOF
                     If objRecordSet(Tech) = "" Then%>
                        <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
<%                   ElseIf Trim(Ucase(objRecordSet(Tech))) <> Trim(Ucase(objTechSet(0))) Then%>
                        <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
<%                   End If
                     objTechSet.MoveNext
                  Loop
               End If%>
               </select>
            </td>
         </tr>
         <tr>
            <td class="showborders">Location:</td>
            <td class="showborders">
               <select size="1" name="Location">

   <%             'If the user is visiting this page for the first time the default item in the location
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strLocation <> objRecordSet(Location)) or bolKeepData Then%>
                     <option value="<%=strLocation%>"><%=strLocation%></option>
   <%                Do Until objLocationSet.EOF
                        If strLocation = "" Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   ElseIf Trim(Ucase(strLocation)) <> Trim(Ucase(objLocationSet(0))) Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   End If
                        objLocationSet.MoveNext
                     Loop
                  Else%>
   					   <option value="<%=objRecordSet(Location)%>"><%=objRecordSet(Location)%></option>
   <%                Do Until objLocationSet.EOF
                        If objRecordSet(Location) = "" Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Location))) <> Trim(Ucase(objLocationSet(0))) Then%>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
   <%                   End If
                        objLocationSet.MoveNext
                     Loop
                  End If%>         
   			   </select>
            </td>
         <tr>
         <tr>
   <%    'Highlight the Category label if it was blank when the form was submitted
         If (strCMD = "Save") And (strCategory = " ") Then %>
            <td class="showborders" width="9%"><font class="missing">Category:</font></td>
   <%    Else %>
            <td class="showborders" width="9%">Category:</td>
   <%    End If %>
            <td class="showborders">
               <select size="1" name="Category">

   <%             'If the user is visiting this page for the first time the default item in the category
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strCategory <> objRecordSet(Category)) or bolKeepData Then%>
                     <option value="<%=strCategory%>"><%=strCategory%></option>
   <%                Do Until objCategorySet.EOF
                        If strCategory = "" Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   ElseIf Trim(Ucase(strCategory)) <> Trim(Ucase(objCategorySet(0))) Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   End If
                        objCategorySet.MoveNext
                     Loop
                  Else%>
                     <option value="<%=objRecordSet(Category)%>"><%=objRecordSet(Category)%></option>
   <%                Do Until objCategorySet.EOF
                        If objRecordSet(Category) = "" Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Category))) <> Trim(Ucase(objCategorySet(0))) Then%>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
   <%                   End If
                        objCategorySet.MoveNext
                     Loop
                  End If %>
   			   </select>
            </td>
         <tr>
         <tr>
         
   <%    'Highlight the Status label if it was blank when the form was submitted
         If (strCMD = "Save") And (strStatus = "New Assignment" or strStatus = "Auto Assigned") Then%>
            <td class="showborders" ><font class="missing">Status:</font></td>
   <%    Else%>
            <td class="showborders" >Status:</td>
   <%    End If%>
            <td class="showborders">
               <select size="1" name="Status">

   <%             'If the user is visiting this page for the first time the default item in the Status
                  'pulldown list will be the current value from the database.  If the user has submitted
                  'the form and their was an error then the default value will be what was submitted.  This
                  'way the form won't reset it's value if there was an error.
                  If ((strCMD = "Save") And strStatus <> objRecordSet(Status)) or bolKeepData Then%>
   					   <option value="<%=strStatus%>"><%=strStatus%></option>
   <%                Do Until objStatusSet.EOF
                        If strStatus = "" Then%>
                           <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   Else%>
                           <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   End If
                        objStatusSet.MoveNext
                     Loop%>
   
   <%             Else%>
   					   <option value="<%=objRecordSet(Status)%>"><%=objRecordSet(Status)%></option>
   <%                Do Until objStatusSet.EOF
                        If objRecordSet(Status) = "" Then
   %>                      <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   ElseIf Trim(Ucase(objRecordSet(Status))) <> Trim(Ucase(objStatusSet(0))) Then
   %>                      <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
   <%                   End If
                        objStatusSet.MoveNext
                     Loop   
                  End If%>   
   				</select>
            </td>
         </tr>
<%       Set objFSO = CreateObject("Scripting.FileSystemObject")
         If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
            strAttachmentFolder = Application("FileLocation") & "\" & intID
            strAttachTech = ""
         End If  
         If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
            strAttachmentFolder = Application("FileLocation") & "\" & intID & "-Tech"
            strAttachTech = "-Tech"
         End If
         If objFSO.FolderExists(strAttachmentFolder) Then %>
            <tr>
               <td class="showborders">Attachment:&nbsp;</td>
               <td class="showborders">
<%          Set objFolder = objFSO.GetFolder(strAttachmentFolder)
            For Each objFile in objFolder.Files
               If UCase(objFile.Name) <> "THUMBS.DB" Then
                  Response.Write "<a href=""download.asp?folder=" & intID & strAttachTech & "&file=" & objFile.Name & """>Click Here</a>&nbsp;"
               End If
            Next %>
            </tr>
<%       End If %>
         <tr>
            <td colspan="2" class="showborders">Problem:</td>
         </tr>
         <tr>
            <td colspan="2" class="showborders">
               <%=Replace(FixURLs(objRecordSet(Problem),1),vbCRLF,"<br />")%>
            </td>
         </tr>
         <tr>
            <td colspan="2" class="showborders">Notes:</td>
         <tr>
         <tr><td colspan="2" class="showborders">
   <%       'If the user is visiting this page for the first time the default item in the Notes
            'text box will be the current value from the database.  If the user has submitted
            'the form and their was an error then the default value will be what was submitted.  This
            'way the form won't reset it's value if there was an error.
            If (strCMD = "Save") or bolKeepData Then%>
               <textarea rows="8" name="Notes" cols="90" style="width: 98%;"><%=strNotes%></textarea>
   <%       Else%>
               <textarea rows="8" name="Notes" cols="90" style="width: 98%;"><%=objRecordSet(Notes)%></textarea>
   <%       End If%>
         </td></tr>
         
         <td class="showborders" colspan="3">
            Send this ticket to an additional email <br />
<%       If Upload.Form("cmdEMail") = "" Then %>
            <input type="text" name="SendEMail" size="<%=intInputSize%>">
<%       Else %>
            <input type="text" name="SendEMail" value="<%=Upload.Form("SendEMail")%>" size="<%=intInputSize%>">
<%       End If %>
            <input type="submit" value="EMail Ticket" name="cmdEMail" style="float: right">
         </td>     

         <tr>
            <td colspan="2" class="showborders">Submitted on <%=objRecordSet(SubmitDate)%> at 
            <%=objRecordSet(SubmitTime)%> 

<%          'Check the last updated date in the database.  Display it if it is different then the default
            If objRecordSet(LastUpdatedDate) = "6/16/1978" Then%>
               - Never updated.
<%                   Else%>
               - Updated <%=objRecordSet(LastUpdatedDate)%> at 
            <%=objRecordSet(LastUpdatedTime)%></td>
<%                   End If %>
         </tr>
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         <tr><td colspan="2" align="center">
<%       If objRecordSet(Status) <> "Complete" Then %>
<%                If bolTracking Then%>                     
                     <input type="submit" value="Stop Tracking" name="cmdSubmit">
<%                Else %>                        
                     <input type="submit" value="Track Ticket" name="cmdSubmit">
<%                End If %>
<%                If bolRequest Then%>
                     <input type="submit" value="Cancel Update Request" name="cmdSubmit">
<%                Else %>
<%                   If objRecordSet(Tech) <> "" Then %>                     
                        <input type="submit" value="Request Update" name="cmdSubmit">
<%                   Else %>   
                        <input type="submit" disabled="disabled" value="Request Update" name="cmdSubmit">
<%                   End If %>   
<%                End If %>
                  </td></tr>
                  <tr><td colspan="2"><hr /></td></tr>
                  <tr><td colspan="2" align="center">
<%                If strRole <> "Data Viewer" Then %> 
<%                   If objRecordSet(Tech) <> "" Then %>
                        <input type="submit" value="Update Tech" name="cmdSubmit">
<%                   Else %>
                        <input type="submit" disabled="disabled" value="Update Tech" name="cmdSubmit">
<%                   End If %>
                     <input type="submit" value="Update User" name="cmdSubmit">
<%                End If %>
                  </td></tr>
                  <tr><td colspan="2"><hr /></td></tr>
                  <tr><td colspan="2"> 
<%             End If %>
<%             If strRole <> "Data Viewer" Then %>
              
<%                If objRecordSet(Status) = "Complete" Then 
                     If bolLog Then%>
                        <input type="submit" value="Hide Log" name="cmdSubmit" style="float: left">
                  <% Else %>
                        <input type="submit" value="Show Log" name="cmdSubmit" style="float: left">
                  <% End If %>
                     <input type="submit" value="User History" name="cmdSubmit">
                     <input type="submit" value="Open Ticket" name="cmdSubmit" style="float: right">
<%                Else 
                     If bolLog Then%>
                        <input type="submit" value="Hide Log" name="cmdSubmit">
                  <% Else %>
                        <input type="submit" value="Show Log" name="cmdSubmit">
                  <% End If %>
                     <input type="submit" value="User History" name="cmdSubmit">
                     <input type="submit" value="Save" name="cmdSubmit" style="float: right">
<%                End If
               End If %>
   <%    objRecordSet.MoveNext
      Loop %>
      </form>
      </td></tr>
      
<%       If bolLog Then %>     
  
            <tr><td colspan="2"><hr /></td></tr>
            <tr>
               <td colspan="2">
                  Activity Log for Ticket #<%=intID%> <br />
<%          Do Until objLog.EOF 
               Select Case objLog(2)
                  Case "Location Changed"%>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                     the location from <%=objLog(4)%> to <%=objLog(5)%>. <br />
<%                Case "EMail Changed" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                     the users email from <%=objLog(4)%> to <%=objLog(5)%>. <br />
<%                Case "Category Changed" 
                     If objLog(4) = " " Then %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the category to <%=objLog(5)%>. <br />
<%                   Else %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                        the category from <%=objLog(4)%> to <%=objLog(5)%>. <br />
<%                   End If                        
                  Case "Tech Changed" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                     the assigned tech from <%=objLog(4)%> to <%=objLog(5)%>. <br />
<%                   If objLog(8) <> "" Then 

                        'Calculate how long a task was assigned to the last t
                        strDays = Int(objLog(8)/1440)
                        strHours = Int((objLog(8)-strDays*1440)/60)
                        strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                        strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                        &nbsp;&nbsp;&nbsp;- The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>.
<%                   End If %>
                     
<%                Case "Status Changed" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> changed
                     the status from <%=objLog(4)%> to <%=objLog(5)%>. <br />
<%                Case "Notes Updated" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> updated
                     the notes. <br />
<%                Case "User Notified" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> sent the user an update. <br />
<%                Case "Ticket EMailed" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> emailed this ticket to <%=objLog(5)%>. <br />
<%                Case "Tech Notified" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> sent <%=objLog(5)%> an update. <br />
<%                Case "Update Requested" %>
                  <% If objLog(4) = "" Then %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> requested an update. <br />
                  <% Else %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> requested an update from
                     <%=objLog(4)%>. <br />
                  <% End If %>
<%                Case "Cancelled Update Request" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> cancelled their request for update. <br />
<%                Case "Request Update Complete" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> updated the ticket after a request for update. <br />
<%                Case "Auto Assigned" %>
                     - Ticket Auto Assigned to <%=objLog(5)%>. <br />
<%                   If objLog(8) <> "" Then 

                        'Calculate how long a task was assigned to the lastt
                        strDays = Int(objLog(8)/1440)
                        strHours = Int((objLog(8)-strDays*1440)/60)
                        strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                        strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                        &nbsp;&nbsp;&nbsp;- The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>. <br />
<%                   End If %>
<%                Case "New Ticket" %>
                     - Ticket Entered. <br />
<%                Case "Ticket Reopened" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> reopened the ticket. <br />
<%                Case "Tech Reassigned" %>
                     - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> assigned the reopened ticket to <%=objLog(5)%>. <br />
<%                   If objLog(8) <> "" Then 

                        'Calculate how long a task was assigned to the last t
                        strDays = Int(objLog(8)/1440)
                        strHours = Int((objLog(8)-strDays*1440)/60)
                        strMinutes = (objLog(8)-(strDays*1440)-(strHours*60))
                        strTaskTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" %>
                        &nbsp;&nbsp;&nbsp;- The task was assigned to <%=objLog(5)%> for <%=strTaskTime%>. <br />
<%                   End If %>      
<%                Case "Assigned" %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> assigned the ticket to <%=objLog(5)%>. <br />
<%                Case "Ticket Tracked" %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> started tracking the ticket. <br />
<%                Case "Ticket Not Tracked" %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> stopped tracking the ticket. <br />
<%                Case "Ticket Viewed" %>
                        - On <%=objLog(6)%> at <%=objLog(7)%>&nbsp;<%=GetFirstandLastName(objLog(3))%> viewed the ticket. <br />
<%             End Select
               objLog.MoveNext
            Loop %>
               </td>
            </tr>
<%       End If%>

      
      
      </table>
      </td></tr>
      </table>
      </center>
   </body>
   </html> 
<%End Sub%>   

<%Sub WatchVersion 

   Const ID = 0
   Const Name = 1
   Const Location = 2
   Const EMail = 3
   Const Problem = 4
   Const SubmitDate = 5
   Const SubmitTime = 6
   Const Notes = 7
   Const Status = 8
   Const Category = 9
   Const Tech = 10
   Const LastUpdatedDate = 11
   Const LastUpdatedTime = 12
   Const Custom1 = 13
   Const Custom2 = 14
   Const TicketViewed = 15
%>

   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>HDL - Admin</title>
   <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>   
      <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <% Else %>
      <link rel="stylesheet" type="text/css" href="../themes/<%=objNameCheckSet(3)%>/<%=objNameCheckSet(3)%>.css" />
   <% End If %>
      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>" />

   <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then %>
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=1.9" />
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %> 
   </head>
      <div align="right">
         <%=intID%></b> - <%=strTimeActive%>
      </div>
      <hr />
   <% If bolUserUpdated Then %>
         <font class="information">EMail Sent to <%=objRecordSet(EMail)%></font>
   <%    bolUpdated = False
      End If
      If bolTechUpdated Then %>
         <font class="information">EMail Sent to <%=strTechEmail%></font>
   <%    bolUpdated = False
      End If %>
         
   <% If strReturnLink = "" Then %>
         <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>">
   <% Else %>
         <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="modify.asp?ID=<%=objRecordSet("ID")%>&<%=Right(strReturnLink,Len(strReturnLink)-1)%>">
   <% End If %>
   
   <% Do Until objRecordSet.EOF %>
         <input type="hidden" name="Name" value="<%=Replace(objRecordSet(Name),"""","")%>">
         <input type="hidden" name="EMail" value="<%=Replace(objRecordSet(EMail),"""","")%>">
         <input type="hidden" name="Problem" value="<%=Replace(objRecordSet(Problem),"""","")%>">
         <input type="hidden" name="Tech" value="<%=Replace(objRecordSet(Tech),"""","")%>">
      <% If Application("UseCustom1") Then %>   
            <input type="hidden" name="Custom1" value="<%=Replace(objRecordSet(13),"""","")%>">
      <% End If %>
      <% If Application("UseCustom2") Then %> 
         <input type="hidden" name="Custom2" value="<%=Replace(objRecordSet(14),"""","")%>">
      <% End If %>
         
         <div align="center"> 
            
      <% If strLinkBack = "Yes" Then %>
            <input type="button" value="Back" onclick="window.location.href='view.asp<%=strReturnLink%>#<%=strLinkID%>'"> <br /> <br />
      <% End If%>
            
      <% If bolTracking Then%>                     
            <input type="submit" value="Stop Tracking" name="cmdSubmit"> <br /> <br />
      <% Else %>                        
            <input type="submit" value="Track Ticket" name="cmdSubmit"> <br /> <br />
      <% End If %>
      <% If bolRequest Then%>
            <input type="submit" value="Cancel Update Request" name="cmdSubmit"> <br /> <br />
      <% Else %>
         <% If objRecordSet(Tech) <> "" Then %>                     
               <input type="submit" value="Request Update" name="cmdSubmit"> <br /> <br />
         <% Else %>   
               <input type="submit" disabled="disabled" value="Request Update" name="cmdSubmit"> <br /> <br />
         <% End If %>
      <% End If %>
      
      <% If strRole <> "Data Viewer" Then %> 
      <%    If objRecordSet(Tech) <> "" Then %>
               <input type="submit" value="Update Tech" name="cmdSubmit"> <br /> <br />
      <%    Else %>
               <input type="submit" disabled="disabled" value="Update Tech" name="cmdSubmit"> <br /> <br />
      <%    End If %>
            <input type="submit" value="Update User" name="cmdSubmit"> <br />
      <% End If %>
         </div>
      
      <% objRecordSet.MoveNext 
      Loop %>
      </form>
      <hr />
   </body>
   </html>
<%End Sub%>

<%Sub Submit

   'This is a simple page that will update the database with the new settings then display
   'a message to the user letting them know.
   
   On Error Resume Next
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserName, strEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strDate, strTime, objRegExp, strNewUserName, strNewEMail, strNewLocation
   Dim strNewCategory, strNewProblem, strNewNotes, strSQL, objShell, strMessage, objMessage
   Dim objTechSet, objConf, strSubmitDate, strSubmitTime, strOpenTime, strDisplayName 
   Dim strCustom1, strCustom2, strOldTech, strCurrentUser, objTrackingSet
   Dim objOldTechSet, strOldTechEmail, objFSO, objFolder, objFile, strAttachment
   Dim intFileCount, strAttachmentTech, objOldAssignment, strOpenAssignmentTime   
   
   strCurrentUser = GetFirstandLastName(strUser)
   
   intID = Request.Querystring("ID")
   
   'If there is an attachment save it to a folder on the server.
   intFileCount = 0
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   objFSO.CreateFolder(Application("FileLocation") & "\" & intID & "-Tech")
   Upload.Save(Application("FileLocation") & "\" & intID & "-Tech")
   Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID & "-Tech")

   For Each objFile in objFolder.Files
      intFileCount = intFileCount + 1
      strAttachment = objFile.Path
   Next
   If intFileCount = 0 Then
      objFSO.DeleteFolder Application("FileLocation") & "\" & intID & "-Tech"
   End If
   Set objFSO = Nothing
   
   'Get the information from the forms and address bar and assign them to variables
   strEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   If bolTicketReOpened Then
      strStatus = "In Progress"
   Else
      strStatus = Upload.Form("Status")
   End If
   strTech = Upload.Form("Tech")
   strDisplayName = Upload.Form("Name")
   strDate = Date
   strTime = Time
   
   If strStatus = "Auto Assigned" Or strStatus = "New Assignment" Then
      strStatus = "In Progress"
   End If

   'Build the SQL string that will get the submitted date and time for the requested ticket
   strSQL = "SELECT SubmitDate, SubmitTime, Name, Custom1, Custom2, Tech" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   
   'Execute the SQL string and assign the results to a Record Set
   Set objRecordSet = Application("Connection").Execute(strSQL)
   
   strSubmitDate = objRecordSet(0)
   strSubmitTime = objRecordSet(1)
   strUsername = objRecordSet(2)
   strCustom1 = objRecordSet(3)
   strCustom2 = objRecordSet(4)
   strOldTech = objRecordSet(5)

   strOpenTime = DateDiff("n",strSubmitDate,strDate)
   strOpenTime = strOpenTime + DateDiff("n",strSubmitTime,strTime)
   
   'Create the Regular Expression object and set it's properties.
   Set objRegExp = New RegExp
   objRegExp.Pattern = "'"
   objRegExp.Global = True

   'Use the regular expression to change a ' to a '' so the SQL Insert command will work.
   'The value will be assigned to a new variable so the old one can still be displayed   
   strNewUserName = objRegExp.Replace(strUserName,"''")
   strNewEMail = objRegExp.Replace(strEMail,"''")
   strNewLocation = objRegExp.Replace(strLocation,"''")
   strNewCategory = objRegExp.Replace(strCategory,"''")
   strNewProblem = objRegExp.Replace(strProblem,"''")
   strNewNotes = objRegExp.Replace(strNotes,"''")
   
   'Find out if the current user is either tracking the ticket or has requested an update
   strSQL = "SELECT Type FROM Tracking WHERE Ticket=" & intID
   Set objTracking = Application("Connection").Execute(strSQL)
   bolTracking = False
   bolRequest = False
   Do Until objTracking.EOF
      Select Case objTracking(0)
         Case "Track"
            bolTracking = True
         Case "Request"
            bolRequest = True
      End Select
      objTracking.MoveNext
   Loop

   'See if anyone has requested an updated on this ticket
   strSQL = "SELECT TrackedBy FROM Tracking WHERE Ticket=" & intID & " And Type='Request'"
   Set objUpdateRequest = Application("Connection").Execute(strSQL)
   
   'Send an email to the tech if the call isn't assigned to them anymore.  Only if someone else changes it.
   If strOldTech <> strTech And strOldTech <> strCurrentUser And strOldTech <> "" Then
      TicketReassigned
   End If
   
   'Send the tech an email if they have just been assigned this call
   If strOldTech <> strTech And strStatus <> "Complete" Then   
      TicketAssigned
      
      If strTech = "Erwin Brace" And Application("TSCHelpDesk") <> "" Then

         'Build the SQL string that will add the data to the TSC database
         strSQL = "Insert Into Main (Name,DisplayName,Email,Location,Problem,SubmitDate,SubmitTime,Category,Status,Tech,LastUpdatedDate,TicketViewed) " & _
         "values ('help','" & Replace(Upload.Form("Name"),"'","''") & "','help@wswheboces.org','BOCES','" & strNewProblem & vbCRLF & vbCRLF & _ 
         "Original Ticket Information" & vbCRLF & _
         "Ticket Number: " & intID & vbCRLF & _
         "User: " & Replace(strDisplayName,"'","''") & vbCRLF & _
         "Location: " & Replace(strNewLocation,"'","''") & vbCRLF & _
         "Room: " & Replace(strCustom1,"'","''") & vbCRLF & _
         "Phone: " & Replace(strCustom2,"'","''") & vbCRLF & _
         "Tech Notes: " & Replace(strNewNotes,"'","''") & _
         "','" & strDate & "','" & strTime & "',' ','New Assignment','','6/16/78',False)"

         Dim objConnection, strConnection
      
         'Create the connection to the TSC database
         Set objConnection = CreateObject("ADODB.Connection")
         strConnection = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & Application("TSCHelpDesk") & ";"
         objConnection.Open strConnection
         objConnection.Execute(strSQL)
         Set objConnection = Nothing
      
      End If
      
   End If
   
   'Send the User an email if the call was just closed
   If strStatus = "Complete" = True Then 
      TicketClosed
   End If
  
   'If someone has requested an update send them the update
   If Not objUpdateRequest.EOF And Not bolUpdateRequested Then
      SendRequestedUpdate
   End If
 
   'If the ticket is being tracked send the person tracking it an update
   If bolTracking And Not bolTrackTicket Then
      SendTrackingEMail 
   End If
   
   UpdateLog
   
   If strStatus = "Complete" Then
      strSQL = "SELECT ID,UpdateDate,UpdateTime" & vbCRLF
      strSQL = strSQL & "FROM Log" & vbCRLF
      strSQL = strSQL & "WHERE (NewValue='" & strOldTech & "' AND Ticket=" & intID & ")" & vbCRLF
      strSQL = strSQL & "ORDER BY ID DESC"
      
      Set objOldAssignment = Application("Connection").Execute(strSQL)
     
      If NOT objOldAssignment.EOF Then
         strOpenAssignmentTime = DateDiff("n",objOldAssignment(1),Date())
         strOpenAssignmentTime = strOpenAssignmentTime + DateDiff("n",objOldAssignment(2),Time())
      
         strSQL = "UPDATE Log" & vbCRLF
         strSQL = strSQL & "SET TaskTime='" & strOpenAssignmentTime & "'" & vbCRLF
         strSQL = strSQL & "WHERE ID=" & objOldAssignment(0)
         Application("Connection").Execute(strSQL)
         
      End If
      
      'Build the SQL string that will remove from the database who is tracking the ticket
      strSQL = "DELETE FROM Tracking" & vbCRLF
      strSQL = strSQL & "WHERE Ticket=" & intID
      Application("Connection").Execute(strSQL)
      
      'Set the ticket as viewed
      strSQL = "UPDATE Main SET TicketViewed=True WHERE ID=" & intID
      Application("Connection").Execute(strSQL)
      
      bolCallClosed = True
   End If
   
   'Build the SQL string that will update the data in the database
   strSQL = "Update Main" & vbCRLF
   strSQL = strSQL & "Set Name = '" & strNewUserName & "',EMail = '" & strNewEMail & "',Location = '" & _
   strNewLocation & "',Category = '" & strNewCategory & "',Notes = '" & _
   strNewNotes & "',Status = '" & strStatus & "',Tech = '" & strTech & "',LastUpdatedDate = '" & _ 
   strDate & "',LastUpdatedTime = '" & strTime & "',OpenTime = '" & strOpenTime & "'" & vbCRLF
   strSQL = strSQL & "WHERE (((Main.ID)=" & intID & "));"

   Application("Connection").Execute(strSQL)
   bolUpdated = True
   
   If strOldTech <> strTech And strStatus <> "Complete" Then   
      API
   End If
   
   Call Main
End Sub%>

<%Sub UpdateLog
   
   Dim objOldData, strCurrentUser, strOldLocation, strOldEmail, strOldCategory
   Dim strOldTech, strOldStatus, strOldNotes, objOldAssignment, strOpenAssignmentTime
   
   On Error Resume Next
     
   strCurrentUser = strUser
   
   'Build the SQL string that will get the old data from the database
   strSQL = "SELECT Location,EMail,Category,Tech,Status,Notes" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE ID=" & intID
   
   Set objOldData = Application("Connection").Execute(strSQL)
   
   strOldLocation = objOldData(0)
   strOldEmail = objOldData(1)
   strOldCategory = objOldData(2)
   strOldTech = objOldData(3)
   strOldStatus = objOldData(4)
   strOldNotes = objOldData(5)
   
   'See what has changed so it can be logged
   If strOldTech <> strTech Then
      strSQL = "SELECT ID,UpdateDate,UpdateTime" & vbCRLF
      strSQL = strSQL & "FROM Log" & vbCRLF
      strSQL = strSQL & "WHERE NewValue='" & strOldTech & "' AND Ticket=" & intID & " AND (Type='Assigned' Or Type='Auto Assigned' Or Type='New Ticket' Or Type='Tech Reassigned' Or NewValue='Complete' or Type='Tech Changed')" & vbCRLF 
      strSQL = strSQL & "ORDER BY ID DESC"
      
      Set objOldAssignment = Application("Connection").Execute(strSQL)
      
      If NOT objOldAssignment.EOF Then
         strOpenAssignmentTime = DateDiff("n",objOldAssignment(1),Date())
         strOpenAssignmentTime = strOpenAssignmentTime + DateDiff("n",objOldAssignment(2),Time())
      
         strSQL = "UPDATE Log" & vbCRLF
         strSQL = strSQL & "SET TaskTime='" & strOpenAssignmentTime & "'" & vbCRLF
         strSQL = strSQL & "WHERE ID=" & objOldAssignment(0)
         Application("Connection").Execute(strSQL)
      End If
      
      If strOldTech = "" Then
         strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
         strSQL = strSQL & "VALUES (" & intID & ",'Assigned','" & strCurrentUser & "','" & strOldTech & "','" & strTech & "','" & Date() & "','" & Time() & "');"
         Application("Connection").Execute(strSQL)
      Else
         strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
         strSQL = strSQL & "VALUES (" & intID & ",'Tech Changed','" & strCurrentUser & "','" & strOldTech & "','" & strTech & "','" & Date() & "','" & Time() & "');"
         Application("Connection").Execute(strSQL)
      End If
      
      strSQL = "UPDATE Main SET TicketViewed=False WHERE ID=" & intID
      Application("Connection").Execute(strSQL)
   End If   
   
   If strOldLocation <> strLocation Then
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'Location Changed','" & strCurrentUser & "','" & strOldLocation & "','" & strLocation & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)
   End If
   
   If strOldEmail <> strEMailTemp Then
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'EMail Changed','" & strCurrentUser & "','" & strOldEMail & "','" & strEMailTemp & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)
   End If
   
   If strOldCategory <> strCategory Then
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'Category Changed','" & strCurrentUser & "','" & strOldCategory & "','" & strCategory & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)
   End If
   
   If (strOldNotes <> strNotes) Or (Len(strNotes) > 1 And IsNull(strOldNotes)) Then
      If IsNull(strOldNotes) Then
         strOldNotes = ""
      End If
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'Notes Updated','" & strCurrentUser & "','" & "" & "','" & "" & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)
   End If
   
   If strOldStatus <> strStatus Then
      strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,NewValue,UpdateDate,UpdateTime)"
      strSQL = strSQL & "VALUES (" & intID & ",'Status Changed','" & strCurrentUser & "','" & strOldStatus & "','" & strStatus & "','" & Date() & "','" & Time() & "');"
      Application("Connection").Execute(strSQL)
   End If
   
End Sub%>

<%Sub UpdateTech
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strFromEMail, objMessageText
   Dim strSubject, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")  
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strCurrentUser = "Heat Help Desk"
      Case "TPERKINS"
         strCurrentUser = "Tech Services"
      Case Else
         strCurrentUser = GetFirstandLastName(strUser)
   End Select
   
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strCurrentUser & "'"
 
   Set objCurrentUser = Application("Connection").Execute(strSQL)
   
   If objCurrentUser.EOF Then
      strCurrentUserEmail = Application("SendFromEMail")
   Else
      strCurrentUserEmail = objCurrentUser(0)
   End If

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
   Else
      strTechEmail = objTechSet(0)
   End If

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Update Tech'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   With objMessage
      .To = strTechEmail
      .From = strCurrentUserEmail 
      .Subject = strSubject
      .TextBody = strMessage
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,NewValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Tech Notified','" & strUser & "','" & strTech & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub%> 

<%Sub UpdateUser

   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objMessageText
   Dim strSubject, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strCurrentUser = "Heat Help Desk"
      Case "TPERKINS"
         strCurrentUser = "Tech Services"
      Case Else
         strCurrentUser = GetFirstandLastName(strUser)
   End Select

   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strCurrentUser & "'"
 
   Set objCurrentUser = Application("Connection").Execute(strSQL)
   
   If objCurrentUser.EOF Then
      strCurrentUserEmail = Application("SendFromEMail")
   Else
      strCurrentUserEmail = objCurrentUser(0)
   End If

   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Update User'"
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
   strMessage = Replace(strMessage,"#NOTES#",HideText(strNotes))
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)

   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
 
   Set objTechSet = Application("Connection").Execute(strSQL)
   
   If objTechSet.EOF Then
      strCCEMail = ""
   Else
      strCCEMail = objTechSet(0)
   End If

   With objMessage
      .To = strUserEMail
      If strCCEMail <> "" Then
         .CC = strCCEMail
      End If
      .From = strCurrentUserEmail 
      .Subject = strSubject
      .TextBody = strMessage
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'User Notified','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub%> 

<%Sub SendTicket
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strFromEMail, objMessageText
   Dim strSubject, strEMail, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2") 
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strCurrentUser = "Heat Help Desk"
      Case "TPERKINS"
         strCurrentUser = "Tech Services"
      Case Else
         strCurrentUser = GetFirstandLastName(strUser)
   End Select
   
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strCurrentUser & "'"
 
   Set objCurrentUser = Application("Connection").Execute(strSQL)
   
   If objCurrentUser.EOF Then
      strCurrentUserEmail = Application("SendFromEMail")
   Else
      strCurrentUserEmail = objCurrentUser(0)
   End If

   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration

   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   strEMail = Upload.Form("SendEMail")

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Send Ticket'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   With objMessage
      .To = strEMail
      .From = strCurrentUserEmail 
      .Subject = strSubject
      .TextBody = strMessage
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,NewValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Ticket EMailed','" & strUser & "','" & strEMail & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub%> 

<%Sub RequestUpdate

   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objMessageText
   Dim strSubject, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")
   strCurrentUser = GetFirstandLastName(strUser)
   
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strCurrentUser & "'"
 
   Set objCurrentUser = Application("Connection").Execute(strSQL)
   
   If objCurrentUser.EOF Then
      strCurrentUserEmail = Application("SendFromEMail")
   Else
      strCurrentUserEmail = objCurrentUser(0)
   End If
  
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Request for Update'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
 
   Set objTechSet = Application("Connection").Execute(strSQL)
   
   If objTechSet.EOF Then
      strTechEMail = ""
   Else
      strTechEMail = objTechSet(0)
   End If
   
   With objMessage
      .To = strTechEMail
      .From = strCurrentUserEmail 
      .CC = strCurrentUserEmail
      .Subject = strSubject
      .TextBody = strMessage
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   'Build the SQL string that will write to the database who is requesting the update.
   strSQL = "INSERT INTO Tracking (Ticket,Type,TrackedBy)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Request','" & strUser & "')" & vbCRLF
   Application("Connection").Execute(strSQL)
   
   'Build the SQL string that will update the log saying who requested the update
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,OldValue,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Update Requested','" & strUser & "','" & strTech & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
   Set objConf = Nothing
   Set objMessage = Nothing

End Sub%> 

<%Sub TicketAssigned
   '*****************************************************************************************
   'Send the tech an email if the ticket was just assigned to them.
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objOldTechSet, objFSO
   Dim objFolder, objFile, strAttachment, strAttachmentTech, objMessageText, strSubject, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")   
   
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'See if there is an attachement
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID)
      For Each objFile in objFolder.Files
         strAttachment = objFile.Path
      Next
   End If
   
   'See if there is an attachement added by the tech
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID & "-Tech")
      For Each objFile in objFolder.Files
         strAttachmentTech = objFile.Path
      Next
   End If
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With

   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
 
   Set objTechSet = Application("Connection").Execute(strSQL)
   strTechEmail = objTechSet(0)
   
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Ticket Assigned'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   With objMessage
      .To = objTechSet(0)
      .From = Application("SendFromEMail") 
      .Subject = strSubject
      .TextBody = strMessage
      If strAttachment <> "" Then
         .AddAttachment strAttachment
      End If
      If strAttachmentTech <> "" Then
         .AddAttachment strAttachmentTech
      End If
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub%>

<%Sub TicketReassigned

   'Send the old tech an email if the ticket was just assigned to someone else.
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objOldTechSet
   Dim strOldTechEmail, objMessageText, strSubject, strName
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")
   
   'This code will fix the display name so it matches what is in the database.
   Select Case UCase(strUser)
      Case "HELPDESK"
         strCurrentUser = "Heat Help Desk"
      Case "TPERKINS"
         strCurrentUser = "Tech Services"
      Case Else
         strCurrentUser = GetFirstandLastName(strUser)
   End Select
   
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strCurrentUser & "'"
 
   Set objCurrentUser = Application("Connection").Execute(strSQL)
   
   If objCurrentUser.EOF Then
      strCurrentUserEmail = Application("SendFromEMail")
   Else
      strCurrentUserEmail = objCurrentUser(0)
   End If

   'Get the old tech's email
   strSQL = "Select Main.ID, Tech.EMail" & vbCRLF
   strSQL = strSQL & "From Main INNER JOIN Tech ON Main.Tech = Tech.Tech" & vbCRLF
   strSQL = strSQL & "Where ((Main.ID)=" & intID & ");"
 
   Set objOldTechSet = Application("Connection").Execute(strSQL)
   strOldTechEmail = objOldTechSet(1)

   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Ticket Reassigned'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)

   With objMessage
      .To = strOldTechEmail
      .From = Application("SendFromEMail") 
      .Subject = strSubject
      .TextBody = strMessage
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   Set objConf = Nothing
   Set objMessage = Nothing

End Sub%>

<%Sub TicketClosed
   
   'Send the user an email if the ticket is closed.
   
   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objOldTechSet, objFSO, strName
   Dim objFolder, objFile, strAttachment, strAttachmentTech, objMessageText, strSubject
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")   
   
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'See if there is an attachement
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID)
      For Each objFile in objFolder.Files
         strAttachment = objFile.Path
      Next
   End If
   
   'See if there is an attachement added by the tech
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID & "-Tech")
      For Each objFile in objFolder.Files
         strAttachmentTech = objFile.Path
      Next
   End If

   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
  
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Ticket Closed'"
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
   strMessage = Replace(strMessage,"#NOTES#",HideText(strNotes))
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)

   With objMessage
      .To = strUserEMail 
      .From = Application("SendFromEMail") 
      .Subject = strSubject
      .TextBody = strMessage
      If strAttachmentTech <> "" Then
         .AddAttachment strAttachmentTech
      End If
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub %>

<%Sub SendRequestedUpdate

   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objFSO, strName
   Dim objFolder, objFile, strAttachment, strAttachmentTech, objTrackingSet, objMessageText
   Dim strSubject
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")

   Set objMessage = CreateObject("CDO.Message")
   
   'See if there is an attachement
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID)
      For Each objFile in objFolder.Files
         strAttachment = objFile.Path
      Next
   End If
   
   'See if there is an attachement added by the tech
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID & "-Tech")
      For Each objFile in objFolder.Files
         strAttachmentTech = objFile.Path
      Next
   End If
   
   strSQL = "SELECT TrackedBy FROM Tracking WHERE Ticket=" & intID & " And Type='Request'"
   Set objTracking = Application("Connection").Execute(strSQL)  
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Requested Update'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   Do Until objTracking.EOF
      With objMessage
         .To = objTracking(0) & Application("EMailSuffix")
         .From = Application("SendFromEMail") 
         .Subject = strSubject
         .TextBody = strMessage
         If strAttachment <> "" Then
            .AddAttachment strAttachment
         End If
         If strAttachmentTech <> "" Then
            .AddAttachment strAttachmentTech
         End If
         If Application("BCC") <> "" Then
            .BCC = Application("BCC")
         End If
         .Send
      End With
      objTracking.MoveNext
   Loop
      
   'Build the SQL string that will remove from the database who is requesting the update.
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE (Ticket=" & intID & " And Type='Request')"
   Application("Connection").Execute(strSQL)   
   
   'Build the SQL string that will update the log saying the update request has been answered
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Request Update Complete','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)

   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub %>

<%Sub CancelUpdateRequest

   Dim intID, strSQL

   intID = request.querystring("ID")

   'Build the SQL string that will remove from the database who is requesting the update.
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE (Ticket=" & intID & " And TrackedBy='" & strUser & "' And Type='Request')"
   Application("Connection").Execute(strSQL)   
   
   'Build the SQL string that will update the log saying the update request has been answered
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Cancelled Update Request','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)
   
End Sub%>

<%Sub TrackTicket  

   intID = Request.QueryString("ID")
  
   'Build the SQL string that will write to the database who is requesting the update.
   strSQL = "INSERT INTO Tracking (Ticket,Type,TrackedBy)" & vbCRLF
   strSQL = strSQL & "VALUES (" & intID & ",'Track','" & strUser & "')" & vbCRLF

   Application("Connection").Execute(strSQL)
  
   'Update the log
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Ticket Tracked','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)

End Sub%> 

<%Sub DontTrackTicket

   intID = Request.QueryString("ID")
  
   'Build the SQL string that will remove from the database who is requesting the update.
   strSQL = "DELETE FROM Tracking" & vbCRLF
   strSQL = strSQL & "WHERE (Ticket=" & intID & " And TrackedBy='" & strUser & "' And Type='Track')"
   Application("Connection").Execute(strSQL)
  
   'Update the log
   strSQL = "INSERT INTO Log (Ticket,Type,ChangedBy,UpdateDate,UpdateTime)"
   strSQL = strSQL & "VALUES (" & intID & ",'Ticket Not Tracked','" & strUser & "','" & Date() & "','" & Time() & "');"
   Application("Connection").Execute(strSQL)

End Sub%> 

<%Sub SendTrackingEMail

   Const cdoSendUsingPickup = 1
   
   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strStatus
   Dim strTech, strCustom1, strCustom2, strCurrentUser, objCurrentUser, strCurrentUserEmail
   Dim objMessage, objConf, strMessage, objTechSet, strCCEMail, objFSO, strName
   Dim objFolder, objFile, strAttachment, strAttachmentTech, objTracking, objMessageText
   Dim strSubject
   
   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")

   Set objMessage = CreateObject("CDO.Message")
   
   'See if there is an attachement
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID) Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID)
      For Each objFile in objFolder.Files
         strAttachment = objFile.Path
      Next
   End If
   
   'See if there is an attachement added by the tech
   If objFSO.FolderExists(Application("FileLocation") & "\" & intID & "-Tech") Then
      Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID & "-Tech")
      For Each objFile in objFolder.Files
         strAttachmentTech = objFile.Path
      Next
   End If
   
   strSQL = "SELECT TrackedBy FROM Tracking WHERE Ticket=" & intID & " And Type='Track'"
   Set objTracking = Application("Connection").Execute(strSQL)
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
   
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With

   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Tracking Update'"
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
   strMessage = Replace(strMessage,"#NOTES#",strNotes)
   strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
   
   strSubject = objMessageText(0)
   strSubject = Replace(strSubject,"#TICKET#",intID)
   
   Do Until objTracking.EOF
      With objMessage
         .To = objTracking(0) & Application("EMailSuffix")
         .From = Application("SendFromEMail") 
         .Subject = strSubject
         .TextBody = strMessage
         If strAttachment <> "" Then
            .AddAttachment strAttachment
         End If
         If strAttachmentTech <> "" Then
            .AddAttachment strAttachmentTech
         End If
         If Application("BCC") <> "" Then
            .BCC = Application("BCC")
         End If
         .Send
      End With
      objTracking.MoveNext
   Loop
   
   If strStatus = "Complete" Then
      'Build the SQL string that will remove from the database who is tracking the ticket
      strSQL = "DELETE FROM Tracking" & vbCRLF
      strSQL = strSQL & "WHERE Ticket=" & intID
      Application("Connection").Execute(strSQL)
   End If
   
   Set objConf = Nothing
   Set objMessage = Nothing
   
End Sub%>

<%Sub API

   Dim intID, strUserEMail, strLocation, strCategory, strProblem, strNotes, strTech
   Dim strCustom1, strCustom2, strName

   intID = request.querystring("ID")
   strName = Upload.Form("Name")
   strUserEMail = Upload.Form("EMail")
   strLocation = Upload.Form("Location")
   strCategory = Upload.Form("Category")
   strProblem = Upload.Form("Problem")
   strNotes = Upload.Form("Notes")
   strStatus = Upload.Form("Status")
   strTech = Upload.Form("Tech")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")

   If Application("TSCAPI") <> "" And Application("TSCTech") = strTech Then %>
      
      <form id="api" method="POST" action="<%=Application("TSCAPI")%>" target="_blank">
         <input type="hidden" name="Name" value="auto" />
         <input type="hidden" name="DisplayName" value="<%=strName%>" />
         <input type="hidden" name="EMail" value="<%=Application("SendFromEMail")%>" />
         <input type="hidden" name="Location" value="<%=Application("TSCSite")%>" />
         <input type="hidden" name="Problem" value="<%=strProblem%>" />
         <input type="hidden" name="Notes" value="<%=strNotes%>" />
         <input type="hidden" name="Category" value="<%=strCategory%>" />
         <input type="hidden" name="Custom1" value="<%=strCustom1%>" />
         <input type="hidden" name="Custom2" value="<%=strCustom2%>" />
         <input type="hidden" name="TicketNumber" value="<%=intID%>" />
      </form>
      
      <script type="text/javascript">
          function autosubmitform () {
              var frm = document.getElementById("api");
              frm.submit();
          }
          window.onload = autosubmitform;
      </script>
      
<% End If 

End Sub %>

<%Sub AccessDenied 

   If bolShowLogout Then
      Response.Redirect("login.asp?action=logout")
   Else
   %>

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
      <center><b>Access Denied</b></center>
   </body>
   </html>
   
<% End If

End Sub%>

<%'IsEMailValid
' Source http://www.aspfree.com/c/a/ASP-Code/VBScript-function-to-validate-Email-Addresses/
' Function IsEmailValid(strEmail)
' Action: checks if an email is correct.
' Parameter: strEmail - the Email address
' Returned value: on success it returns True, else False.
Function IsEmailValid(strEmail)
 
    Dim strArray
    Dim strItem
    Dim i
    Dim c
    Dim blnIsItValid
 
    ' assume the email address is correct 
    blnIsItValid = True
   
    ' split the email address in two parts: name@domain.ext
    strArray = Split(strEmail, "@")
 
    ' if there are more or less than two parts 
    If UBound(strArray) <> 1 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' check each part
    For Each strItem In strArray
        ' no part can be void
        If Len(strItem) <= 0 Then
            blnIsItValid = False
            IsEmailValid = blnIsItValid
            Exit Function
        End If
       
        ' check each character of the part
        ' only following "abcdefghijklmnopqrstuvwxyz_-."
        ' characters and the ten digits are allowed
        For i = 1 To Len(strItem)
               c = LCase(Mid(strItem, i, 1))
               ' if there is an illegal character in the part
               If InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 And Not IsNumeric(c) Then
                   blnIsItValid = False
                   IsEmailValid = blnIsItValid
                   Exit Function
               End If
        Next
  
      ' the first and the last character in the part cannot be . (dot)
        If Left(strItem, 1) = "." Or Right(strItem, 1) = "." Then
           blnIsItValid = False
           IsEmailValid = blnIsItValid
           Exit Function
        End If
    Next
 
    ' the second part (domain.ext) must contain a . (dot)
    If InStr(strArray(1), ".") <= 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' check the length oh the extension 
    i = Len(strArray(1)) - InStrRev(strArray(1), ".")
    ' the length of the extension can be only 2, 3, or 4
    ' to cover the new "info" extension
    If i <> 2 And i <> 3 And i <> 4 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If

    ' after . (dot) cannot follow a . (dot)
    If InStr(strEmail, "..") > 0 Then
        blnIsItValid = False
        IsEmailValid = blnIsItValid
        Exit Function
    End If
 
    ' finally it's OK 
    IsEmailValid = blnIsItValid
   
 End Function
%>

<%'FreeASPUpload
'  For examples, documentation, and your own free copy, go to:
'  http://www.freeaspupload.net
'  Note: You can copy and use this script for free and you can make changes
'  to the code, but you cannot remove the above comment.

'Changes:
'Aug 2, 2005: Add support for checkboxes and other input elements with multiple values
'Jan 6, 2009: Lars added ASP_CHUNK_SIZE
'Sep 3, 2010: Enforce UTF-8 everywhere; new function to convert byte array to unicode string

const DEFAULT_ASP_CHUNK_SIZE = 200000

const adModeReadWrite = 3
const adTypeBinary = 1
const adTypeText = 2
const adSaveCreateOverWrite = 2

Class FreeASPUpload
	Public UploadedFiles
	Public FormElements

	Private VarArrayBinRequest
	Private StreamRequest
	Private uploadedYet
	Private internalChunkSize

	Private Sub Class_Initialize()
		Set UploadedFiles = Server.CreateObject("Scripting.Dictionary")
		Set FormElements = Server.CreateObject("Scripting.Dictionary")
		Set StreamRequest = Server.CreateObject("ADODB.Stream")
		StreamRequest.Type = adTypeText
		StreamRequest.Open
		uploadedYet = false
		internalChunkSize = DEFAULT_ASP_CHUNK_SIZE
	End Sub
	
	Private Sub Class_Terminate()
		If IsObject(UploadedFiles) Then
			UploadedFiles.RemoveAll()
			Set UploadedFiles = Nothing
		End If
		If IsObject(FormElements) Then
			FormElements.RemoveAll()
			Set FormElements = Nothing
		End If
		StreamRequest.Close
		Set StreamRequest = Nothing
	End Sub

	Public Property Get Form(sIndex)
		Form = ""
		If FormElements.Exists(LCase(sIndex)) Then Form = FormElements.Item(LCase(sIndex))
	End Property

	Public Property Get Files()
		Files = UploadedFiles.Items
	End Property
	
    Public Property Get Exists(sIndex)
            Exists = false
            If FormElements.Exists(LCase(sIndex)) Then Exists = true
    End Property
        
    Public Property Get FileExists(sIndex)
        FileExists = false
            if UploadedFiles.Exists(LCase(sIndex)) then FileExists = true
    End Property
        
    Public Property Get chunkSize()
		chunkSize = internalChunkSize
	End Property

	Public Property Let chunkSize(sz)
		internalChunkSize = sz
	End Property

	'Calls Upload to extract the data from the binary request and then saves the uploaded files
	Public Sub Save(path)
		Dim streamFile, fileItem, filePath

		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload

		For Each fileItem In UploadedFiles.Items
			filePath = path & fileItem.FileName
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position=fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile filePath, adSaveCreateOverWrite
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = filePath
		 Next
	End Sub
	
	public sub SaveOne(path, num, byref outFileName, byref outLocalFileName)
		Dim streamFile, fileItems, fileItem, fs

        set fs = Server.CreateObject("Scripting.FileSystemObject")
		if Right(path, 1) <> "\" then path = path & "\"

		if not uploadedYet then Upload
		if UploadedFiles.Count > 0 then
			fileItems = UploadedFiles.Items
			set fileItem = fileItems(num)
		
			outFileName = fileItem.FileName
			outLocalFileName = GetFileName(path, outFileName)
		
			Set streamFile = Server.CreateObject("ADODB.Stream")
			streamFile.Type = adTypeBinary
			streamFile.Open
			StreamRequest.Position = fileItem.Start
			StreamRequest.CopyTo streamFile, fileItem.Length
			streamFile.SaveToFile path & outLocalFileName, adSaveCreateOverWrite
			streamFile.close
			Set streamFile = Nothing
			fileItem.Path = path & filename
		end if
	end sub

	Public Function SaveBinRequest(path) ' For debugging purposes
		StreamRequest.SaveToFile path & "\debugStream.bin", 2
	End Function

	Public Sub DumpData() 'only works if files are plain text
		Dim i, aKeys, f
		response.write "Form Items:<br>"
		aKeys = FormElements.Keys
		For i = 0 To FormElements.Count -1 ' Iterate the array
			response.write aKeys(i) & " = " & FormElements.Item(aKeys(i)) & "<BR>"
		Next
		response.write "Uploaded Files:<br>"
		For Each f In UploadedFiles.Items
			response.write "Name: " & f.FileName & "<br>"
			response.write "Type: " & f.ContentType & "<br>"
			response.write "Start: " & f.Start & "<br>"
			response.write "Size: " & f.Length & "<br>"
		 Next
   	End Sub

	Public Sub Upload()
		Dim nCurPos, nDataBoundPos, nLastSepPos
		Dim nPosFile, nPosBound
		Dim sFieldName, osPathSep, auxStr
		Dim readBytes, readLoop, tmpBinRequest
		
		'RFC1867 Tokens
		Dim vDataSep
		Dim tNewLine, tDoubleQuotes, tTerm, tFilename, tName, tContentDisp, tContentType
		tNewLine = String2Byte(Chr(13))
		tDoubleQuotes = String2Byte(Chr(34))
		tTerm = String2Byte("--")
		tFilename = String2Byte("filename=""")
		tName = String2Byte("name=""")
		tContentDisp = String2Byte("Content-Disposition")
		tContentType = String2Byte("Content-Type:")

		uploadedYet = true

		On Error resume next
			' Copy binary request to a byte array, on which functions like InstrB and others can be used to search for separation tokens
			readBytes = internalChunkSize
			VarArrayBinRequest = Request.BinaryRead(readBytes)
			VarArrayBinRequest = midb(VarArrayBinRequest, 1, lenb(VarArrayBinRequest))
			Do Until readBytes < 1
				tmpBinRequest = Request.BinaryRead(readBytes)
				if readBytes > 0 then
					VarArrayBinRequest = VarArrayBinRequest & midb(tmpBinRequest, 1, lenb(tmpBinRequest))
				end if
			Loop
			StreamRequest.WriteText(VarArrayBinRequest)
			StreamRequest.Flush()
			if Err.Number <> 0 then 
				response.write "<br><br><B>System reported this error:</B><p>"
				response.write Err.Description & "<p>"
				response.write "The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
				Exit Sub
			end if
		On Error goto 0 'reset error handling

		nCurPos = FindToken(tNewLine,1) 'Note: nCurPos is 1-based (and so is InstrB, MidB, etc)

		If nCurPos <= 1  Then Exit Sub
		 
		'vDataSep is a separator like -----------------------------21763138716045
		vDataSep = MidB(VarArrayBinRequest, 1, nCurPos-1)

		'Start of current separator
		nDataBoundPos = 1

		'Beginning of last line
		nLastSepPos = FindToken(vDataSep & tTerm, 1)

		Do Until nDataBoundPos = nLastSepPos
			
			nCurPos = SkipToken(tContentDisp, nDataBoundPos)
			nCurPos = SkipToken(tName, nCurPos)
			sFieldName = ExtractField(tDoubleQuotes, nCurPos)

			nPosFile = FindToken(tFilename, nCurPos)
			nPosBound = FindToken(vDataSep, nCurPos)
			
			If nPosFile <> 0 And  nPosFile < nPosBound Then
				Dim oUploadFile
				Set oUploadFile = New UploadedFile
				
				nCurPos = SkipToken(tFilename, nCurPos)
				auxStr = ExtractField(tDoubleQuotes, nCurPos)
                ' We are interested only in the name of the file, not the whole path
                ' Path separator is \ in windows, / in UNIX
                ' While IE seems to put the whole pathname in the stream, Mozilla seem to 
                ' only put the actual file name, so UNIX paths may be rare. But not impossible.
                osPathSep = "\"
                if InStr(auxStr, osPathSep) = 0 then osPathSep = "/"
				oUploadFile.FileName = Right(auxStr, Len(auxStr)-InStrRev(auxStr, osPathSep))

				if (Len(oUploadFile.FileName) > 0) then 'File field not left empty
					nCurPos = SkipToken(tContentType, nCurPos)
					
                    auxStr = ExtractField(tNewLine, nCurPos)
                    ' NN on UNIX puts things like this in the stream:
                    '    ?? python py type=?? python application/x-python
					oUploadFile.ContentType = Right(auxStr, Len(auxStr)-InStrRev(auxStr, " "))
					nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
					
					oUploadFile.Start = nCurPos+1
					oUploadFile.Length = FindToken(vDataSep, nCurPos) - 2 - nCurPos
					
					If oUploadFile.Length > 0 Then UploadedFiles.Add LCase(sFieldName), oUploadFile
				End If
			Else
				Dim nEndOfData, fieldValueUniStr
				nCurPos = FindToken(tNewLine, nCurPos) + 4 'skip empty line
				nEndOfData = FindToken(vDataSep, nCurPos) - 2
				fieldValueuniStr = ConvertUtf8BytesToString(nCurPos, nEndOfData-nCurPos)
				If Not FormElements.Exists(LCase(sFieldName)) Then 
					FormElements.Add LCase(sFieldName), fieldValueuniStr
				else
                    FormElements.Item(LCase(sFieldName))= FormElements.Item(LCase(sFieldName)) & ", " & fieldValueuniStr
                end if 

			End If

			'Advance to next separator
			nDataBoundPos = FindToken(vDataSep, nCurPos)
		Loop
	End Sub

	Private Function SkipToken(sToken, nStart)
		SkipToken = InstrB(nStart, VarArrayBinRequest, sToken)
		If SkipToken = 0 then
			Response.write "Error in parsing uploaded binary request. The most likely cause for this error is the incorrect setup of AspMaxRequestEntityAllowed in IIS MetaBase. Please see instructions in the <A HREF='http://www.freeaspupload.net/freeaspupload/requirements.asp'>requirements page of freeaspupload.net</A>.<p>"
			Response.End
		end if
		SkipToken = SkipToken + LenB(sToken)
	End Function

	Private Function FindToken(sToken, nStart)
		FindToken = InstrB(nStart, VarArrayBinRequest, sToken)
	End Function

	Private Function ExtractField(sToken, nStart)
		Dim nEnd
		nEnd = InstrB(nStart, VarArrayBinRequest, sToken)
		If nEnd = 0 then
			Response.write "Error in parsing uploaded binary request."
			Response.End
		end if
		ExtractField = ConvertUtf8BytesToString(nStart, nEnd-nStart)
	End Function

	'String to byte string conversion
	Private Function String2Byte(sString)
		Dim i
		For i = 1 to Len(sString)
		   String2Byte = String2Byte & ChrB(AscB(Mid(sString,i,1)))
		Next
	End Function

	Private Function ConvertUtf8BytesToString(start, length)	
		StreamRequest.Position = 0
	
	    Dim objStream
	    Dim strTmp
	    
	    ' init stream
	    Set objStream = Server.CreateObject("ADODB.Stream")
	    objStream.Charset = "utf-8"
	    objStream.Mode = adModeReadWrite
	    objStream.Type = adTypeBinary
	    objStream.Open
	    
	    ' write bytes into stream
	    StreamRequest.Position = start+1
	    StreamRequest.CopyTo objStream, length
	    objStream.Flush
	    
	    ' rewind stream and read text
	    objStream.Position = 0
	    objStream.Type = adTypeText
	    strTmp = objStream.ReadText
	    
	    ' close up and return
	    objStream.Close
	    Set objStream = Nothing
	    ConvertUtf8BytesToString = strTmp	
	End Function
End Class

Class UploadedFile
	Public ContentType
	Public Start
	Public Length
	Public Path
	Private nameOfFile

    ' Need to remove characters that are valid in UNIX, but not in Windows
    Public Property Let FileName(fN)
        nameOfFile = fN
        nameOfFile = SubstNoReg(nameOfFile, "\", "_")
        nameOfFile = SubstNoReg(nameOfFile, "/", "_")
        nameOfFile = SubstNoReg(nameOfFile, ":", "_")
        nameOfFile = SubstNoReg(nameOfFile, "*", "_")
        nameOfFile = SubstNoReg(nameOfFile, "?", "_")
        nameOfFile = SubstNoReg(nameOfFile, """", "_")
        nameOfFile = SubstNoReg(nameOfFile, "<", "_")
        nameOfFile = SubstNoReg(nameOfFile, ">", "_")
        nameOfFile = SubstNoReg(nameOfFile, "|", "_")
    End Property

    Public Property Get FileName()
        FileName = nameOfFile
    End Property

    'Public Property Get FileN()ame
End Class


' Does not depend on RegEx, which is not available on older VBScript
' Is not recursive, which means it will not run out of stack space
Function SubstNoReg(initialStr, oldStr, newStr)
    Dim currentPos, oldStrPos, skip
    If IsNull(initialStr) Or Len(initialStr) = 0 Then
        SubstNoReg = ""
    ElseIf IsNull(oldStr) Or Len(oldStr) = 0 Then
        SubstNoReg = initialStr
    Else
        If IsNull(newStr) Then newStr = ""
        currentPos = 1
        oldStrPos = 0
        SubstNoReg = ""
        skip = Len(oldStr)
        Do While currentPos <= Len(initialStr)
            oldStrPos = InStr(currentPos, initialStr, oldStr)
            If oldStrPos = 0 Then
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, Len(initialStr) - currentPos + 1)
                currentPos = Len(initialStr) + 1
            Else
                SubstNoReg = SubstNoReg & Mid(initialStr, currentPos, oldStrPos - currentPos) & newStr
                currentPos = oldStrPos + skip
            End If
        Loop
    End If
End Function

Function GetFileName(strSaveToPath, FileName)
'This function is used when saving a file to check there is not already a file with the same name so that you don't overwrite it.
'It adds numbers to the filename e.g. file.gif becomes file1.gif becomes file2.gif and so on.
'It keeps going until it returns a filename that does not exist.
'You could just create a filename from the ID field but that means writing the record - and it still might exist!
'N.B. Requires strSaveToPath variable to be available - and containing the path to save to
    Dim Counter
    Dim Flag
    Dim strTempFileName
    Dim FileExt
    Dim NewFullPath
    dim objFSO, p
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Counter = 0
    p = instrrev(FileName, ".")
    FileExt = mid(FileName, p+1)
    strTempFileName = left(FileName, p-1)
    NewFullPath = strSaveToPath & "\" & FileName
    Flag = False
    
    Do Until Flag = True
        If objFSO.FileExists(NewFullPath) = False Then
            Flag = True
            GetFileName = Mid(NewFullPath, InstrRev(NewFullPath, "\") + 1)
        Else
            Counter = Counter + 1
            NewFullPath = strSaveToPath & "\" & strTempFileName & Counter & "." & FileExt
        End If
    Loop
End Function 

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
      
      'If a session isn't found for then kick them out
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
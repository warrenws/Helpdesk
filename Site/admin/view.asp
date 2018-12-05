<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/12/04
'Last Updated 6/16/14

'This page will list all the calls from the requested query and display them to the user.
'There is an icon to the left that the user can click on to edit that call.

Option Explicit

On Error Resume Next

Dim strFilter, strLocation, strStatus, strTech, strSQL, strTitle, strSQLLocation
Dim objRecordSet, strIcon, intIconSpan, strCategory, strSQLStatus, strSQLTech, strDateSQL
Dim strMessage, intRecordCount, strSQLCategory, objRegExp, strNotes, strSort, strSQLSort
Dim strUser, objNetwork, strProblem, strSQLUser, strSearchField, strSearchString
Dim bolSearch, strSearchSQL, objNameCheckSet, strReturnLink, strEMail, bolMobileVersion
Dim strSQLProblem, strSQLNotes, strSQLEMail, strStatusBar, strViewed, strSQLViewed
Dim strSQLDateRange, intDays, bolDateRange, strUserAgent, strRole, strDate, strType
Dim objTechSet, strTechSQL, bolShowLogout, strSearchUser, intZoom, bolDataSubmitted

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

Response.Buffer = False

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

'Build the SQL string
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

   On Error Resume Next

   'Send you to the home page if home is clicked on the iPhone version
   If Request.Form("home") = "Home" or Request.Form("back") = "Back" Then
      Response.Redirect "index.asp"
   End If
   
   If Request.Form("Log Out") = "Log Out" Then
      Response.Redirect "login.asp?action=logout"
   End If   
   
   'Get the information from the forms and address bar and assign them to variables
   strFilter = Request.QueryString("Filter")
   strSort = Request.QueryString("Sort")
   strLocation = Request.Form("Location")
   strStatus = Request.Form("Status")
   strTech = Request.Form("Tech")
   strCategory = Request.Form("Category")
   strSearchUser = Request.Form("User")
   strProblem = Request.Form("Problem")
   strNotes = Request.Form("Notes")
   strEMail = Request.Form("EMail")
   strViewed = Request.Form("Viewed")
   intDays = Request.Form("Days")
   
   If strFilter = "" Then
      strFilter = Replace(Request.Form("Filter")," ","")
   End If
   
   If intDays = "" or intDays = 0 Then
      bolDateRange = False
   Else
      bolDateRange = True
   End If
   
   'If the information wasn't passed in the form look at the URL
   If strLocation = "" Then
      If Request.QueryString("Location") = "" Then
         strLocation = "Any"
      Else
         strLocation = Request.QueryString("Location")
      End If
   End If
   If strStatus = "" Then
      If Request.QueryString("Status") = "" Then
         strStatus = "Any"
      Else
         strStatus = Request.QueryString("Status")
      End If
   End If
   If strTech = "" Then
      If Request.QueryString("Tech") = "" Then
         strTech = "Any"
      Else
         strTech = Request.QueryString("Tech")
      End If
   End If
   If strCategory = "" Then
      If Request.QueryString("Category") = "" Then
         strCategory = "Any"
      Else
         strCategory = Request.QueryString("Category")
      End If
   End If
   If strSearchUser = "" Then
      If Request.QueryString("User") = "" Then
         strSearchUser = "Any"
      Else
         strSearchUser = Request.QueryString("User")
      End If
   End If
   If strProblem = "" Then
      If Request.QueryString("Problem") = "" Then
         strProblem = "Any"
      Else
         strProblem = Request.QueryString("Problem")
      End If
   End If
   If strNotes = "" Then
      If Request.QueryString("Notes") = "" Then
         strNotes = "Any"
      Else
         strNotes = Request.QueryString("Notes")
      End If
   End If
   If strEMail = "" Then
      If Request.QueryString("EMail") = "" Then
         strEMail = "Any"
      Else
         strEMail = Request.QueryString("EMail")
      End If
   End If
   If strSort = "" Then
      If Request.Form("Sort") = "" Then
         strSort = ""
      Else
         strSort = Request.Form("Sort")
      End If
   End If
   If intDays = "" Then
      If Request.QueryString("Days") = "" Then
         intDays = ""
      Else
         intDays = Request.QueryString("Days")
      End If
   End If
   If strViewed = "" Then
      If Request.QueryString("Viewed") = "" Then
         strViewed = ""
      Else
         strViewed = Request.QueryString("Viewed")
      End If
   End If

   'Build the return portion of the link
   If strLocation <> "Any" Then
      strReturnLink = strReturnLink & "&Location=" & Replace(strLocation," ","%20")
   End If
   If strStatus <> "Any" Then
      strReturnLink = strReturnLink & "&Status=" & Replace(strStatus," ","%20")
   End If
   If strTech <> "Any" Then
      strReturnLink = strReturnLink & "&Tech=" & Replace(strTech," ","%20")
   End If
   If strCategory <> "Any" Then
      strReturnLink = strReturnLink & "&Category=" & Replace(strCategory," ","%20")
   End If
   If strSearchUser <> "Any" Then
      strReturnLink = strReturnLink & "&User=" & Replace(strSearchUser," ","%20")
   End If
   If strFilter <> "Any" Then
      strReturnLink = strReturnLink & "&Filter=" & Replace(strFilter," ","%20")
   End If
   If strProblem <> "Any" Then
      strReturnLink = strReturnLink & "&Problem=" & Replace(strProblem," ","%20")
   End If
   If strNotes <> "Any" Then
      strReturnLink = strReturnLink & "&Notes=" & Replace(strNotes," ","%20")
   End If
   If strEMail <> "Any" Then
      strReturnLink = strReturnLink & "&EMail=" & Replace(strEMail," ","%20")
   End If
   If strSort <> "" Then
      strReturnLink = strReturnLink & "&Sort=" & Replace(strSort," ","%20")
   End If
   If intDays <> "" And intDays <> "0" Then
      strReturnLink = strReturnLink & "&Days=" & intDays
      bolDateRange = True
   Else
      bolDateRange = False
   End If
   If strViewed <> "Any" Then
      strReturnLink = strReturnLink & "&Viewed=" & strViewed
   End If
   
   strReturnLink = strReturnLink & "&Back=Yes"

   'Clean Inputs
   strSearchUser = Replace(strSearchUser,"'","''")
   strProblem = Replace(strProblem,"'","''")
   strNotes = Replace(strNotes,"'","''")
   strEMail = Replace(strEMail,"'","''")

   'This will allow the user to put a sort string in the URL.
   Select Case strSort
      Case ""
         strSort = "Date - Newest on Top"
      Case "DN"
         strSort = "Date - Newest on Top"
      Case "DO"
         strSort = "Date - Oldest on Top"
      Case "LA"
         strSort = "Location - A to Z"
      Case "LZ"
         strSort = "Location - Z to A"
      Case "TA"
         strSort = "Tech - A to Z"
      Case "TZ"
         strSort = "Tech - Z to A"
   End Select

   strTitle = "Help Desk Tickets"
   
   If strFilter = "AllOpenTickets" or strFilter = "OpenTickets" Then
      strStatus = "Any Open Ticket"
      strLocation = "Any"
      strTech = "Any"
      strCategory = "Any"
      strTitle = "All Open Tickets"
   End If

   If strFilter = "MyOpenTickets" or strFilter = "YourOpenTickets" or strFilter = "YourTickets" Then

      'Create the SQL string that will get the users name
      strSQL = "Select Tech.Tech" & vbCRLF
      strSQL = strSQL & "From Tech" & vbCRLF
      strSQL = strSQL & "Where Tech.Username='" &  strUser & "'"

      'Execute the SQL string
      Set objRecordSet = Application("Connection").Execute(strSQL)

      strStatus = "Any Open Ticket"
      strLocation = "Any"
      strTech = objRecordSet(0)
      strCategory = "Any"
      strTitle = "My Open Tickets"

   End If

   'If the user is looking for tickets assigned to no one then set the status to New
   'Assignment and set the title.
   If strTech = "Nobody" Then
      strStatus = "New Assignment"
      strTech = "Any"
      strTitle = "Ticket's Not Assigned to Anyone"
   End If

   'Make sure something was submitted.  If so then figure out what they wanted and run the query
   bolDataSubmitted = False
   If strStatus <> "Any" Then
      bolDataSubmitted = True
   End If
   If strCategory <> "Any" Then
      bolDataSubmitted = True
   End If
   If strTech <> "Any" Then
      bolDataSubmitted = True
   End If
   If strLocation <> "Any" Then
      bolDataSubmitted = True
   End If
   If intDays <> "" And intDays <> "0" Then
      bolDataSubmitted = True
   End If
   If strSearchUser <> "Any" Then
      bolDataSubmitted = True
   End If
   If strProblem <> "Any" Then
      bolDataSubmitted = True
   End If
   If strNotes <> "Any" Then
      bolDataSubmitted = True
   End If
   If strEMail <> "Any" Then
      bolDataSubmitted = True
   End If
   If strViewed <> "Any" And strViewed <> "" Then
      bolDataSubmitted = True
   End If
   If Request.QueryString("Date") <> "" Then
      bolDataSubmitted = True
   End If

   If bolDataSubmitted Then
      
      'Build the start of the SQL string
      strSQL = "SELECT ID,DisplayName,Location,Status,Category,Tech,SubmitDate,SubmitTime,Problem,Notes,LastUpdatedDate,LastUpdatedTime,Custom1,Custom2,TicketViewed" & vbCRLF
      strSQL = strSQL & "FROM Main" & vbCRLF
      
      'If the status is "Any" then don't filter, if we are looking for only open calls then we will
      'look for calls that are not complete.  Otherwise filter by the submitted value.
      If strStatus = "Any" Then
         strSQLStatus = ""      
      ElseIf StrStatus = "Any Open Ticket" Then
         strSQLStatus = "((Main.Status)<>""Complete"") AND " 
         strStatusBar = strStatusBar & " | Status = " & strStatus
      Else
         strSQLStatus = "((Main.Status)=""" & strStatus & """) AND "
         strStatusBar = strStatusBar & " | Status = " & strStatus
      End If
      
      'Set the location portion of the SQL string
      If strLocation <> "Any" Then
         strSQLLocation = "((Main.Location)=""" & strLocation & """) AND "
         strStatusBar = strStatusBar & " | Location = " & strLocation
      Else
         strSQLLocation = ""
      End If
      
      'Set the tech portion of the SQL string
      If strTech <> "Any" Then
         strSQLTech = "((Main.Tech)=""" & strTech & """) AND "
         strStatusBar = strStatusBar & " | Tech = " & strTech
      Else
         strSQLTech = ""
      End If
      
      'Set the category portion of the SQL string
      If strCategory <> "Any" Then
         strSQLCategory = "((Main.Category)=""" & strCategory & """) AND "
         strStatusBar = strStatusBar & " | Category = " & strCategory
      Else
         strSQLCategory = ""
      End If
      
      'Set the user portion of the SQL string
      If strSearchUser <> "Any" Then
         strSQLUser = "((Main.DisplayName) LIKE '%" & strSearchUser & "%') AND "
         strStatusBar = strStatusBar & " | User = " & strSearchUser
      Else
         strSQLUser = ""
      End If
      
      'Set the problem portion of the SQL string
      If strProblem <> "Any" Then
         strSQLProblem = "((Main.Problem) LIKE '%" & strProblem & "%') AND "
         strStatusBar = strStatusBar & " | Problem = " & strProblem
      Else
         strSQLProblem = ""
      End If
      
      'Set the notes portion of the SQL string
      If strNotes <> "Any" Then
         strSQLNotes = "((Main.Notes) LIKE '%" & strNotes & "%') AND "
         strStatusBar = strStatusBar & " | Notes = " & strNotes
      Else
         strSQLNotes = ""
      End If
      
      'Set the EMail portion of the SQL string
      If strEMail <> "Any" Then
         strSQLEMail = "((Main.EMail) LIKE '%" & strEMail & "%') AND "
         strStatusBar = strStatusBar & " | EMail = " & strEMail
      Else
         strSQLEMail = ""
      End If
      
      'Set the date range to use in the query
      If bolDateRange Then
         If intDays > 0 Then
            strSQLDateRange = "((Main.SubmitDate > Date()-" & intDays & ")) AND "
            If intDays = "1" Then
               strStatusBar = strStatusBar & " | " & intDays & " Day"
            Else 
               strStatusBar = strStatusBar & " | " & intDays & " Days"
            End If
         Else
            strSQLDateRange = "((Main.SubmitDate < Date()-" & Abs(intDays) & ")) AND "
               strStatusBar = strStatusBar & " | Older than " &  Abs(intDays) & " Days"
         End If
      Else
         strSQLDateRange = ""
         
      End If

      'See if someone entered a date in the URL.  If so override the date range above.
      If Request.QueryString("Date") <> "" And IsDate(Request.QueryString("Date")) Then
         strDate = Request.QueryString("Date")
         strType = Request.QueryString("Type")
         Select Case strType
            Case "submitted"
               strSQLDateRange = "((Main.SubmitDate = #" & strDate & "#)) AND "
               strStatusBar = "   Tickes Submitted on " & strDate
            Case "workedon"
               If strTech <> "Any" Then
                  
                  strTechSQL = "SELECT UserName FROM Tech WHERE Tech='" & strTech & "'"
                  Set objTechSet = Application("Connection").Execute(strTechSQL)
                  
                  If objTechSet.EOF Then
                     strDateSQL = "SELECT Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate" & vbCRLF
                     strDateSQL = strDateSQL & "FROM Main INNER JOIN Log ON Main.ID = Log.Ticket" & vbCRLF
                     strDateSQL = strDateSQL & "GROUP BY Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate" & vbCRLF
                     strDateSQL = strDateSQL & "HAVING (((Log.UpdateDate)=#" & strDate & "#));"
                  Else
                     strDateSQL = "SELECT Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate" & vbCRLF
                     strDateSQL = strDateSQL & "FROM Main INNER JOIN Log ON Main.ID = Log.Ticket" & vbCRLF
                     strDateSQL = strDateSQL & "GROUP BY Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate, Log.ChangedBy" & vbCRLF
                     strDateSQL = strDateSQL & "HAVING (((Log.UpdateDate)=#" & strDate & "#) AND ((Log.ChangedBy)='" & objTechSet(0) & "'));"
                  End If
               Else
                  strDateSQL = "SELECT Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate" & vbCRLF
                  strDateSQL = strDateSQL & "FROM Main INNER JOIN Log ON Main.ID = Log.Ticket" & vbCRLF
                  strDateSQL = strDateSQL & "GROUP BY Main.ID, Main.DisplayName, Main.Location, Main.Status, Main.Category, Main.Tech, Main.SubmitDate, Main.SubmitTime, Main.Problem, Main.Notes, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2, Main.TicketViewed, Log.UpdateDate" & vbCRLF
                  strDateSQL = strDateSQL & "HAVING (((Log.UpdateDate)=#" & strDate & "#));"
               End If
               strStatusBar = "   Tickets Worked on " & strDate
            Case "completed"
               strSQLDateRange = "((Main.LastupdatedDate = #" & strDate & "#)) AND Status='Complete' And "
               strStatusBar = "   Tickets Complete on " & strDate
         End Select
         strReturnLink = ""
      End If

      'Set the viewed status portion of the SQL string
      If strViewed <> "Any" And strViewed <> "" Then
         If strViewed = "Yes" Then
            strSQLViewed = "((Main.TicketViewed)=True) AND "
         Else
            strSQLViewed = "((Main.TicketViewed)=False) AND "
         End If
         strStatusBar = strStatusBar & " | Viewed = " & strViewed
      Else
         strSQLViewed = ""
      End If
      
      If strStatusBar = "" Then
         strStatusBar = "   All Tickets"
      End If
      
      strStatusBar = strStatusBar & " | Sort = " & strSort
      
      Select Case strSort
         Case "Date - Newest on Top"
            strSQLSort = "ORDER BY Main.SubmitDate DESC, Main.SubmitTime DESC;"
         Case "Date - Oldest on Top"
            strSQLSort = "ORDER By Main.SubmitDate ASC, Main.SubmitTime ASC;"
         Case "Location - A to Z"
            strSQLSort = "ORDER By Main.Location ASC;"
         Case "Location - Z to A"
            strSQLSort = "ORDER By Main.Location DESC;"
         Case "Tech - A to Z"
            strSQLSort = "ORDER By Main.Tech ASC;"
         Case "Tech - Z to A"
            strSQLSort = "ORDER By Main.Tech DESC;"
         Case Else
            strSort = "Date - Newest on Top"
            strSQLSort = "ORDER BY Main.SubmitDate DESC, Main.SubmitTime DESC;"
      End Select

      'Complete the rest of the SQL string with the information from above
      If strSQLStatus <> "" Or strSQLLocation <> "" Or strSQLTech <> "" Or strSQLCategory <> "" Or strSQLUser <> "" Or strSQLProblem <> "" Or strSQLNotes <> "" Or strSQLEMail <> "" Or strSQLDateRange <> "" Or strSQLViewed <> "" Then
         strSQL = strSQL & "WHERE (" & strSQLStatus & strSQLLocation & strSQLTech & strSQLCategory & strSQLUser & strSQLProblem & strSQLNotes & strSQLEMail & strSQLDateRange & strSQLViewed 
         strSQL = Left(strSQL,(Len(strSQL)-5)) & ")"
      End If
      strSQL = strSQL & strSQLSort

      'Execute the SQL string and get the number of returned fields.
      If strDateSQL = "" Then 
         Set objRecordSet = Application("Connection").Execute(strSQL)
      Else   
         Set objRecordSet = Application("Connection").Execute(strDateSQL)
      End If
      
      'Count the number of returned records.  If none are returned a message will be
      'displayed to the user.
      intRecordCount = 0
      Do Until objRecordSet.EOF
         intRecordCount = intRecordCount + 1
         objRecordSet.MoveNext
      Loop
      objRecordSet.MoveFirst
      
      If intRecordCount = 0 Then
         strMessage = "No Tickets Found"
      End If
      strTitle = strTitle & " (" & intRecordCount & ")"
      
      strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

      'Set the zoom level
      If Request.Cookies("ZoomLevel") = "ZoomIn" Then
         If InStr(strUserAgent,"Silk") Then
            intZoom = 1.4
         Else
            intZoom = 1.9
         End If
      End If
      
      If IsMobile Then
         MobileVersion
      ElseIf IsWatch Then
         WatchVersion
      Else
         MainVersion
      End If 
      
   Else 
      Response.Redirect("index.asp")   
   %>
   
<% End If   
 End Sub%>
 
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
 
 <%Sub MainVersion %>
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
         <%=strTitle%>
      </div>
      
      <div class="version">
         Version <%=Application("Version")%>
      </div>
      
      <hr class="admintopbar" />
      <div class="admintopbar">
         <ul class="topbar">
            <li class="topbar"><a href="index.asp">Home</a><font class="separator"> | </font></li>
      <%    If Left(strTitle,16) = "All Open Tickets" Then %>            
               <li class="topbar">Open Tickets <font class="separator"> | </font></li>
      <%    Else %>      
               <li class="topbar"><a href="view.asp?Filter=AllOpenTickets">Open Tickets</a><font class="separator"> | </font></li>
      <%    End If %>
      <%    If strRole <> "Data Viewer" Then %>
      <%       If Left(strTitle,15) = "My Open Tickets" Then %>
                  <li class="topbar">Your Tickets <font class="separator"> | </font></li> 
      <%       Else %>
                  <li class="topbar"><a href="view.asp?Filter=MyOpenTickets">Your Tickets</a><font class="separator"> | </font></li>
      <%       End If %>
      <%    End If %>
      <%    If Application("UseTaskList") And objNameCheckSet(5) <> "Deny" Then %>
               <li class="topbar"><a class="linkbar" href="tasklist.asp">Tasks</a><font class="separator"> | </font></li>
      <%    End If %>
      <% If Application("UseStats") Then %>
			<li class="topbar"><a class="linkbar" href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% End If %>
      <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="docs.asp">Docs</a><font class="separator"> | </font></li>
      <% End If %>
         <li class="topbar"><a href="settings.asp">Settings</a>
      <% If objNameCheckSet(1) = "Administrator" Then %> 
         <font class="separator"> | </font></li>
         <li class="topbar"><a class="linkbar" href="setup.asp">Admin Mode</a>
      <% Else %>
         </li>
      <% End If %>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a class="linkbar" href="login.asp?action=logout">Log Out</a></li>
      <% Else %>
         </li>
      <% End If %>
		</ul>
      </div>
   
   <% If InStr(strUserAgent,"MSIE") Then %>
         <hr class="adminbottombarIE"/>
   <% Else %>   
         <hr class="adminbottombar"/>
   <% End If %>
      
      <div align="center">
      <table border="0" width="750" cellspacing="0" cellpadding="0">
         <tr>
            <td colspan="4"><center>  
            <%=Right(strStatusBar,Len(strStatusBar)-3)%>     
            </center></td>
         </tr>
         <tr><td colspan="4"><hr></td></tr>

      <% 'If no tickets are found then display a message to the user
         If strMessage <> "" Then %>
         <tr><td colspan "4"><center><%=strMessage%></center></td></tr>
      <% End If%>
      
   <% 'Loop through each returned call and display it
      Do  Until objRecordSet.EOF

         'Create the Regular Expression object and set it's properties.
         Set objRegExp = New RegExp
         objRegExp.Pattern = vbCRLF
         objRegExp.Global = True

         'Change a carriage return to a <br /> so it will display properly in HTML.
         If objRecordSet(9) <> "" Then
            strNotes = FixURLs(objRecordSet(9),1)
            strNotes = objRegExp.Replace(strNotes,"<br />")
         End If
         strProblem = FixURLs(objRecordSet(8),1)
         strProblem = objRegExp.Replace(strProblem,"<br />")
         
         'If the notes field contains data then change the icon to the page with the N on it
         If objRecordSet(9) <> "" Then
            strIcon = "nedit"
            
            'Set the number of rows that the icon section should span.  This will change if the call'
            'is closed
            If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then
               intIconSpan = 6
            Else
               intIconSpan = 5
            End If
         Else
            strIcon = "edit"
            
            'Set the number of rows that the icon section should span.  This will change if the call'
            'is closed
            If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then
               intIconSpan = 5
            Else
               intIconSpan = 4
            End If
         End If 
        
        'Drop one more row down in a custom variable is used.
        If Not Application("UseCustom1") And Not Application("UseCustom2") Then 
           intIconSpan = intIconSpan - 1
        End If
        
        %>
         <tr>
            <td rowspan="<%=intIconSpan%>" width="10%" valign="top">
            <a name="<%=objRecordSet(0)%>"></a>
            <center>
         <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
            <a href="modify.asp?ID=<%=objRecordSet(0)%><%=strReturnLink%>"><img border="0" src="../themes/<%=Application("Theme")%>/images/<%=strIcon%>.gif"></a></center>
         <% Else %>
            <a href="modify.asp?ID=<%=objRecordSet(0)%><%=strReturnLink%>"><img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/<%=strIcon%>.gif"></a></center>
         <% End If %>
         <% If objRecordSet(14) Then %>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <center><img border="0" src="../themes/<%=Application("Theme")%>/images/viewed.gif" alt="Viewed by Tech" width="20" height="20"></center>
            <% Else %>
               <center><img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/viewed.gif" alt="Viewed by Tech" width="20" height="20"></center>
            <% End If %>  
         <% Else %>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <center><img border="0" src="../themes/<%=Application("Theme")%>/images/notviewed.gif" alt="Not Viewed by Tech" width="20" height="20"></center>
            <% Else %>
               <center><img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/notviewed.gif" alt="Not Viewed by Tech" width="20" height="20"></center>
            <% End If %>
         <% End If%>
          <center><a href="modify.asp?ID=<%=objRecordSet(0)%><%=strReturnLink%>"><%=objRecordSet(0)%></a></center>
         <% If objRecordSet(3) = "Complete" Then %>
            <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <center><img border="0" src="../themes/<%=Application("Theme")%>/images/closed.gif" width="30" height="30"></center>
            <% Else %>
               <center><img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/closed.gif" width="30" height="30"></center>
            <% End If %>
         <% End If %>
            </td>
            <td width="30%" valign="top"><b>User:</b> <a href="view.asp?User=<%=objRecordSet(1)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(1)%></a></td>
            <td width="30%" valign="top"><b>Location:</b> <a href="view.asp?Location=<%=objRecordSet(2)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(2)%></a></td>
            <td width="30%" valign="top"><b>Category:</b> <a href="view.asp?Category=<%=objRecordSet(4)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(4)%></a></td>
         </tr>
         <tr>
            <td width="30%" valign="top"><b>Tech:</b> <a href="view.asp?Tech=<%=objRecordSet(5)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(5)%></td>
            <td width="30%" valign="top"><b>Status:</b> <%=objRecordSet(3)%></td>
            <td width="30%" valign="top"><b>Received:</b> <%=objRecordSet(6)%></td>
         </tr>
      <%If Application("UseCustom1") or Application("UseCustom2") Then %>
       <tr>
          <%If Application("UseCustom1") Then%>
               <td width="30%" valign="top"><b><%=Application("Custom1Text")%>:</b> <%=objRecordSet(12)%></td>
         <%End If%>
         
         <%If Application("UseCustom2") Then%>
               <td width="30%" valign="top"><b><%=Application("Custom2Text")%>:</b> <%=objRecordSet(13)%></td>
         <%Else%>
            <td width="30%" valign="top">&nbsp;</td>
         <%End If%>
            <td width="30%" valign="top">&nbsp;</td>
         </tr>
      <%End If%>
         <tr>
            <td colspan="3">Problem: <%=strProblem%> </td>
         </tr>
         
   <%    'Display the notes if they are in the database
         If objRecordSet(9) <> "" Then %>
            <tr>
               <td colspan="3"><hr><b>Notes</b>: <%=strNotes%></td>
            </tr>
   <%    End If %>

   <%    'Display the date completed if the call is closed
         If objRecordSet(10) <> "6/16/1978" And objRecordSet(3) = "Complete" Then %>
            <tr>
               <td colspan="3"><hr><b>Date Closed</b>: <%=objRecordSet(10)%> -  
               <%=objRecordSet(11)%></td>
            </tr>
   <%    End If %>

         <tr>
            <td colspan="4"><hr></td>
         </tr>
   <%    objRecordSet.MoveNext
      Loop %>
      </table>   
      </body>   
      </html>
<%End Sub%>

<%Sub MobileVersion %>
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
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %>
      <meta name="theme-color" content="#<%=Application("AndroidBarColor")%>">
   </head>
   <body>
      <center><b><%=strTitle%></b></center>
      <center>
      <table align="center">
         <tr><td width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         <tr>
            <td>
               <div align="center">
               <form method="Post" action="view.asp">
                  <input type="submit" value="Home" name="home">
                  <input type="submit" value="Open Tickets" name="filter">
            <% If strRole <> "Data Viewer" Then %>   
                  <input type="submit" value="Your Tickets" name="filter">
            <% End If %>
               </form>
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table>   
   <% Do  Until objRecordSet.EOF

         'Create the Regular Expression object and set it's properties.
         Set objRegExp = New RegExp
         objRegExp.Pattern = vbCRLF
         objRegExp.Global = True

         'Change a carriage return to a <br /> so it will display properly in HTML.
         If objRecordSet(9) <> "" Then
            strNotes = FixURLs(objRecordSet(9),1)
            strNotes = objRegExp.Replace(strNotes,"<br />")
         End If
         strProblem = FixURLs(objRecordSet(8),1)
         strProblem = objRegExp.Replace(strProblem,"<br />")   
   %>
      <table align="center">
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               <% If objRecordSet(3) = "Complete" Then %>
                  <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
                     <img border="0" src="../themes/<%=Application("Theme")%>/images/closed.gif" width="15" height="15">
                  <% Else %>
                     <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/closed.gif" width="15" height="15">
                  <% End If %>
               <% End If %> 

               <% If objRecordSet(14) Then %>
                  <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
                     <img border="0" src="../themes/<%=Application("Theme")%>/images/viewed.gif" alt="Viewed by Tech" width="15" height="15">
                  <% Else %>
                     <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/viewed.gif" alt="Viewed by Tech" width="15" height="15">
                  <% End If %>  
               <% Else %>
                  <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
                     <img border="0" src="../themes/<%=Application("Theme")%>/images/notviewed.gif" alt="Not Viewed by Tech" width="15" height="15">
                  <% Else %>
                     <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/notviewed.gif" alt="Not Viewed by Tech" width="15" height="15">
                  <% End If %>
               <% End If%>
               
               <a name="<%=objRecordSet(0)%>"></a>
               <a href="modify.asp?ID=<%=objRecordSet(0)%><%=strReturnLink%>">Ticket #<%=objRecordSet(0)%></a> -
               <a href="view.asp?User=<%=objRecordSet(1)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(1)%></a>
          

            </td>
         </tr>
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               &nbsp;&nbsp;&nbsp;- Received: <%=objRecordSet(6)%>
            </td>
         </tr>  
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               &nbsp;&nbsp;&nbsp;- Location: <a href="view.asp?Location=<%=objRecordSet(2)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(2)%></a>
            </td>
         </tr>
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               &nbsp;&nbsp;&nbsp;- Status: <%=objRecordSet(3)%>
            </td>
         </tr>
   <%    If objRecordSet(4) <> " " Then   %>   
            <tr>
               <td width="<%=Application("MobileSiteWidth")%>">
                  &nbsp;&nbsp;&nbsp;- Category: <a href="view.asp?Category=<%=objRecordSet(4)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(4)%>
               </td>
            </tr>
   <%    End If %>      
   <%    If Application("UseCustom1") Then %>      
            <tr>
               <td width="<%=Application("MobileSiteWidth")%>">
                  &nbsp;&nbsp;&nbsp;- <%=Application("Custom1Text")%>: <%=objRecordSet(12)%> 
               </td>
            </tr>
   <%    End If %>      
   <%    If Application("UseCustom2") Then %> 			
            <tr>
               <td width="<%=Application("MobileSiteWidth")%>">
                  &nbsp;&nbsp;&nbsp;- <%=Application("Custom2Text")%>: <%=objRecordSet(13)%> 
               </td>
            </tr>
   <%    End If %>
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               &nbsp;&nbsp;&nbsp;- Tech: <a href="view.asp?Tech=<%=objRecordSet(5)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(5)%></a>
            </td>
         </tr>
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               &nbsp;&nbsp;&nbsp;- Problem: <%=strProblem%>
            </td>
         </tr>
   <%    If objRecordSet(9) <> "" Then %>
            <tr>
               <td width="<%=Application("MobileSiteWidth")%>">
                  &nbsp;&nbsp;&nbsp;- Notes: <%=strNotes%>
               </td>
            </tr>
   <%    End If%>   
         <tr><td width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         </table>
   <%    objRecordSet.MoveNext
      Loop %>

      </center>
   </body>
   </html>
<%End Sub%>

<%Sub WatchVersion %>

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
   <body>
      <div align="right">
         (<%=intRecordCount%>) Help Desk
      </div>
      <hr />
      
      <form method="Post" action="view.asp">
      <div align="center"> 
         <input type="submit" value="Home" name="home">
      </div>
      </form>
      <hr />
      
      <div align="left">
      <% Do  Until objRecordSet.EOF 
      
            'Create the Regular Expression object and set it's properties.
            Set objRegExp = New RegExp
            objRegExp.Pattern = vbCRLF
            objRegExp.Global = True

            'Change a carriage return to a <br /> so it will display properly in HTML.
            If objRecordSet(9) <> "" Then
               strNotes = FixURLs(objRecordSet(9),1)
               strNotes = objRegExp.Replace(strNotes,"<br />")
            End If
            strProblem = FixURLs(objRecordSet(8),1)
            strProblem = objRegExp.Replace(strProblem,"<br />")    
      
      %>
         <a href="modify.asp?ID=<%=objRecordSet(0)%><%=strReturnLink%>"><%=objRecordSet(0)%></a> - 
         <a href="view.asp?User=<%=objRecordSet(1)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(1)%></a> <br />
         
         - <a href="view.asp?Location=<%=objRecordSet(2)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(2)%></a> <br />
         
   <%    If Application("UseCustom1") Then %>      
            - <%=objRecordSet(12)%> <br />
   <%    End If %>      
   <%    If Application("UseCustom2") Then %> 			
            - <%=objRecordSet(13)%> <br />
   <%    End If %>

         - <a href="view.asp?Tech=<%=objRecordSet(5)%>&Status=Any%20Open%20Ticket"><%=objRecordSet(5)%></a> <br />
         - <%=strProblem%> <br />

   <%    If objRecordSet(9) <> "" Then %>
            <%=strNotes%> <br />
   <%    End If%>  
         <hr />
      <%    objRecordSet.MoveNext
         Loop %>
      </div>
      
   </body>
   </html>
<%End Sub%>

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
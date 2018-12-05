<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/12/04
'Last Updated 6/16/14

'This is the administrators main page.  From here you can run some of the pre-made queries.
 
Option Explicit

On Error Resume Next

Dim strSQL, objLocationSet, objStatusSet, objTechSet, strLocation, strTech
Dim objCategorySet, objRecordSet, intCallCount, intOpenCallCount, strCategory
Dim intPercentOpen, intTotalTime, intAvgOpenTime, intEvents, strAvgTicketTime
Dim strDays, strHours, strMinutes, objNetwork, strUser, objNameCheckSet
Dim objUpdateRequested, objYourUpdateRequests, bolNotifications, bolShowEvents
Dim intID, strDate, strTime, strChangeType, strChangedBy, strShowEvents
Dim strUserName, strCMD, strNewTicketMessage, strCurrentUser, strUserAgent
Dim strRole, objMessage, strMessageFont, objTicketCheck, strTicketCheck
Dim objTracking, objUnViewedTickets, objAvgTicketTime, objTodaysTicketCount
Dim intTicketCount, objCompleteTicketCount, intCompleteCount, bolMobileVersion
Dim objCheckIns, bolShowLogout, intZoom, intInputSize

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

strCMD = Request.Form("Submit")

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

strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

Select Case strCMD
   Case "New Ticket"
      strUserName = Request.Form("UserName")
      If strUserName = "" Then
         Response.Redirect("../index.asp?username=<empty>" & strUserName)
      Else
         If strUserName = GetFirstandLastName(strUserName) Then
            strNewTicketMessage = "Invalid Username"
         Else
            Response.Redirect("../index.asp?username=" & strUserName)
         End If
      End If
   Case "Submit"
      intID = Request.Form("ID")
      strSQL = "SELECT ID FROM Main WHERE ID=" & intID
      Set objTicketCheck = Application("Connection").Execute(strSQL)
      If objTicketCheck.EOF Then
         strTicketCheck = "Invalid Ticket Number"
      Else
         Response.Redirect("modify.asp?ID=" & intID)
      End If
   Case "Mobile Site"
      Response.Cookies("SiteVersion") = "Mobile"
      Response.Cookies("SiteVersion").Expires = Date() + 14
      GetUser
   Case "Full Site"
      Response.Cookies("SiteVersion") = "Full"
      Response.Cookies("SiteVersion").Expires = Date() + 14
      GetUser
      
End Select

If Request.QueryString("Created") = "True" Then
   strNewTicketMessage = "Submitted"
End If

'Figure out if the log should be displayed.
strShowEvents = Request.Form("cmdSubmit")
If strShowEvents = "Show Log" Then
   bolShowEvents = True
Else
   bolShowEvents = False
End If

intEvents = Request.Form("Events")
If intEvents = "" Then
   intEvents = 10
End If

'Create the SQL string that will count the number off total calls and open calls
strSQL = "Select Main.Status,Main.OpenTime" & vbCRLF
strSQL = strSQL & "From Main"

'Execute the SQL string
Set objRecordSet = Application("Connection").Execute(strSQL)

intCallCount = 0
intOpenCallCount = 0
intTotalTime = 0
Do Until objRecordSet.EOF
   intCallCount = intCallCount + 1
   If objRecordSet(0) <> "Complete" Then
      intOpenCallCount = intOpenCallCount + 1
   Else
      If objRecordSet(1) <> "" Then
         intTotalTime = intTotalTime + objRecordSet(1)
      End If
   End If
   objRecordSet.MoveNext
Loop

intAvgOpenTime = Round(intTotalTime  / (intCallCount - intOpenCallCount),0)
intPercentOpen = Round(((intOpenCallCount / intCallCount) * 100),5)

strSQL = "SELECT Avg(OpenTime) AS AvgOfOpenTime" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "Where Main.OpenTime<>''"
Set objAvgTicketTime = Application("Connection").Execute(strSQL) 

strDays = Int(objAvgTicketTime(0)/1440)
strHours = Int((objAvgTicketTime(0)-strDays*1440)/60)
strMinutes = (objAvgTicketTime(0)-(strDays*1440)-(strHours*60))
strAvgTicketTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" 

'Build the SQL string for the Location pull down box
strSQL = "Select Location.Location" & vbCRLF
strSQL = strSQL & "From Location" & vbCRLF
strSQL = strSQL & "Order By Location.Location;"

'Execute the SQL string
Set objLocationSet = Application("Connection").Execute(strSQL)

'Build the SQL string for the Status pull down box
strSQL = "Select Status.Status" & vbCRLF
strSQL = strSQL & "From Status" & vbCRLF
strSQL = strSQL & "Order By Status.Status;"

'Execute the SQL string
Set objStatusSet = Application("Connection").Execute(strSQL)

'Build the SQL string for the Tech pull down box
strSQL = "Select Tech" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "Where UserLevel<>'Data Viewer'" & vbCRLF
strSQL = strSQL & "Order By Tech;"

'Execute the SQL string
Set objTechSet = Application("Connection").Execute(strSQL)

'Build the SQL string for the Tech pull down box
strSQL = "Select Category.Category" & vbCRLF
strSQL = strSQL & "From Category" & vbCRLF
strSQL = strSQL & "Order By Category.Category;"

'Execute the SQL string
Set objCategorySet = Application("Connection").Execute(strSQL)

'This code will fix the display name so it matches what is in the database.
Select Case UCase(strUser)
   Case "HELPDESK"
      strCurrentUser = "Heat Help Desk"
   Case "TPERKINS"
      strCurrentUser = "Tech Services"
   Case Else
      strCurrentUser = GetFirstandLastName(strUser)
End Select

'Go to the correct site if the buttons are used on the mobile site
Select Case Request.Form("Site")
   Case "Settings"
      Response.Redirect("settings.asp")
   Case "Search"
      Response.Redirect("msearch.asp")
   Case "Stats"
      Response.Redirect("mstats.asp")
   Case "Zoom In"
      Response.Cookies("ZoomLevel") = "ZoomIn"
      Response.Cookies("ZoomLevel").Expires = Date() + 14
   Case "Zoom Out"
      Response.Cookies("ZoomLevel") = "ZoomOut"
      Response.Cookies("ZoomLevel").Expires = Date() + 14
End Select

'Set the zoom level
If Request.Cookies("ZoomLevel") = "ZoomIn" Then
   If InStr(strUserAgent,"Silk") Then
      intZoom = 1.4
   Else
      intZoom = 1.9
   End If
End If

'If an update has been requested let them know
strSQL = "SELECT Tracking.TrackedBy, Tracking.Ticket" & vbCRLF 
strSQL= strSQL & "FROM Tracking INNER JOIN Main ON Tracking.Ticket = Main.ID" & vbCRLF
strSQL = strSQL & "WHERE (Tracking.Type='Request') AND (Main.Tech='" & strCurrentUser & "')"
Set objUpdateRequested = Application("Connection").Execute(strSQL)

If Not objUpdateRequested.EOF Then
   bolNotifications = True
End If

'If the current user is still waiting for an update get the info
strSQL = "SELECT Main.Tech, Tracking.Ticket" & vbCRLF 
strSQL= strSQL & "FROM Tracking INNER JOIN Main ON Tracking.Ticket = Main.ID" & vbCRLF
strSQL = strSQL & "WHERE (Tracking.Type='Request') AND (Tracking.TrackedBy='" & strUser & "')"
Set objYourUpdateRequests = Application("Connection").Execute(strSQL)

If Not objYourUpdateRequests.EOF Then
   bolNotifications = True
End If

'Check and see if the current tech is tracking any tickets
strSQL = "SELECT Main.Tech, Tracking.Ticket" & vbCRLF
strSQL = strSQL & "FROM Tracking INNER JOIN Main ON Tracking.Ticket = Main.ID" & vbCRLF
strSQL = strSQL & "WHERE (Tracking.Type='Track') And (Tracking.TrackedBy='" & strUser & "')"
Set objTracking = Application("Connection").Execute(strSQL)

If Not objTracking.EOF Then
   bolNotifications = True
End If

'Check and see if the current tech has any unviwed tickets
strSQL = "SELECT ID FROM Main Where TicketViewed=False And Tech='" & strCurrentUser & "'"
Set objUnViewedTickets = Application("Connection").Execute(strSQL)

If Not objUnViewedTickets.EOF Then
   bolNotifications = True
End If

'See if there are any system messages for the techs.
strSQL = "SELECT Message,Recipient,Type,PositionOnPage,Enabled" & vbCRLF
strSQL = strSQL & "FROM Message" & vbCRLF
strSQL = strSQL & "WHERE ID=1"

Set objMessage = Application("Connection").Execute(strSQL)

strSQL = "SELECT Count(ID) AS CountOfTickets" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "GROUP BY SubmitDate" & vbCRLF
strSQL = strSQL & "HAVING SubmitDate=Date()"
Set objTodaysTicketCount = Application("Connection").Execute(strSQL)
If objTodaysTicketCount.EOF Then
   intTicketCount = 0
Else
   intTicketCount = objTodaysTicketCount(0)
End If

strSQL = "SELECT Count(ID) AS CountOfID" & vbCRLF
strSQL = strSQL & "FROM Main" & vbCRLF
strSQL = strSQL & "GROUP BY LastUpdatedDate, Status" & vbCRLF
strSQL = strSQL & "HAVING LastUpdatedDate=Date() AND Status='Complete'"
Set objCompleteTicketCount = Application("Connection").Execute(strSQL)
If objCompleteTicketCount.EOF Then
   intCompleteCount = 0
Else
   intCompleteCount = objCompleteTicketCount(0)
End If

strSQL = "SELECT Tech.Tech, CheckIns.Location, CheckIns.CheckInTime" & vbCRLF
strSQL = strSQL & "FROM CheckIns INNER JOIN Tech ON CheckIns.Tech = Tech.Username" & vbCRLF
strSQL = strSQL & "WHERE (((CheckIns.[CheckInDate])=Date()));"
Set objCheckIns = Application("Connection").Execute(strSQL)

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
%>

<%Sub AccessGranted

   On Error Resume Next 

   If IsMobile Then
      MobileVersion
   ElseIf IsWatch Then
      WatchVersion
   Else
      MainVersion
   End If   
   
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

<%
Function IsTablet
   If InStr(strUserAgent,"Nexus 7") Then
      IsTablet = True
   ElseIf InStr(strUserAgent,"Nexus 9") Then  
      IsTablet = True
   ElseIf InStr(strUserAgent,"iPad") Then
      IsTablet = True
   ElseIf InStr(strUserAgent,"Silk") Then
      IsTablet = True
   ElseIf InStr(strUserAgent,"GT-N5110") Then
      IsTablet = True
   Else
      IsTablet = False
   End If
End Function
%>

<%Sub MainVersion 

   Dim intInputSize
   
   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

   If InStr(strUserAgent,"Android") Or InStr(strUserAgent,"Silk") Then
      intInputSize = 18
   Else
      intInputSize = 25
   End If
   
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
      <%=Application("SchoolName")%> Help Desk
   </div>
   
   <div class="version">
      Version <%=Application("Version")%>
   </div>
   
   <hr class="admintopbar" />
   <div class="admintopbar">
      <ul class="topbar">
			<li class="topbar">Home<font class="separator"> | </font></li>
         <li class="topbar"><a class="linkbar" href="view.asp?Filter=AllOpenTickets">Open Tickets</a><font class="separator"> | </font></li>
      <% If strRole <> "Data Viewer" Then %>
         <li class="topbar"><a class="linkbar" href="view.asp?Filter=MyOpenTickets">Your Tickets</a><font class="separator"> | </font></li>  
      <% End If %>
      <% If Application("UseTaskList") And objNameCheckSet(5) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="tasklist.asp">Tasks</a><font class="separator"> | </font></li>
      <% End If %>
      <% If Application("UseStats") Then %>
			<li class="topbar"><a class="linkbar" href="stats.asp">Stats</a><font class="separator"> | </font></li> 
      <% End If %>
      <% If Application("UseDocs") And objNameCheckSet(6) <> "Deny" Then %>
         <li class="topbar"><a class="linkbar" href="docs.asp">Docs</a><font class="separator"> | </font></li>
      <% End If %>   
         <li class="topbar"><a class="linkbar" href="settings.asp">Settings</a>
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
         <td width="22%" valign="top">
         <table>
            <tr>
         <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <td><img border="0" src="../themes/<%=Application("Theme")%>/images/admin.gif"></td>
         <% Else %>
               <td><img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/admin.gif"></td>
         <% End If %>
            </tr>
            <tr>
               <td>
               <% If Application("UseStats") Then %>
                  <a href="stats.asp">Help Desk Statistics</a><br />
                  Total tickets = <%=intCallCount%> <br />
                  Open tickets = <a href="view.asp?Filter=AllOpenTickets"><%=intOpenCallCount%></a> <br />
                  Submitted today = <a href="view.asp?date=<%=Date()%>&type=submitted"><%=intTicketCount%></a> <br />
                  Closed today = <a href="view.asp?date=<%=Date()%>&type=completed"><%=intCompleteCount%></a> <br />
                  Avg time = <%=strAvgTicketTime%> <br />
                  <hr />
               <% End If %>
                  <form method="POST" action="index.asp">
<%                If Not bolShowEvents Then%>
                     <input type="submit" value="Show Log" name="cmdSubmit">
<%                   If InStr(Request.ServerVariables("HTTP_USER_AGENT"),"iPad") or InStr(Request.ServerVariables("HTTP_USER_AGENT"),"iPhone") Then %>
                        <input type="text" name="Events" value="10" size="4"> Events
<%                   Else %>
                        <input type="text" name="Events" value="10" size="1"> Events
<%                   End If %>                     
                 
<%                Else %>
                     <input type="submit" value="Hide Log" name="cmdSubmit"></td>
<%                End If %>
                                    
               </td>
            </tr>
               <% If IsTablet Then %>
                     <tr><td><hr /></td></tr>
                     <tr><td align="center"><input type="submit" value="Mobile Site" name="Submit"></td></tr>
               <% End If %>
               </form>
         </table>

         <td width="2%">&nbsp;</td>
 
         <td width="71%" valign="top">
<%          If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Techs") And (objMessage(3) = "Top" or objMessage(3) = "Both") Then 
               Select Case objMessage(2)
                  Case "Normal"
                     strMessageFont = ""
                  Case "Alert"
                     strMessageFont = "<font class=""information"">"
               End Select%>
               <table>
                  <tr><td colspan="2">
                     <%=strMessageFont%><%=Replace(objMessage(0),vbCRLF,"<br />")%> </font>
                  </td></tr>
               </table>
               <hr /> 
<%          End If %>  

         Search for tickets or enter the ticket number below.
         <hr>		

      <table border="0" width="100%" cellspacing="0" cellpadding="0">
      <form method="POST" action="view.asp">
         <tr>
            <td height="31" width="84">Status: </td>
            <td height="31">
               <select name="Status">			
                  <option>Any</option>
                  <option selected="selected">Any Open Ticket</option>

         <% 'Populates the status pulldown list
            Do Until objStatusSet.EOF
               If Trim(Ucase(objStatusSet(0))) <> Trim(Ucase(strLocation)) Then
            %>    <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
         <%    End If
               objStatusSet.MoveNext
            Loop
            %>
            
               </select>
            </td>
            <td>User:</td>
            <td><input type="text" name="User" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td>
         </tr>
         <tr>
            <td height="31" width="84">Category:</td>
            <td height="31"><select size="1" name="Category">
            <option>Any</option>

         <% 'Populates the category pulldown list
            Do Until objCategorySet.EOF      
               If Trim(Ucase(objCategorySet(0))) <> Trim(Ucase(strCategory)) Then
         %>       <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
         <%    End If
               objCategorySet.MoveNext
            Loop
         %> 
         
            </select></td>
            <td>Problem:</td>
            <td><input type="text" name="Problem" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td>
         </tr>
         <tr>
            <td height="31" width="84">Tech:</td>
            <td height="31"><select size="1" name="Tech">
            <option>Any</option>
            <option>Nobody</option>

         <% 'Populates the tech pulldown list
            Do Until objTechSet.EOF      
               If Trim(Ucase(objTechSet(0))) <> Trim(Ucase(strTech)) Then
         %>       <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
         <%    End If
               objTechSet.MoveNext
            Loop
         %>  
         
            </select></td>
            <td>Notes:</td>
            <td><input type="text" name="Notes" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td>
         </tr>
         <tr>
            <td height="31" width="84">Location:</td>
            <td height="31"><select size="1" name="Location">
            <option>Any</option>

         <% 'Populates the location pulldown list
            Do Until objLocationSet.EOF
               If Trim(Ucase(objLocationSet(0))) <> Trim(Ucase(strLocation)) Then
         %>       <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
         <%    End If
               objLocationSet.MoveNext
            Loop
         %>
            </select></td>
            <td>EMail:</td>
            <td><input type="text" name="EMail" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td>
            </tr>
            <tr>
            <td height="31" width="84">Sort By:</td>
            <td height="31"><select size="1" name="Sort">
                           <option>Date - Newest on Top</option>
                           <option>Date - Oldest on Top</option>
                           <option>Location - A to Z</option>
                           <option>Location - Z to A</option>
                           <option>Tech - A to Z</option>
                           <option>Tech - Z to A</option>
            </select>

            <td>Viewed:</td>
            <td>
               <select size="1" name="Viewed">
                  <option>Any</option>
                  <option>Yes</option>
                  <option>No</option>
               </select>
            </td>
         </tr>
         <tr>
         <td>Submitted:</td>
         <td>
            <select size="1" name="Days">
               <option value="0">Any</option>
               <option value="1">Today</option>
               <option value="7">Within the Past Week</option>
               <option value="14">Within the Past 2 Weeks</option>
               <option value="30">Within the Past 30 Days</option>
               <option value="90">Within the Past 90 Days</option>
               <option value="180">Within the Past 180 Days</option>
               <option value="-7">Over a Week Ago</option>
               <option value="-14">Over Two Weeks Ago</option>
               <option value="-30">Over 30 Days Ago</option>
               <option value="-90">Over 90 Days Ago</option>
               <option value="-180">Over 180 Days Ago</option>
            </select>
         </td>
         <td colspan="2">
            <input type="submit" value="Submit" name="Submit" style="float: right">
         </td></tr>  

         </form>		
      </table>
      <hr />
      <table width="100%" cellspacing="0" cellpadding="0">
            <form method="Post" action="index.asp">		
               <tr>
                  <td>
                     View ticket #:&nbsp; <input type="text" name="ID" value="<%=intID%>" size="14">
               <% If strTicketCheck <> "" Then %>
                  <font class="missing"><%=strTicketCheck%></font>
               <% End If %>
                  </td>
                  <td>
                  <input type="submit" value="Submit" name="Submit" style="float: right"></td>
               </tr>
            </form>
      <% If strRole <> "Data Viewer" AND Application("UseAD") Then %>       
            <tr><td colspan="2"><hr /></td></tr>
            <form method="POST" action="index.asp">		
               <tr>
                  <td colspan="2">
                     Enter a new ticket.  
                  <% If strNewTicketMessage = "Submitted" Then %>
                     <font class="information"><%=strNewTicketMessage%></font>
                  <% Else %>
                     <font class="missing"><%=strNewTicketMessage%></font>
                  <% End If %>
                  </td>
               </tr>
               <tr>
                  <td>
                     Username:&nbsp;<input type="text" name="UserName" value="<%=strUserName%>" size="25">
                  </td>
                  <td>
                  <input type="submit" value="New Ticket" name="Submit" style="float: right"></td>
               </tr>  
            </form>
      <% End If %>
      
      <% If Not objCheckIns.EOF Then %>
            <tr><td colspan="2"><hr /></td></tr>
            <tr><td colspan="2">Current Tech Locations</td></tr>
            <tr><td colspan="2">
               <table  width="500">
                  <tr>
                     <td>Tech</td>
                     <td>Location</td>
                     <td>Time</td>
                  </tr>
               <% Do Until objCheckIns.EOF %>
                     <tr>
                        <td class="Showborders"><%=objCheckIns(0)%></td>
                        <td class="Showborders"><%=objCheckIns(1)%></td>
                        <td class="Showborders"><%=objCheckIns(2)%></td>
                     </tr>
               <%    objCheckIns.Movenext
                  Loop %>
               </table>
            </td></tr>
      <% End If %>
      
<%          If bolNotifications Then %>
               <tr><td colspan="2"><hr /></td></tr>
               <tr>
                  <td colspan="2">Notifications:<br />
                  <ul>
<%                Do Until objUpdateRequested.EOF%>
                     <li><%=GetFirstandLastName(objUpdateRequested(0))%> has requested an update on 
                     <a href="modify.asp?ID=<%=objUpdateRequested(1)%>">Ticket <%=objUpdateRequested(1)%></a>.</li>
<%                   objUpdateRequested.MoveNext
                  Loop 
                  Do Until objYourUpdateRequests.EOF%>
                     <li>   
                        <%=objYourUpdateRequests(0)%> has not updated 
                        <a href="modify.asp?ID=<%=objYourUpdateRequests(1)%>">Ticket <%=objYourUpdateRequests(1)%></a> since your request.
                     </li>
<%                   objYourUpdateRequests.MoveNext
                  Loop 
                  Do Until objTracking.EOF%>
                     <li>
                     You are tracking 
                        <a href="modify.asp?ID=<%=objTracking(1)%>&ShowLog=Yes">Ticket <%=objTracking(1)%></a>.
                        Assigned to <%=objTracking(0)%>.
                     </li>
<%                   objTracking.MoveNext 
                  Loop 
                  Do Until objUnViewedTickets.EOF%>
                     <li>
                     You have not viewed  
                        <a href="modify.asp?ID=<%=objUnViewedTickets(0)%>&ShowLog=Yes">Ticket <%=objUnViewedTickets(0)%></a>.
                     </li>
<%                   objUnViewedTickets.MoveNext
                  Loop%>                   
                  </ul>
                  </td>
               </tr>
<%          End If %>

<%          If bolShowEvents Then
               strSQL = "SELECT TOP " & intEvents & " Log.Ticket, Log.UpdateDate, Log.UpdateTime, Log.Type, Log.ChangedBy, Main.Tech, Main.Status, Main.Name, Main.Location" & vbCRLF
               strSQL = strSQL & "FROM Log INNER JOIN Main ON Log.Ticket = Main.ID" & vbCRLF
               strSQL = strSQL & "ORDER BY Log.ID DESC;" 

               Set objRecordSet = Application("Connection").Execute(strSQL)%>
               <tr><td colspan="2"><hr /></td></tr>
               <tr><td colspan="2">
               <table>
                  <tr>
                     <th class="showborders">Ticket</th>
                     <th class="showborders">Date</th>
                     <th class="showborders">Time</th>
                     <th class="showborders">Change Type</th>
                     <th class="showborders">Changed By</th>
                   </tr>
<%                Do  Until objRecordSet.EOF
                     intID = objRecordSet(0)
                     strDate = objRecordSet(1)
                     strTime = objRecordSet(2)
                     strChangeType = objRecordSet(3)
                     Select Case strChangeType
                        Case "Request Update Complete"
                           strChangeType = "Update Done"
                        Case "Cancelled Update Request"
                           strChangetype = "Cancel Update"
                     End Select
                     
                     If strChangeType = "Request Update Complete" Then
                        strChangeType = "Update Done"
                     End If
                     strChangedBy = GetFirstandLastName(objRecordSet(4)) %>
                     <tr>
                        <td class="showborders"><center><a href="modify.asp?ID=<%=intID%>&ShowLog=Yes"><%=intID%></a></center></td>
                        <td class="showborders"><%=strDate%>&nbsp;&nbsp;</td>
                        <td class="showborders"><%=strTime%>&nbsp;&nbsp;</td>
                        <td class="showborders"><%=strChangeType%>&nbsp;&nbsp;</td>
                        <td class="showborders"><%=strChangedBy%>&nbsp;&nbsp;</td>
                     </tr>
<%                   objRecordSet.MoveNext  
                  Loop%>
               </table>
               </td></tr>
<%          End If %>
            <tr><td colspan="2"><hr /></td></tr> 
<%          If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Techs") And (objMessage(3) = "Bottom" or objMessage(3) = "Both")Then 
               Select Case objMessage(2)
                  Case "Normal"
                     strMessageFont = ""
                  Case "Alert"
                     strMessageFont = "<font class=""information"">"
               End Select%>
               <tr><td colspan="2">
                  <%=strMessageFont%><%=Replace(objMessage(0),vbCRLF,"<br />")%> </font>
               </td></tr>
               <tr><td colspan="2"><hr /></td></tr> 
<%          End If %>            
      </table>
         </td>
      </tr>

   </table>

   </div>
   </body>

   </html>
<%End Sub%>

<%Sub MobileVersion 
   
   intInputSize = 20

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
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   
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
      <center><b><%=Application("SchoolName")%> Help Desk Admin</b></center>
      <center>
      <table align="center">
         <tr><td><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               <div align="center">
                  <input type="submit" value="Open Tickets" name="filter">
            <% If strRole <> "Data Viewer" Then %>   
                  <input type="submit" value="Your Tickets" name="filter">
            <% End If %>
            
            <% If bolShowLogout Then %>   
                  <input type="submit" value="Log Out" name="Log Out">
            <% End If %> 
            
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table>
      <table align="center">
<%    If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Techs") And (objMessage(3) = "Top" or objMessage(3) = "Both") Then 
         Select Case objMessage(2)
            Case "Normal"
               strMessageFont = ""
            Case "Alert"
               strMessageFont = "<font class=""information"">"
         End Select%>
         <tr><td colspan="2">
            <%=strMessageFont%><%=Replace(objMessage(0),vbCRLF,"<br />")%> </font>
         </td></tr>
         <tr><td colspan="2"><hr /></td></tr> 
<%    End If %>         
      <form method="POST" action="view.asp">
         <tr>
            <td>
               Status: 
            </td>
            <td>
               <select name="Status">			
                  <option>Any</option>
                  <option selected="selected">Any Open Ticket</option>

            <% 'Populates the status pulldown list
               Do Until objStatusSet.EOF
                  If Trim(Ucase(objStatusSet(0))) <> Trim(Ucase(strLocation)) Then
               %>    <option value="<%=objStatusSet(0)%>"><%=objStatusSet(0)%></option>
            <%    End If
                  objStatusSet.MoveNext
               Loop
               %>
               
               </select>
            </td>
         </tr>
         <tr>
            <td>
               Category:
            </td>
            <td>
               <select size="1" name="Category">
                  <option>Any</option>

            <% 'Populates the category pulldown list
               Do Until objCategorySet.EOF      
                  If Trim(Ucase(objCategorySet(0))) <> Trim(Ucase(strCategory)) Then
            %>       <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
            <%    End If
                  objCategorySet.MoveNext
               Loop
            %> 
               </select>
            </td>
         </tr>
         
         <tr>
            <td>
               Tech:
            </td>
            <td>
               <select size="1" name="Tech">
                  <option>Any</option>
                  <option>Nobody</option>

               <% 'Populates the tech pulldown list
                  Do Until objTechSet.EOF      
                     If Trim(Ucase(objTechSet(0))) <> Trim(Ucase(strTech)) Then
               %>       <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
               <%    End If
                     objTechSet.MoveNext
                  Loop
               %>  
               
               </select>
            </td>
         </tr>
         <tr>
            <td>
               Location:
            </td>
            <td>
               <select size="1" name="Location">
                  <option>Any</option>

            <% 'Populates the location pulldown list
               Do Until objLocationSet.EOF
                  If Trim(Ucase(objLocationSet(0))) <> Trim(Ucase(strLocation)) Then
            %>       <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
            <%    End If
                  objLocationSet.MoveNext
               Loop
            %>
               </select>
            </td>
         </tr>
         <tr>
            <td>
               Sort By:
            </td>
            <td>
               <select size="1" name="Sort">
                  <option>Date - Newest on Top</option>
                  <option>Date - Oldest on Top</option>
                  <option>Location - A to Z</option>
                  <option>Location - Z to A</option>
                  <option>Tech - A to Z</option>
                  <option>Tech - Z to A</option>
               </select>
            </td>
         </tr>
         <tr>
            <td>
               Viewed:
            </td>
            <td>
               <select size="1" name="Viewed">
                  <option>Any</option>
                  <option>Yes</option>
                  <option>No</option>
               </select>
               
            </td>
         </tr>
         <tr>
            <td>Submitted:</td>
            <td>
               <select size="1" name="Days">
                  <option value="0">Any</option>
                  <option value="1">Today</option>
                  <option value="7">Within the Past Week</option>
                  <option value="14">Within the Past 2 Weeks</option>
                  <option value="30">Within the Past 30 Days</option>
                  <option value="90">Within the Past 90 Days</option>
                  <option value="180">Within the Past 180 Days</option>
                  <option value="-7">Over a Week Ago</option>
                  <option value="-14">Over Two Weeks Ago</option>
                  <option value="-30">Over 30 Days Ago</option>
                  <option value="-90">Over 90 Days Ago</option>
                  <option value="-180">Over 180 Days Ago</option>
               </select>
            </td>
         </tr>
         <tr><td>User:</td><td><input type="text" name="User" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td></tr>
         <tr><td>Problem:</td><td><input type="text" name="Problem" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"></td></tr>
         <tr><td>Notes:</td><td><input type="text" name="Notes" value="Any" size="<%=intInputSize%>" onfocus="if(this.value=='Any') this.value='';"><input type="submit" value="Submit" style="float: right"></td></tr>
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         </form>
      </table>      
      <table align="center">
         <form method="post" action="index.asp">
         <tr>
            <td colspan="2">Ticket # <input type="text" name="ID" size="10">
            <input type="submit" value="Submit" name="Submit" style="float: right"> </td>
         </tr>
         </form>
   <% If strTicketCheck <> "" Then %>
         <tr><td colspan="2"><font class="missing"><%=strTicketCheck%></font></td></tr>
   <% End If %>
         
   <% If strRole <> "Data Viewer" AND Application("UseAD") Then %>       
         <tr><td colspan="2"><hr /></td></tr>
         <form method="POST" action="index.asp">		
            <tr>
               <td colspan="2">
                  Enter a new ticket.  
               <% If strNewTicketMessage = "Submitted" Then %>
                  <font class="information"><%=strNewTicketMessage%></font>
               <% Else %>
                  <font class="missing"><%=strNewTicketMessage%></font>
               <% End If %>
               </td>
            </tr>
            <tr>
               <td>
                  Username:&nbsp;<input type="text" name="UserName" value="<%=strUserName%>" size="15">
               </td>
               <td>
               <input type="submit" value="New Ticket" name="Submit" style="float: right"></td>
            </tr>  
         </form>
   <% End If %>
         
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
<%    If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Techs") And (objMessage(3) = "Bottom" or objMessage(3) = "Both") Then 
         Select Case objMessage(2)
            Case "Normal"
               strMessageFont = ""
            Case "Alert"
               strMessageFont = "<font class=""information"">"
         End Select%>
         <tr><td colspan="2">
            <%=strMessageFont%><%=Replace(objMessage(0),vbCRLF,"<br />")%> </font>
         </td></tr>
         <tr><td colspan="2"><hr /></td></tr> 
<%    End If %> 
      <form method="Post" action="index.asp">
         <tr>
            <td>
               <!-- <input type="submit" value="Search" name="Site"> -->
         <% If Application("UseStats") Then %>
               <input type="submit" value="Stats" name="Site"> 
         <% End If %>
               <input type="submit" value="Settings" name="Site">
            </td>
            <td align="right">
               <%=Application("Version")%>
            </td>
         </tr>   
         <% If IsTablet Then %>
               <tr><td><input type="submit" value="Full Site" name="Submit">
            <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then%>
               <input type="submit" value="Zoom Out" name="Site">
            <% Else %>
               <input type="submit" value="Zoom In" name="Site">
         <%    End If %>
               </td></tr>
         <% End If %>
      </form>
      </table>
      </center>
   </body>
   
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
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />

   <% If Request.Cookies("ZoomLevel") = "ZoomIn" Then %>
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=1.9" />
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %> 
   </head>
   <body>
      <div align="right">
         Help Desk
      </div>
      <hr />
      <form method="Post" action="view.asp">
      <div align="center"> 
         <input type="submit" value="Open Tickets" name="filter"> <br /> <br />
   <% If strRole <> "Data Viewer" Then %>   
         <input type="submit" value="Your Tickets" name="filter">
   <% End If %>
         
   <% If bolShowLogout Then %>   
          <br /> <br /><input type="submit" value="Log Out" name="Log Out">
   <% End If %> 
      </div>
      </form>
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

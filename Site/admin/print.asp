<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 11/20/04
'Last Updated 6/16/14

'This page will list all the fields in one call so it can be printed.

Option Explicit

On Error Resume Next

Dim objNetwork, strSQL, objNameCheckSet, strRole, bolMobileVersion, strUser, bolShowLogout

'If the database and the website are not the same version then let them know
If Application("VersionError") Then
   VersionProblem
End If

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
   'Const OpenTime = 13
   Const Custom1 = 13
   Const Custom2 = 14

   Dim intID, intTest, strType, strSQL, objRecordSet, intCount, strDate, strTime, strDays
   Dim strMinutes, strHours, strTimeActive, objCategorySet, objTechSet, objStatusSet
   Dim objLocationSet, objRegExp, strNotes, strProblem

   intID = Request.Querystring("ID")
      
   'Verify that the intID is a number.  If not then the user enter a non numeric value in for 
   'a ticket number.  If that is the case then set intID to 0 so it will kick out as an error
   'to the user.
   intTest = CInt(intID)
   strType = TypeName(intTest)
   If UCase(strType) <> "INTEGER" Then
      AccessDenied
      Exit Sub
   End If
   
   If intID = "" Then
      AccessDenied
      Exit Sub
   End If
   
   'Build the SQL string that will get the data for the requested ticket
   strSQL = "SELECT Main.ID, Main.Name, Main.Location, Main.Email, Main.Problem, Main.SubmitDate, Main.SubmitTime, Main.Notes, Main.Status, Main.Category, Main.Tech, Main.LastUpdatedDate, Main.LastUpdatedTime, Main.Custom1, Main.Custom2" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE (((Main.ID)=" & intID & "));"

   'Execute the SQL string and assign the results to a Record Set
   Set objRecordSet = Application("Connection").Execute(strSQL)

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

   'Create the Regular Expression object and set it's properties.
   Set objRegExp = New RegExp
   objRegExp.Pattern = vbCRLF
   objRegExp.Global = True

   'Change a carriage return to a <br> so it will disply properly in HTML.
   If Not IsNull(objRecordSet(Notes)) Then
      strNotes = objRegExp.Replace(objRecordSet(Notes),"<br>")
   End If
   If Not IsNull(objRecordSet(Problem)) Then
      strProblem = objRegExp.Replace(objRecordSet(Problem),"<br>")
   End If

   %>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <head>
      <title>Help Desk - Admin</title>
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>" />
   </head>
   <body bgcolor="<%=Application("BGColor")%>">

      <b><font face="Arial" size="4" color="<%=Application("TxtColor")%>">
      Ticket #<%=intID%>
      </font></b><br>
      <font face="Arial" color="<%=Application("TxtColor")%>">
      <b>User</b>: <%=objRecordSet(Name)%> <br>
      <b>EMail</b>: <%=objRecordSet(EMail)%> <br>
      <b>Location</b>: <%=objRecordSet(Location)%> <br>
      
   <% If Application("UseCustom1") Then %>
         <b><%=Application("Custom1Text")%></b>: <%=objRecordSet(Custom1)%> <br>
   <% End If %>

   <% If Application("UseCustom2") Then %>
         <b><%=Application("Custom2Text")%></b>: <%=objRecordSet(Custom2)%> <br>
   <% End If %>

      <b>Submitted</b>: <%=objRecordSet(SubmitDate)%> - <%=objRecordSet(SubmitTime)%> <br>
      
   <% 'If the call has been modified then display the modification date.
      If objRecordSet(LastUpdatedDate) = "6/16/1978" Then %>
         <b>Updated</b>: Never <br>
   <% Else %>
         <b>Updated</b>: <%=objRecordSet(LastUpdatedDate)%> - <%=objRecordSet(LastUpdatedTime)%> <br>
   <% End If %>

      <b>Open Time</b>: <%=strTimeActive%> <br>
      <b>Status</b>: <%=objRecordSet(Status)%> <br>
      <b>Category</b>: <%=objRecordSet(Category)%> <br>
      <b>Tech</b>: <%=objRecordSet(Tech)%> <br>
      <b>Problem</b>: <%=strProblem%> <br> <br>
      <b>Notes</b>: <%=strNotes%> <br>
      
      </font>

      <script language="JavaScript"><!--
         window.print();
      //--></script>
   </body>
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
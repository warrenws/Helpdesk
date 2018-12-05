<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 11/20/04
'Last Updated 6/16/14

'This page is the popup that will show all the items in the database for the modify page.

Option Explicit

On Error Resume Next

Dim strItemType, strSQL, objRecordSet, strEnabled, strRole, strRoleDisplay, objNameCheckSet, strUser
Dim objNetwork, strUserAgent, strTaskListRole, strDocumentationRole, bolMobileVersion, bolShowLogout

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
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

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
If objNameCheckSet(1) = "Administrator" And objNameCheckSet(2) Then

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

   'Get the item type from the URL
   strItemType = Request.QueryString("Item")

   If LCase(strItemType) <> "category" And LCase(strItemType) <> "location" And LCase(strItemType) <> "tech" And LCase(strItemType) <> "assignment" And LCase(strItemType) <> "subnet" Then
      AccessDenied
      Exit Sub
   End If
   
   Select Case strItemType
      Case "Category"
         strSQL = "SELECT Category,Active FROM Category Order By Category"
      Case "Location"
         strSQL = "SELECT Location,Active FROM Location Order By Location"
      Case "Tech"
         strSQL = "SELECT Tech,Active,UserLevel,TaskListRole,DocumentationRole FROM Tech Order By Tech"
      Case "Assignment"
         strSQL = "SELECT Location,Tech FROM Location Order By Location"
      Case "Subnet"
         strSQL = "SELECT Subnet,Location FROM Subnets Order By Location"
      Case Else
         strSQL = "SELECT Location,Active FROM Location Order By Location"
   End Select

   'Execute the SQL string built above to populate the table
   Set objRecordSet = Application("Connection").Execute(strSQL)

   'Get the item type from the URL
   strItemType = Request.QueryString("Item")

   strSQL = "Select Username, UserLevel, Active, Theme" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

   Set objNameCheckSet = Application("Connection").Execute(strSQL)

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
   </head>

   <body>
      <center>
      <b>Current <%=strItemType%>s</b> &nbsp;&nbsp;&nbsp;&nbsp;
      <a href="javascript:window.close()">Close</a>


      <table border="1" width="100%">
         <tr>
         <% Select Case strItemType 
               Case "Assignment" %>
                  <td class="showborders"><b>Location</b></td>
                  <td class="showborders"><b>Tech</b></td>
            <% Case "Subnet" %>
                  <td class="showborders"><b>Subnet</b></td>
                  <td class="showborders"><b>Location</b></td>
            <% Case "Tech" %>
                  <td class="showborders"><b><%=strItemType%></b></td>
                  <td class="showborders"><b>Enabled</b></td>
                  <td class="showborders"><b>HD Role</b></td>
                  <td class="showborders"><b>TL Role</b></td>
                  <td class="showborders"><b>Doc Role</b></td>
            <% Case Else %>
                  <td class="showborders"><b><%=strItemType%></b></td>
                  <td class="showborders"><b>Enabled</b></td>
            <% End Select%>
         </tr>

      <% 'Loop through the record set and display all the items
         Do Until objRecordSet.EOF %>
            <tr>
               <td class="showborders"><%=objRecordSet(0)%></td>
      <%       Select Case objRecordSet(1)
                  Case "True"
                     strEnabled = "Yes"
                  Case "False"
                     strEnabled = "No"
                  Case Else
                     strEnabled = objRecordSet(1)
               End Select %>
               <td class="showborders"><%=strEnabled%></td>
            <%If strItemType = "Tech" Then 
            
               Select Case objRecordSet(2)
                  Case "Administrator"
                     strRoleDisplay = "Admin"
                  Case "User"
                     strRoleDisplay = "User"
                  Case "Data Viewer"
                     strRoleDisplay = "Viewer"
                  Case Else
                     strRoleDisplay = objRecordSet(2)
               End Select
               
               Select Case objRecordSet(3)
                  Case "User"
                     strTaskListRole = "User"
                  Case "Viewer"
                     strTaskListRole = "Viewer"
                  Case "Deny"
                     strTaskListRole = "Deny"
               End Select
                 
               Select Case objRecordSet(4)
                  Case "Full"
                     strDocumentationRole = "Full"
                  Case "Read Only"
                     strDocumentationRole = "Read Only"
                  Case "Deny"
                     strDocumentationRole = "Deny"
               End Select
               
               %>
               <td class="showborders"><%=strRoleDisplay%></td>
               <td class="showborders"><%=strTaskListRole%></td>
               <td class="showborders"><%=strDocumentationRole%></td>

         <% End If %>

            </tr>
      <%    objRecordSet.MoveNext
         Loop %>
         
      </table>
      </center>
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
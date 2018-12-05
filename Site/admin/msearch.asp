<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 12/16/13
'Last Updated 6/16/14

'This is a mobile search page.

Option Explicit

On Error Resume Next

Dim objNetwork, strUser, strSQL, strRole, objNameCheckSet, strUserAgent, bolShowLogout, intZoom

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
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

'Check and see if anonymous access is enabled
If LCase(Left(strUser,4)) = "iusr" Then
   strUser = GetUser
   bolShowLogout = True
Else
   bolShowLogout = False
End If

'Set the zoom level
If Request.Cookies("ZoomLevel") = "ZoomIn" Then
   If InStr(strUserAgent,"Silk") Then
      intZoom = 1.4
   Else
      intZoom = 1.9
   End If
End If

'Build the SQL string, this will check the userlevel of the user.
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"
Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)

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

   Dim objLocationSet, objStatusSet, objTechSet, objCategorySet
   Dim strCurrentUser, strLocation, strCategory, strTech, intInputSize
   
   intInputSize = 25
   
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
         strCurrentUser = GetFirstandLastName(objNetwork.UserName)
   End Select

   On Error Resume Next 
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
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %> 
      
   </head>
   <body>
      <center><b><%=Application("SchoolName")%> Help Desk Admin</b></center>
      <center>
      <table align="center">
         <tr><td width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td>
               <div align="center">
                  <input type="submit" value="Home" name="home">
                  <input type="submit" value="Open Tickets" name="filter">
            <% If strRole <> "Data Viewer" Then %>   
                  <input type="submit" value="Your Tickets" name="filter">
            <% End If %>
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table> 

      <table align="center">    
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
         <tr><td>User:</td><td><input type="text" name="User" value="Any" size="<%=intInputSize%>"></td></tr>
         <tr><td>Problem:</td><td><input type="text" name="Problem" value="Any" size="<%=intInputSize%>"></td></tr>
         <tr><td>Notes:</td><td><input type="text" name="Notes" value="Any" size="<%=intInputSize%>"></td></tr>
         <tr><td>EMail:</td><td><input type="text" name="EMail" value="Any" size="<%=intInputSize%>"></td></tr>
         <tr><td colspan="2"><hr /></td></tr>
         <tr><td colspan="2" align="right" width="<%=Application("MobileSiteWidth")%>">
            <!--
            <input type="submit" value="Back" name="back">
            -->
            <input type="submit" value="Search"> 
         </td></tr>
         </form>
      </table>       
   
<%End Sub %>

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
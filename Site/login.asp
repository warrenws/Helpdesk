<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 4/18/14
'Last Updated 6/16/14

On Error Resume Next

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

'Delete the old session if they are logging out.
If Request.QueryString("Action") = "logout" Then
   strSQL = "DELETE FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
   Application("Connection").Execute(strSQL)
      
   'Clear the cookies
   Response.Cookies("SessionID") = ""
   
End If

'Get needed information
strUserName = Request.Form("UserName")
strPassword = Request.Form("Password")
strLogin = Request.Form("Login")
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
strIPAddress = Request.ServerVariables("REMOTE_ADDR")

'Remove all old sessions
DeleteOldSessions

'If they are already logged in send them back
If Request.Cookies("SessionID") <> "" Then
   strSQL = "SELECT UserAgent FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
   Set objActiveSession = Application("Connection").Execute(strSQL)
   
   If Not objActiveSession.EOF Then
   
      'This line fixes a redirect loop, the user would be sent back and fourth if the useragent changed since last time
      If Left(Replace(objActiveSession(0),"'","''"),250) = Left(Replace(strUserAgent,"'","''"),250) Then

         'Redirect the user to the page they came from, or to the default page
         strSourcePage = Request.QueryString("SourcePage")
         If strSourcePage = "" or strSourcePage = "view.asp" Then
            Response.Redirect("index.asp")
         Else
            Response.Redirect(strSourcePage & BuildReturnLink)
         End If
         
      Else
         
         'The user agent has changed since the last login, we're going to delete the old session
         strSQL = "DELETE FROM Sessions WHERE SessionID='" & Replace(Request.Cookies("SessionID"),"'","''") & "'"
         Application("Connection").Execute(strSQL)
         
      End If
      
   End If
End If

'Build return string
If Request.ServerVariables("QUERY_STRING") = "" Then
   strReturnLink = ""
Else
   strReturnLink = "?" & Request.ServerVariables("QUERY_STRING")
End If

'If they hit the login button
If strLogin = "Login" Then

   'Fix the username if it's an email address
   If InStr(strUserName,"@") Then
      strUserName = Left(strUserName,InStr(strUserName,"@") - 1)
   End If

   'Fix the username if it's in legacy form
   If InStr(strUserName,"\") Then
      strUserName = Right(strUserName,Len(strUserName) - InStr(strUserName,("\")))
   End If

   'Create objects required to connect to AD
   Set objConnection = CreateObject("ADODB.Connection")
   Set objCommand = CreateObject("ADODB.Command")
   Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

   'Create a connection to AD
   objConnection.Provider = "ADSDSOObject"

   objConnection.Open "Active Directory Provider",strUserName & "@" & Application("Domain"), strPassword
   objCommand.ActiveConnection = objConnection
   strDNSDomain = objRootDSE.Get("DefaultNamingContext")
   objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); GivenName,SN,name ;subtree"

   'Initiate the LDAP query and return results to a RecordSet object.
   Set objRecordSet = objCommand.Execute
   
   If Trim(strUserName) = "" Then
   
      strMessage = "Invalid Password"
      strMessageType = "Error"
   
   Else
  
      'If the connection works then we have the correct username and password
      If Err.Number = 0 Then
         
         'See if they created a session in the past 10 seconds.
         strSQL = "SELECT SessionID,LoginTime FROM Sessions WHERE Username='" & strUserName & "' AND UserAgent='" & strUserAgent &"' And LoginDate=Date()"
         Set objActiveSession = Application("Connection").Execute(strSQL)
         
         If objActiveSession.EOF Then
            CreateNewSession
         Else
            intSeconds = DateDiff("s",objActiveSession(1),Time())
            If intSeconds > 10 Then
               CreateNewSession
            Else
               Response.Redirect("index.asp")
            End If
         End If
         
         'Redirect the user to the page they came from, or to the default page
         strSourcePage = Request.QueryString("SourcePage")
         If strSourcePage = "" or strSourcePage = "view.asp" Then
            Response.Redirect("index.asp")
         Else
            Response.Redirect(strSourcePage & BuildReturnLink)
         End If
         
      Else
         Err.Clear
         strMessage = "Invalid Password"
         strMessageType = "Error"
      End If
      
   End If
   
End If

'Set the size of the input boxes, they show up wrong on Windows Phone
If InStr(strUserAgent,"Windows Phone 8") Then
   intUserNameBoxSize = 20
   intPasswordBoxSize = 25
ElseIf InStr(strUserAgent,"Windows Phone OS 7") Then
   intUserNameBoxSize = 20
   intPasswordBoxSize = 23
Else
   intUserNameBoxSize = 20
   intPasswordBoxSize = 20
End If

%>

<html>

<head>
   <title>HDL - Admin</title>
   <link rel="stylesheet" type="text/css" href="../themes/<%=Application("Theme")%>/<%=Application("Theme")%>.css" />
   <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
   <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   
</head>

<body>
   <center>
   <table>
   <form method="post" action="login.asp<%=strReturnLink%>">
      <tr><th colspan="2" align="center"><%=Application("SchoolName")%> Help Desk</th></tr>
      <tr><td colspan="2"><hr /></td></tr>
      <tr>
         <td>Username</td>
         <td><input type="textbox" name="Username" size="<%=intUserNameBoxSize%>" /></td>
      </tr>
      <tr>
         <td>Password</td>
         <td><input type="password" name="Password" size="<%=intPasswordBoxSize%>" /></td>
      </tr>
      <tr><td colspan="2" align="right"><input type="submit" value="Login" name="Login" /></td></tr>
      <tr>
         <td colspan="2" align="center">
            <div class="<%=strMessageType%>"><%=strMessage%></div>
         </td>
      </tr>
   </table>
   </form>
   </center>
</body>

</html>

<%
Sub CreateNewSession

   intSessionID = GenerateSessionID
   strSQL = "INSERT INTO Sessions " & _
   "(Username,SessionID,IPAddress,UserAgent,LoginDate,LoginTime,ExpirationDate) VALUES " & _
   "('" & strUserName & "','" & intSessionID & "','" & strIPAddress & "','" & _
   Left(Replace(strUserAgent,"'","''"),250) & "',Date(),Time(),#" & Date() + Application("UserLogInDays") & "#)"
   Application("Connection").Execute(strSQL)
   Response.Cookies("SessionID") = intSessionID
   Response.Cookies("SessionID").Expires = Date() + Application("UserLogInDays")
End Sub 
%>

<%
Sub DeleteOldSessions

   strSQL = "DELETE FROM Sessions WHERE Date() >= ExpirationDate"
   Application("Connection").Execute(strSQL)

End Sub
%>

<%
Function GenerateSessionID
   
   'Get a random number 
   GenerateSessionID = GetRandomNumber(1000000000,9999999999)

   'See if it's already in use in the database
   strSQL = "SELECT ID FROM Sessions WHERE SessionID ='" & GenerateSessionID & "'"
   Set objSessionCheck = Application("Connection").Execute(strSQL)
   If Not objSessionCheck.EOF Then
      GenerateSessionID = GenerateSessionID()
   End If
   
End Function
%>

<%
Function GetRandomNumber(intLow,intHigh)
	Randomize
	GetRandomNumber = (Int(RND * (intHigh - intLow + 1))) + intLow
End Function
%>

<%
Function BuildReturnLink

   strQueryString = Request.ServerVariables("QUERY_STRING")

   'Build the return link
   strLinkID = Request.QueryString("ID")
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
   strLinkRating = Request.QueryString("Rating")
   strLinkTicket = Request.QueryString("Ticket")

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
   If strLinkRating <> "" Then
      BuildReturnLink = BuildReturnLink & "&Rating=" & strLinkRating
   End If
   If strLinkTicket <> "" Then
      BuildReturnLink = BuildReturnLink & "&Ticket=" & strLinkTicket
   End If
   
   If BuildReturnLink <> "" Then
      BuildReturnLink = "?" & Right(BuildReturnLink,(Len(BuildReturnLink)-1))
   End If
   
End Function
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
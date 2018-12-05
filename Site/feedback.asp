<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 1/18/12
'Last Updated 6/16/14

'This is the feedback page.
 
Option Explicit

'On Error Resume Next

Dim intRating, intTicket, strSQL, objOldFeedback, objTicketData, strComment, objNetwork
Dim strName, strSubmit, bolShowLogout, strUser, strUserAgent

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

intTicket = Request.QueryString("Ticket")
intRating = Request.QueryString("Rating")
strComment = Request.Form("Comment")
strSubmit = Request.Form("cmdSubmit")

If intTicket = "" Then
   intTicket = " "
End If

If intRating = "" Then
   intRating = " "
End If

If IsNumeric(intTicket) And IsNumeric(intRating) Then
   Select Case intRating
      Case 1, 2, 3, 4, 5
         strSQL = "SELECT Tech, Location, Name FROM Main WHERE ID =" & intTicket
         Set objTicketData = Application("Connection").Execute(strSQL)
         
         If objTicketData.EOF Then
            Response.Redirect("index.asp")
         Else
            If LCase(objTicketData(2)) <> LCase(strUser) Then
               Response.Redirect("index.asp")
            Else
               Main
            End If
         End If 
      Case Else
         Response.Redirect("index.asp")
   End Select
Else
   Response.Redirect("index.asp")
End If
%>

<%Sub Main 
   strSQL = "SELECT ID FROM Feedback WHERE Ticket =" & intTicket
   Set objOldFeedback = Application("Connection").Execute(strSQL)
   
   If objOldFeedback.EOF Then
      strSQL = "INSERT INTO Feedback (Ticket,Rating,Tech,Location,DateSubmitted,TimeSubmitted)" & vbCRLF
      strSQL = strSQL & "VALUES (" & intTicket & "," & intRating & ",'" & objTicketData(0) & "','" & objTicketData(1) & "','" & Date() & "','" & Time() & "')"

      Application("Connection").Execute(strSQL)
   Else
      If strSubmit = "Submit" Then
         strSQL = "UPDATE Feedback" & vbCRLF
         strSQL = strSQL & "SET Comment='" & Replace(strComment,"'","''") & "',DateSubmitted='" & Date() & "',TimeSubmitted='" & Time() & "'" & vbCRLF
         strSQL = strSQL & "WHERE ID=" & objOldFeedback(0)
      Else
         strSQL = "UPDATE Feedback" & vbCRLF
         strSQL = strSQL & "SET Rating=" & intRating & vbCRLF & ",DateSubmitted='" & Date() & "',TimeSubmitted='" & Time() & "'" & vbCRLF
         strSQL = strSQL & "WHERE ID=" & objOldFeedback(0)
      End If
      Application("Connection").Execute(strSQL)
   End If
   
   If IsMobile Then
      MobileVersion
   Else
      FullVersion
   End If
End Sub%>

<%Sub FullVersion %>

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
      <%=Application("SchoolName")%> Help Desk
   </div>
   
   <div class="version">
      Version <%=Application("Version")%>
   </div>
   <hr class="admintopbar" />
   
   <center>
   <table width="750">
      <form method="POST" action="feedback.asp?rating=<%=intRating%>&ticket=<%=intTicket%>">
      <input type="hidden" name="ticket" value="<%=intTicket%>">
      <input type="hidden" name="rating" value="<%=intRating%>">
      <tr><td width="33%"><h1>Thank you!</h1></td>
         <td>
         <% If intRating > 3 Then %>
               We are pleased to hear you had a positive experience. Please let us know if you have any additional comments. 
               Your feedback will allow us to continue to improve our service. 
               Thank you again for your time.
         <% Else %>
               We're sorry to hear you didn't have a great experience. Can you tell us more about it? 
               Your feedback will allow us to continue to improve our service. Thank you again for your time.
         <% End If %>
         </td>
      </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td>&nbsp;</td><td><b>Comments</b></td></tr>
      <tr><td>&nbsp;</td>
         <td>
          <textarea rows="8" name="Comment" cols="90" style="width: 500px;"><%=strComment%></textarea>
         </td>
      </tr>
      <tr><td align="right" colspan="2"><input type="submit" value="Submit" name="cmdSubmit"></td></tr>
   <% If strSubmit = "Submit" Then %>
         <tr><td>&nbsp;</td><td>Comment submitted - Thank you...</td></tr>
   <% End If %>   
      </form>
   </table>
   </center>
   
<%End Sub%>

<%Sub MobileVersion %>

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
         <meta name="viewport" content="width=100,user-scalable=no,initial-scale=1.9" />
      <% Else %>
         <meta name="viewport" content="width=device-width,user-scalable=no" /> 
      <% End If %> 

      </head>
  <body>  
      <center><b><%=Application("SchoolName")%> Help Desk</b></center>
      <center>
      <table align="center">
         <tr><td><hr /></td></tr>
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">         
               <h1>Thank you!</h1>
            </td>
         <tr>
            <td>
            <% If intRating > 3 Then %>
                  We are pleased to hear you had a positive experience. Please let us know if you have any additional comments. 
                  Your feedback will allow us to continue to improve our service. 
                  Thank you again for your time.
            <% Else %>
                  We're sorry to hear you didn't have a great experience. Can you tell us more about it? 
                  Your feedback will allow us to continue to improve our service. Thank you again for your time.
            <% End If %>
            </td>
         </tr>
      <tr><td>&nbsp;</td></tr>
      <tr><td><b>Comments</b></td></tr>
      <tr>
         <td>
                         <form method="POST" action="feedback.asp?rating=<%=intRating%>&ticket=<%=intTicket%>">
               <input type="hidden" name="ticket" value="<%=intTicket%>">
               <input type="hidden" name="rating" value="<%=intRating%>">  
          <textarea rows="8" name="Comment" cols="90" style="width: 99%;"><%=strComment%></textarea>
         </td>
      </tr>
      <tr><td align="right" colspan="2"><input type="submit" value="Submit" name="cmdSubmit"></td></tr>
   <% If strSubmit = "Submit" Then %>
         <tr><td>&nbsp;</td><td>Comment submitted - Thank you...</td></tr>
   <% End If %>   
      </form>
   </table>
   </center>
   
<%End Sub%>

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
Function BuildReturnLink(bolIncludeID)

   Dim strLinkID, strLinkLocation, strLinkStatus, strLinkTech, strLinkCategory, strLinkUser, strLinkFilter, strLinkProblem
   Dim strLinkNotes, strLinkEMail, strLinkSort, strLinkDays, strLinkBack, strLinkViewed, strLinkRating, strLinkTicket

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
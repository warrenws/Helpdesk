<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 3/28/12
'Last Update 3/28/12

'This page is designed to run on a TV to give important information on one page and auto update.

Option Explicit

On Error Resume Next

Dim objNetwork, strUserName, strSQL, objNameCheckSet, strRole, objRecentTickets, strTicketDate
Dim strUserAgent, intLatestTicket, intOldLatestTicket, bolNewTicketArrived, strTicketTime, strFixSite
Dim objTotalOpenTickets, intTotalOpenTickets, intSiteNameLength, intMaxNameLength

'Get the users logon name
Set objNetwork = CreateObject("WSCRIPT.Network")   
strUserName = objNetwork.UserName
strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

'Build the SQL string
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUserName & "'));"

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

   On Error Resume Next

   strSQL = "SELECT Top 5 ID, DisplayName, Location, SubmitTime, SubmitDate, TicketViewed" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status <> 'Complete'" & vbCRLF
   strSQL = strSQL & "ORDER BY ID DESC;"
   Set objRecentTickets = Application("Connection").Execute(strSQL)
   
   intLatestTicket = objRecentTickets(0)
   intOldLatestTicket = Request.QueryString("LatestTicket")
   
   If CInt(intLatestTicket) > CInt(intOldLatestTicket) Then
      bolNewTicketArrived = True
   Else
      bolNewTicketArrived = False
   End If
      
   If IsNull(intOldLatestTicket) Or intOldLatestTicket = "" Then
      bolNewTicketArrived = False
   End If

   strSQL = "SELECT Count(Name) AS CountOfName" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete"""
   Set objTotalOpenTickets = Application("Connection").Execute(strSQL) 
   
   If objTotalOpenTickets.EOF Then
      intTotalOpenTickets = 0
   Else
      intTotalOpenTickets = objTotalOpenTickets(0)
   End If   

   intMaxNameLength = 0
   If NOT objRecentTickets.EOF Then
      Do Until objRecentTickets.EOF  
         If Len(objRecentTickets(1)) > intMaxNameLength Then
            intMaxNameLength = Len(objRecentTickets(1))
         End If
         objRecentTickets.MoveNext
      Loop
   End If
   objRecentTickets.MoveFirst
   intSiteNameLength = 35 - intMaxNameLength
   
%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>Help Desk</title>
      <link rel="stylesheet" type="text/css" href="style.css" />

      <link rel="apple-touch-icon-precomposed" href="<%=Application("IconLocation")%>/ipadadminicon.png" />
      <link rel="shortcut icon" href="<%=Application("IconLocation")%>/helpdesk.ico" />
   <% If InStr(strUserAgent,"iPad") or InStr(strUserAgent,"Transformer") Then %>
         <meta name="viewport" content="width=device-width" />
   <% End If %>
   <% If InStr(strUserAgent,"iPhone") Then %>
      <meta name="viewport" content="initial-scale=.41" />
   <% End If %>
      <meta http-equiv="refresh" content="60;url=index.asp?LatestTicket=<%=intLatestTicket%>" >
   </head>

   <body style="overflow:auto">
   
   <% If bolNewTicketArrived Then %>
         <audio src="sounds/alert.wav" autoplay="autoplay"> </audio>
   <% End If %>
      
      <div id="schoolHeader"><%=Application("SchoolName")%> (<%=intTotalOpenTickets%>)</a></div>
      <table width="650px">
         <tr><td valign="top">
         <table>
            <tr><td valign="top">
               <table>
               <% If NOT objRecentTickets.EOF Then
                     Do Until objRecentTickets.EOF 
                  
                        'Fix the date and time
                        strTicketDate = Left(objRecentTickets(4),Len(objRecentTickets(4))-5)
                        strTicketTime = Left(objRecentTickets(3),Len(objRecentTickets(3))-6) & " " & LCase(Right(objRecentTickets(3),2))
                        
                        Select Case objRecentTickets(2)
                           Case "Mary J Tanner"
                              strFixSite = "MJT"
                           Case "Elementary School"
                              strFixSite = "ES"
                           Case Else
                              strFixSite = objRecentTickets(2)
                        End Select
                        %>
                        <tr>
                           <td>
                        <% If objRecentTickets(5) Then %>
                              <center><img border="0" src="images/viewed.gif" alt="Viewed by Tech" width="20" height="20"></center>
                        <% Else %>
                              <center><img border="0" src="images/notviewed.gif" alt="Not Viewed by Tech" width="20" height="20"></center>
                        <% End If%>
                           </td>
                        <% If CInt(intOldLatestTicket) >= CInt(objRecentTickets(0)) Or Not bolNewTicketArrived Then %>
                              <td align="center">&nbsp;&nbsp;<a href="../modify.asp?ID=<%=objRecentTickets(0)%>" target="_blank"><%=objRecentTickets(0)%></a>&nbsp;&nbsp;</td>
                              <td><%=objRecentTickets(1)%>&nbsp;&nbsp;&nbsp;</td>
                              <td><%=Left(strFixSite,intSiteNameLength)%>&nbsp;&nbsp;&nbsp;</td>
                              <td><%=strTicketDate%>&nbsp;&nbsp;&nbsp;</td>
                              <td><%=strTicketTime%>&nbsp;&nbsp;&nbsp;</td>
                        <% Else %>
                              <td class="highlight" align="center">&nbsp;&nbsp;<a href="../modify.asp?ID=<%=objRecentTickets(0)%>" target="_blank"><%=objRecentTickets(0)%></a>&nbsp;&nbsp;</td>
                              <td class="highlight"><%=objRecentTickets(1)%>&nbsp;&nbsp;&nbsp;</td>
                              <td class="highlight"><%=strFixSite%>&nbsp;&nbsp;&nbsp;</td>
                              <td class="highlight"><%=strTicketDate%>&nbsp;&nbsp;&nbsp;</td>
                              <td class="highlight"><%=strTicketTime%>&nbsp;&nbsp;&nbsp;</td>
                        <% End If %>
                        </tr>
                  <%    objRecentTickets.MoveNext 
                     Loop
                  Else %>
                     <tr><td>No Open Tickets</td></tr>
               <% End If %>
                  
               </table>
            </td>
            <td valign="top">   
            </td>
            <td valign="top">
            </td></tr>
         </table>
			</td></tr>
      </table>                     
   </body>
   </html>
<%End Sub%>

<%Sub AccessDenied %>
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
   
<%End Sub%>
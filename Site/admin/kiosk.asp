<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 3/28/12
'Last Updated 6/16/14

'This page is designed to run on a TV to give important information on one page and auto update.

Option Explicit

'On Error Resume Next

Dim objNetwork, strUser, strSQL, objNameCheckSet, strRole, objRecentTickets, objNumOpenTicketsPerLocation
Dim objNumOpenTicketsPerTech, objCompletedTickets, objTodaysNewTickets, intTodaysNewTickets
Dim objCompleteTicketCount, intCompleteCount, objTodaysTicketCount, intTicketCount, objTotalOpenTickets
Dim intTotalOpenTickets, intMinute, strUserAgent, intMoveScreen, objFeedback, strURLRoot, strRating
Dim intLatestTicket, intOldLatestTicket, bolNewTicketArrived, intActiveLocationCount, intOldID
Dim intActiveTechCount, intLatestOpenCount, objOverallRating, intIndex, objFSO, strRoot, objFolder, objFile
Dim intFileCount, strImageName, bolFound, intFeedbackCount, bolEnbableFeedback, bolShowLogout
Dim bolShowCharts, objTicketsOpenPerTech, objTicketsOpenPerLocation, strTechName

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

'Build the SQL string
strSQL = "Select Username, UserLevel, Active, Theme, MobileVersion, TaskListRole, DocumentationRole" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"

bolEnbableFeedback = False

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

   'On Error Resume Next
   
   bolShowCharts = True
   
   strSQL = "SELECT Location, Count(Location) AS OpenTicketsperSite" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Location" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Location) DESC"
   Set objNumOpenTicketsPerLocation = Server.CreateObject("ADODB.Recordset")
   objNumOpenTicketsPerLocation.CursorLocation = 3
   objNumOpenTicketsPerLocation.Open strSQL, Application("Connection")
   intActiveLocationCount = objNumOpenTicketsPerLocation.RecordCount
   
   strSQL = "SELECT Tech,Count(Tech) AS OpenTicketsperTech" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status<>""Complete""" & vbCRLF
   strSQL = strSQL & "GROUP BY Tech" & vbCRLF
   strSQL = strSQL & "ORDER BY Count(Tech) DESC"
   Set objNumOpenTicketsPerTech = Server.CreateObject("ADODB.Recordset")
   objNumOpenTicketsPerTech.CursorLocation = 3
   objNumOpenTicketsPerTech.Open strSQL, Application("Connection")   
   intActiveTechCount = objNumOpenTicketsPerTech.RecordCount

   If intActiveTechCount >= intActiveLocationCount Then
      intLatestOpenCount = intActiveTechCount + 6
   Else
      intLatestOpenCount = intActiveLocationCount + 6
   End if
   intLatestOpenCount = 12
   intFeedbackCount = 15 - intLatestOpenCount
   
   strSQL = "SELECT Top " & intLatestOpenCount & " ID, DisplayName, Location, SubmitTime, SubmitDate, TicketViewed" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "WHERE Status <> 'Complete'" & vbCRLF
   strSQL = strSQL & "ORDER BY ID DESC;"
   Set objRecentTickets = Application("Connection").Execute(strSQL)
   
   If Not objRecentTickets.EOF Then
      intLatestOpenCount = 0
      Do Until objRecentTickets.EOF
         objRecentTickets.MoveNext
         intLatestOpenCount = intLatestOpenCount + 1
      Loop
      objRecentTickets.MoveFirst
   End If
   
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
   
   strSQL = "SELECT Top 10 Log.Ticket, Main.DisplayName, Main.Location, Main.Tech, Log.UpdateTime, Log.UpdateDate" & vbCRLF
   strSQL = strSQL & "FROM Log INNER JOIN Main ON Log.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "WHERE Log.NewValue='Complete'" & vbCRLF
   strSQL = strSQL & "ORDER BY Log.ID DESC"
   Set objCompletedTickets = Application("Connection").Execute(strSQL)
   
   strSQL = "SELECT Count(Main.ID) AS Tickets" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "GROUP BY SubmitDate" & vbCRLF
   strSQL = strSQL & "HAVING SubmitDate=Date()"
   Set objTodaysNewTickets = Application("Connection").Execute(strSQL) 
   
   If objTodaysNewTickets.EOF Then
      intTodaysNewTickets = 0
   Else
      intTodaysNewTickets = objTodaysNewTickets(0)
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
   
   strSQL = "SELECT Count(ID) AS CountOfID" & vbCRLF
   strSQL = strSQL & "FROM Main"
   Set objTodaysTicketCount = Application("Connection").Execute(strSQL)
   
   If objTodaysTicketCount.EOF Then
      intTicketCount = 0
   Else
      intTicketCount = objTodaysTicketCount(0)
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

   strSQL = "SELECT Avg(Rating) AS AvgOfRating, Count(Rating) AS CountOfRating" & vbCRLF
   strSQL = strSQL & "FROM Feedback"
   Set objOverallRating = Application("Connection").Execute(strSQL)
  
   strSQL = "SELECT TOP " & intFeedbackCount & " Feedback.Ticket, Main.DisplayName, Feedback.Rating, Feedback.Tech, Feedback.Location, Feedback.TimeSubmitted, Feedback.DateSubmitted" & vbCRLF
   strSQL = strSQL & "FROM Feedback INNER JOIN Main ON Feedback.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "ORDER BY Feedback.ID DESC"
   Set objFeedback = Application("Connection").Execute(strSQL) 
   
   If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then
      strURLRoot = "../themes/" & Application("Theme") & "/images/stars/"
   Else
      strURLRoot = "../themes/" & objNameCheckSet(3) & "/images/stars/"
   End If   
   
   intMinute = Minute(Time())
   intMoveScreen = intMinute Mod 2
   
   'Override screen shifting
   intMoveScreen = 0
   
   'Get the next image  
   Set objFSO = CreateObject("Scripting.FileSystemObject")
   If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then
      strRoot = Application("ThemeLocation") & "/" & Application("Theme") & "/images/kiosk"
   Else
      strRoot = Application("ThemeLocation") & "/" & objNameCheckSet(3) & "/images/kiosk"
   End If

   Set objFolder = objFSO.GetFolder(strRoot)

   If Request.Cookies("NextImage") = "" Then
      Response.Cookies("NextImage") = 0
   End If
   
   intFileCount = 0
   For Each objFile in objFolder.Files
      If objFile.Name <> "Thumbs.db" Then
         intFileCount = intFileCount + 1
      End If
   Next
   
   If CInt(intFileCount) = CInt(Request.Cookies("NextImage")) Then
      Response.Cookies("NextImage") = 1
   Else
      Response.Cookies("NextImage") = Request.Cookies("NextImage") + 1
   End If

   bolFound = False
   If CInt(intFileCount) > 0 Then
      intIndex = 0
      For Each objFile in objFolder.Files
         intIndex = intIndex + 1
         If CInt(intIndex) = CInt(Request.Cookies("NextImage")) Then
            strImageName = objFile.Name
            bolFound = True
         End If
      Next
   End If

   If Not bolFound Then
      Response.Cookies("NextImage") = 0
   End If
   
%>
   <!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" 
   "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
   <html>
   <head>
      <title>Help Desk Kiosk</title>
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
   <% If InStr(strUserAgent,"Nexus") Then %>
      <meta name="viewport" content="width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no, target-densitydpi=device-dpi"/>
   <% End If %>
   <% If InStr(strUserAgent,"GT-N5110") Then %>
         <meta name="viewport" content="width=device-width, initial-scale=.77, maximum-scale=1.0, user-scalable=no"/>
   <% End If %>
      <meta http-equiv="refresh" content="60;url=kiosk.asp?LatestTicket=<%=intLatestTicket%>" >
      
   <% If bolShowCharts Then 

         strSQL = "SELECT Tech, Count(ID) AS CountOfID" & vbCRLF
         strSQL = strSQL & "FROM Main" & vbCRLF
         strSQL = strSQL & "WHERE Status<>'Complete'" & vbCRLF
         strSQL = strSQL & "GROUP BY Tech" & vbCRLF
         strSQL = strSQL & "ORDER BY Tech"
         Set objTicketsOpenPerTech = Application("Connection").Execute(strSQL) 

         strSQL = "SELECT Location, Count(ID) AS CountOfID" & vbCRLF
         strSQL = strSQL & "FROM Main" & vbCRLF
         strSQL = strSQL & "WHERE Status<>'Complete'" & vbCRLF
         strSQL = strSQL & "GROUP BY Location" & vbCRLF
         strSQL = strSQL & "ORDER BY Location"
         Set objTicketsOpenPerLocation = Application("Connection").Execute(strSQL) 
   %>   
      <script type="text/javascript" src="https://www.google.com/jsapi"></script>
      <script type="text/javascript">
         google.load("visualization", "1", {packages:["corechart"]});
         google.setOnLoadCallback(drawOpenPerTechChart);
         google.setOnLoadCallback(drawOpenPerLocationChart);
         
         function drawOpenPerTechChart() {
            var data = google.visualization.arrayToDataTable([ 
               ['Tech', 'Open Tickets'],
      <% If Not objTicketsOpenPerTech.EOF Then
            Do Until objTicketsOpenPerTech.EOF 
               If objTicketsOpenPerTech(0) = "" or objTicketsOpenPerTech(0) = " " Then 
                  strTechName = "Unassigned"
               ElseIf InStr(objTicketsOpenPerTech(0)," ") Then
                  strTechName = Trim(Left(objTicketsOpenPerTech(0), InStr(objTicketsOpenPerTech(0)," ")))
               Else
                  strTechName = objTicketsOpenPerTech(0)
               End If%>
               ['<%=strTechName & " (" & objTicketsOpenPerTech(1) & ")"%>', <%=objTicketsOpenPerTech(1)%>],
         <% objTicketsOpenPerTech.MoveNext 
            Loop 
         End If%>                
            ]);

            var options = {
               title: 'Number of Open Tickets Per Tech',
               is3D: true,
               pieSliceText: 'percent',
               chartArea:{left:25,top:20,width:'95%',height:'95%'},
               titleTextStyle:{fontSize: 14},
               legend:{textStyle: {fontSize: 14}, alignment: 'center'},
            };
            
            var chart = new google.visualization.PieChart(document.getElementById('openPerTech'));
            chart.draw(data, options);
         }
         
         function drawOpenPerLocationChart() {
            var data = google.visualization.arrayToDataTable([ 
               ['Location', 'Open Tickets'],
         <% If Not objTicketsOpenPerLocation.EOF Then
               Do Until objTicketsOpenPerLocation.EOF %>
                  ['<%=objTicketsOpenPerLocation(0) & " (" & objTicketsOpenPerLocation(1) & ")"%>', <%=objTicketsOpenPerLocation(1)%>],
            <% objTicketsOpenPerLocation.MoveNext 
               Loop 
            End If%>   

            ]);

            var options = {
               title: 'Number of Open Tickets Per Location',
               is3D: true,
               pieSliceText: 'percent',
               chartArea:{left:25,top:20,width:'95%',height:'95%'},
               titleTextStyle:{fontSize: 14},
               legend:{textStyle: {fontSize: 14}, alignment: 'center'},
            };
            
            var chart = new google.visualization.PieChart(document.getElementById('openPerLocation'));
            chart.draw(data, options);
         }    
      </script>
   <%End If %>

   </head>

   <body style="overflow:auto">
   <center>
   <% If bolNewTicketArrived Then %>
         <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
               <audio src="../themes/<%=Application("Theme")%>/sounds/alert.mp3" autoplay="autoplay"> </audio>
         <% Else %>
               <audio src="../themes/<%=objNameCheckSet(3)%>/sounds/alert.mp3" autoplay="autoplay"> </audio>
         <% End If %>
   <% End If %>
   
      <center><%=Application("SchoolName")%></center>
      <table align="center">
      <% If intMoveScreen = 1 Then %>
         <tr><td align="center">&nbsp;</td></tr>
      <% End If %>
         <tr><td valign="top">
         <table>
            <tr><td valign="top">
               <table>
                  <tr><th colspan="6"><%=intLatestOpenCount%> Latest Open Tickets</th></tr>
               <% Do Until objRecentTickets.EOF %>
                     <tr>
                        <td class="showborders">
                     <% If objRecentTickets(5) Then %>
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
                        </td>
                        <td class="showborders" align="center">&nbsp;&nbsp;<a href="modify.asp?ID=<%=objRecentTickets(0)%>" target="_blank"><%=objRecentTickets(0)%></a>&nbsp;&nbsp;</td>
                        <td class="showborders"><%=Left(objRecentTickets(1),15)%></td>
                        <td class="showborders"><%=objRecentTickets(2)%>&nbsp;&nbsp;&nbsp;</td>
                        <td class="showborders"><%=objRecentTickets(4)%>&nbsp;&nbsp;&nbsp;</td>
                        <td class="showborders"><%=objRecentTickets(3)%>&nbsp;&nbsp;&nbsp;</td>
                     </tr>
               <%    objRecentTickets.MoveNext 
                  Loop%>
                  
               </table>
            </td>
            
      <% If bolShowCharts Then %>
      
            <td>&nbsp;</td>
            
            <td valign="top">
            
               <div id="openPerTech" style="height: 140px;"></div> <br />
               <div id="openPerLocation" style="height: 140px;"></div> <br />
            </td>
      <% Else %>
            
         <% If intMoveScreen = 1 Then %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% Else %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% End If %>
            <td valign="top">
               <table>
               <tr><th colspan="2">Tickets Per Location</th></tr>
               <% Do Until objNumOpenTicketsPerLocation.EOF %>
                  <tr>
                     <td class="showborders"><a href="view.asp?Location=<%=objNumOpenTicketsPerLocation(0)%>&Status=Any%20Open%20Ticket" target="_blank"><%=objNumOpenTicketsPerLocation(0)%></a></td>
                     <td class="showborders" align="center"><%=objNumOpenTicketsPerLocation(1)%></td>
                  </tr>
               <%    objNumOpenTicketsPerLocation.MoveNext
                  Loop%>
               </table>   
            </td>
         <% If intMoveScreen = 1 Then %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% Else %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% End If %>
            <td valign="top">
               <table>
               <tr><th colspan="2">Tickets Per Tech</th></tr>
               <% Do Until objNumOpenTicketsPerTech.EOF %>
                  <tr>
                     <% If objNumOpenTicketsPerTech(0) = "Jacqueline Chromczak" Then %>
                           <td class="showborders"><a href="view.asp?Tech=<%=objNumOpenTicketsPerTech(0)%>&Status=Any%20Open%20Ticket" target="_blank">Jackie Chromczak</a></td>
                     <% Else %>
                           <td class="showborders"><a href="view.asp?Tech=<%=objNumOpenTicketsPerTech(0)%>&Status=Any%20Open%20Ticket" target="_blank"><%=objNumOpenTicketsPerTech(0)%></a></td>
                     <% End If %>
                     <td class="showborders" align="center">&nbsp;<%=objNumOpenTicketsPerTech(1)%>&nbsp;</td>
                  </tr>
               <%    objNumOpenTicketsPerTech.MoveNext
                  Loop%>
               </table>
      <% End If %>
            </td></tr>
         </table>
         </td></tr>
         <tr><td><hr /></td></tr>
            <tr><td valign="top">
            <table>
            <tr><td valign="top">
               <table>
                  <tr><th colspan="6">Last 5 Tickets Completed</th></tr>
               <% For intIndex = 1 to 5 
                     If NOT objCompletedTickets.EOF AND objCompletedTickets(1) <> intOldID Then %>
                        <tr>
                           <td class="showborders" align="center">&nbsp;&nbsp;<a href="modify.asp?ID=<%=objCompletedTickets(0)%>" target="_blank"><%=objCompletedTickets(0)%></a>&nbsp;&nbsp;</td>
                           <td class="showborders"><%=objCompletedTickets(1)%>&nbsp;&nbsp;&nbsp;</td>
                           <td class="showborders"><%=objCompletedTickets(2)%>&nbsp;&nbsp;&nbsp;</td>
                           <td class="showborders"><%=objCompletedTickets(3)%>&nbsp;&nbsp;&nbsp;</td>
                           <td class="showborders"><%=objCompletedTickets(5)%>&nbsp;&nbsp;&nbsp;</td>
                           <td class="showborders"><%=objCompletedTickets(4)%>&nbsp;&nbsp;&nbsp;</td>
                        </tr>
                  <% Else 
                        intIndex = intIndex - 1
                     End If 
                     intOldID = objCompletedTickets(1)
                     objCompletedTickets.MoveNext 
                  Next%>
                  
               </table>
            </td>
         <% If intMoveScreen = 1 Then %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% Else %>
               <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
         <% End If %>
            <td valign="top">
               <table>
                  <tr>
                     <th colspan="2">Database Stats</th>
                  <tr>
                     <td class="showborders">Total tickets</td>
                     <td class="showborders" align="center">&nbsp;&nbsp;<%=intTicketCount%>&nbsp;&nbsp;</td>
                  </tr>
                  <tr>
                     <td class="showborders"><a href="view.asp?Filter=AllOpenTickets">Open tickets</a></td>
                     <td class="showborders" align="center"><%=intTotalOpenTickets%></td>
                  </tr>
                  <tr>
                     <td class="showborders"><a href="view.asp?date=<%=Date()%>&type=submitted">Submitted today</a></td>
                     <td class="showborders" align="center"><%=intTodaysNewTickets%></td>
                  </tr>
                  <tr>
                     <td class="showborders"><a href="view.asp?date=<%=Date()%>&type=completed">Completed today</a>&nbsp;&nbsp;</td>
                     <td class="showborders" align="center"><%=intCompleteCount%></td>
                  </tr>
            <% If IsNumeric(objOverallRating(0)) Then %>      
                  <tr>
                     <td class="showborders"><a href="feedback.asp">Feedback Score</a>&nbsp;&nbsp;</td>
                     <td class="showborders" align="center"><%=(Round(objOverallRating(0),2)/5)*100%>%</td>                 
                  </tr>   
            <% End If %>
                </table>
            </td></tr>
            </table>
         </td></tr>
         <tr><td><hr /></td></tr>
   <% If Not objFeedback.EOF Then %>
      <% If bolEnbableFeedback Then %>
         <tr><td align="center">
            <table>
               <tr>

            <% If (bolFound And intFeedbackCount >= 7) And  bolEnbableFeedback Then %>
                  <td valign="top">
                     <table>
                        <tr><td>
                        <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>
                              <img width="180" src="../themes/<%=Application("Theme")%>/images/kiosk/<%=strImageName%>" />
                        <% Else %>
                              <img width="180" src="../themes/<%=objNameCheckSet(3)%>/images/kiosk/<%=strImageName%>" />
                        <% End If %>
                        </td></tr>
                     </table>
                  </td>
            <% End If %>
               <td>&nbsp;</td> 
               <td valign="top">
               <% If bolEnbableFeedback Then %>
                  <table>
                     <tr>
                        <th colspan="7">Recent Feedback</th>
                     </tr>
                  <% Do Until objFeedback.EOF %>
                     <tr>
                        <td class="showborders" align="center">&nbsp;&nbsp;<a href="modify.asp?ID=<%=objFeedback(0)%>" target="_blank"><%=objFeedback(0)%></a>&nbsp;&nbsp;</td>
                        <td class="showborders"><%=objFeedback(1)%></td>
                        <td class="showborders">
                        <% If objFeedback(2) >= 0 And objFeedback(2) < .25 Then 
                              strRating = "src=""" & strURLRoot & "0star.png"""
                           ElseIf objFeedback(2) >= .25 And objFeedback(2) < 1 Then 
                              strRating = "src=""" & strURLRoot & "0star-half.png"""
                           ElseIf objFeedback(2) >= 1 And objFeedback(2) < 1.25 Then
                              strRating = "src=""" & strURLRoot & "1star.png"""
                           ElseIf objFeedback(2) >= 1.25 And objFeedback(2) < 2 Then 
                              strRating = "src=""" & strURLRoot & "1star-half.png"""
                           ElseIf objFeedback(2) >= 2 And objFeedback(2) < 2.25 Then
                              strRating = "src=""" & strURLRoot & "2star.png"""   
                           ElseIf objFeedback(2) >= 2.25 And objFeedback(2) < 3 Then 
                              strRating = "src=""" & strURLRoot & "2star-half.png"""
                           ElseIf objFeedback(2) >= 3 And objFeedback(2) < 3.25 Then
                              strRating = "src=""" & strURLRoot & "3star.png"""   
                           ElseIf objFeedback(2) >= 3.25 And objFeedback(2) < 4 Then 
                              strRating = "src=""" & strURLRoot & "3star-half.png"""
                           ElseIf objFeedback(2) >= 4 And objFeedback(2) < 4.25 Then
                              strRating = "src=""" & strURLRoot & "4star.png"""   
                           ElseIf objFeedback(2) >= 4.25 And objFeedback(2) < 5 Then 
                              strRating = "src=""" & strURLRoot & "4star-half.png"""
                           ElseIf objFeedback(2) >= 5  Then
                              strRating = "src=""" & strURLRoot & "5star.png"""      
                           End If %>
                  <img <%=strRating%> />
                        </td>
                        <td class="showborders"><%=objFeedback(3)%>&nbsp;</td>
                        <td class="showborders"><%=objFeedback(4)%>&nbsp;</td>
                        <td class="showborders"><%=objFeedback(6)%>&nbsp;</td>
                        <td class="showborders"><%=objFeedback(5)%>&nbsp;</td>
                     </tr>
                  <%    objFeedback.MoveNext
                     Loop %>
                  </table>
               <% End If %>
               </td></tr>
            </table>
         </td></tr>
      <% End If %>   
   <% End If %>
         <tr><td align="center">Last Refreshed: <%=Date()%> - <%=Time()%></td></tr>
      </table>
   </center>
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
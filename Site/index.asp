<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/12/04
'Last Updated 6/16/14

'This is the main index page that the user goes to when entering a help desk ticket.  
'Before it displays anything to the user it will check some conditions.  This page is broken
'down into two subs, one is the Main sub which is the user form.  The other is the Submit
'sub which is what the user sees after they successfully submit the form.

Option Explicit

On Error Resume Next

Dim strUserTemp, strLocation, strProblem, strEMailTemp, bolValidEMail, strEMail
Dim strCustom1, strCustom2, objMessage, objConf, strSQL, Upload, strIP
Dim strUserAgent, bolAdminEnter, strMessageFont, strUser, objNetwork, bolShowLogout
Dim strTech, intID, strAttachment, objStats, intTotalTickets, intOpenTickets
Dim intClosedTickets, strAvgTicketTime, objAvgTicketTime, strDays, strHours, strMinutes
Dim strName, objRecordSet, strUserDisplayName, strCMD, strError, strIPAddress
Dim bolNotifications, objYourUpdateRequests, objTracking, objSubnets

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

Const cdoSendUsingPickup = 1

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

Set Upload = New FreeASPUpload
Upload.Save(Application("FileLocation"))

strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")

strCMD = Upload.Form("cmdSubmit")

Select Case strCMD
   Case "Mobile Site"
      Response.Cookies("SiteVersion") = "Mobile"
      Response.Cookies("SiteVersion").Expires = Date() + 14
      GetUser
   Case "Full Site"
      Response.Cookies("SiteVersion") = "Full"
      Response.Cookies("SiteVersion").Expires = Date() + 14
      GetUser
      
End Select

'Go to the correct site if the buttons are used on the mobile site
Select Case Upload.Form("Site")
   Case "Zoom In"
      Response.Cookies("ZoomLevel") = "ZoomIn"
      Response.Cookies("ZoomLevel").Expires = Date() + 14
   Case "Zoom Out"
      Response.Cookies("ZoomLevel") = "ZoomOut"
      Response.Cookies("ZoomLevel").Expires = Date() + 14
End Select

'Check and see if this page was already submitted.  If so then make sure all the
'fields were filled in, if everything is ok then submit the page
If (Upload.Form("cmdSubmit") = "Submit") Then

   strUserTemp = Upload.Form("UserName")
   strLocation = Upload.Form("Location")
   strProblem = Upload.Form("Problem")
   strEMailTemp = Upload.Form("EMail")
   
   If Application("UseCustom1") Then
      strCustom1 = Upload.Form("Custom1")
   Else
      strCustom1 = "6/16/78"
   End If
   
   If Application("UseCustom2") Then
      strCustom2 = Upload.Form("Custom2")
   Else
      strCustom2 = "6/16/78"
   End If
   
   bolValidEMail = IsEmailValid(strEMailTemp)
   
   'Set strMail to nothing if the box was empty, otherwise set the mail setting
   If Not bolValidEMail And strEMailTemp = "" Then
      strEMail = ""
   ElseIf bolValidEMail Then
      strEMail = strEMailTemp
   End If
   
   'Make sure all the fields were filled out, if not show the form again
   If strUserTemp = "" Or strLocation = "" Or strEMail = "" Or strProblem = "" Or strCustom1 = "" or strCustom2 = "" Then
      Call Main()
   Else
      Call Submit()
   End If
Else  
   Call Main()
End If

Sub Main()

   'This is the Main sub, the user will see this on their first visit to this page.  If
   'they submit the form and something on it is missing or incorrect this sub will run
   'again but will highlight the fields that had errors and display a message at the
   'top of the screen explaining what is wrong.   
   
   On Error Resume Next
   
   strSQL = "SELECT Subnet,Location FROM Subnets ORDER BY Subnet"
   Set objSubnets = Application("Connection").Execute(strSQL)   
   
   'Set the location based on the IP address
   If strLocation = "" Then
      strIP = Request.ServerVariables("REMOTE_ADDR")
      If Not objSubnets.EOF Then
         Do Until objSubnets.EOF
            If InSubnet(strIP,objSubnets(0)) Then
               strLocation = objSubnets(1)
            End If
            objSubnets.MoveNext
         Loop
      End If
   End If
   
   Set objNetwork = CreateObject("WSCRIPT.Network")   
   
   'Set the username to match what was entered in the email field.  This overrides the default
   'username setting allowing someone else to enter a ticket for someone.
   If strEMail <> "" Then
      strName = Left(strEmail,InStr(strEMail,"@")-1)
   End If

   'Get the users logon name
   If strName = "" Then
      strName = objNetwork.UserName
   End If
   
   'Check and see if a username was sent in the URL
   If Request.QueryString("UserName") <> "" Then
      strName = Request.QueryString("UserName")
      bolAdminEnter = True
      If strName = "<empty>" Then
         strName = ""
      End If
   End If
   
   'Check and see if anonymous access is enabled
   If LCase(Left(strName,4)) = "iusr" Then
      strName = GetUser
      bolShowLogout = True
   Else
      bolShowLogout = False
   End If

   If LCase(Left(objNetwork.UserName,4)) = "iusr" Then
      bolShowLogout = True
   End If
   
   'If we successfully got the username then generate the e-mail address
   If strName <> "" Then
      strEMailTemp = LCase(GetEMail(strName))
   Else 
      strName = strUserTemp
   End If

   'See if there are any system messages for the users.
   strSQL = "SELECT Message,Recipient,Type,PositionOnPage,Enabled" & vbCRLF
   strSQL = strSQL & "FROM Message" & vbCRLF
   strSQL = strSQL & "WHERE ID=1"

   Set objMessage = Application("Connection").Execute(strSQL)
   
   'If the current user is still waiting for an update get the info
   strSQL = "SELECT Main.Tech, Tracking.Ticket" & vbCRLF 
   strSQL= strSQL & "FROM Tracking INNER JOIN Main ON Tracking.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "WHERE (Tracking.Type='Request') AND (Tracking.TrackedBy='" & strUser & "')"
   Set objYourUpdateRequests = Application("Connection").Execute(strSQL)

   If Not objYourUpdateRequests.EOF Then
      bolNotifications = True
   End If

   'Check and see if the current user is tracking any tickets
   strSQL = "SELECT Main.Tech, Tracking.Ticket" & vbCRLF
   strSQL = strSQL & "FROM Tracking INNER JOIN Main ON Tracking.Ticket = Main.ID" & vbCRLF
   strSQL = strSQL & "WHERE (Tracking.Type='Track') And (Tracking.TrackedBy='" & strUser & "')"
   Set objTracking = Application("Connection").Execute(strSQL)

   If Not objTracking.EOF Then
      bolNotifications = True
   End If
   
   'Build the SQL string
   strSQL = "Select Location.Location,Location.Active" & vbCRLF
   strSQL = strSQL & "From Location" & vbCRLF
   strSQL = strSQL & "Order By Location.Location;"
   Set objRecordSet = Application("Connection").Execute(strSQL)
   
   'Get the users tickets from the database.
   strSQL = "SELECT Status FROM Main WHERE Name='" & strUser & "'"
   Set objStats = Application("Connection").Execute(strSQL)

   'Count the tickets
   intTotalTickets = 0
   intOpenTickets = 0
   intClosedTickets = 0
   Do Until objStats.EOF
      intTotalTickets = intTotalTickets + 1
      If objStats(0) <> "Complete" Then
         intOpenTickets = intOpenTickets + 1
      Else
         intClosedTickets = intClosedTickets + 1
      End If
      objStats.MoveNext
   Loop

   strSQL = "SELECT Avg(OpenTime) AS AvgOfOpenTime" & vbCRLF
   strSQL = strSQL & "FROM Main" & vbCRLF
   strSQL = strSQL & "Where OpenTime<>'' And Name='" & strUser & "'"
   Set objAvgTicketTime = Application("Connection").Execute(strSQL) 

   strDays = Int(objAvgTicketTime(0)/1440)
   strHours = Int((objAvgTicketTime(0)-strDays*1440)/60)
   strMinutes = (objAvgTicketTime(0)-(strDays*1440)-(strHours*60))
   strAvgTicketTime = Int(strDays) & "d " & Int(strHours) & "h " & Int(strMinutes) & "m" 
   If Trim(strAvgTicketTime) = "d h m" Then
      strAvgTicketTime = "N/A"
   End If

   If IsMobile Then
      MobileVersionMain
   Else
      FullVersionMain
   End If
End Sub
%>

<%
Sub FullVersionMain 

   Dim intNameBoxLength, intEMailBoxLength
   
   If InStr(strUserAgent,"Android") Then
      intNameBoxLength = 20
      intEMailBoxLength = 25
   ElseIf InStr(strUserAgent,"Windows NT") Then
      intNameBoxLength = 20
      intEMailBoxLength = 30
   Else
      intNameBoxLength = 25
      intEMailBoxLength = 35
   End If

%>

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

   <hr class="usertopbar"/>
      
   <% 'Check the global setting to see if the user has the rights to view old calls.
   If Application("UserCanViewCallStatus") = True Then %>
      <div class="usertopbar">
         <ul class="topbar">
            <li class="topbar">Home<font class="separator"> | </font></li>
            <li class="topbar"><a href="view.asp?filter=open">Open Tickets</a><font class="separator"> | </font></li>
            <li class="topbar"><a href="view.asp?filter=closed">Closed Tickets</a></li>
         <% If bolShowLogout Then %>
            <font class="separator"> | </font></li>
            <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
         <% Else %>
            </li>
         <% End If %>
         </ul>
      </div>
      <hr class="userbottombar"/>
   <%End If%>

   <div class="mainarea">
      <table>
         <tr><td valign="top">
         <table>
            <tr><td><img src="themes/<%=Application("Theme")%>/images/user.gif"/></td></tr>
         <% If Application("ShowUserStats") Then %>
            <tr><td>
               Your Help Desk Statistics <br />
               Total tickets = <%=intTotalTickets%> <br />
            <% If Application("UserCanViewCallStatus") Then %>   
                  Open tickets = <a href="view.asp?filter=open"><%=intOpenTickets%></a> <br />
            <% Else %>
                  Open tickets = <%=intOpenTickets%> <br />
            <% End If %>
            
            <% If Application("UserCanViewCallStatus") Then %>   
                  Closed tickets = <a href="view.asp?filter=closed"><%=intClosedTickets%></a> <br />
            <% Else %>
                  Closed tickets = <%=intClosedTickets%> <br />
            <% End If %>
            
               Avg Time = <%=strAvgTicketTime%> <br />
            
            </td></tr>
         <% End If %>
         <% If Application("RemoteSupportLink") <> "" Then %>
            <tr><td><hr /></td></tr>
            <tr>
               <td align="center">
                  <a href="<%=Application("RemoteSupportLink")%>">Remote Support</a>
                </td>
            </tr>
         <% End If %>
         
         <% If IsTablet Then %>
               <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="index.asp">
               <tr><td><hr /></td></tr>
               <tr><td align="center"><input type="submit" value="Mobile Site" name="cmdSubmit"></td></tr>
               </form>
         <% End If %>
         
         </table>
         </td>
         <td>&nbsp;</td>
         <td valign="top">
         <table width="100%">
            <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="index.asp">
            
      <%       If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Users") And (objMessage(3) = "Top" or objMessage(3) = "Both") Then 
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
   <%       End If %> 
         <%=Application("MainPageText")%><br/>
         
         <%'If the form was submitted and something was missing then display a message to the user
         If (Upload.Form("cmdSubmit") = "Submit") And (strUserTemp = "" Or strLocation = "" Or strEMailTemp = "" Or strProblem = "" Or strCustom1 = "" or strCustom2 = "") Then%>
            <font class="missing">Please fill out the highlighted fields...</font>
         <%End If%>

         <%'If the form is submitted and the email address was incorrect it will let the user know
         If (Upload.Form("cmdSubmit") = "Submit") And Not bolValidEMail And strEMailTemp <> "" Then %>
            <font class="missing">E-mail address is not valid...</font>
         <%End If %>
         
         <hr/>
            <tr>
               <td>
                  <%If (Upload.Form("cmdSubmit") = "Submit") And (strUserTemp = "") Then %>
                     <font class="missing">Name:</font>
                  <%Else%>
                     Name:
                  <%End If%>
                  <input type="text" size="<%=intNameBoxLength%>" name="Username" value="<%=GetFirstandLastName(strName)%>">
               </td>
               <td>
                  <%If (Upload.Form("cmdSubmit") = "Submit") And (strEMail = "") Then %>
                     <font class="missing">EMail:</font>
                  <%Else%>
                     EMail:
                  <%End If%>
                  <input type="text" size="<%=intEMailBoxLength%>" name="EMail" value="<%=strEMailTemp%>">
               </td>
            </tr>
            <tr>
               <td colspan="2"> 
                  <%If (Upload.Form("cmdSubmit") = "Submit") And (strLocation = "") Then %>
                     <font class="missing">Location:</font>
                  <%Else%>
                     Location:
                  <%End If%>
                  <select name="Location">
                     <option value="<%=strLocation%>"><%=strLocation%></option>
                     <%'Populate the Location pulldown list
                     Do Until objRecordSet.EOF
                        If Trim(Ucase(objRecordSet(0))) <> Trim(Ucase(strLocation)) Then
                           If objRecordSet(1) = True Then%>
                              <option value="<%=objRecordSet(0)%>"><%=objRecordSet(0)%></option>
                     <%    End If
                        End If
                        objRecordSet.MoveNext
                     Loop%>
                  </select>
               </td>
            </tr>
            <%If Application("UseCustom1") Then%>
               <tr>
                  <td colspan="2">
                     <%If (Upload.Form("cmdSubmit") = "Submit") And (strCustom1 = "") Then %>
                        <font class="missing"><%=Application("Custom1Text")%>:</font>
                     <%Else%>
                        <%=Application("Custom1Text")%>:
                     <%End If%>
                     <input type="text" size="30" name="Custom1" value="<%=strCustom1%>">
                  </td>
               </tr>
            <%End If%>
            <%If Application("UseCustom2") Then%>
               <tr>
                  <td colspan="2">
                     <%If (Upload.Form("cmdSubmit") = "Submit") And (strCustom2 = "") Then %>
                        <font class="missing"><%=Application("Custom2Text")%>:</font>
                        <input type="text" size="30" name="Custom2" value="<%=strCustom2%>">
                     <%Else%>
                        <%=Application("Custom2Text")%>:
                     <% If strCustom2 = "" Then 
                           If Application("Custom2Text") = "Phone" Then %>
                              <input type="text" size="30" name="Custom2" value="<%=GetPhoneNumber(strName)%>">
                        <% Else %>
                              <input type="text" size="30" name="Custom2" >
                        <% End If %>
                     <% Else%>
                           <input type="text" size="30" name="Custom2" value="<%=strCustom2%>">
                     <% End If%>
                     <%End If%>
                     
                  </td>
               </tr>
            <%End If%>
            <tr><td>&nbsp;</td></tr>

         <% If inStr(strUserAgent,"iPad") = False And inStr(strUserAgent,"iPhone") = False Then
               If Application("UseUpload") Then
                  If InStr(strUserAgent,"Chrome") or InStr(strUserAgent,"Safari") Then %>
                     <tr><td colspan="2">Upload a File: <input class="fileuploadchrome" type="file" name="Attachment" size="50"></td></tr>
                     <tr><td>&nbsp;</td></tr>
               <% Else%>
                     <tr><td colspan="2">Upload a File: <input class="fileupload" type="file" name="Attachment" size="50"></td></tr>
                     <tr><td>&nbsp;</td></tr>
               <% End If
               End If
            End If%>
            <tr>
               <td colspan="2">
                  <%If (Upload.Form("cmdSubmit") = "Submit") And (strProblem = "") Then %>
                     <font class="missing">Describe the problem:</font>
                  <%Else%>
                     Describe the problem:
                  <%End If%>
               </td>
            </tr>
            <tr>
               <td colspan="2"><textarea rows="6" cols="63" name="Problem"><%=strProblem%></textarea></td>
            </tr>
            <tr>
               <td colspan="2">
                  <input type="submit" value="Submit" name="cmdSubmit" style="float: right">

               </td>
            </tr>
            
<%          If bolNotifications Then %>
               <tr><td colspan="2"><hr /></td></tr>
               <tr>
                  <td colspan="2">Notifications:<br />
                  <ul>
<%                Do Until objYourUpdateRequests.EOF%>
                     <li>   
                        <%=GetFirstandLastName(objYourUpdateRequests(0))%> has not updated 
                        Ticket <%=objYourUpdateRequests(1)%> since your request.
                     </li>
<%                   objYourUpdateRequests.MoveNext
                  Loop 
                  Do Until objTracking.EOF%>
                     <li>
                     You are tracking Ticket <%=objTracking(1)%>
                     </li>
<%                   objTracking.MoveNext 
                  Loop %>                  
                  </ul>
                  </td>
               </tr>
<%          End If %>
            <tr><td colspan="2"><hr /></td></tr>
            
   <%       If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Users") And (objMessage(3) = "Bottom" or objMessage(3) = "Both") Then 
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
   <%       End If %> 
            
            </form>
         </table>
         </td></tr>
      </table>
   </div>


   </body>
   </html>
<%End Sub%>

<%Sub Submit()

   'If all fields in the Main sub's form are filled out properly then this sub will run.
   'This will add the data entered by the user to the database.  Then it will send an e-mail
   'to all the email addresses in the E-Mail section of the settings page in the Admins
   'section.  Then it sends a conformation email to the user.
   
   On Error Resume Next
   
   Dim strUserName, strLocation, strProblem, strEMail, strDate, strTime
   Dim objRegExp, strNewProblem, strNewUserName, strNewEMail, strNewLocation, strSQL
   Dim objRecordSet, objAdminEMail, intIndex, strTempEMail, strAdminEMail
   Dim strMessage, objMessage, objConf, objNetwork, strCustom1, strCustom2
   Dim strNewCustom1, strNewCustom2, strNewDisplayName, objSiteTech
   Dim objTechSet, strNewTech, strStatus, strNewStatus, objFSO, objFolder, objFile
   Dim intFileCount, objAdminCheck, strName
   Dim bolDBWorking, strDisplayName
   
   'Get the values from the form on the index page and assign them to variables
   strLocation = Upload.Form("Location")
   strProblem = Upload.Form("Problem")
   strEMail = Upload.Form("EMail")
   strDisplayName = Upload.Form("Username")
   strIPAddress = Request.ServerVariables("REMOTE_ADDR")
   strDate = Date
   strTime = Time
   
   If GetFirstandLastName(strUser) <> strDisplayName Then
      bolAdminEnter = True
   End If

   If Application("UseCustom1") Then
      strCustom1 = Upload.Form("Custom1")
   Else
      strCustom1 = ""
   End If
   
   If Application("UseCustom2") Then
      strCustom2 = Upload.Form("Custom2")
   Else
      strCustom2 = ""
   End If   
   
   'Get the users logon name from the email address then use that to get the first and
   'last name from Active Directory.
   If LCase(Application("EMailSuffix")) = LCase(Right(strEmail,Len(strEMail)-InStr(strEMail,"@")+1)) Then
      strUserName = Left(strEmail,InStr(strEMail,"@")-1)
      strDisplayName = GetFirstandLastName(strUserName)
    
      'If the Display name changed to the first part of the users email change it back to what was submitted.
      If strUserName = strDisplayName Then
         strDisplayName = Upload.Form("Username")
      End If
   End If
   
   'Reset the display name if the help email address is used.
   If strEMail = LCase("help@wswheboces.org") Then
      strDisplayName = Upload.Form("Username")
   End If
 
   'Create the Regular Espression object and set it's properties.
   Set objRegExp = New RegExp
   objRegExp.Pattern = "'"
   objRegExp.Global = True
   
   'Assign the ticket to the tech linked to the location
   strSQL = "Select Tech From Location Where Location='" & strLocation & "'"
   Set objSiteTech = Application("Connection").Execute(strSQL)
   
   If objSiteTech(0) <> "" Then
      strTech = objSiteTech(0)
      strStatus = "Auto Assigned"
   Else
      strStatus = "New Assignment"
   End If
   
   If bolAdminEnter Then
      strProblem = strProblem & vbCRLF & vbCRLF & "Ticket entered by " & GetFirstandLastName(strUser) & "."
   Else
      'strProblem = strProblem & vbCRLF & vbCRLF & "IP Address = " & strIPAddres
   End If
   
   'Use the regular expression to change a ' to a '' so the SQL Insert command will work.
   'The value will be assigned to a new variable so the old one can still be used in the
   'emails.
   strNewProblem = objRegExp.Replace(strProblem,"''")
   strNewUserName = objRegExp.Replace(strUserName,"''")
   strNewDisplayName = objRegExp.Replace(strDisplayName,"''")
   strNewEMail = objRegExp.Replace(strEMail,"''")
   strNewLocation = objRegExp.Replace(strLocation,"''")
   strNewCustom1 = objRegExp.Replace(strCustom1,"''")
   strNewCustom2 = objRegExp.Replace(strCustom2,"''") 
   strNewTech = objRegExp.Replace(strTech,"''")
   strNewStatus = objRegExp.Replace(strStatus,"''")
   
   'Build the SQL string that will add the data to the database
   strSQL = "Insert Into Main (Name,DisplayName,Email,Location,Problem,SubmitDate,SubmitTime,Category,Status,Tech,LastUpdatedDate,Custom1,Custom2,TicketViewed) " & _
   "values ('" & strNewUserName & "','" & strNewDisplayName & "','" & strNewEMail & "','" & strNewLocation & "','" & strNewProblem & "','" & _
   strDate & "','" & strTime & "',' ','" & strNewStatus & "','" & strNewTech & "','6/16/78','" & strNewCustom1 & "','" & strNewCustom2 & "',False)"
   
   Err.Clear
   bolDBWorking = True
   
   'Execute the SQL string
   Application("Connection").Execute(strSQL)
   
   If Err Then
      strError = strError & "There was an error connection to the database. <br/>" 
      bolDBWorking = False
      Err.Clear
   End If
   
   If bolDBWorking Then
      'Build a new SQL string, this will be used to get the ID of the recently added ticket
      strSQL = "SELECT Main.ID, Main.SubmitDate, Main.SubmitTime, Main.Name" & vbCRLF
      strSQL = strSQL & "FROM Main" & vbCRLF
      strSQL = strSQL & "WHERE (((Main.SubmitDate)=#" & strDate& "#) AND ((Main.SubmitTime)=#" & strTime & "#) AND ((Main.Name)=""" & strUserName & """));"
      
      'Execute the SQL string and get the ticket number of the new ticket
      Set objRecordSet = Application("Connection").Execute(strSQL)
      intID = objRecordSet("ID")
      
      'Send a reminder about using the new system if the ticket was entered by someone else.
      If Application("SendReminder") And bolAdminEnter Then
         SendReminder
      End If

      'Update the log
      If strStatus = "Auto Assigned" Then
         strSQL = "INSERT INTO Log (Ticket,Type,NewValue,UpdateDate,UpdateTime)"
         strSQL = strSQL & "VALUES (" & intID & ",'Auto Assigned','" & strTech & "','" & Date() & "','" & Time() & "');"
         Application("Connection").Execute(strSQL)   
      End If      
      If strStatus = "New Assignment" Then
         strSQL = "INSERT INTO Log (Ticket,Type,NewValue,UpdateDate,UpdateTime)"
         strSQL = strSQL & "VALUES (" & intID & ",'New Ticket','" & strTech & "','" & Date() & "','" & Time() & "');"
         Application("Connection").Execute(strSQL)   
      End If
      
      Err.Clear
      
      If Application("UseUpload") Then
      
         'If there is an attachment save it to a folder on the server.
         intFileCount = 0
         Set objFSO = CreateObject("Scripting.FileSystemObject")
         objFSO.CreateFolder(Application("FileLocation") & "\" & intID)
         Upload.Save(Application("FileLocation") & "\" & intID)
         Set objFolder = objFSO.GetFolder(Application("FileLocation") & "\" & intID)
         For Each objFile in objFolder.Files
            intFileCount = intFileCount + 1
            strAttachment = objFile.Path
         Next
         If intFileCount = 0 Then
            objFSO.DeleteFolder Application("FileLocation") & "\" & intID
         End If
         Set objFSO = Nothing
         
         If Err Then
            strError = strError & "There is a problem with the uploads folder. <br/>"
            Err.Clear
         End If
      End If
      
      '*****************************************************************************************
      'Send the tech an email if the ticket is automatically assigned.
      If strTech <> "" Then
         EMailTech
      End If
      
      '*****************************************************************************************
      'Send an email to all the admins.
      EMailAdmins

      '*****************************************************************************************   
      'Had to kill the objMessage to wipe out the attachment.
      Set objMessage = Nothing
      
      EMailUser
      '*****************************************************************************************
      
      If Err Then
         strError = strError & "Unable to send email. <br />"
         Err.Clear
      End If
      
      If bolAdminEnter Then
      
         strSQL = "Select Username, UserLevel, Active" & vbCRLF
         strSQL = strSQL & "From Tech" & vbCRLF
         strSQL = strSQL & "WHERE UserName='" & strUser & "'"
         Set objAdminCheck = Application("Connection").Execute(strSQL)
      
         If Not objAdminCheck.EOF Then
            If objAdminCheck(2) Then
               Response.Redirect("admin/modify.asp?ID=" & intID)
            End If
         End If
      End If
      
      'Close open object
      Set objMessage = Nothing
      
   End If 'End to If bolDBWorking statement
   
   If IsMobile Then
      MobileVersionSubmit
   Else
      FullVersionSubmit
   End If
   
End Sub
%>

<%
Sub FullVersionSubmit %>

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

   <hr class="usertopbar"/>

<% 'Check the global setting to see if the user has the rights to view old calls.
If Application("UserCanViewCallStatus") = True Then %>
   <div class="usertopbar">
      <ul class="topbar">
         <li class="topbar">Home<font class="separator"> | </font></li>
         <li class="topbar"><a href="view.asp?filter=open">Open Tickets</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="view.asp?filter=closed">Closed Tickets</a></li>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
      <% Else %>
         </li>
      <% End If %>
      </ul>
   </div>
   <hr class="userbottombar"/>
<%End If%>


<div class="mainarea">

	<img class="mainareaimage" src="themes/<%=Application("Theme")%>/images/submit.gif"/>
	<div class="mainareatext">
   <% If strError = "" Then %>
         Your request has been submitted. Your ticket number is <%=intID%>. <br/>
         A confirmation message has been sent to <%=strEMail%> <br/>
         You will be contacted shortly. <br/> <br/>
   <% Else %>
         <%=strError%><br/> <br/>
   <% End If %>
      <a href="<%=Left(Application("AdminURL"),(Len(Application("AdminURL"))-5))%>">Return to Previous Page</a>
   </div>
</div>

</body>
</html>
<%End Sub%>

<%
Sub MobileVersionSubmit %>
   
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
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %> 

   </head>

   <body>  
      <center><b><%=Application("SchoolName")%> Help Desk</b></center>
      <center>
      <table align="center">
         <tr><td><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               <div align="center">
                  <input type="submit" value="Open Tickets" name="filter"> 
                  <input type="submit" value="Closed Tickets" name="filter">
            
            <% If bolShowLogout Then %>   
                  <input type="submit" value="Log Out" name="Log Out">
            <% End If %> 
            
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table>
      <table><tr><td width="<%=Application("MobileSiteWidth")%>">
      <table align="center">
         <tr>
            <td>
               Ticket submitted, your ticket number is <%=intID%>. <br /><br />
            
               <a href="<%=Left(Application("AdminURL"),(Len(Application("AdminURL"))-5))%>">Return to Previous Page</a>
            </td>
         </tr>
   </body>
   </html>
<%End Sub%>

<%
Sub MobileVersionMain 
%>

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
      <meta name="viewport" content="width=100,user-scalable=no,initial-scale=<%=intZoom%>" />
   <% ElseIf InStr(strUserAgent,"Windows Phone") Then %>
      <meta name="viewport" content="width=375,user-scalable=no" /> 
   <% Else %>
      <meta name="viewport" content="width=device-width,user-scalable=no" /> 
   <% End If %> 

   </head>
   
   <body>  
      <center><b><%=Application("SchoolName")%> Help Desk</b></center>
      <center>
      <table align="center">
         <tr><td><hr /></td></tr>               
         <form method="Post" action="view.asp">
         <tr>
            <td width="<%=Application("MobileSiteWidth")%>">
               <div align="center">
                  <input type="submit" value="Open Tickets" name="filter"> 
                  <input type="submit" value="Closed Tickets" name="filter">
            
            <% If bolShowLogout Then %>   
                  <input type="submit" value="Log Out" name="Log Out">
            <% End If %> 
            
               </div>
            </td>
         </tr>
         </form>
         <tr><td><hr /></td></tr>
      </table>
      <table><tr><td width="<%=Application("MobileSiteWidth")%>">
      <table align="center">
<%    If objMessage(4) And (objMessage(1) = "Both" Or objMessage(1) = "Users") And (objMessage(3) = "Top" or objMessage(3) = "Both") Then 
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
      <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="index.asp">
         <tr>
            <td colspan="2" width="<%=Application("MobileSiteWidth")%>">
               <%=Application("MainPageText")%><br/>
               
               <%'If the form was submitted and something was missing then display a message to the user
               If (Upload.Form("cmdSubmit") = "Submit") And (strUserTemp = "" Or strLocation = "" Or strEMailTemp = "" Or strProblem = "" Or strCustom1 = "" or strCustom2 = "") Then%>
                  <font class="missing">Please fill out the highlighted fields...</font>
               <%End If%>

               <%'If the form is submitted and the email address was incorrect it will let the user know
               If (Upload.Form("cmdSubmit") = "Submit") And Not bolValidEMail And strEMailTemp <> "" Then %>
                  <font class="missing">E-mail address is not valid...</font>
               <%End If %>
            </td>
         </tr>
         <tr><td colspan="2"><hr /></td></tr>
         <tr>
            <td>
               <%If (Upload.Form("cmdSubmit") = "Submit") And (strUserTemp = "") Then %>
                  <font class="missing">Name:</font>
               <%Else%>
                  Name:
               <%End If%>
            </td>
            <td>
               <input type="text" style="width: 95%;" name="Username" value="<%=GetFirstandLastName(strName)%>">
            </td>
         </tr>
         <tr>
            <td>
               <%If (Upload.Form("cmdSubmit") = "Submit") And (strEMail = "") Then %>
                  <font class="missing">EMail:</font>
               <%Else%>
                  EMail:
               <%End If%>
            </td>
            <td>
               <input type="text" style="width: 95%;" name="EMail" value="<%=strEMailTemp%>">
            </td>
         </tr>
         
         <tr>
            <td>
               <%If (Upload.Form("cmdSubmit") = "Submit") And (strLocation = "") Then %>
                  <font class="missing">Location:</font>
               <%Else%>
                  Location:
               <%End If%>

            </td>
            <td>
               <select name="Location">
                  <option value="<%=strLocation%>"><%=strLocation%></option>
                  <%'Populate the Location pulldown list
                  Do Until objRecordSet.EOF
                     If Trim(Ucase(objRecordSet(0))) <> Trim(Ucase(strLocation)) Then
                        If objRecordSet(1) = True Then%>
                           <option value="<%=objRecordSet(0)%>"><%=objRecordSet(0)%></option>
                  <%    End If
                     End If
                     objRecordSet.MoveNext
                  Loop%>

               </select>
            </td>
         </tr>
      <%If Application("UseCustom1") Then%>
         <tr>
            <td>
               <%If (Upload.Form("cmdSubmit") = "Submit") And (strCustom1 = "") Then %>
                  <font class="missing"><%=Application("Custom1Text")%>:</font>
               <%Else%>
                  <%=Application("Custom1Text")%>:
               <%End If%>
            </td>
            <td>
               <input type="text" style="width: 95%;" name="Custom1" value="<%=strCustom1%>">
            </td>
         </tr>
      <%End If%>
            <%If Application("UseCustom2") Then%>
               <tr>
                  <td>
                  <% If (Upload.Form("cmdSubmit") = "Submit") And (strCustom2 = "") Then %>
                        <font class="missing"><%=Application("Custom2Text")%>:</font>
                        </td><td>
                        <input type="text" style="width: 95%;" name="Custom2" value="<%=strCustom2%>">
                  <% Else%>
                        <%=Application("Custom2Text")%>:
                        </td><td>
                     <% If strCustom2 = "" Then 
                           If Application("Custom2Text") = "Phone" Then %>
                              <input type="text" style="width: 95%;" name="Custom2" value="<%=GetPhoneNumber(strName)%>">
                        <% Else %>
                              <input type="text" style="width: 95%;" name="Custom2" >
                        <% End If %>
                     <% Else%>
                           <input type="text" style="width: 95%;" name="Custom2" value="<%=strCustom2%>">
                     <% End If%>
                  <% End If%>
                     
                  </td>
               </tr>
            <%End If%>
            
            <tr>
               <td colspan="2">
                  <%If (Upload.Form("cmdSubmit") = "Submit") And (strProblem = "") Then %>
                     <font class="missing">Describe the problem:</font>
                  <%Else%>
                     Describe the problem:
                  <%End If%>
               </td>
            </tr>
            <tr>
               <td colspan="2"><textarea name="Problem" rows="8" cols="90" style="width: 98%;"><%=strProblem%></textarea></td>
            </tr>
            <tr>
               <td><%=Application("Version")%></td>
               <td>
                  <input type="submit" value="Submit" name="cmdSubmit" style="float: right">
               </td>
            </tr>
            
         <tr><td colspan="2" width="<%=Application("MobileSiteWidth")%>"><hr /></td></tr>
         </form>
      </table>  
      </table>
      <table align="center">
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
      <form enctype="multipart/form-data" method="POST" accept-charset="utf-8" action="index.asp">
         <tr>
            <td colspan="2">
               
            </td>
         </tr>   
         <% If IsTablet Then %>
               <tr><td><input type="submit" value="Full Site" name="cmdSubmit">
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

<%
Sub EMailTech

   Dim objMessage, objConf, strSQL, strMessage, objTechSet, strLocation, strProblem
   Dim strEMail, strDisplayName, strCustom1, strCustom2, objMessageText
   Dim strCurrentUser, strStatus, strNotes, strSubject
   
   'Get the values from the form on the index page and assign them to variables
   strLocation = Upload.Form("Location")
   strProblem = Upload.Form("Problem")
   strEMail = Upload.Form("EMail")
   strDisplayName = Upload.Form("Username")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2") 
   
   If bolAdminEnter Then
      strProblem = strProblem & vbCRLF & vbCRLF & "Ticket entered by " & GetFirstandLastName(strUser) & "."
   Else
      'strProblem = strProblem & vbCRLF & vbCRLF & "IP Address = " & strIPAddres
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
   
   'Get the tech's email address
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
   Set objTechSet = Application("Connection").Execute(strSQL)
      
   'Create the body of the e-mail to the Tech
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='New Ticket Assigned'"
   Set objMessageText = Application("Connection").Execute(strSQL)
   
   strMessage = objMessageText(1)
   strMessage = Replace(strMessage,"#TICKET#",intID)
   strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
   strMessage = Replace(strMessage,"#USER#",strDisplayName)
   strMessage = Replace(strMessage,"#TECH#",strTech)
   strMessage = Replace(strMessage,"#STATUS#",strStatus)
   strMessage = Replace(strMessage,"#USEREMAIL#",strEMail)
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
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With

   Set objMessage = Nothing
   Set objConf = Nothing
End Sub
%>

<%
Sub EMailAdmins

   Dim strLocation, strProblem, strEMail, strDisplayName, strCustom1, strCustom2, arrAdminEMail
   Dim intIndex, strTempEMail, objMessage, objConf, strAdminEMail, strMessage, objMessageText
   Dim strCurrentUser, strStatus, strNotes, strSubject, objTechSet, strTechEMail
   
   'Get the values from the form on the index page and assign them to variables
   strLocation = Upload.Form("Location")
   strProblem = Upload.Form("Problem")
   strEMail = Upload.Form("EMail")
   strDisplayName = Upload.Form("Username")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2") 

   'Get the tech's email address, this will be used to prevent a tech who is also in the admin
   'list from getting two emails when a new ticket is auto assigned to them.
   strSQL = "Select EMail" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Where Tech='" & strTech & "'"
   Set objTechSet = Application("Connection").Execute(strSQL) 

   'Set the tech's email address
   If objTechSet.EOF Then
      strTechEMail = ""
   Else
      strTechEMail = objTechSet(0)
   End If
   
   'If someone else is entering a ticket for a user append to the problem who is entering the ticket
   If bolAdminEnter Then
      strProblem = strProblem & vbCRLF & vbCRLF & "Ticket entered by " & GetFirstandLastName(strUser) & "."
   Else
      'strProblem = strProblem & vbCRLF & vbCRLF & "IP Address = " & strIPAddres
   End If
   
   'Add each email address from the semicolon separated variable to an array
   arrAdminEMail = Split(Application("AdminEMail"),";")
   
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
  
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   'Send an email to each of the help desk administrators
   For Each strAdminEMail in arrAdminEMail
      
      If strAdminEMail <> "" Then
         If InStr(LCase(strTechEMail),LCase(strAdminEMail)) = 0 Then
      
            'Create the body of the e-mail to the administrator
            strSQL = "SELECT Subject, Message FROM EMail WHERE Title='New Ticket Admin'"
            Set objMessageText = Application("Connection").Execute(strSQL)
            
            strMessage = objMessageText(1)
            strMessage = Replace(strMessage,"#TICKET#",intID)
            strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
            strMessage = Replace(strMessage,"#USER#",strDisplayName)
            strMessage = Replace(strMessage,"#TECH#",strTech)
            strMessage = Replace(strMessage,"#STATUS#",strStatus)
            strMessage = Replace(strMessage,"#USEREMAIL#",strEMail)
            strMessage = Replace(strMessage,"#LOCATION#",strLocation)
            strMessage = Replace(strMessage,"#CUSTOM1#",strCustom1)
            strMessage = Replace(strMessage,"#CUSTOM2#",strCustom2)
            strMessage = Replace(strMessage,"#PROBLEM#",strProblem)
            strMessage = Replace(strMessage,"#NOTES#",strNotes)
            strMessage = Replace(strMessage,"#LINK#",Application("AdminURL") & "/modify.asp?ID=" & intID)
            
            strSubject = objMessageText(0)
            strSubject = Replace(strSubject,"#TICKET#",intID)
            
            'Send the email to the administrators
            With objMessage
               .To = strAdminEMail
               .Subject = strSubject
               .TextBody = strMessage
               .From = Application("SendFromEMail")
               If strAttachment <> "" Then
                  .AddAttachment strAttachment
               End If         
               If Application("BCC") <> "" Then
                  .BCC = Application("BCC")
               End If         
               .Send
            End With  
         End If
      End If
   Next
   
   Set objMessage = Nothing
   Set objConf = Nothing

End Sub
%>

<%
Sub EMailUser

   Dim strLocation, strProblem, strEMail, strDisplayName, strCustom1, strCustom2
   Dim objMessage, objConf, strMessage, objMessageText
   Dim strCurrentUser, strStatus, strNotes, strSubject

   'Get the values from the form on the index page and assign them to variables
   strLocation = Upload.Form("Location")
   strProblem = Upload.Form("Problem")
   strEMail = Upload.Form("EMail")
   strDisplayName = Upload.Form("Username")
   strCustom1 = Upload.Form("Custom1")
   strCustom2 = Upload.Form("Custom2")
   
   If bolAdminEnter Then
      strProblem = strProblem & vbCRLF & vbCRLF & "Ticket entered by " & GetFirstandLastName(strUser) & "."
   Else
      'strProblem = strProblem & vbCRLF & vbCRLF & "IP Address = " & strIPAddres
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
   
   'Create the body of the e-mail to the Tech
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='New Ticket User'"
   Set objMessageText = Application("Connection").Execute(strSQL)
   
   strMessage = objMessageText(1)
   strMessage = Replace(strMessage,"#TICKET#",intID)
   strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
   strMessage = Replace(strMessage,"#USER#",strDisplayName)
   strMessage = Replace(strMessage,"#TECH#",strTech)
   strMessage = Replace(strMessage,"#STATUS#",strStatus)
   strMessage = Replace(strMessage,"#USEREMAIL#",strEMail)
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
      .Subject = strSubject
      .TextBody = strMessage
      .From = Application("SendFromEMail")
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With

End Sub
%>

<%
Sub SendReminder

   On Error Resume Next

   Const cdoSendUsingPickup = 1

   Dim strMessage, objMessage, objConf, strURL, strDirections

   'Had to kill the objMessage to wipe out the attachment.
   Set objMessage = Nothing
   
   'Create the message object.
   Set objMessage = CreateObject("CDO.Message")
   
   'Create the configuration object.
   Set objConf = objMessage.Configuration
  
   With objConf.Fields
      .item("http://schemas.microsoft.com/cdo/configuration/sendusing") = cdoSendUsingPickup
      .item("http://schemas.microsoft.com/cdo/configuration/smtpserverpickupdirectory") = Application("SMTPPickupFolder")
      .Update
   End With
   
   'Create the body of the e-mail to the Tech
   strSQL = "SELECT Subject, Message FROM EMail WHERE Title='Reminder'"
   Set objMessageText = Application("Connection").Execute(strSQL)
   
   strMessage = objMessageText(1)
   strMessage = Replace(strMessage,"#TICKET#",intID)
   strMessage = Replace(strMessage,"#CURRENTUSER#",strCurrentUser)
   strMessage = Replace(strMessage,"#USER#",strDisplayName)
   strMessage = Replace(strMessage,"#TECH#",strTech)
   strMessage = Replace(strMessage,"#STATUS#",strStatus)
   strMessage = Replace(strMessage,"#USEREMAIL#",strEMail)
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
      .Subject = strSubject
      .TextBody = strMessage
      .From = Application("SendFromEMail")
      If Application("BCC") <> "" Then
         .BCC = Application("BCC")
      End If
      .Send
   End With
End Sub %>

<%
Function ToBinary(intNumber)

   'Convert the number to binary
   Do 
      ToBinary = intNumber Mod 2 & ToBinary
      intNumber = intNumber \ 2
   Loop Until intNumber = 0
   
   'Add the zero's to make the result 8 bits
   If Len(ToBinary) < 8 Then
      Do 
         ToBinary = "0" & ToBinary
      Loop Until Len(ToBinary) = 8
   End If

End Function %>

<%
Function IPtoBinary(strIP)
   
   Dim arrOctets, intOctet
   
   'Split the IP address into each octet
   arrOctets = Split(strIP,".")
   For Each intOctet in arrOctets
      IPtoBinary = IPtoBinary & ToBinary(intOctet)
   Next

End Function %>

<%
Function InSubnet(strIP,strSubnet)

   Dim arrSubnet, strNetworkID, intNetworkBits

   'Split the subnet into the network ID and the number of network bits
   arrSubnet = Split(strSubnet,"/")
   strNetworkID = arrSubnet(0)
   intNetworkBits = arrSubnet(1)
   
   'Compare the network bits from the IP to the network bits from the network ID
   If Left(IPtoBinary(strNetworkID),intNetworkBits) = Left(IPtoBinary(strIP),intNetworkBits) Then
      InSubnet = True
   Else
      InSubnet = False
   End If

End Function %>

<%
Function GetFirstandLastName(strUserName)

   On Error Resume Next

   Dim objConnection, objCommand, objRootDSE, objRecordSet,strDNSDomain

   If Application("UseAD") Then
   
      If strUserName <> "" Then
   
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

         If Not objRecordSet.EOF Then
            If objRecordSet(0) = "" Then
               GetFirstandLastName = strUserName
            Else
               GetFirstandLastName = objRecordSet(0) & " " & objRecordSet(1)
            End If
         Else
            GetFirstandLastName = strUserName
         End If
      
      Else
         GetFirstandLastName = strUserName
      End If

   Else
      GetFirstandLastName= strUserName
   End If

End Function
%>

<%
Function GetEMail(strUserName)

   On Error Resume Next

   Dim objConnection, objCommand, objRootDSE, objRecordSet,strDNSDomain

   If Application("UseAD") Then
      'Create objects requred to connect to AD
      Set objConnection = CreateObject("ADODB.Connection")
      Set objCommand = CreateObject("ADODB.Command")
      Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

      'Create a connection to AD
      objConnection.Provider = "ADSDSOObject"

      objConnection.Open "Active Directory Provider", Application("ADUsername"), Application("ADPassword")
      objCommand.ActiveConnection = objConnection
      strDNSDomain = objRootDSE.Get("DefaultNamingContext")
      objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); mail ;subtree"

      'Initiate the LDAP query and return results to a RecordSet object.
      Set objRecordSet = objCommand.Execute

      If NOT objRecordSet.EOF Then
         If objRecordSet(0) = "" Then
            GetEMail = ""
         Else
            GetEMail = objRecordSet(0)
         End If
      Else
         GetEMail= strUserName & Application("EMailSuffix")
      End If

   Else
      GetEMail= strUserName & Application("EMailSuffix")
   End If

End Function
%>

<%
Function GetPhoneNumber(strUserName)

   On Error Resume Next

   Dim objConnection, objCommand, objRootDSE, objRecordSet,strDNSDomain

   If Application("UseAD") Then
      
      If Not IsNumeric(Left(strUserName,2)) Then
   
         'Create objects required to connect to AD
         Set objConnection = CreateObject("ADODB.Connection")
         Set objCommand = CreateObject("ADODB.Command")
         Set objRootDSE = GetObject("LDAP://" & Application("Domain") & "/rootDSE")

         'Create a connection to AD
         objConnection.Provider = "ADSDSOObject"

         objConnection.Open "Active Directory Provider", Application("ADUsername"), Application("ADPassword")
         objCommand.ActiveConnection = objConnection
         strDNSDomain = objRootDSE.Get("DefaultNamingContext")
         objCommand.CommandText = "<LDAP://" & Application("DomainController") & "/" & strDNSDomain & ">;(&(objectCategory=person)(objectClass=user)(samaccountname=" & strUserName & ")); TelephoneNumber,distinguishedname,name ;subtree"

         'Initiate the LDAP query and return results to a RecordSet object.
         Set objRecordSet = objCommand.Execute

         If NOT objRecordSet.EOF Then
            If objRecordSet(0) = "" Then
               GetPhoneNumber = ""
            Else
               GetPhoneNumber = objRecordSet(0)
            End If
         Else
            GetPhoneNumber= ""
         End If

      Else
         GetPhoneNumber = "N/A"
      End If
         
   Else
      GetPhoneNumber= ""
   End If

End Function
%>

<%
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
        ' only following "abcdefghijklmnopqrstuvwxyz_-.'"
        ' characters and the ten digits are allowed
        For i = 1 To Len(strItem)
               c = LCase(Mid(strItem, i, 1))
               ' if there is an illegal character in the part
               If InStr("abcdefghijklmnopqrstuvwxyz_-.'", c) <= 0 And Not IsNumeric(c) Then
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

<%
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

		'''On Error resume next
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
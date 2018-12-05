<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 10/23/04
'Last Updated 6/16/14

'This page allows you to add categories, locations and techs to the database.

Option Explicit

On Error Resume Next

Dim strCategory, strLocation, strTech, strMessage, objNetwork, strUser, strSQL
Dim objNameCheckSet, objModifyTech, strModify, intID, strModifyTech, strUserAgent
Dim strRole, bolUseAD, bolShowLogout, strSubnet, objSubnets, intSubnetCount

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

'Build the SQL string, this will check the userlevel of the user.
strSQL = "Select Username, UserLevel, Active, Theme" & vbCRLF
strSQL = strSQL & "From Tech" & vbCRLF
strSQL = strSQL & "WHERE (((Tech.UserName)='" & strUser & "'));"
Set objNameCheckSet = Application("Connection").Execute(strSQL)
strRole = objNameCheckSet(1)

'See if the user has the rights to visit this page
If objNameCheckSet(1) = "Administrator" AND objNameCheckSet(2) Then

   'An error would be generated if the user has NTFS rights to get in but is not found
   'in the database.  In this case the user is denied access.
   If Err Then
      Err.Clear
      Call AccessDenied
   Else
      Call AccessGranted
   End If
Else
   Call AccessDenied
End If

Sub AccessGranted

   'The user has a userlevel of administrator and is allowed to access this site.

   'Get the information from the forms and assign them to variables
   strCategory = Request.Form("Category")
   strLocation = Request.Form("Location")
   strTech = Request.Form("Tech")
   strModify = Request.Form("cmdModifyTech")
   intID = Request.Form("ID")
   strModifyTech = Request.Form("ModifyTech")
   strSubnet = Request.Form("Subnet")

   strUserAgent = Request.ServerVariables("HTTP_USER_AGENT")
  
   'Figure out what needs to be done then call the main page.  A message will show
   'at the botom of the page that lets you know it's status.
   If (Request("cmdCategory") = "Add") Then
      strMessage = AddItem(strCategory,"category")
      Call Main
   ElseIf (Request("cmdLocation") = "Add") Then
      strMessage = AddItem(strLocation,"location")
      Call Main
   ElseIf (Request("cmdTech") = "Add") Then
      strMessage = AddItem(strTech,"tech")
      Call Main
   ElseIf (Request("cmdTech") = "Use AD") Then
      bolUseAD = True
      strMessage = AddItem(strTech,"tech")
      Call Main
   ElseIf (Request("cmdDisableCategory") = "Disable") Then
      strMessage = DisableItem(strCategory,"category")
      Call Main
   ElseIf (Request("cmdDisableCategory") = "Delete") Then
      strMessage = DeleteItem(strCategory, "category")
      Call Main
   ElseIf (Request("cmdDisableLocation") = "Disable") Then
      strMessage = DisableItem(strLocation,"location")
      Call Main
   ElseIf (Request("cmdDisableLocation") = "Delete") Then
      strMessage = DeleteItem(strLocation,"location")
      Call Main
   ElseIf (Request("cmdDisableTech") = "Disable") Then
      strMessage = DisableItem(strTech,"tech")
      Call Main
   ElseIf (Request("cmdDisableTech") = "Delete") Then
      strMessage = DeleteItem(strTech,"tech")
      Call Main
   ElseIf (Request("cmdEnableCategory") = "Enable") Then
      strMessage = EnableItem(strCategory,"category")
      Call Main
   ElseIf (Request("cmdEnableLocation") = "Enable") Then
      strMessage = EnableItem(strLocation,"location")
      Call Main
   ElseIf (Request("cmdEnableTech") = "Enable") Then
      strMessage = EnableItem(strTech,"tech")
      Call Main
   ElseIf (strModify = "Select") Then
      If strModifyTech = "" Then
         strMessage = "You must select a user."
         strModify = ""
         Call Main
      Else
         strSQL = "SELECT Tech,Username,EMail,UserLevel,ID,TaskListRole,DocumentationRole FROM Tech WHERE Tech='" & strModifyTech & "'"
         Set objModifyTech = Application("Connection").Execute(strSQL)      
         Call Main
      End If
   ElseIf (strModify = "Modify") Then
      strMessage = ModifyUser
      Call Main
   ElseIf (strModify = "Use AD") Then
      bolUseAD = True
      strMessage = ModifyUser  
      strModify = "Modify" 
      Call Main      
   ElseIf Request("cmdAssignTech") = "Assign" Then
      strMessage = AssignTechToLocation(strLocation,strTech)
      Call Main
   ElseIf Request("cmdAssignSubnet") = "Assign" Then
      strMessage = AssignSubnetToLocation(strSubnet,strLocation)
      Call Main
   ElseIf Request("cmdDeleteSubnet") = "Delete" Then
      strMessage = DeleteSubnet(strSubnet)
      Call Main
   Else
      Call Main
   End If
End Sub

Function AddItem(strItem,strItemType)

   'This function will make sure the value doesn't already exist in the database.  If
   'not it will add it to the database.  It returns a string that contains the status 
   'of the function.

   Dim objRegExp, strSQL, objRecordSet, strItemTemp, strEMail, strEMailTemp
   Dim strUserLevel, strUserLevelTemp, strUsername, strUsernameTemp, strTaskListRole
   Dim strTaskListRoleTemp, strDocRole, strDocRoleTemp
   
   'If they are adding a tech then this will get the email address
   strEMail = Request.Form("EMail")
   strUserLevel = Request.Form("UserLevel")
   strUsername = Request.Form("Username")
   strTaskListRole = Request.Form("TaskListRole")
   strDocRole = Request.Form("DocRole")
   
   If bolUseAD Then
      strItem = GetFirstandLastName(strUsername)
      If strItem = "" Then
         AddItem = "User not found"
         Exit Function
      End If
      
      strEMail = LCase(GetEMail(strUserName))
      If strEMail = "" Or IsNull(strEMail) Then
         AddItem = "No email address found for user"
         Exit Function
      End If
   End If
   
   If strEMail <> "" Then
      strEMail = ",'" & strEMail & "'"
      strEMailTemp = ",EMail"
      If strUserLevel = "" Then
         strUserLevel = "User"
      End If
      If strTaskListRole = "" Then
         strTaskListRole = "User"
      End If
      If strDocRole = "" Then 
         strDocRole = "Full"
      End If
      strUserLevel = ",'" & strUserLevel & "'"
      strUserLevelTemp = ",UserLevel"
      strTaskListRole = ",'" & strTaskListRole & "'"
      strTaskListRoleTemp = ",TaskListRole"
      strDocRole = ",'" & strDocRole & "'"
      strDocRoleTemp = ",DocumentationRole"
      strUsername = ",'" & strUsername & "'"
      strUsernameTemp = ",Username"
   Else
      strEMailTemp = ""
      strUserLevelTemp = ""
   End If
   
   'Create the Regular Expression object and set it's properties.
   Set objRegExp = New RegExp
   objRegExp.Pattern = "'"
   objRegExp.Global = True
   
   'Use the regular expression to change a ' to a '' so the SQL Insert command will work.
   'The value will be assigned to a new variable so the old one can still be displayed  
   strItemTemp = objRegExp.Replace(strItem,"''")
      
   'Build the SQL string
   strSQL = "Select " & strItemType & "." & strItemType & vbCRLF
   strSQL = strSQL & "From " & strItemType & vbCRLF
   strSQL = strSQL & "Order By " & strItemType & "." & strItemType & ";"

   'Execute the SQL String
   Set objRecordSet = Application("Connection").Execute(strSQL)
   
   'Loop through the record set and make sure the category doesn't already exists
   Do Until objRecordSet.EOF
      If UCase(objRecordSet(0)) = UCase(strItem) Then
         AddItem = """" & strItem & """ Already exists - If it is disabled you can enable it." 
         strItem = ""         
      End If
      objRecordSet.MoveNext
   Loop   

   'Add the data to the database if it isn't blank
   If strItem <> "" Then 
      If strItemType = "tech" Then
         If strItemTemp = "" Or strEMail = "" Or strUserLevel = "" Or strUserName = ",''" Then
            AddItem = "You didn't fill out all fields"
         Else
            strSQL = "Insert Into " & strItemType & " (" & strItemType & ",Active" & strEMailTemp & strUserLevelTemp & strUsernameTemp & strTaskListRoleTemp & strDocRoleTemp & _
            ",MobileVersion) values ('" & strItemTemp & "',True" & strEMail & strUserLevel & strUsername & strTaskListRole & strDocRole & ",True)"   
            Application("Connection").Execute(strSQL)
            AddItem = """" & strItem & """" & " added to the " & strItemType & " list"
         End If
      Else
         strSQL = "Insert Into " & strItemType & " (" & strItemType & ",Active" & strEMailTemp & strUserLevelTemp & strUsernameTemp & _
         ") values ('" & strItemTemp & "',True" & strEMail & strUserLevel & strUsername & ")"   
         Application("Connection").Execute(strSQL)
         AddItem = """" & strItem & """" & " added to the " & strItemType & " list"       
      End If
   Else
      If AddItem = "" Then   
         AddItem = "You must enter a value"
      End If
   End If
   
   'Close all open object
   Set objRegExp = Nothing
   
End Function

Function DisableItem(strItem,strItemType)

   'This function will disable the selected item by setting the Active boolean to false.
   'It returns a string that contains the status of the function.

   Dim strSQL

   'Disable the data in the database if the input isn't blank
   If strItem <> "" Then            
      strSQL = "Update " & strItemType & vbCRLF
      strSQL = strSQL & "Set Active = False" & vbCRLF
      strSQL = strSQL & "Where " & strItemType & " = '" & strItem & "'"
      Application("Connection").Execute(strSQL)
      DisableItem = """" & strItem & """" & " has been disabled"
      
      If strItemType = "tech" Then
         strSQL = "UPDATE Main" & vbCRLF
         strSQL = strSQL & "SET Tech = '', Status='New Assignment'" & vbCRLF
         strSQL = strSQL & "WHERE Tech='" & strItem &"' AND Not Status='Complete';"
         Application("Connection").Execute(strSQL)
      End If
      
   Else
      If DisableItem = "" Then   
         DisableItem = "You must choose a value"
      End If
   End If
   
End Function

Function DeleteItem(strItem,strItemType)

   'This function will disable the selected item by setting the Active boolean to false.
   'It returns a string that contains the status of the function.

   Dim strSQL

   'Disable the data in the database if the input isn't blank
   If strItem <> "" Then            
      strSQL = "DELETE FROM " & strItemType & vbCRLF
      strSQL = strSQL & "Where " & strItemType & " = '" & strItem & "'"
      Application("Connection").Execute(strSQL)
      DeleteItem = """" & strItem & """" & " has been deleted"
      
      If strItemType = "tech" Then
         strSQL = "UPDATE Main" & vbCRLF
         strSQL = strSQL & "SET Tech = '', Status='New Assignment'" & vbCRLF
         strSQL = strSQL & "WHERE Tech='" & strItem &"' AND Not Status='Complete';"
         Application("Connection").Execute(strSQL)
      End If
      
   Else
      If DeleteItem = "" Then   
         DeleteItem = "You must choose a value"
      End If
   End If
   
End Function

Function EnableItem(strItem,strItemType)

   'This function will disable the selected item by setting the Active boolean to false.
   'It returns a string that contains the status of the function.

   Dim strSQL

   'Disable the data in the database if the input isn't blank
   If strItem <> "" Then            
      strSQL = "Update " & strItemType & vbCRLF
      strSQL = strSQL & "Set Active = True" & vbCRLF
      strSQL = strSQL & "Where " & strItemType & " = '" & strItem & "'"
      Application("Connection").Execute(strSQL)
      EnableItem = """" & strItem & """" & " has been enabled"
   Else
      If EnableItem = "" Then   
         EnableItem = "You must choose a value"
      End If
   End If
   
End Function

Function ModifyUser
   
   Dim strModifyTech, strUserName, strEMail, strUserLevel, strOldTech, objOldTech, strTaskListRole
   Dim strDocRole
   
   strModifyTech = Replace(Request.Form("ModifyTech"),"'","''")
   strUserName = Replace(Request.Form("ModifyUserName"),"'","''")
   strEMail = Replace(Request.Form("ModifyEMail"),"'","''")
   strUserLevel = Replace(Request.Form("ModifyUserLevel"),"'","''")
   strTaskListRole = Replace(Request.Form("ModifyTaskListRole"),"'","''")
   strDocRole = Replace(Request.Form("ModifyDocsRole"),"'","''")
   intID = Request.Form("ID")
   
   If strModifyTech = "" Or strUserName = "" Or strEMail = "" or strUserLevel = "" or strTaskListRole = "" Then
      ModifyUser = "You cannot have blank values."
   Else
      If bolUseAD Then
         strModifyTech = GetFirstandLastName(strUserName)
         strEMail = LCase(GetEMail(strUserName))
      End If
   
      strSQL = "SELECT Tech FROM Tech WHERE ID=" & intID
      Set objOldTech = Application("Connection").Execute(strSQL)
      strOldTech = objOldTech(0)

      strSQL = "UPDATE Tech" & vbCRLF
      strSQL = strSQL & "SET Tech='" & strModifyTech & "',UserName='" & strUserName & "',EMail='" & strEMail & "',UserLevel='" & strUserLevel & "',TaskListRole='" & strTaskListRole & "',DocumentationRole='" & strDocRole & "'" & vbCRLF
      strSQL = strSQL & "Where ID=" & intID
      Application("Connection").Execute(strSQL)
      
      strSQL = "UPDATE Main" & vbCRLF
      strSQL = strSQL & "SET Tech='" & strModifyTech & "'" & vbCRLF
      strSQL = strSQL & "WHERE Tech='" & strOldTech & "'"
      Application("Connection").Execute(strSQL)
      
      strSQL = "UPDATE Log" & vbCRLF
      strSQL = strSQL & "SET OldValue='" & strModifyTech & "'" & vbCRLF
      strSQL = strSQL & "WHERE OldValue='" & strOldTech & "'"
      Application("Connection").Execute(strSQL)
      
      strSQL = "UPDATE Log" & vbCRLF
      strSQL = strSQL & "SET NewValue='" & strModifyTech & "'" & vbCRLF
      strSQL = strSQL & "WHERE NewValue='" & strOldTech & "'"
      Application("Connection").Execute(strSQL)
      
      ModifyUser = "User Updated" 
   End If

End Function

Function AssignTechToLocation(strLocation,strTech)

   If strLocation <> "" Then
      strSQL = "UPDATE Location SET Tech='" & Replace(strTech,"'","''") & "' WHERE Location='" & Replace(strLocation,"'","''") & "'" 
      Application("Connection").Execute(strSQL)
      If strTech = "" Then
         AssignTechToLocation = "The assigned tech has been removed from " & strLocation
      Else
         AssignTechToLocation = strTech & " has been assigned to " & strLocation
      End If
   Else
      AssignTechToLocation = "Location cannot be blank."
   End If

End Function

Function AssignSubnetToLocation(strSubnet,strLocation)

   If strLocation = "" Then
      AssignSubnetToLocation = "Location cannot be blank."
   ElseIf strSubnet = "" Then
      AssignSubnetToLocation  = "Subnet cannot be blank."
   Else
   
      'See if the subnet is already in the database
      strSQL = "SELECT ID FROM Subnets WHERE Subnet='" & strSubnet & "'"
      Set objSubnets = Application("Connection").Execute(strSQL)
      
      If objSubnets.EOF Then
         
         If SubnetValid(strSubnet) Then
         
            strSQL = "INSERT INTO Subnets (Subnet,Location) "
            strSQL = strSQL & " Values ('" & Replace(strSubnet,"'","''") & "','" & Replace(strLocation,"'","''") & "')" 
            Application("Connection").Execute(strSQL)

            AssignSubnetToLocation = strSubnet & " has been assigned to " & strLocation
         
         Else
            AssignSubnetToLocation  = "Invalid Subnet."
         End If
      Else
         AssignSubnetToLocation  = "Subnet already exists in the database."
      End If
   End If

End Function

Function DeleteSubnet(intID)

   If IsNumeric(intID) Then
      strSQL = "DELETE FROM Subnets WHERE ID=" & intID
      Application("Connection").Execute(strSQL)
      DeleteSubnet = "Subnet Deleted" 
   Else
      DeleteSubnet = "You must select a subnet."
   End If
   
End Function

Function SubnetValid(strSubnet)
   
   Dim arrSubnet, strNetworkID, intNetworkBits
   
   SubnetValid = True
   
   'Split the subnet into the network ID and the number of network bits
   If InStr(strSubnet,"/") Then
      arrSubnet = Split(strSubnet,"/")
      strNetworkID = arrSubnet(0)
      intNetworkBits = arrSubnet(1)
      
      If Not IPValid(strNetworkID) Then
         SubnetValid = False
      End If
      
      If IsNumeric(intNetworkBits) Then
         If intNetworkBits > 32 Or intNetworkBits < 0 Then
            SubnetValid = False
         End If
      Else
         SubnetValid = False
      End If
      
   Else 
      SubnetValid = False
   End If
   
End Function

Function IPValid(strIP)

   Dim arrOctets, intOctet, intOctetCount

   IPValid = True
   intOctetCount = 0

   'Split the IP address into each octet
   arrOctets = Split(strIP,".")
   For Each intOctet in arrOctets
      intOctetCount = intOctetCount + 1
      If IsNumeric(intOctet) Then  
         If intOctet > 255 Or intOctet < 0 Then
            IPValid = False
         End If
      Else 
         IPValid = False
      End If
   Next
   
   'If there aren't four octets then it isn't valid
   If intOctetCount <> 4 Then
      IPValid = False
   End If

End Function

Sub Main

   Dim objStatusSet, objLocationSet, objCategorySet, objTechSet, strSQL, intCategoryCount
   Dim intTechCount, intLocationCount, objUserLevelSet, objTaskListRoleSet, objDocRoleSet

   'Build the SQL string and execute it to populate the category pulldown list
   strSQL = "Select Category.Category,Category.Active" & vbCRLF
   strSQL = strSQL & "From Category" & vbCRLF
   strSQL = strSQL & "Order By Category.Category;"
   Set objCategorySet = Application("Connection").Execute(strSQL)
   
   'Count the number of category's for the size of the popup window
   intCategoryCount = 0
   Do Until objCategorySet.EOF
      intCategoryCount = intCategoryCount + 1
      objCategorySet.MoveNext
   Loop
   If intCategoryCount <> 0 Then
      objCategorySet.MoveFirst
   End If
   
   'Build the SQL string and execute it to populate the tech pulldown list
   strSQL = "Select Tech.Tech,Tech.Active,Tech.Username" & vbCRLF
   strSQL = strSQL & "From Tech" & vbCRLF
   strSQL = strSQL & "Order By Tech.Tech;"
   Set objTechSet = Application("Connection").Execute(strSQL)
   
   'Count the number of techs for the size of the popup window
   intTechCount = 0
   Do Until objTechSet.EOF
      intTechCount = intTechCount + 1
      objTechSet.MoveNext
   Loop
   objTechSet.MoveFirst
   
   'Build the SQL string and execute it to populate the userlevel pulldown list
   strSQL = "Select UserLevel.UserLevel" & vbCRLF
   strSQL = strSQL & "From UserLevel" & vbCRLF
   strSQL = strSQL & "Order By UserLevel.UserLevel;"
   Set objUserLevelSet = Application("Connection").Execute(strSQL)
   
   'Build the SQL string and execute it to populate the TaskListRole pulldown list
   strSQL = "Select Role" & vbCRLF
   strSQL = strSQL & "From TaskListRoles" & vbCRLF  
   Set objTaskListRoleSet = Application("Connection").Execute(strSQL)
   
   'Build the SQL string and execute it to populate the DocRole pulldown list
   strSQL = "Select Role" & vbCRLF
   strSQL = strSQL & "From DocRoles" & vbCRLF  
   Set objDocRoleSet = Application("Connection").Execute(strSQL)
   
   'Build the SQL string and execute it to populate the status pulldown list
   strSQL = "Select Status.Status" & vbCRLF
   strSQL = strSQL & "From Status" & vbCRLF
   Set objStatusSet = Application("Connection").Execute(strSQL)
   
   'Build the SQL string and execute it to populate the location pulldown list
   strSQL = "Select Location.Location,Location.Active" & vbCRLF
   strSQL = strSQL & "From Location" & vbCRLF
   strSQL = strSQL & "Order By Location.Location;"
   Set objLocationSet = Application("Connection").Execute(strSQL)
   
  'Count the number of locations for the size of the popup window
   intLocationCount = 0
   Do Until objLocationSet.EOF
      intLocationCount = intLocationCount + 1
      objLocationSet.MoveNext
   Loop
   If intLocationCount <> 0 Then
      objLocationSet.MoveFirst  
   End If      
   
   strSQL = "SELECT ID,Subnet,Location FROM Subnets ORDER BY Location"
   Set objSubnets = Application("Connection").Execute(strSQL)   
   intSubnetCount = 0
   Do Until objSubnets.EOF
      intSubnetCount = intSubnetCount + 1
      objSubnets.MoveNext
   Loop
   If intSubnetCount <> 0 Then
      objSubnets.MoveFirst  
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
   
   <SCRIPT LANGUAGE="JavaScript" TYPE="text/javascript">
   <!--
   function popitup(url)
   {
      if (url == 'popup.asp?Item=Category') 
      {  
         intHeight = (<%=intCategoryCount%> * 25) + 80;
      }
      else if (url == 'popup.asp?Item=Location')
      {
         intHeight = (<%=intLocationCount%> * 25) + 80;   
      }
      else if (url == 'popup.asp?Item=Tech')
      {
         intHeight = (<%=intTechCount%> * 25) + 80;
      }
      else if (url == 'popup.asp?Item=Assignment')
      {
         intHeight = (<%=intLocationCount%> * 25) + 80;
      }
      else if (url == 'popup.asp?Item=Subnet')
      {
         intHeight = (<%=intSubnetCount%> * 25) + 80;
      }
   
   
      if (intHeight > 640)
      {
         intHeight = 640;
      }

   	newwindow=window.open(url,'name','height=' + intHeight + ',width=500,scrollbars=yes,top=100,resizable=yes,status=no');
   	if (window.focus) {newwindow.focus()}
   	return false;
   }
   
   // -->
   </SCRIPT>       
   
   <body>
   <div class="header">
      <%=Application("SchoolName")%> Help Desk Admin
   </div>
   
   <div class="version">
      Version <%=Application("Version")%>
   </div>
   
   <hr class="admintopbar" />
   <div class="admintopbar">
      <ul class="topbar">
			<li class="topbar"><a href="setup.asp">Setup</a><font class="separator"> | </font></li> 
         <li class="topbar">Users<font class="separator"> | </font></li>
         <li class="topbar"><a href="messages.asp">Messages</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="dbtools.asp">Database Tools</a><font class="separator"> | </font></li>
         <li class="topbar"><a href="index.asp">User Mode</a></li>
      <% If bolShowLogout Then %>
         <font class="separator"> | </font></li>
         <li class="topbar"><a href="login.asp?action=logout">Log Out</a></li>
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
   <table border="0" width="750" id="table1" cellspacing="0" cellpadding="0">
   	<tr>
   		<td width="22%" valign="top">
      <% If objNameCheckSet(3) = "" or IsNull(objNameCheckSet(3)) Then %>      
   		<img border="0" src="../themes/<%=Application("Theme")%>/images/admin.gif"><p>
      <% Else %>
         <img border="0" src="../themes/<%=objNameCheckSet(3)%>/images/admin.gif"><p>
      <% End If %>

   <%    'Display a message when the form is submitted
         If strMessage <> "" Then %>
            <p align="center"><font class="information"><%=strMessage%></font></p>
   <%    Else%>
   	      &nbsp;</p>
   <%    End If%>
   
         </td>
   		<td width="2%" valign="top">&nbsp;</td>
   		<td width="71%" valign="top">
         Add elements to the help desk.
   <%    If Application("UseAD") Then %>
            When adding a tech you can enter their username and hit Use AD and it will get the rest of the 
            information from Active Diretory.
   <%    End If %>

         <table width="100%">
            <tr><td>
            <table width="100%">
            <form method="POST" action="users.asp">
               <tr>
                  <td>Add Category:</td>
                  <td><input type="text" name="Category" size="35"></td>
                  <td><a href="popup.asp?Item=Category" onClick="return popitup('popup.asp?Item=Category')">View</a></td>
                  <td><input type="submit" value="Add" name="cmdCategory" style="float: right"></td>
               </tr>
            </form>
            <form method="POST" action="users.asp">
               <tr>
                  <td>Add Location:</td>
                  <td><input type="text" name="Location" size="35"></td>
                  <td><a href="popup.asp?Item=Location" onClick="return popitup('popup.asp?Item=Location')">View</a></td>
                  <td><input type="submit" value="Add" name="cmdLocation" style="float: right"></td>
               </tr>
            </form>
            <form method="POST" action="users.asp">
               <tr>
                  <td>Add Tech:</td>
                  <td>&nbsp;</td>
                  <td><a href="popup.asp?Item=Tech" onClick="return popitup('popup.asp?Item=Tech')">View</a></td>
               </tr>
               <tr>
                  <td><div align="right">Name:</div></td>
                  <td colspan="2"><input type="text" name="Tech" size="35"></td>
               </tr>
               <tr>
                  <td ><div align="right">EMail:</div></td>
                  <td><input type="text" name="EMail" size="35"></td>
                  <td>&nbsp;</td>
               </tr>
               <tr>
                  <td><div align="right">Username:</div></td>
                  <td colspan="2"><input type="text" name="Username" size="35"></td>
                  <td>&nbsp;</td>
               </tr>
               <tr>
                  <td><div align="right">Help Desk Role:</div></td>
                  <td>
                     <select size="1" name="UserLevel">
                        <option value=""></option>
               
         <%       'Populate the Userlevel pulldown list
                  Do Until objUserLevelSet.EOF %>
                        <option value="<%=objUserLevelSet(0)%>"><%=objUserLevelSet(0)%></option>
         <%          objUserLevelSet.MoveNext
                  Loop
                  objUserLevelSet.MoveFirst%>      
                     </select></td>
                  <td>&nbsp;</td>
                  <td align="right">
                  </td>
               </tr>
               <tr>
                  <td><div align="right">Task List Role:</div></td>
                  <td>
                     <select size="1" name="TaskListRole">
                        <option value=""></option>
               
         <%       'Populate the Userlevel pulldown list
                  Do Until objTaskListRoleSet.EOF %>
                        <option value="<%=objTaskListRoleSet(0)%>"><%=objTaskListRoleSet(0)%></option>
         <%          objTaskListRoleSet.MoveNext
                  Loop
                  objTaskListRoleSet.MoveFirst%>      
                     </select></td>
                  <td>&nbsp;</td>
                  <td align="right">
                  </td>
               </tr>
               <tr>
                  <td><div align="right">Docs Role:</div></td>
                  <td>
                     <select size="1" name="DocRole">
                        <option value=""></option>
               
         <%       'Populate the Userlevel pulldown list
                  Do Until objDocRoleSet.EOF %>
                        <option value="<%=objDocRoleSet(0)%>"><%=objDocRoleSet(0)%></option>
         <%          objDocRoleSet.MoveNext
                  Loop
                  objDocRoleSet.MoveFirst%>      
                     </select></td>
                  <td>&nbsp;</td>
                  <td align="right">
               <% If Application("UseAD") Then %>   
                     <input type="submit" value="Use AD" name="cmdTech">
               <% End If %>
                     <input type="submit" value="Add" name="cmdTech">
                  </td>
               </tr>
            </form>
            </table>
            </td></tr>
               <tr><td colspan="4"><hr /></td></tr>
   <%       If strModify <> "Select" Then %>
               <table width="100%">
               <form method="POST" action="users.asp">
                  <tr>
                     <td colspan="4">Modify properties of a user.  Select the user below to load their propereties.</td>
                  </tr>
                  <tr>
                     <td width="150">Tech: </td>
                     <td>
                        <select size="1" name="ModifyTech">
                           <option value=""></option>   
               <%          'Populate the Tech pulldown list
                           objTechSet.MoveFirst
                           Do Until objTechSet.EOF %>
                              <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
               <%             objTechSet.MoveNext
                           Loop
                           objTechSet.MoveFirst%> 
                        </select>
                     </td>
                     <td width="14%">
                        <input type="submit" value="Select" name="cmdModifyTech" style="float: right">
                     </td>
                  </tr>
                  </form>
               </table>

   <%       Else %>
               <tr><td colspan="2">
               <table width="100%">
               <form method="POST" action="users.asp">
                  <tr>
                     <td colspan="4">Modify properties of a user.
            <%    If Application("UseAD") Then %>
                        Enter their username and hit Use AD to get the user's information from Active Directory.
            <%    End If %>
                     </td>
                  </tr>
                  <tr>
                     <td width="15%">Name:</td><td><input type="text" name="ModifyTech" value ="<%=objModifyTech(0)%>" size="30"></td>
                  </tr>
                  <tr>
                     <td>EMail:</td><td><input type="text" name="ModifyEMail" value ="<%=objModifyTech(2)%>" size="30"></td>
                  </tr>
                  <tr>
                     <td>Username:</td><td><input type="text" name="ModifyUserName" value ="<%=objModifyTech(1)%>" size="30">
                     </td>
                  </tr>
                  <tr><td colspan="4">
                     <table width="100%">
                     <td width="30%">Help Desk Role:</td>
                     
                        <td><select size="1" name="ModifyUserLevel">
                              <option value="<%=objModifyTech(3)%>"><%=objModifyTech(3)%></option>         
                  <%       'Populate the Userlevel pulldown list
                           Do Until objUserLevelSet.EOF
                              If IsNull(objModifyTech(3)) Then %>
                                 <option value="<%=objUserLevelSet(0)%>"><%=objUserLevelSet(0)%></option>
                  <%          Else
                                 If objModifyTech(3) <> objUserLevelSet(0) Then %>
                                    <option value="<%=objUserLevelSet(0)%>"><%=objUserLevelSet(0)%></option>
                  <%             End If
                              End If
                              objUserLevelSet.MoveNext
                           Loop
                           objUserLevelSet.MoveFirst%>      
                        </select>
                     </table>
                  </td></tr>
                  <tr><td colspan="4">
                     <table width="100%">
                     <td width="30%">Task List Role:</td>
                     
                        <td><select size="1" name="ModifyTaskListRole">
                              <option value="<%=objModifyTech(5)%>"><%=objModifyTech(5)%></option>         
                  <%       'Populate the Userlevel pulldown list
                           Do Until objTaskListRoleSet.EOF
                              If IsNull(objModifyTech(5)) Then %>
                                 <option value="<%=objTaskListRoleSet(0)%>"><%=objTaskListRoleSet(0)%></option>
                  <%          Else
                                 If objModifyTech(5) <> objTaskListRoleSet(0) Then %>
                                    <option value="<%=objTaskListRoleSet(0)%>"><%=objTaskListRoleSet(0)%></option>
                  <%             End If
                              End If
                              objTaskListRoleSet.MoveNext
                           Loop
                           objTaskListRoleSet.MoveFirst%>      
                        </select>

                     </table>
                  </td></tr>
                  
                  <tr><td colspan="4">
                     <table width="100%">
                     <td width="30%">Docs Role:</td>
                     
                        <td><select size="1" name="ModifyDocsRole">
                              <option value="<%=objModifyTech(6)%>"><%=objModifyTech(6)%></option>         
                  <%       'Populate the Userlevel pulldown list
                           Do Until objDocRoleSet.EOF
                              If IsNull(objModifyTech(6)) Then %>
                                 <option value="<%=objDocRoleSet(0)%>"><%=objDocRoleSet(0)%></option>
                  <%          Else
                                 If objModifyTech(6) <> objDocRoleSet(0) Then %>
                                    <option value="<%=objDocRoleSet(0)%>"><%=objDocRoleSet(0)%></option>
                  <%             End If
                              End If
                              objDocRoleSet.MoveNext
                           Loop
                           objDocRoleSet.MoveFirst%>      
                        </select>
                        <input type="hidden" name="ID" value="<%=objModifyTech(4)%>">
                        </td>
                        <td colspan="2" align="right">
                           <input type="submit" value="Cancel" name="cmdModifyTech">
                  <%    If Application("UseAD") Then %>
                           <input type="submit" value="Use AD" name="cmdModifyTech">
                  <%    End If %>
                           <input type="submit" value="Modify" name="cmdModifyTech">
                        </td>
                     </table>
                  </td></tr>

               </form>
               </table>
               </td></tr>
   <%       End If %>
            <table>
            <form method="POST" action="users.asp">
               <tr><td colspan="4"><hr /></td></tr>
               <tr>
               
                  <td colspan="4">
                     You can delete or disable an item.  If you disabled an item it can still be used in queries.
                  </td>
               </tr>
               
               <tr>
                  <td width="150">Remove Category:</td>
                  <td>
                     <select size="1" name="Category">
                        <option value=""></option>       
      <%          'Populate the Category pulldown list
                  If Not objCategorySet.EOF Then 
                     Do Until objCategorySet.EOF 
                        If objCategorySet(1) = True Then %>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
         <%             End If
                        objCategorySet.MoveNext
                     Loop
                     If intCategoryCount <> 0 Then
                        objCategorySet.MoveFirst
                     End If
                  End If %>
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td align="right">
                     <input type="submit" value="Delete" name="cmdDisableCategory">
                     <input type="submit" value="Disable" name="cmdDisableCategory">
                  </td>
               </tr>  
            </form>
            
            <form method="POST" action="users.asp">
               <tr>
                  <td width="100">Remove Location:</td>
                  <td>
                     <select size="1" name="Location">
                        <option value=""></option>
      <%          'Populate the Location pulldown list
                  If Not objLocationSet.EOF Then
                     Do Until objLocationSet.EOF 
                        If objLocationSet(1) = True Then %>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
         <%             End If
                        objLocationSet.MoveNext
                     Loop
                     If intLocationCount <> 0 Then
                        objLocationSet.MoveFirst
                     End If 
                  End If %>
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td align="right">
                     <input type="submit" value="Delete" name="cmdDisableLocation">
                     <input type="submit" value="Disable" name="cmdDisableLocation">
                  </td>
               </tr>
            </form>
            
            <form method="POST" action="users.asp">
               <tr>
                  <td>Remove Tech:</td>
                  <td>
                     <select size="1" name="Tech">
                        <option value=""></option>                 
      <%          'Populate the Tech pulldown list
                  If Not objTechSet.EOF Then
                     Do Until objTechSet.EOF 
                        If objTechSet(1) = True And LCase(objTechSet(2)) <> LCase(strUser) Then %>
                           <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
         <%             End If
                        objTechSet.MoveNext
                     Loop
                     objTechSet.MoveFirst
                  End If %>       
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td align="right">
                     <input type="submit" value="Delete" name="cmdDisableTech">
                     <input type="submit" value="Disable" name="cmdDisableTech">
                  </td>
               </tr>
            </form>
            </table>

            <table width="100%">
            <form method="POST" action="users.asp">
               <tr><td colspan="4"><hr /></td></tr>
               <tr>
                  <td colspan="4">Enabling an item allows you to assign it's value to new tickets.</td>
               </tr>
               <tr>
                  <td width="150">Enable Category:</td>
                  <td>
                     <select size="1" name="Category">
                        <option value=""></option>
      <%          'Populate the Category pulldown list
                  If Not objCategorySet.EOF Then 
                     Do Until objCategorySet.EOF 
                        If objCategorySet(1) = False Then %>
                           <option value="<%=objCategorySet(0)%>"><%=objCategorySet(0)%></option>
         <%             End If
                        objCategorySet.MoveNext
                     Loop
                     objCategorySet.MoveFirst
                  End If %>
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td><input type="submit" value="Enable" name="cmdEnableCategory" style="float: right"></td>
               </tr>
             </form>
            <form method="POST" action="users.asp">
               <tr>
                  <td>Enable Location:</td>
                  <td>
                     <select size="1" name="Location">
                        <option value=""></option>
                        
      <%          'Populate the Location pulldown list
                  If Not objLocationSet.EOF Then
                     Do Until objLocationSet.EOF
                        If objLocationSet(1) = False Then %>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
         <%             End If
                        objLocationSet.MoveNext
                     Loop
                     objLocationSet.MoveFirst
                  End If %>   
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td><input type="submit" value="Enable" name="cmdEnableLocation" style="float: right"></td>
               </tr>
            </form>
            <form method="POST" action="users.asp">
               <tr>
                  <td>Enable Tech:</td>
                  <td>
                     <select size="1" name="Tech">
                           <option value=""></option>    
         <%          'Populate the Tech pulldown list
                     If Not objTechSet.EOF Then
                        Do Until objTechSet.EOF
                           If objTechSet(1) = False Then %>
                              <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
            <%             End If
                           objTechSet.MoveNext
                        Loop
                        objTechSet.MoveFirst
                     End If %>
                     </select>
                  </td>
                  <td>&nbsp;</td>
                  <td><input type="submit" value="Enable" name="cmdEnableTech" style="float: right"></td>
               </tr>
            </form>
            <tr><td colspan="4"><hr /></td></tr>
            </table>
            <table width="100%">
            <form method="POST" action="users.asp">
               <tr>
                  <td colspan="3">
                     Assign a tech to a location.  Ticket's will be automatically assigned to the tech when entered.
                  </td>
               </tr>
               <tr>
                  <td width="1">
                     Location: 
                  </td>
                  <td colspan="2">
                     <select size="1" name="Location">
                        <option value=""></option>
                        
      <%          'Populate the Location pulldown list
                  If Not objLocationSet.EOF Then
                     Do Until objLocationSet.EOF
                        If objLocationSet(1) = True Then %>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
         <%             End If
                        objLocationSet.MoveNext
                     Loop
                     objLocationSet.MoveFirst
                  End If%>   
                     </select>
                  </td>
               </tr>
               <tr>
                  <td width="1">
                     Tech:
                  </td>
                  <td width="325">
                     <select size="1" name="Tech">
                           <option value=""></option>    
         <%          'Populate the Tech pulldown list
                     If Not objTechSet.EOF Then
                        Do Until objTechSet.EOF
                           If objTechSet(1) = True Then %>
                              <option value="<%=objTechSet(0)%>"><%=objTechSet(0)%></option>
            <%             End If
                           objTechSet.MoveNext
                        Loop
                        objTechSet.MoveFirst 
                     End If %>
                     </select>
                  </td>
                  <td>
                     <a href="popup.asp?Item=Assignment" onClick="return popitup('popup.asp?Item=Assignment')">View</a>
                     <input type="submit" value="Assign" name="cmdAssignTech" style="float: right">
                  </td>
               </tr>
               <tr><td colspan="3"><hr /></td></tr>
            </form>
            </table>
            <table width="100%">
            <form method="POST" action="users.asp">
               <tr>
                  <td colspan="3">
                     Assign a subnet to a location.  Location information will be automatically populated when entering
                     a new ticket using the client's IP address.  Enter the subnet information using CIDR notation. 
                     (192.168.1.0/24)
                  </td>
               </tr>
               <tr>
                  <td width="1">
                     Subnet: 
                  </td>
                  <td colspan="2">
                     <input type="text" name="Subnet" size="25">
                  </td>
               </tr>
               <tr>
                  <td width="1">
                     Location: 
                  </td>
                  <td width="325">
                     <select size="1" name="Location">
                        <option value=""></option>
   
      <%          'Populate the Location pulldown list
                  If NOt objLocationSet.EOF Then
                     Do Until objLocationSet.EOF
                        If objLocationSet(1) = True Then %>
                           <option value="<%=objLocationSet(0)%>"><%=objLocationSet(0)%></option>
         <%             End If
                        objLocationSet.MoveNext
                     Loop
                     objLocationSet.MoveFirst
                  End If %>   
                     </select>
                  </td>
                  <td>
                     <a href="popup.asp?Item=Subnet" onClick="return popitup('popup.asp?Item=Subnet')">View</a>
                     <input type="submit" value="Assign" name="cmdAssignSubnet" style="float: right">
                  </td>
               </tr>
               <tr><td colspan="3"><hr /></td></tr>
            </form>
            </table>
            <table width="100%">
            <form method="POST" action="users.asp">
               <tr>
                  <td colspan="3">
                     Delete a subnet.
                  </td>
               </tr>
               <tr>
                  <td width="1">
                     Subnet: 
                  </td>
                  <td>
                     <select size="1" name="Subnet">
                        <option value=""></option>
   
      <%          'Populate the Subnet pulldown list
                  If Not objSubnets.EOF Then
                     Do Until objSubnets.EOF %>
                        <option value="<%=objSubnets(0)%>"><%=objSubnets(1) & " - " & objSubnets(2)%></option>
         <%             objSubnets.MoveNext
                     Loop
                     objSubnets.MoveFirst
                  End If %>   
                     </select>
                  </td>
                  <td>
                     <input type="submit" value="Delete" name="cmdDeleteSubnet" style="float: right">
                  </td>
               </tr>
               <tr><td colspan="3"><hr /></td></tr>
            </form>
            </table>
            </td></tr>
            </td></tr>
         </table>
      </td></tr>
   </table>
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
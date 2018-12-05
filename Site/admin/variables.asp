<%@LANGUAGE="VBSCRIPT" CODEPAGE="1252"%>
<%
'Created by Matthew Hull on 1/9/12
'Last Updated 6/16/14

'This page is the popup that will show all the items in the database for the modify page.

Option Explicit

On Error Resume Next

Dim strSQL, objNameCheckSet

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
   <meta name="viewport" content="width=device-width" />
</head>

<body>
   
   <b>Availible Variables</b> &nbsp;&nbsp;&nbsp;&nbsp; <a href="javascript:window.close()">Close</a>

	<table border="1" width="100%">
		<tr>
			<td class="showborders"><b>Variable</b></td>
			<td class="showborders"><b>Value</b></td>
		</tr>
		<tr>
			<td class="showborders">#TICKET#</td>
			<td class="showborders">The current ticket number.</td>
		</tr> 
      <tr>
			<td class="showborders">#CURRENTUSER#</td>
			<td class="showborders">The current tech logged into the help desk.</td>
		</tr>
      <tr>
			<td class="showborders">#USER#</td>
			<td class="showborders">The user on the ticket.</td>
		</tr>
      <tr>
			<td class="showborders">#TECH#</td>
			<td class="showborders">The tech assigned the ticket.</td>
		</tr>
      <tr>
			<td class="showborders">#STATUS#</td>
			<td class="showborders">The current ticket status.</td>
		</tr>
      <tr>
			<td class="showborders">#USEREMAIL#</td>
			<td class="showborders">The user's email from the ticket.</td>
		</tr>
      <tr>
			<td class="showborders">#LOCATION#</td>
			<td class="showborders">The user's location.</td>
		</tr>
      <tr>
			<td class="showborders">#CUSTOM1#</td>
			<td class="showborders">The value of your first custom variable</td>
		</tr>
      <tr>
			<td class="showborders">#CUSTOM2#</td>
			<td class="showborders">The value of your second custom variable</td>
		</tr>
      <tr>
			<td class="showborders">#PROBLEM#</td>
			<td class="showborders">The problem reported by the user.</td>
		</tr>
      <tr>
			<td class="showborders">#NOTES#</td>
			<td class="showborders">The notes entered by the tech.</td>
		</tr>
      <tr>
			<td class="showborders">#LINK#</td>
			<td class="showborders">The link to the current ticket.</td>
		</tr>
	</table>
</body>
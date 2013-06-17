<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN" "http://www.w3.org/TR/html4/loose.dtd">
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Instrument Spec Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Instrument Spec Database</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<link rel="stylesheet" href="equipmentstyle.css" type="text/css">
<style type="text/css">
  div {font-family:verdana;}
  input {font-family:verdana;}
  select {font-family:verdana;
		width:99%;}
  textarea {font-family:verdana;}
  input[type=button] {font-size:10pt;
		width:140px;
		height:25px;}
</style> 
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, March 30, 2009
'   Creation
' Keith Brooks - Monday, July 30, 2012
'	Cleaned up html to remove deprecated items.
'*************

dim HitCounts
Dim currentuser
Dim access
Dim access2

'Set/get hit counts.
HitCounts = HitCounter("equipment_home")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
'access = UserAllowed(currentuser, "masterbatchentry")
access = UserAccess("equipment", "default", currentuser)
If access <> "none" Then

%>
	<div class="center">
		<table style="width:100%">
			<tr>
				<td style="width:25%">&nbsp;</td>
				<td class="center" style="width:50%"><h1>Instrument Spec Database</h1></td>
				<td class="right top" style="width:25%"><a class="noprint" href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
			<tr>
				<td>&nbsp;</td>
				<td class="center"><input type="button" name="findinstrument" value="Find Spec" title="Select an instrument spec to view"  onclick="window.location='findinstrument.asp'" /></td>
				<td>&nbsp;</td>
			</tr>
<%
		access2 = UserAccess("equipment","addinstrument",currentuser)
		If access2 <> "none" Then
%>
			<tr>
				<td>&nbsp;</td>
				<td class="center"><input type="button" name="addinstrument" value="Add Instrument" title="Create a new instrument spec" onclick="window.location='addinstrument.asp'" /></td>
				<td>&nbsp;</td>
			</tr>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","adminmenu",currentuser)
		If access2 <> "none" Then
%>
			<tr>
				<td>&nbsp;</td>
				<td class="center"><input type="button" name="adminmenu" value="Administration" title="Open the administration main menu" onclick="window.location='adminmenu.asp'" /></td>
				<td>&nbsp;</td>
			</tr>
<%
		End If
%>
		</table>
	</div>
<%
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</body>
</html>

<%@ language="vbscript" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
function doDelete(id) {
	if (confirm("Are you sure you want to delete record number "+id+"?")) {
		document.form1.action="adminaction.asp?action=delete&RECORD="+id;
		document.form1.submit()
	}
}
function openhelp() {
 window.open("Instrument Spec Database Administrators Guide.doc","userguide");
}
</script>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Administer Classifications</title>
<link rel="stylesheet" href="equipmentstyle.css" type="text/css">
<style>
  input {font-family:verdana;
		font-size:8pt;
		background-color:#DBF5F5}
<!--
  a     { text-decoration:underline; }
-->
</style> 
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, April 22, 2010
'   Creation.
' Keith Brooks - Monday, July 30, 2012
'	Cleaned up html to remove deprecated items.
'*************

'on error resume next

dim cn
dim rs
dim recordid
Dim currentuser
Dim access
Dim recid
Dim sqlString
Dim sortkey
Dim sortdir

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment","classifications",currentuser)
If access <> "none" Then

	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	set rs = CreateObject("adodb.recordset")

	If session("err") <> "" And session("err") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("err") & ".focus();'>"
	ElseIf session("focus") <> "" And session("focus") <> "NONE" Then
	  Response.Write "<body onload='document.form1." & session("focus") & ".focus();'>"
	  session("focus") = "NONE"
	Else
	  Response.Write "<body>"
	End If

	If request("record_id") <> "" Then
	  If IsNumeric(request("record_id")) Then
	    recordid = request("record_id")
	  Else
	    recordid = 0
	  End If
	Else
	  recordid = 0
	End If
	If request("sort") <> "" Then
		sortkey = request("sort")
	Else
		sortkey = "classification_id"
	End If
	If request("direction") <> "" Then
		sortdir = request("direction")
	Else
		sortdir = "ASC"
	End If

	response.write "<table class='headertable' style='width:100%'>"
	response.write "<tr>"
	response.write "<td class='headertd left top' style='width:20%'><a class='noprint' href='adminmenu.asp' title='Open the administration main menu'>Menu</a></td>"
	response.write "<td class='headertd center middle' style='width:60%'><h1/>Edit Classifications</td>"
	response.write "<td class='headertd right top' style='width:20%'><a class='noprint' href='#' onClick='openhelp();return false;' title='Open the Admin Guide'>Help</a></td>"
	response.write "</tr>"

	response.write "<tr class='noprint'>"
	response.write "<td class='headertd'>&nbsp;</td>"
	If access = "write" Or access = "delete" Then
		response.write "<td class='headertd center'><a href='classifications.asp?record_id=-1&sort=" & sortkey & "&direction=" & sortdir & "' title='Add a new classification record'>Add new record</a></td>"
	Else
		Response.Write "<td class='headertd'>&nbsp;</td>"
	End If
	response.write "<td class='headertd'>&nbsp;</td>"
	response.write "</tr>"
	response.write "</table>"

	Response.Write "<br />"

	'Draw the header
	Response.Write "<div class='center'>"
	response.Write "<table style='width:55%'>"
	response.Write "<tr>"
	response.Write "<th class='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td class='headertd'><a class='noprint' href='classifications.asp?sort=classification_id&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Classification ID&nbsp;</th>"
	Response.Write "<td class='headertd'><a class='noprint' href='classifications.asp?sort=classification_id&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	response.Write "<th class='mediumth'>"
	Response.Write "<table><tr>"
	Response.Write "<td class='headertd'><a class='noprint' href='classifications.asp?sort=classification_desc&direction=DESC'>"
	response.Write "<img src='mdownarrow.gif' alt='DESC'></a></td>"
	Response.Write "<th style='font-size:8pt'>&nbsp;Classification Description&nbsp;</th>"
	Response.Write "<td class='headertd'><a class='noprint' href='classifications.asp?sort=classification_desc&direction=ASC'>"
	Response.Write "<img src='muparrow.gif' alt='ASC'></a></td></tr></table></th>"
	If access = "write" Or access = "delete" Then
 	  response.Write "<th class='mediumth noprint'>&nbsp;</th>"
 	End If
	If access = "delete" Or recordid < 0 Then
 	  response.Write "<th class='mediumth noprint'>&nbsp;</th>"
 	End If
	response.Write "</tr>"

	response.Write "<form action='adminaction.asp' id='form1' method='post' name='form1'>"

	response.write "<input type='hidden' name='RECORD' value='" & recordid & "'>"
	response.write "<input type='hidden' name='SORT' value='" & sortkey & "'>"
	response.write "<input type='hidden' name='DIRECTION' value='" & sortdir & "'>"

	sqlString = "SELECT * FROM classifications " & _
				"ORDER BY " & sortkey & " " & sortdir
	set rs = cn.Execute(sqlString)

	If Not rs.BOF Then
	  rs.MoveFirst
	End If
	  
	'If recordid<0, the user has selected "Add new record" so insert a blank data entry line
	'at the top of the form.
	If access = "write" Or access = "delete" Then
		If recordid < 0 Then
			Response.Write "<tr>"

			response.write "<td class='mediumtd'>&nbsp;</td>"

			If session("err") = "classification_desc" Then
				If session("classification_desc") <> "" Then
					response.write "<td class='mediumtd error'><input type='text' name='classification_desc' size='50' value='" & session("classification_desc") & "'></td>"
				Else
					response.write "<td class='mediumtd error'><input type='text' name='classification_desc' size='50' value=''></td>"
				End If
			Else
				If session("classification_desc") <> "" Then
					response.write "<td class='mediumtd'><input type='text' name='classification_desc' size='50' value='" & session("classification_desc") & "'></td>"
				Else
					response.write "<td class='mediumtd'><input type='text' name='classification_desc' size='50' value=''></td>"
				End If
			End If

			response.write "<td class='mediumtd center noprint'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"

			Response.Write "<td class='mediumtd center noprint'><a href='classifications.asp?sort=" & sortkey & "&direction=" & sortdir & "' title='Cancel changes to this record'>Cancel</a></td>"

			Response.Write "</tr>"
		End If
	End If

	Do While Not rs.EOF
		Response.Write "<tr>"
		If CLng(rs("classification_id")) = CLng(recordid) Then
			'Draw the data entry line
			response.write "<td class='mediumtd'>" & rs("classification_id") & "</td>"

			If session("err") = "classification_desc" Then
				If session("classification_desc") <> "" Then
					response.write "<td class='mediumtd' style='background-color:red'><input type='text' name='classification_desc' size='50' value='" & session("classification_desc") & "'></td>"
				Else
					response.write "<td class='mediumtd' style='background-color:red'><input type='text' name='classification_desc' size='50' value='" & rs("classification_desc") & "'></td>"
				End If
			Else
				If session("classification_desc") <> "" Then
					response.write "<td class='mediumtd'><input type='text' name='classification_desc' size='50' value='" & session("classification_desc") & "'></td>"
				Else
					response.write "<td class='mediumtd'><input type='text' name='classification_desc' size='50' value='" & rs("classification_desc") & "'></td>"
				End If
			End If

			If access = "write" Or access = "delete" Then
				response.write "<td class='mediumtd' style='text-align:center'><input type='submit' value='Submit' id='submit1' name='submit1' style='font-size:8pt;background-color:white' title='Apply changes to this record'></td>"
			End If

			If access = "delete" Then
				recid = rs("classification_id")
				response.write "<td class='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			End If
		Else
			'Draw the history records
			response.write "<tr>"
			response.write "<td class='mediumtd'>" & rs("classification_id") & "</td>"
			If rs("classification_desc") <> "" Then
				response.write "<td class='mediumtd'>" & rs("classification_desc") & "</td>"
			Else
				Response.Write "<td class='mediumtd'>&nbsp;</td>"
			End If
			If access = "write" Or access = "delete" Then
				response.write "<td class='mediumtd'><a href='classifications.asp?record_id=" & rs("classification_id") & "&sort=" & sortkey & "&direction=" & sortdir & "' title='Edit this record'>Edit</a></td>"
			End If
			If access = "delete" Then
				recid = rs("classification_id")
				response.write "<td class='mediumtd'><a href='javascript:doDelete(" & recid & ");' title='Delete this record'>Delete</a></td>"
			ElseIf recordid < 0 Then
				Response.Write "<td class='mediumtd'>&nbsp;</td>"
			End If
			response.write "</tr>"
		End If
		rs.Movenext
	loop
	rs.Close
	
	Response.Write "</form>"
	Response.Write "</table>"
	Response.Write "</div>"
	Response.Write "</body>"

	'session("err") = "NONE"

	Set rs = Nothing
	cn.Close
	Set cn = Nothing

Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If
%>
</html>

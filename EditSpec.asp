<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<script language="javascript">
var needToConfirm = true;

window.onbeforeunload = confirmExit;

function confirmExit() {
 document.getElementById('PleaseWait').style.display = 'block';
 if ((needToConfirm)&&(document.form1.changedtags.value!='')) {
  document.getElementById('PleaseWait').style.display = 'none';
  return "You have made changes to this form that have not been saved.";
 }
}
function reloadpage(id,page,ctl) {
 needToConfirm = false;
 document.getElementById('PleaseWait').style.display = 'block';
 document.form1.action="editspec.asp?instrID=" + id + "&page_num=" + page + "&ctl=" + ctl;
 document.form1.submit();
}
function saveChange(id) {
 if (document.form1.changedtags.value=='') {
  document.form1.changedtags.value=id;
 } else {
  document.form1.changedtags.value=document.form1.changedtags.value+','+id;
 }
}
function openRev(id) {
 window.open("updaterev.asp?instrID="+id,"updaterev","width=500px,height=300px");
}
function openEquip(id) {
 window.open("updateequipid.asp?instrumentID="+id,"updateequip","width=500px,height=300px");
}
function demoInstrument(stat) {
 var answer;
 if (stat==1) {
  answer = confirm("Are you sure you want to void this instrument?");
 } else {
  answer = confirm("Are you sure you want to return this instrument to active service?");
 }
 if (answer) {
  document.form1.action="updatespec.asp?demo="+stat;
  document.form1.submit();
 }
}
function openhelp() {
 window.open("Instrument Spec Database Users Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Spec Sheet</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<style>
  div {font-family:verdana;}
  input {font-family:verdana;}
  select {font-family:verdana;
		width:99%;}
  textarea {font-family:verdana;}
  @media print { .noprint {display: none; } }
</style>
</head>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, March 23, 2010
'   Creation
' Keith Brooks - Monday, December 6, 2010
'	Changed "demoed" text to "voided" and changed water-mark file from 
'	"demoed2.gif to "voided3.gif".
'*************

'This function emulates the VB iif function to return the first string if the
'expression evaluates True and the second string otherwise.
Function iif(boolEval, trueStr, falseStr)
	If boolEval Then
		iif = trueStr
	Else
		iif = falseStr
	End If
End Function


'Dim strSitePath
Dim objConnection
Dim objRecordset
Dim objRecordset2
Dim objRecordset3
Dim strSQL
Dim instrumentID
Dim strArray
Dim strValue
Dim strField
Dim fieldArray()
Dim valueArray()
Dim fieldCount
Dim lineHeight
Dim formID
Dim pageID
Dim missingValue
Dim page_num
Dim last_page_num
Dim count
Dim labelFormatID
Dim fieldFormatID
Dim currentuser
Dim access
Dim access2

fieldCount = 0
ReDim fieldArray(fieldCount)
ReDim valueArray(fieldCount)
If Request("instrID") <> "" Then
	instrumentID = Request("instrID")
Else
	instrumentID = 0
End If
If Request("page_num") <> "" Then
	page_num = Request("page_num")
Else
	page_num = 1
End If
fieldArray(fieldCount) = "instr_id"
valueArray(fieldCount) = instrumentID
fieldCount = fieldCount + 1
ReDim Preserve fieldArray(fieldCount)
ReDim Preserve valueArray(fieldCount)
fieldArray(fieldCount) = "page_num"
valueArray(fieldCount) = page_num

If Request("ctl") <> "" Then
	Response.Write "<body onload='javascript:document.form1." & Request("ctl") & ".focus();'>"
Else
	Response.Write "<body>"
End If

set objConnection = CreateObject("adodb.connection")
'objConnection.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "\SpecDB.mdb"
objConnection.Open = DBString
set objRecordset = CreateObject("adodb.recordset")
Set objRecordset2 = CreateObject("adodb.recordset")
Set objRecordset3 = CreateObject("adodb.recordset")

Response.Write "<form id='form1' name='form1' action='updatespec.asp' method='post'>"

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
access = UserAccess("equipment", "editspec", currentuser)
If access <> "none" Then

	If Request("changedtags") = "" Then
		Response.Write "<input type='hidden' name='changedtags' id='changedtags' />"
	Else
		Response.Write "<input type='hidden' name='changedtags' id='changedtags' value='" & Request("changedtags") & "' />"
	End If
	Response.Write "<input type='hidden' name='instrumentID' id='instrumentID' value='" & instrumentID & "' />"
	Response.Write "<input type='hidden' name='pagenum' id='pagenum' value='" & page_num & "' />"

	'Get the spec form id number.
	strSQL = "SELECT spec_id FROM instruments WHERE instr_id=" & CStr(instrumentID)
	'strSQL = "SELECT instr_func_type_spec_form_id FROM specforminfoview WHERE instr_id=" & instrumentID
	Set objRecordset = objConnection.Execute(strSQL)
	If Not objRecordset.BOF Then
		objRecordset.MoveFirst
		If Not IsNull(objRecordset(0)) Then
			formID = objRecordset(0)
		Else
			formID = 0
		End If
	Else
		formID = 0
	End If
	objRecordset.Close

	If formID > 0 Then
		fieldCount = fieldCount + 1
		ReDim Preserve fieldArray(fieldCount)
		ReDim Preserve valueArray(fieldCount)
		fieldArray(fieldCount) = "formID"
		valueArray(fieldCount) = formID

		'Get the page id number.
		strSQL = "SELECT spec_page_id FROM spec_form_pages WHERE spec_form_id=" & formID & " AND spec_page_seq=" & page_num
		Set objRecordset = objConnection.Execute(strSQL)
		If Not objRecordset.BOF Then
			objRecordset.MoveFirst
			If Not IsNull(objRecordset(0)) Then
				pageID = objRecordset(0)
			Else
				pageID = 0
			End If
		Else
			pageID = 0
		End If
		fieldCount = fieldCount + 1
		ReDim Preserve fieldArray(fieldCount)
		ReDim Preserve valueArray(fieldCount)
		fieldArray(fieldCount) = "pageID"
		valueArray(fieldCount) = pageID
		objRecordset.Close

		'Get the last page number for this spec form.
		strSQL = "SELECT MAX(spec_page_seq) FROM spec_form_pages WHERE spec_form_id=" & formID
		Set objRecordset = objConnection.Execute(strSQL)
		If Not objRecordset.BOF Then
			objRecordset.MoveFirst
			If Not IsNull(objRecordset(0)) Then
				last_page_num = objRecordset(0)
			Else
				last_page_num = 0
			End If
		Else
			last_page_num = 0
		End If
		fieldCount = fieldCount + 1
		ReDim Preserve fieldArray(fieldCount)
		ReDim Preserve valueArray(fieldCount)
		fieldArray(fieldCount) = "last_page_num"
		valueArray(fieldCount) = last_page_num
		objRecordset.Close
		
		'Draw the base form.
		strSQL = "SELECT * FROM spec_base_format ORDER BY spec_base_format_id"
		Set objRecordset = objConnection.Execute(strSQL)
		If Not objRecordset.BOF Then
			objRecordset.MoveFirst
			Do While Not objRecordset.EOF
				If Not IsNull(objRecordset("label_image_path")) Then
					Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;text-align:" & objRecordset("label_align") & ";padding:1px;z-index:2;overflow:hidden'><img src='" & objRecordset("label_image_path") & "' width='" & objRecordset("label_width") & "' height='" & objRecordset("label_height") & "' /></div>"
				ElseIf Not IsNull(objRecordset("label_text")) Then
					strSQL = objRecordset("label_text")
					Do While InStr(strSQL,"[") > 0
						'Extract the field name from the string.
						strField = Mid(strSQL,InStr(strSQL,"[") + 1,InStr(strSQL,"]") - InStr(strSQL,"[") - 1)
						'strValue = Execute("Response.Write " & strField & ".value")
						For count = 0 To fieldCount
							If fieldArray(count) = strField Then
								strValue = valueArray(count)
								Exit For
							End If
						Next
						strSQL = Replace(strSQL,"[" & strField & "]",strValue)
					Loop
					lineHeight = objRecordset("label_height")
					Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight - 1 & "px;padding-left:1px;z-index:2;overflow:hidden'>" & strSQL & "</div>"
				End If
				If Not IsNull(objRecordset("field_table")) And Not IsNull(objRecordset("field_field")) Then
					strSQL = "SELECT " & objRecordset("field_field") & " FROM " & objRecordset("field_table") & " WHERE instr_id=" & instrumentID
					Set objRecordset2 = objConnection.Execute(strSQL)
					If Not objRecordset2.BOF Then
						objRecordset2.MoveFirst
						If objRecordset("field_type") = "Textbox" Then
							Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='text' id='" & objRecordset("field_field") & "' value='" & objRecordset2(0) & "' size='" & objRecordset("field_length") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";height:" & objRecordset("field_height") & "px;padding-top:0px;vertical-align:top' /></div>"
						Else
							If objRecordset("field_height") > 20 Then
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";text-align:center;padding-top:5px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
							Else
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;text-align:" & objRecordset("field_align") & ";padding-left:2px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
							End If
						End If
		'				fieldCount = fieldCount + 1
					End If
					objRecordset2.Close
				ElseIf Not IsNull(objRecordset("field_field")) Then
					If objRecordset("field_field") = "@" Then
						If CInt(objRecordset("field_height") > 1000) Then
							Response.Write "<div style=border-style:solid;border-width:2px;border-color:black;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;background-color:lightgray;z-index:0'>"
						Else
							Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";z-index:1'></div>"
						End If
					Else
						Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;text-align:" & objRecordset("field_align") & ";padding-left:5px;z-index:2;overflow:hidden'>" & objRecordset("field_field") & "</div>"
					End If
				ElseIf objRecordset("field_type") = "Lookup" Then
					If Not IsNull(objRecordset("field_source_sql")) Then
						strSQL = objRecordset("field_source_sql")
						missingValue = False
						Do While InStr(strSQL,"[") > 0
							'Extract the field name from the string.
							strField = Mid(strSQL,InStr(strSQL,"[") + 1,InStr(strSQL,"]") - InStr(strSQL,"[") - 1)
							'strValue = Execute("Response.Write " & strField & ".value")
							strValue = ""
							For count = 0 To fieldCount
								If fieldArray(count) = strField Then
									strValue = valueArray(count)
									Exit For
								End If
							Next
							If strValue <> "" Then
								strSQL = Replace(strSQL,"[" & strField & "]",strValue)
							Else
								missingValue = True
								Exit Do
							End If
						Loop
						If Not missingValue Then
							Set objRecordset2 = objConnection.Execute(strSQL)
							If Not objRecordset2.BOF Then
								objRecordset2.MoveFirst
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;text-align:" & objRecordset("field_align") & ";padding-left:2px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
							End If
						End If
					End If
				End If
				objRecordset.MoveNext
			Loop
		End If
		objRecordset.Close

		'Draw the links for other pages.
		Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:0px;left:860px;width:50px;height:20px;font-size:8pt;font-weight:bold;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><a href='default.asp' title='Open the main menu'>Home</a></div>"
		Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:0px;left:920px;width:50px;height:20px;font-size:8pt;font-weight:bold;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><a href='' onclick='openhelp();return false;' title='Open the User Guide'>Help</a></div>"
		For count = 1 To last_page_num
			If CInt(count) = CInt(page_num) Then
				Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:" & count * 20 & "px;left:860px;width:50px;height:20px;font-size:8pt;font-weight:bold;text-align:left;padding-left:2px;z-index:2;overflow:hidden'>Page " & count & "</div>"
			Else
				Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:" & count * 20 & "px;left:860px;width:50px;height:20px;font-size:8pt;font-weight:bold;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><a href='editspec.asp?instrID=" & instrumentID & "&page_num=" & count & "' title='Display page " & count & " of this instrument spec form'>Page " & count & "</a></div>"
			End If
		Next
		
		'Draw link to the printer-friendly version of the spec.
		Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:" & (last_page_num + 1) * 20 & "px;left:860px;width:100px;height:20px;font-size:8pt;font-weight:bold;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><a href='printspec.asp?instrID=" & instrumentID & "&page_num=1' title='Open the printer-friendly version of this instrument spec form'>View Spec</a></div>"
		
		'Draw link to allow the revision information to be updated for this instrument.
		access2 = UserAccess("equipment","updaterev",currentuser)
		If access2 = "write" Or access2 = "delete" Then
			Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:" & (last_page_num + 2) * 20 & "px;left:860px;width:100px;height:20px;font-size:8pt;font-weight:bold;text-aligh:left;padding-left:2px;z-index:2;overflow:hidden'><a href='' onclick='openRev(" & instrumentID & ");return false;' title='Open a form to update the revision information for this instrument spec'>Update Rev</a></div>"
		End If
		
		'Draw link to allow the equipment to be updated for this instrument.
		access2 = UserAccess("equipment","updateequipid",currentuser)
		If access2 = "write" Or access2 = "delete" Then
			Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:" & (last_page_num + 3) * 20 & "px;left:860px;width:100px;height:20px;font-size:8pt;font-weight:bold;text-aligh:left;padding-left:2px;z-index:2;overflow:hidden'><a href='' onclick='openEquip(" & instrumentID & ");return false;' title='Open a form to update the equipment for this instrument spec'>Update Equip</a></div>"
		End If
		
		'Draw "Please Wait..." message that will be displayed when this page is
		'reloading, saving data, or moving to another page.
	%>
			<div class="helptext" id="PleaseWait" style="display: none; text-align:center; color:White; vertical-align:top;border-style:none;position:absolute;top:500px;left:860px">
				<table id="MyTable" bgcolor="blue">
					<tr><td style="width: 95px"><b><font color="white">Please Wait...</font></b></td></tr>
				</table>
			</div>
	<%
		'Draw the "Save" button.
		Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:1020px;left:860px;width:50px;height:30px;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><input type='submit' id='submit1' name='submit1' title='Save the changes made to this instrument spec form' value='Save' onclick='needToConfirm = false;' /></div>"

		'Draw the "Demo" button.
		'Get the current status of the instrument.
		strSQL = "SELECT void_sequence FROM instruments WHERE instr_id=" & instrumentID
		Set objRecordset = objConnection.Execute(strSQL)
		If Not objRecordset.BOF Then
			objRecordset.MoveFirst
			If Not IsNull(objRecordset(0)) Then
				If objRecordset(0) = 0 Then
					Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:1055px;left:860px;width:100px;height:30px;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><input type='button' id='demo' name='demo' title='Void this instrument spec so that it is no longer active' value='Void' onclick='demoInstrument(1);' /></div>"
				Else
					Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:1055px;left:860px;width:100px;height:30px;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><input type='button' id='demo' name='demo' title='Return this instrument to active service' value='Un-Void' onclick='demoInstrument(0);' /></div>"
					'Draw a label status for demoed item.
					'Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:450px;left:860px;width:80px;height:20px;text-align:left;padding-left:2px;z-index:2;overflow:hidden;background-color:DarkSlateBlue;color:White;font-weight:bold'>Demoed</div>"
					'Response.Write "<div style='filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=3);border-style:none;position:absolute;top:0px;left:0px;width:840;height:1080;line-height:840px;z-index:3;overflow:hidden;font-size:144pt;color:LightGrey'>Demoed</div>"
					Response.Write "<div style='border-style:none;position:absolute;top:50px;left:100px;width:840;height:1080;z-index:3;overflow:hidden'><img src='voided4.gif' /></div>"
				End If
			Else
				Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:1055px;left:860px;width:100px;height:30px;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><input type='button' id='demo' name='demo' title='Return this instrument to active service' value='Un-Void' onclick='demoInstrument(0);' /></div>"
			End If
		Else
'			Response.Write "<div class='noprint' style='border-style:none;position:absolute;top:1055px;left:860px;width:100px;height:30px;text-align:left;padding-left:2px;z-index:2;overflow:hidden'><input type='button' id='demo' name='demo' title='Return this instrument to active service' value='Un-Demo' onclick='demoInstrument(0);' /></div>"
		End If
		objRecordset.Close

		'Get the formatting data.
		strSQL = "SELECT * FROM spec_formats WHERE spec_page_id=" & pageID & " ORDER BY spec_format_id"
		Set objRecordset = objConnection.Execute(strSQL)
		If Not objRecordset.BOF Then
			objRecordset.MoveFirst
			Do While Not objRecordset.EOF
				labelFormatID = "L" & CStr(objRecordset("spec_format_id"))
				fieldFormatID = "F" & CStr(objRecordset("spec_format_id"))
				If Not IsNull(objRecordset("num_text")) Then
					Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("num_top") & "px;left:" & objRecordset("num_left") & "px;width:" & objRecordset("num_width") & "px;height:" & objRecordset("num_height") & "px;font-size:" & objRecordset("num_font_size") & ";font-weight:" & objRecordset("num_font_weight") & ";line-height:" & objRecordset("num_height") - 1 & "px;text-align:right;padding-right:2px;z-index:2'>" & objRecordset("num_text") & "</div>"
				End If
				If Not IsNull(objRecordset("label_text")) Then
					'This is a work-around for the large grouping labels to allow the
					'text to be vertically centered.
					If objRecordset("label_height") > 20 Then
						If objRecordset("label_orientation") = "vertical" Then
							If Len(objRecordset("label_text")) < objRecordset("label_width") / 7 Then
								lineHeight = objRecordset("label_height")
							Else
								lineHeight = objRecordset("label_height") / 2
							End If
							Response.Write "<div style='filter: progid:DXImageTransform.Microsoft.BasicImage(rotation=3);border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight & "px;z-index:2'>" & objRecordset("label_text") & "</div>"
						Else
		'					If Len(objRecordset("label_text")) < 12 Then
							If Len(objRecordset("label_text")) < objRecordset("label_width") / 7.2 Then
								lineHeight = objRecordset("label_height")
							ElseIf Len(objRecordset("label_text")) <= objRecordset("label_width") / 3 And objRecordset("label_height") > 119 Then
								lineHeight = objRecordset("label_height") / 2
							Else
								lineHeight = objRecordset("label_height") / 3
							End If
							Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight & "px;z-index:2'>" & objRecordset("label_text") & "</div>"
						End If
					Else
						lineHeight = objRecordset("label_height")
						Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight & "px;padding-left:2px;padding-right:2px;z-index:2;overflow:hidden'>" & objRecordset("label_text") & "</div>"
					End If
				ElseIf Not IsNull(objRecordset("label_table")) And Not IsNull(objRecordset("label_field")) Then
					strSQL = "SELECT " & objRecordset("label_field") & " FROM " & objRecordset("label_table") & " WHERE instr_id=" & instrumentID
					Set objRecordset2 = objConnection.Execute(strSQL)
					If Not objRecordset2.BOF Then
						objRecordset2.MoveFirst
						lineHeight = objRecordset("label_height")
						If Request(labelFormatID) <> "" Then
							Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;z-index:2'><input type='text' id='" & labelFormatID & "' name='" & labelFormatID & "' value='" & Request(labelFormatID) & "' size='" & objRecordset("label_length") & "' style='font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";height:" & objRecordset("label_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("label_tabindex") & "' /></div>"
						Else
			'				Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight & "px;padding-left:2px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
			'				Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;z-index:2'><input type='text' id='" & objRecordset("label_field") & "' value='" & objRecordset2(0) & "' size='42' style='font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";height:" & objRecordset("label_height") & "px;padding-top:0px;vertical-align:top' /></div>"
							Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;z-index:2'><input type='text' id='" & labelFormatID & "' name='" & labelFormatId & "' value='" & objRecordset2(0) & "' size='" & objRecordset("label_length") & "' style='font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";height:" & objRecordset("label_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("label_tabindex") & "' onchange='saveChange(this.name);' /></div>"
						End If
					End If
					objRecordset2.Close
				ElseIf Not IsNull(objRecordset("label_field")) Then
					strSQL = objRecordset("label_field")
					Do While InStr(strSQL,"[") > 0
						'Extract the field name from the string.
						strField = Mid(strSQL,InStr(strSQL,"[") + 1,InStr(strSQL,"]") - InStr(strSQL,"[") - 1)
						'strValue = Execute("Response.Write " & strField & ".value")
						For count = 0 To fieldCount
							If fieldArray(count) = strField Then
								strValue = valueArray(count)
								Exit For
							End If
						Next
						If Not IsNull(strValue) Then
							strSQL = Replace(strSQL,"[" & strField & "]",strValue)
						Else
							strSQL = Replace(strSQL,"[" & strField & "]","null")
						End If
					Loop
					'Evaluate the resulting string.
		'			Response.Write "result = " & Eval(strSQL)
					Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("label_top") & "px;left:" & objRecordset("label_left") & "px;width:" & objRecordset("label_width") & "px;height:" & objRecordset("label_height") & "px;font-size:" & objRecordset("label_font_size") & ";font-weight:" & objRecordset("label_font_weight") & ";text-align:" & objRecordset("label_align") & ";line-height:" & lineHeight & "px;padding-left:2px;z-index:2;overflow:hidden'>" & Eval(strSQL) & "</div>"
				End If
				If Not IsNull(objRecordset("field_table")) And Not IsNull(objRecordset("field_field")) Then
					strSQL = "SELECT " & objRecordset("field_field") & " FROM " & objRecordset("field_table") & " WHERE instr_id=" & instrumentID
					Set objRecordset2 = objConnection.Execute(strSQL)
					If Not objRecordset2.BOF Then
						objRecordset2.MoveFirst
						fieldCount = fieldCount + 1
						ReDim Preserve fieldArray(fieldCount)
						ReDim Preserve valueArray(fieldCount)
						fieldArray(fieldCount) = objRecordset("field_field")
						If Request(fieldFormatID) <> "" Then
							valueArray(fieldCount) = Request(fieldFormatID)
						Else
							valueArray(fieldCount) = objRecordset2(0)
						End If
		'				Response.Write objRecordset("field_field") & " = " & objRecordset2(0)
						If objRecordset("field_type") = "Textbox" Then
							If Request(fieldFormatID) <> "" Then
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='text' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='" & Request(fieldFormatID) & "' size='" & objRecordset("field_length") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";height:" & objRecordset("field_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
							Else
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='text' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='" & objRecordset2(0) & "' size='" & objRecordset("field_length") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";height:" & objRecordset("field_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
							End If
						ElseIf objRecordset("field_type") = "Textarea" Then
							If Request(fieldFormatID) <> "" Then
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><textarea id='" & fieldFormatID & "' name='" & fieldFormatID & "' cols='" & objRecordset("field_length") & "' rows='" & objRecordset("field_rows") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & "' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);'>" & Request(fieldFormatID) & "</textarea></div>"
							Else
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><textarea id='" & fieldFormatID & "' name='" & fieldFormatID & "' cols='" & objRecordset("field_length") & "' rows='" & objRecordset("field_rows") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & "' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);'>" & objRecordset2(0) & "</textarea></div>"
							End If
						ElseIf objRecordset("field_type") = "Dropdown" Then
							If objRecordset("reload") = True Then
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><select style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";padding-top:0px' id='" & fieldFormatID & "' name='" & fieldFormatID & "' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);reloadpage(" & instrumentID & "," & page_num & ",""" & fieldFormatID & """);'>"
							Else
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><select style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";padding-top:0px' id='" & fieldFormatID & "' name='" & fieldFormatID & "' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);'>"
							End If
							Response.Write "<option value='null'>"
							If Not IsNull(objRecordset("field_source_list")) Then
								strArray = Split(objRecordset("field_source_list"),",")
								For Each strValue In strArray
									If Request(fieldFormatID) <> "" Then
										If Trim(strValue) = Trim(Request(fieldFormatID)) Then
											Response.Write "<option value='" & strValue & "' selected>" & strValue
										Else
											Response.Write "<option value='" & strValue & "'>" & strValue
										End If
									Else
										If Trim(strValue) = Trim(objRecordset2(0)) Then
											Response.Write "<option value='" & strValue & "' selected>" & strValue
										Else
											Response.Write "<option value='" & strValue & "'>" & strValue
										End If
									End If
								Next
							ElseIf Not IsNull(objRecordset("field_source_sql")) Then
								If InStr(objRecordset("field_source_sql"),"[") > 0 Then
									'Extract the field name from the sql string.
									strField = Mid(objRecordset("field_source_sql"),InStr(objRecordset("field_source_sql"),"[") + 1,InStr(objRecordset("field_source_sql"),"]") - InStr(objRecordset("field_source_sql"),"[") - 1)
									'strValue = Execute("Response.Write " & strField & ".value")
									strValue = ""
									For count = 0 To fieldCount
										If fieldArray(count) = strField Then
											strValue = valueArray(count)
											Exit For
										End If
									Next
									If strValue <> "" Then
										strSQL = Replace(objRecordset("field_source_sql"),"[" & strField & "]",strValue)
										Set objRecordset3 = objConnection.Execute(strSQL)
										If Not objRecordset3.BOF Then
											objRecordset3.MoveFirst
											Do While Not objRecordset3.EOF
												'If the source SQL contains a comma, there are 2 fields in the
												'query so set the first as the value and display the second.
												'Otherwise set the value and display the first.
												If InStr(objRecordset("field_source_sql"),",") > 0 Then
													If Request(fieldFormatID) <> "" Then
														If Trim(objRecordset3(0)) = Trim(Request(fieldFormatID)) Then
															Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(1)
														Else
															Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(1)
														End If
													Else
														If Trim(objRecordset3(0)) = Trim(objRecordset2(0)) Then
															Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(1)
														Else
															Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(1)
														End If
													End If
												Else
													If Request(fieldFormatID) <> "" Then
														If Trim(objRecordset3(0)) = Trim(Request(fieldFormatID)) Then
															Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(0)
														Else
															Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(0)
														End If
													Else
														If Trim(objRecordset3(0)) = Trim(objRecordset2(0)) Then
															Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(0)
														Else
															Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(0)
														End If
													End If
												End If
												objRecordset3.MoveNext
											Loop
											objRecordset3.Close
										End If
									End If
								Else
									Set objRecordset3 = objConnection.Execute(objRecordset("field_source_sql"))
									If Not objRecordset3.BOF Then
										objRecordset3.MoveFirst
										Do While Not objRecordset3.EOF
											'If the source SQL contains a comma, there are 2 fields in the
											'query so set the first as the value and display the second.
											'Otherwise set the value and display the first.
											If InStr(objRecordset("field_source_sql"),",") > 0 And InStr(objRecordset("field_source_sql"),"CONCAT") <= 0 Then
												If Request(fieldFormatID) <> "" Then
													If Trim(objRecordset3(0)) = Trim(Request(fieldFormatID)) Then
														Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(1)
													Else
														Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(1)
													End If
												Else
													If Trim(objRecordset3(0)) = Trim(objRecordset2(0)) Then
														Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(1)
													Else
														Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(1)
													End If
												End If
											Else
												If Request(fieldFormatID) <> "" Then
													If Trim(objRecordset3(0)) = Trim(Request(fieldFormatID)) Then
														Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(0)
													Else
														Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(0)
													End If
												Else
													If Trim(objRecordset3(0)) = Trim(objRecordset2(0)) Then
														Response.Write "<option value='" & objRecordset3(0) & "' selected>" & objRecordset3(0)
													Else
														Response.Write "<option value='" & objRecordset3(0) & "'>" & objRecordset3(0)
													End If
												End If
											End If
											objRecordset3.MoveNext
										Loop
									End If
									objRecordset3.Close
								End If
							End If
							Response.Write "</select></div>"
						ElseIf objRecordset("field_type") = "Checkbox" Then
							If Request(fieldFormatID) <> "" Then
								If Request(fieldFormatID) = "1" Then
									Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='checkbox' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='1' style='height:" & objRecordset("field_height") & "px' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' checked /></div>"
								Else
									Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='checkbox' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='1' style='height:" & objRecordset("field_height") & "px' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
								End If
							Else
								If objRecordset2(0) = "1" Then
									Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='checkbox' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='1' style='height:" & objRecordset("field_height") & "px' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' checked /></div>"
								Else
									Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='checkbox' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='1' style='height:" & objRecordset("field_height") & "px' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
								End If
							End If
						Else
							'Add the special case for users with "delete" privilege
							'to allow them to change the instrument name and the
							'equipment.
							If access = "delete" Then
								If objRecordset("field_field") = "instr_name" Then
									If Request(fieldFormatID) <> "" Then
										Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='text' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='" & Request(fieldFormatID) & "' size='" & objRecordset("field_length") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";height:" & objRecordset("field_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
									Else
										Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;z-index:2'><input type='text' id='" & fieldFormatID & "' name='" & fieldFormatID & "' value='" & objRecordset2(0) & "' size='" & objRecordset("field_length") & "' style='font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";height:" & objRecordset("field_height") & "px;padding-top:0px;vertical-align:top' tabindex='" & objRecordset("field_tabindex") & "' onchange='saveChange(this.name);' /></div>"
									End If
								End If
							Else
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;padding-left:2px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
							End If
						End If
		'				fieldCount = fieldCount + 1
					End If
					objRecordset2.Close
				ElseIf Not IsNull(objRecordset("field_field")) Then
					'If the field is "@" then write a blank label with a border.
					If objRecordset("field_field") = "@" Then
						Response.Write "<div style='border-style:solid;border-width:1px;border-color:black;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";z-index:1'></div>"
					ElseIf InStr(objRecordset("field_field"),"[") > 0 Then
						strSQL = objRecordset("field_field")
						Do While InStr(strSQL,"[") > 0
							'Extract the field name from the string.
							strField = Mid(strSQL,InStr(strSQL,"[") + 1,InStr(strSQL,"]") - InStr(strSQL,"[") - 1)
							'strValue = Execute("Response.Write " & strField & ".value")
							For count = 0 To fieldCount
								If fieldArray(count) = strField Then
									strValue = valueArray(count)
									Exit For
								End If
							Next
							If Not IsNull(strValue) Then
								strSQL = Replace(strSQL,"[" & strField & "]",strValue)
							Else
								strSQL = Replace(strSQL,"[" & strField & "]","null")
							End If
						Loop
						'Evaluate the resulting string.
			'			Response.Write "result = " & Eval(strSQL)
						Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;padding-left:5px;z-index:2;overflow:hidden'>" & Eval(strSQL) & "</div>"
					Else
						Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;padding-left:5px;z-index:2;overflow:hidden'>" & objRecordset("field_field") & "</div>"
					End If
				ElseIf objRecordset("field_type") = "Lookup" Then
					If Not IsNull(objRecordset("field_source_sql")) Then
						strSQL = objRecordset("field_source_sql")
						missingValue = False
						Do While InStr(strSQL,"[") > 0
							'Extract the field name from the string.
							strField = Mid(strSQL,InStr(strSQL,"[") + 1,InStr(strSQL,"]") - InStr(strSQL,"[") - 1)
							'strValue = Execute("Response.Write " & strField & ".value")
							strValue = ""
							For count = 0 To fieldCount
								If fieldArray(count) = strField Then
									strValue = valueArray(count)
									Exit For
								End If
							Next
							If strValue <> "" Then
								strSQL = Replace(strSQL,"[" & strField & "]",strValue)
							Else
								missingValue = True
								Exit Do
							End If
						Loop
						If Not missingValue Then
							Set objRecordset2 = objConnection.Execute(strSQL)
							If Not objRecordset2.BOF Then
								objRecordset2.MoveFirst
								Response.Write "<div style='border-style:none;position:absolute;top:" & objRecordset("field_top") & "px;left:" & objRecordset("field_left") & "px;width:" & objRecordset("field_width") & "px;height:" & objRecordset("field_height") & "px;font-size:" & objRecordset("field_font_size") & ";font-weight:" & objRecordset("field_font_weight") & ";line-height:" & objRecordset("field_height") & "px;padding-left:3px;z-index:2;overflow:hidden'>" & objRecordset2(0) & "</div>"
							End If
						End If
					End If
				End If
		'		Response.Write "<br />"
				objRecordset.MoveNext
			Loop
			objRecordset.Close
		'	For count = 0 To UBound(fieldArray)
		'		Response.Write "field = " & fieldArray(count) & " - value = " & valueArray(count) & "<br />"
		'	Next
		Else
			Response.Write "No records found"
		End If

	'If a spec form type has not already been specified for this instrument, redirect
	'to another form to select one.
	Else
		Response.Redirect "SelectForm.asp?instrID=" & instrumentID & "&page=editspec"
	End If
Else
	response.write "<h1>You don't have permission to access this page.</h1>"
	response.write "<br />"
	response.write "<a href='" & request("http_referer") & "'>Return to previous page</a>"
End If

Set objRecordset = Nothing
Set objRecordset2 = Nothing
Set objRecordset3 = Nothing
objConnection.Close
Set objConnection = Nothing
%>
</form>
</body>
</html>
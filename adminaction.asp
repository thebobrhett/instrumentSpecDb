<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<meta http-equiv='Expires' content='0'></meta>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>Instrument Spec Admin Action</title>
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
<!--#include file="EquipmentFunctions.asp"-->
<%
Function GetInsertFields(fieldNames,fieldVals)
	Dim insertFields
	Dim count
	insertFields = ""
	For count = 0 To UBound(fieldNames)
		If fieldVals(count) <> "" Then
			If insertFields = "" Then
				insertFields = fieldNames(count)
			Else
				insertFields = insertFields & "," & fieldNames(count)
			End If
		End If
	Next
	GetInsertFields = insertFields
End Function

Function GetInsertVals(fieldTypes,fieldVals)
	Dim insertVals
	Dim count
	insertVals = ""
	For count = 0 To UBound(fieldNames)
		If insertVals = "" Then
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					insertVals = "'" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					insertVals = "'" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					insertVals = fieldVals(count)
				End If
			End If
		Else
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					insertVals = insertVals & ",'" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					insertVals = insertVals & ",'" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					insertVals = insertVals & "," & fieldVals(count)
				End If
			End If
		End If
	Next
	GetInsertVals = insertVals
End Function

Function GetUpdateString(fieldNames,fieldVals,fieldTypes)
	Dim updateString
	Dim count
	updateString = ""
	For count = 0 To UBound(fieldNames)
		If updateString = "" Then
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					updateString = fieldNames(count) & "='" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					updateString = fieldNames(count) & "='" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					updateString = fieldNames(count) & "=" & fieldVals(count)
				End If
			Else
				updateString = fieldNames(count) & "=null"
			End If
		Else
			If fieldVals(count) <> "" Then
				If fieldTypes(count) = "Text" Then
					updateString = updateString & "," & fieldNames(count) & "='" & fieldVals(count) & "'"
				ElseIf fieldTypes(count) = "Date" Then
					updateString = updateString & "," & fieldNames(count) & "='" & FormatMySQLDateTime(fieldVals(count)) & "'"
				Else
					updateString = updateString & "," & fieldNames(count) & "=" & fieldVals(count)
				End If
			Else
				updateString = updateString & "," & fieldNames(count) & "=null"
			End If
		End If
	Next
	GetUpdateString = updateString
End Function

Function WriteDeleteQuery(tableName,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteDeleteQuery = "DELETE FROM " & tableName & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteDeleteQuery = "DELETE FROM " & tableName & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteInsertQuery(tableName,insertFields,insertVals)
	WriteInsertQuery = "INSERT INTO " & tableName & " (" & insertFields & ") VALUES (" & insertVals & ")"
End Function

Function WriteSelectQuery(fieldName,tableName,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteSelectQuery = "SELECT " & fieldName & " FROM " & tableName & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteSelectQuery = "SELECT " & fieldName & " FROM " & tableName & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteUpdateQuery(tableName,updateString,idName)
	If IsNumeric(Request("RECORD")) Then
		WriteUpdateQuery = "UPDATE " & tableName & " SET " & updateString & " WHERE " & idName & "=" & Request("RECORD")
	Else
		WriteUpdateQuery = "UPDATE " & tableName & " SET " & updateString & " WHERE " & idName & "='" & Request("RECORD") & "'"
	End If
End Function

Function WriteAuditInsertQuery(user,idVal,tableName,fieldName,oldVal,newVal,auditType)
	WriteAuditInsertQuery = "INSERT INTO admin_audit_trail (change_user,change_table_id,change_table,change_field,old_value,new_value,change_type) " & _
					"VALUES ('" & user & "','" & idVal & "','" & tableName & "','" & fieldName & "','" & oldVal & "','" & newVal & "','" & auditType & "')"
End Function

Function ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
	Dim cn
	Dim rs
	Dim status
	On Error Resume Next
	status = True
	set cn = CreateObject("adodb.connection")
	cn.Open = DBString
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs = CreateObject("adodb.recordset")
	If Request.QueryString("action") = "delete" Then
		'Delete the record.
		Set rs = cn.Execute(WriteDeleteQuery(tableName,idName))
		If Err.number <> 0 Then
			Session("err") = "db"
			Exit Function
		End If
		'Write the change to the audit trail table.
		Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,idName,Request("RECORD"),"null","delete"))
		If Err.number <> 0 Then
			status = False
			Exit Function
		End If
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = ""
		Next
	Else
		'If the record number < 0, insert a record; otherwise, update the specified record.
		If request("RECORD") <> "" Then
			If request("RECORD") = -1 Or Request("RECORD") = "-1" Then
'				If 1=2 Then
				Set rs = cn.Execute(WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals)))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				'Get the id number that was just assigned.
				sqlString = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sqlString)
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				If Not rs.BOF Then
					tableID = rs(0)
				Else
					tableID = 0
				End If
				rs.Close
				'If the ID is not an autoincrement, use the first field value.
				If tableID = 0 Then
					tableID = fieldVals(0)
				End If
				'Write the changes to the audit trail.
				Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,idName,"null",tableID,"insert"))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,fieldNames(count),"null",fieldVals(count),"insert"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
'				Else
'					Response.Write "sqlString = " & WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals))
'				End If
			Else
				'Get the existing field values for the audit trail.
				ReDim oldVals(UBound(fieldNames))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteSelectQuery(fieldNames(count),tableName,idName))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							oldVals(count) = rs(0)
						Else
							oldVals(count) = "null"
						End If
					Else
						oldVals(count) = "null"
					End If
					rs.Close
				Next
				
				'Update the record.
				Set rs = cn.Execute(WriteUpdateQuery(tableName,GetUpdateString(fieldNames,fieldVals,fieldTypes),idName))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If

				'Write the changes to the audit trail table.
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,fieldNames(count),oldVals(count),fieldVals(count),"update"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			End If
				
		End If
		Session.Contents.RemoveAll
'		For count = 0 To UBound(fieldNames)
'			Session(fieldNames(count)) = ""
'		Next
	End If
	Set rs = Nothing
	cn.Close
	Set cn = Nothing
	On Error Goto 0
	ProcessChange = status
End Function

Function ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
	Dim cn
	Dim cn2
	Dim rs
	Dim rs2
	Dim status
	On Error Resume Next
	status = True
	set cn = CreateObject("adodb.connection")
	cn.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;User=assetmgtuser;password=asset;DATABASE=asset_management;"
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs = CreateObject("adodb.recordset")
	set cn2 = CreateObject("adodb.connection")
	cn2.Open = DBString
	If Err.number <> 0 Then
		status = False
		Exit Function
	End If
	set rs2 = CreateObject("adodb.recordset")
	If Request.QueryString("action") = "delete" Then
		'Delete the record.
		Set rs = cn.Execute(WriteDeleteQuery(tableName,idName))
		If Err.number <> 0 Then
			Session("err") = "db"
			Exit Function
		End If
		'Write the change to the audit trail table.
		Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,idName,Request("RECORD"),"null","delete"))
		If Err.number <> 0 Then
			status = False
			Exit Function
		End If
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = ""
		Next
	Else
		'If the record number < 0, insert a record; otherwise, update the specified record.
		If request("RECORD") <> "" Then
			If request("RECORD") = -1 Or Request("RECORD") = "-1" Then
				Set rs = cn.Execute(WriteInsertQuery(tableName,GetInsertFields(fieldNames,fieldVals),GetInsertVals(fieldTypes,fieldVals)))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				'Get the id number that was just assigned.
				sqlString = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sqlString)
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If
				If Not rs.BOF Then
					tableID = rs(0)
				Else
					tableID = 0
				End If
				rs.Close
				'If the ID is not an autoincrement, use the first field value.
				If tableID = 0 Then
					tableID = fieldVals(0)
				End If
				'Write the changes to the audit trail.
				Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,idName,"null",tableID,"insert"))
				For count = 0 To UBound(fieldNames)
					Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,tableID,tableName,fieldNames(count),"null",fieldVals(count),"insert"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			Else
				'Get the existing field values for the audit trail.
				ReDim oldVals(UBound(fieldNames))
				For count = 0 To UBound(fieldNames)
					Set rs = cn.Execute(WriteSelectQuery(fieldNames(count),tableName,idName))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							oldVals(count) = rs(0)
						Else
							oldVals(count) = "null"
						End If
					Else
						oldVals(count) = "null"
					End If
					rs.Close
				Next
				
				'Update the record.
				Set rs = cn.Execute(WriteUpdateQuery(tableName,GetUpdateString(fieldNames,fieldVals,fieldTypes),idName))
				If Err.number <> 0 Then
					status = False
					Exit Function
				End If

				'Write the changes to the audit trail table.
				For count = 0 To UBound(fieldNames)
					Set rs2 = cn2.Execute(WriteAuditInsertQuery(currentuser,Request("RECORD"),tableName,fieldNames(count),oldVals(count),fieldVals(count),"update"))
					If Err.number <> 0 Then
						status = False
						Exit Function
					End If
				Next
			End If
				
		End If
		Session.Contents.RemoveAll
'		For count = 1 To UBound(fieldNames)
'			Session(fieldNames(count)) = ""
'		Next
	End If
	Set rs = Nothing
	Set rs2 = Nothing
	cn.Close
	cn2.Close
	Set cn = Nothing
	Set cn2 = Nothing
	On Error Goto 0
	ProcessUserChange = status
End Function
%>
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Thursday, April 22, 2010
'   Creation
'*************

'on error resume next
'Dim cn
'Dim cn2
'Dim rs
Dim fieldVals()
Dim fieldNames()
Dim fieldTypes()
Dim oldVals()
Dim oldVal
Dim sqlString
Dim reload
Dim samePage
Dim NewID
Dim currentuser
Dim count
Dim insertFields
Dim insertVals
Dim updateString
Dim tableID
Dim tableName
Dim idName
Dim status

reload = "NONE"
session("err") = "NONE"
samePage = False
NewID = 0

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

If InStr(request("http_referer"),"useraccess") > 0 Then
	tableName = "application_permissions"
	idName = "permission_id"
	ReDim fieldNames(6)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "application_name"
	fieldNames(1) = "user_name"
	fieldNames(2) = "form_name"
	fieldNames(3) = "role_id"
	fieldNames(4) = "write_access"
	fieldNames(5) = "delete_access"
	fieldNames(6) = "disabled"
	For count = 0 To UBound(fieldNames)
		If fieldNames(count) = "application_name" Then
			fieldVals(count) = "equipment"
		ElseIf fieldNames(count) = "user_name" Or fieldNames(count) = "form_name" Or fieldNames(count) = "role_id" Then
			fieldVals(count) = Request(fieldNames(count))
		Else
			If Request(fieldNames(count)) <> "" Then
				fieldVals(count) = 1
			Else
				fieldVals(count) = 0
			End If
		End If
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	fieldTypes(4) = "Number"
	fieldTypes(5) = "Number"
	fieldTypes(6) = "Number"

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("form_name") = "" Then
			session("err") = "form_name"
			session("saveval") = request("form_name")
		ElseIf Request("role_id") = "" And request("user_name") = "" Then
			session("err") = "role_id"
			session("saveval") = request("role_id")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 1 To UBound(fieldNames)
				If count > 2 And fieldVals(count) = "0" Then
'					Session(fieldNames(count)) = ""
					Session.Contents.Remove(FieldNames(count))
				Else
					Session(fieldNames(count)) = fieldVals(count)
				End If
			Next
		End If
	Else
		For count = 1 To UBound(fieldNames)
			If count > 2 And fieldVals(count) = "0" Then
'				Session(fieldNames(count)) = ""
				Session.Contents.Remove(FieldNames(count))
			Else
				Session(fieldNames(count)) = fieldVals(count)
			End If
		Next
	End If
	
ElseIf InStr(request("http_referer"),"rolemembers") > 0 Then
	tableName = "application_role_members"
	idName = "role_member_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "application_role_id"
	fieldNames(1) = "user_name"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"

	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("application_role_id") = "" Then
			session("err") = "application_role_id"
			session("saveval") = request("application_role_id")
		ElseIf Request("user_name") = "" Then
			session("err") = "user_name"
			session("saveval") = request("user_name")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessUserChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 1 To UBound(fieldNames)
				If count > 2 And fieldVals(count) = "0" Then
'					Session(fieldNames(count)) = ""
					Session.Contents.Remove(FieldNames(count))
				Else
					Session(fieldNames(count)) = fieldVals(count)
				End If
			Next
		End If
	Else
		For count = 1 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If
	
ElseIf InStr(request("http_referer"),"classifications") > 0 Then
	'Load arrays of field names and values.
	tableName = "classifications"
	idName = "classification_id"
	ReDim fieldNames(0)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "classification_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("classification_desc") = "" Then
			session("err") = "classification_desc"
			session("saveval") = request("classification_desc")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"cvtypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "control_valve_types"
	idName = "cv_type_id"
	ReDim fieldNames(0)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "cv_type_name"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("cv_type_name") = "" Then
			session("err") = "cv_type_name"
			session("saveval") = request("cv_type_name")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"drawings") > 0 Then
	'Load arrays of field names and values.
	tableName = "drawings"
	idName = "dwg_id"
	ReDim fieldNames(4)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "dwg_file_name"
	fieldNames(1) = "plant_area_id"
	fieldNames(2) = "dwg_name"
	fieldNames(3) = "dwg_desc"
	fieldNames(4) = "dwg_type_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Number"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Text"
	fieldTypes(4) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_id") = "" Then
			session("err") = "plant_area_id"
			session("saveval") = request("plant_area_id")
		ElseIf request("dwg_type_id") = "" Then
			session("err") = "dwg_type_id"
			session("saveval") = request("dwg_type_id")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"drawingtypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "drawing_types"
	idName = "dwg_type_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "dwg_type_name"
	fieldNames(1) = "dwg_type_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("dwg_type_name") = "" Then
			session("err") = "dwg_type_name"
			session("saveval") = request("dwg_type_name")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"equipmenttype") > 0 Then
	'Load arrays of field names and values.
	tableName = "equipment_type"
	idName = "equip_type_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "equip_type_name"
	fieldNames(1) = "equip_type_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("equip_type_name") = "" Then
			session("err") = "equip_type_name"
			session("saveval") = request("equip_type_name")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"equipment.asp") > 0 Then
	'Load arrays of field names and values.
	tableName = "equipment"
	idName = "equip_id"
	ReDim fieldNames(3)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "plant_area_id"
	fieldNames(1) = "equip_name"
	fieldNames(2) = "equip_desc"
	fieldNames(3) = "equip_type_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_id") = "" Then
			session("err") = "plant_area_id"
			session("saveval") = request("plant_area_id")
		ElseIf request("equip_name") = "" Then
			session("err") = "equip_name"
			session("saveval") = request("equip_name")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"fmsubtypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "flow_meter_subtypes"
	idName = "fm_subtype_id"
	ReDim fieldNames(2)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "fm_subtype_desc"
	fieldNames(1) = "fm_type_id"
	fieldNames(2) = "fm_subtype_number"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Number"
	fieldTypes(2) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("fm_subtype_desc") = "" Then
			session("err") = "fm_subtype_desc"
			session("saveval") = request("fm_subtype_desc")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"fmtypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "flow_meter_types"
	idName = "fm_type_id"
	ReDim fieldNames(0)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "fm_type_name"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("fm_type_name") = "" Then
			session("err") = "fm_type_name"
			session("saveval") = request("fm_type_name")
		End If
	End If

	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"fluidphases") > 0 Then
	'Load arrays of field names and values.
	tableName = "fluid_phases"
	idName = "fluid_phase_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "fluid_phase_name"
	fieldNames(1) = "fluid_phase_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("fluid_phase_name") = "" Then
			session("err") = "fluid_phase_name"
			session("saveval") = request("fluid_phase_name")
		ElseIf request("fluid_phase_desc") = "" Then
			session("err") = "fluid_phase_desc"
			session("saveval") = request("fluid_phase_desc")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"instrfunctiontypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "instrument_function_types"
	idName = "instr_func_type_id"
	ReDim fieldNames(3)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "instr_func_type_name"
	fieldNames(1) = "instr_func_type_desc"
	fieldNames(2) = "instr_func_type_spec_form_id"
	fieldNames(3) = "proc_func_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Number"
	fieldTypes(3) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("instr_func_type_name") = "" Then
			session("err") = "instr_func_type_name"
			session("saveval") = request("instr_func_type_name")
		ElseIf request("instr_func_type_desc") = "" Then
			session("err") = "instr_func_type_desc"
			session("saveval") = request("instr_func_type_desc")
		ElseIf request("proc_func_id") = "" Then
			session("err") = "proc_func_id"
			session("saveval") = request("proc_func_id")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"instrlocations") > 0 Then
	'Load arrays of field names and values.
	tableName = "instrument_locations"
	idName = "instr_location_id"
	ReDim fieldNames(2)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "plant_area_id"
	fieldNames(1) = "instr_location_name"
	fieldNames(2) = "instr_location_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_id") = "" Then
			session("err") = "plant_area_id"
			session("saveval") = request("plant_area_id")
		ElseIf request("instr_location_name") = "" Then
			session("err") = "instr_location_name"
			session("saveval") = request("instr_location_name")
		End If
	End If
		
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"instrmanufacturers") > 0 Then
	'Load arrays of field names and values.
	tableName = "instrument_manufacturers"
	idName = "instr_mfr_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "instr_mfr_name"
	fieldNames(1) = "instr_mfr_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("instr_mfr_name") = "" Then
			session("err") = "instr_mfr_name"
			session("saveval") = request("instr_mfr_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"instrmodels") > 0 Then
	'Load arrays of field names and values.
	tableName = "instrument_models"
	idName = "instr_mod_id"
	ReDim fieldNames(2)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "instr_mfr_id"
	fieldNames(1) = "instr_mod_name"
	fieldNames(2) = "instr_mod_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("instr_mfr_id") = "" Then
			session("err") = "instr_mfr_id"
			session("saveval") = request("instr_mfr_id")
		ElseIf request("instr_mod_name") = "" Then
			session("err") = "instr_mod_name"
			session("saveval") = request("instr_mod_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"linetypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "line_types"
	idName = "line_type_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "line_type_name"
	fieldNames(1) = "line_type_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("line_type_name") = "" Then
			session("err") = "line_type_name"
			session("saveval") = request("line_type_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"lines.asp") > 0 Then
	'Load arrays of field names and values.
	tableName = "process_lines"
	idName = "line_id"
	ReDim fieldNames(14)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "plant_area_id"
	fieldNames(1) = "line_num"
	fieldNames(2) = "line_size"
	fieldNames(3) = "wall_thick"
	fieldNames(4) = "rtg"
	fieldNames(5) = "line_type_id"
	fieldNames(6) = "line_sched"
	fieldNames(7) = "line_internal_dia"
	fieldNames(8) = "pipe_std_id"
	fieldNames(9) = "ansi_din"
	fieldNames(10) = "pipe_size"
	fieldNames(11) = "pipe_orif_mat_id"
	fieldNames(12) = "stream_num"
	fieldNames(13) = "dwg_id"
	fieldNames(14) = "pipe_class_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	fieldTypes(4) = "Text"
	fieldTypes(5) = "Number"
	fieldTypes(6) = "Text"
	fieldTypes(7) = "Number"
	fieldTypes(8) = "Number"
	fieldTypes(9) = "Text"
	fieldTypes(10) = "Number"
	fieldTypes(11) = "Number"
	fieldTypes(12) = "Text"
	fieldTypes(13) = "Number"
	fieldTypes(14) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_id") = "" Then
			session("err") = "plant_area_id"
			session("saveval") = request("plant_area_id")
		ElseIf request("line_num") = "" Then
			session("err") = "line_num"
			session("saveval") = request("line_num")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"addloops.asp") > 0 Then
	'Load arrays of field names and values.
	tableName = "loops"
	idName = "loop_id"
	ReDim fieldNames(10)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "plant_area_id"
	fieldNames(1) = "loop_name"
	fieldNames(2) = "loop_num"
	fieldNames(3) = "loop_proc_id"
	fieldNames(4) = "loop_type_id"
	fieldNames(5) = "loop_func_id"
	fieldNames(6) = "loop_suff"
	fieldNames(7) = "loop_desc"
	fieldNames(8) = "loop_note"
	fieldNames(9) = "dwg_id"
	fieldNames(10) = "loop_equip_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Number"
	fieldTypes(4) = "Number"
	fieldTypes(5) = "Number"
	fieldTypes(6) = "Text"
	fieldTypes(7) = "Text"
	fieldTypes(8) = "Text"
	fieldTypes(9) = "Number"
	fieldTypes(10) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_id") = "" Then
			session("err") = "plant_area_id"
			session("saveval") = request("plant_area_id")
		ElseIf request("loop_name") = "" Then
			session("err") = "loop_name"
			session("saveval") = request("loop_name")
		ElseIf request("loop_num") = "" Then
			session("err") = "loop_num"
			session("saveval") = request("loop_num")
		ElseIf request("loop_proc_id") = "" Then
			session("err") = "loop_proc_id"
			session("saveval") = request("loop_proc_id")
		End If		
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"loops.asp") > 0 Then
	'Load arrays of field names and values.
	tableName = "loops"
	idName = "loop_id"
	ReDim fieldNames(5)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "loop_type_id"
	fieldNames(1) = "loop_func_id"
	fieldNames(2) = "loop_suff"
	fieldNames(3) = "loop_desc"
	fieldNames(4) = "loop_note"
	fieldNames(5) = "dwg_id"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Number"
	fieldTypes(1) = "Number"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Text"
	fieldTypes(4) = "Text"
	fieldTypes(5) = "Number"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"loopfunctions") > 0 Then
	'Load arrays of field names and values.
	tableName = "loop_functions"
	idName = "loop_func_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "loop_func_name"
	fieldNames(1) = "loop_func_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("loop_func_name") = "" Then
			session("err") = "loop_func_name"
			session("saveval") = request("loop_func_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"loopprocesses") > 0 Then
	'Load arrays of field names and values.
	tableName = "loop_processes"
	idName = "loop_proc_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "loop_proc_name"
	fieldNames(1) = "loop_proc_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("loop_proc_name") = "" Then
			session("err") = "loop_proc_name"
			session("saveval") = request("loop_proc_name")
		ElseIf request("loop_proc_desc") = "" Then
			session("err") = "loop_proc_desc"
			session("saveval") = request("loop_proc_desc")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"looptypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "loop_types"
	idName = "loop_type_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "loop_type_name"
	fieldNames(1) = "loop_type_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("loop_type_name") = "" Then
			session("err") = "loop_type_name"
			session("saveval") = request("loop_type_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"pipeclasses") > 0 Then
	'Load arrays of field names and values.
	tableName = "pipe_classes"
	idName = "pipe_class_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "pipe_class_name"
	fieldNames(1) = "pipe_class_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("pipe_class_name") = "" Then
			session("err") = "pipe_class_name"
			session("saveval") = request("pipe_class_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"pipematerials") > 0 Then
	'Load arrays of field names and values.
	tableName = "pipe_materials"
	idName = "pipe_mat_id"
	ReDim fieldNames(7)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "pipe_mat_name"
	fieldNames(1) = "linear_exp_coef"
	fieldNames(2) = "linear_exp_coef_add"
	fieldNames(3) = "linear_exp_coef_uid"
	fieldNames(4) = "border_temp"
	fieldNames(5) = "temp_min"
	fieldNames(6) = "temp_max"
	fieldNames(7) = "temp_uid"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Number"
	fieldTypes(2) = "Number"
	fieldTypes(3) = "Text"
	fieldTypes(4) = "Number"
	fieldTypes(5) = "Number"
	fieldTypes(6) = "Number"
	fieldTypes(7) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("pipe_mat_name") = "" Then
			session("err") = "pipe_mat_name"
			session("saveval") = request("pipe_mat_name")
		End If
		For count = 0 To UBound(fieldNames)
			If fieldTypes(count) = "Number" And Request(fieldNames(count)) <> "" Then
				If Not IsNumeric(Request(fieldNames(count))) Then
					Session("err") = fieldNames(count)
					Session("saveval") = Request(fieldNames(count))
					Exit For
				End If
			End If
		Next
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"plantareas") > 0 Then
	'Load arrays of field names and values.
	tableName = "plant_areas"
	idName = "plant_area_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "plant_area_name"
	fieldNames(1) = "plant_area_num"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("plant_area_name") = "" Then
			session("err") = "plant_area_name"
			session("saveval") = request("plant_area_name")
		ElseIf request("plant_area_num") = "" Then
			session("err") = "plant_area_num"
			session("saveval") = request("plant_area_num")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"processfunctions") > 0 Then
	'Load arrays of field names and values.
	tableName = "process_functions"
	idName = "proc_func_id"
	ReDim fieldNames(0)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "proc_func_name"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("proc_func_name") = "" Then
			session("err") = "proc_func_name"
			session("saveval") = request("proc_func_name")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"unitsofmeasure") > 0 Then
	'Load arrays of field names and values.
	tableName = "units_of_measure"
	idName = "uom_id"
	ReDim fieldNames(4)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "uom_code"
	fieldNames(1) = "uom_name"
	fieldNames(2) = "uom_desc"
	fieldNames(3) = "uom_type"
	fieldNames(4) = "uom_kind"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	fieldTypes(2) = "Text"
	fieldTypes(3) = "Text"
	fieldTypes(4) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("uom_code") = "" Then
			session("err") = "uom_code"
			session("saveval") = request("uom_code")
		ElseIf request("uom_name") = "" Then
			session("err") = "uom_name"
			session("saveval") = request("uom_name")
		ElseIf request("uom_type") = "" Then
			session("err") = "uom_type"
			session("saveval") = request("uom_type")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

ElseIf InStr(request("http_referer"),"unitofmeasuretypes") > 0 Then
	'Load arrays of field names and values.
	tableName = "unit_of_measure_types"
	idName = "uom_type_id"
	ReDim fieldNames(1)
	ReDim fieldVals(UBound(fieldNames))
	ReDim fieldTypes(UBound(fieldNames))
	fieldNames(0) = "uom_type_id"
	fieldNames(1) = "uom_type_desc"
	For count = 0 To UBound(fieldNames)
		fieldVals(count) = Request(fieldNames(count))
	Next
	fieldTypes(0) = "Text"
	fieldTypes(1) = "Text"
	
	'Validate the data.
	If Request.QueryString("action") <> "delete" Then
		If request("uom_type_id") = "" Then
			session("err") = "uom_type_id"
			session("saveval") = request("uom_type_id")
		ElseIf request("uom_type_desc") = "" Then
			session("err") = "uom_type_desc"
			session("saveval") = request("uom_type_desc")
		End If
	End If
	
	'Execute the change and log to the audit table.
	If session("err") = "" Or session("err") = "NONE" Then
		status = ProcessChange(tableName,idName,fieldNames,fieldVals,fieldTypes)
		If Not status Then
			session("err") = "db"
			For count = 0 To UBound(fieldNames)
				Session(fieldNames(count)) = fieldVals(count)
			Next
		End If
	Else
		For count = 0 To UBound(fieldNames)
			Session(fieldNames(count)) = fieldVals(count)
		Next
	End If

End If

'Pop up a message and return to the calling page.
If session("err") = "db" Then
	Response.Write "<script language='javascript'>alert('A database error occurred during update: " & FixString(Err.Description) & "'); window.location.href='" & Request("http_referer") & "';</script>"
ElseIf session("err") <> "" And session("err") <> "NONE" Then
	Response.Write "<script language='javascript'>alert('One or more errors occurred during update. Field [" & session("err") & "] is a required field or has an invalid value.'); window.location.href='" & Request("http_referer") & "';</script>"
Else
	If samePage = True Then
		Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?record_id=" & NewID & "';</script>"
	ElseIf InStr(Request("http_referer"),"addloops.asp") > 0 Then
		Response.Write "<script language='javascript'>window.location.href='loops.asp?sort=loop_id&direction=DESC';</script>"
	ElseIf InStr(Request("http_referer"),"?") > 0 Then
		If InStr(Request("http_referer"),"sort=") > 0 Then
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "?sort=" & Request("SORT") & "&direction=" & Request("DIRECTION") & "&limit=" & Request("limit") & "';</script>"
		Else
			Response.Write "<script language='javascript'>window.location.href='" & Left(Request("http_referer"),InStr(Request("http_referer"),"?") - 1) & "';</script>"
		End If
	Else
		Response.Write "<script language='javascript'>window.location.href='" & Request("http_referer") & "';</script>"
	End If
End If

'Set rs = Nothing
'cn.Close
'Set cn = Nothing
%>

<body bgcolor='#d0d0d0' link='black' vLink='black'>
</body>
</html>
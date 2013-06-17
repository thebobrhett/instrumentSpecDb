<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<!--#include file="EquipmentFunctions.asp"-->
<html>
<head>
<meta http-equiv='Expires' content='0'></meta>
<meta http-equiv='Content-Type' content='text/html; charset=windows-1252'></meta>
<title>Update Spec</title>
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
</head>

<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, March 30, 2009
'   Creation
' Keith Brooks - Monday, December 6, 2010
'	Changed the "demo" update to also increment the "void_sequence" field if the
'	instrument is being voided (to allow re-use of instrument tags).  Also added
'	"void_sequence" criteria to check for duplicate tag on instrument creation.
'*************

Dim sqlString
Dim cn
Dim rs
Dim cmd
Dim tags
Dim tag
Dim idNum
Dim tableName
Dim fieldName
Dim errorFlag
Dim errorMsg
Dim oldValue
Dim newValue
Dim oldVoid
Dim newVoid
Dim currentuser

On Error Resume Next

errorFlag = False
set cn = CreateObject("adodb.connection")
cn.Open = DBString
If Err.number = 0 Then
	set rs = CreateObject("adodb.recordset")
	Set cmd = CreateObject("adodb.command")
	cmd.ActiveConnection = cn

	'Get the current user.
	currentuser = Request.ServerVariables("LOGON_USER")
	currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

	If InStr(request("http_referer"),"editspec") > 0 Then
		If Request("demo") <> "" Then
			'Get the current value of the void_sequence field and increment it.
			sqlString = "SELECT void_sequence FROM instruments WHERE instr_id=" & Request("InstrumentID")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					oldVoid = rs(0)
				Else
					oldVoid = 0
				End If
			Else
				oldVoid = 0
			End If
			rs.Close
			If Request("demo") = "1" Then
				oldValue = 1
				newValue = 0
				newVoid = oldVoid + 1
			Else
				oldValue = 0
				newValue = 1
				newVoid = 0
			End If
			'Set the "active" flag and "void_sequence" fields for this instrument
			'to reflect whether or not it is voided.
			sqlString = "UPDATE instruments SET active=" & CStr(newValue) & _
						", void_sequence=" & CStr(newVoid) & _
						" WHERE instr_id=" & Request("InstrumentID")
			cn.BeginTrans
			cmd.CommandText = sqlString
			cmd.Execute
			If cn.Errors.Count > 0 Then
				If cn.Errors.Item(0).Number <> 0 Then
					cn.RollbackTrans
					errorFlag = True
					errorMsg = "Transaction rolled back: " & Err.Description
				End If
			End If

			If Not errorFlag Then
				'Write the change to the audit trail table.
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & Request("instrumentID") & ",'instruments','active','" & oldValue & "','" & newValue & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "Transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & Request("instrumentID") & ",'instruments','void_sequence','" & oldVoid & "','" & newVoid & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "Transaction rolled back: " & Err.Description
					End If
				End If
			End If

			Set rs = Nothing
			Set cmd = Nothing
			If errorFlag Then
				Response.Write "An error occurred while saving the spec: " & errorMsg
				cn.Close
				Set cn = Nothing
			Else
				cn.CommitTrans
				cn.Close
				Set cn = Nothing
				Response.Redirect "editspec.asp?instrID=" & Request("instrumentID") & "&page_num=" & Request("pagenum")
			End If

		ElseIf Request("changedtags") <> "" Then
			'Separate the string that holds the tags that have been changed
			'into an array.
			tags = Split(Request("changedtags"),",")
			
			'Begin the database transaction.
			cn.BeginTrans
			
			For Each tag In tags
				idNum = Right(tag, Len(tag) - 1)
				'Get the field information for the tag from the spec_formats table.
				'If the tag name starts with an F, get the info from the field_table and
				'field_field fields.
				If UCase(Left(tag,1)) = "F" Then
					sqlString = "SELECT field_table,field_field FROM spec_formats WHERE spec_format_id=" & idNum
					Set rs = cn.Execute(sqlString)
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							tableName = rs(0)
						Else
							errorFlag = True
							errorMsg = "Null field_table found in spec_format record " & idNum
						End If
						If Not IsNull(rs(1)) Then
							fieldName = rs(1)
						Else
							errorFlag = True
							errorMsg = "Null field_field found in spec_format record " & idNum
						End If
					End If
					rs.Close
				'If the tag name starts with an L, get the info from the label_table and
				'label_field fields.
				ElseIf UCase(Left(tag,1)) = "L" Then
					sqlString = "SELECT label_table,label_field FROM spec_formats WHERE spec_format_id=" & idNum
					Set rs = cn.Execute(sqlString)
					If Not rs.BOF Then
						rs.MoveFirst
						If Not IsNull(rs(0)) Then
							tableName = rs(0)
						Else
							errorFlag = True
							errorMsg = "Null label_table found in spec_format record " & idNum
						End If
						If Not IsNull(rs(1)) Then
							fieldName = rs(1)
						Else
							errorFlag = True
							errorMsg = "Null label_field found in spec_format record " & idNum
						End If
					End If
					rs.Close
				End If
				
				'If no error has occurred, update the field with the new data.
				If Not errorFlag Then
					If Request("instrumentID") <> "" Then
					
						'Read the old value from the database for the audit trail.
						sqlString = "SELECT " & fieldName & " FROM " & tableName & " WHERE instr_id=" & Request("instrumentID")
						Dim rs2
						set rs2 = CreateObject("adodb.recordset")
						Set rs2 = cn.Execute(sqlString)
						If Not rs2.BOF Then
							If Not IsNull(rs2(0)) Then
								oldValue = rs2(0)
							Else
								oldValue = "null"
							End If
						Else
							oldValue = "null"
						End If
						rs2.Close
						Set rs2 = Nothing

						'Find out the data type for the field and create the update string.
						sqlString = "SELECT data_type FROM information_schema.columns WHERE table_schema='instruments' AND table_name='" & tableName & "' AND column_name='" & fieldName & "'"
						Set rs = cn.Execute(sqlString)
						If Not rs.BOF Then
							rs.MoveFirst
'							Response.Write tag & " = " & Request(tag) & " and table = " & tableName & " and field = " & fieldName & " and datatype = " & rs(0) & "<br />"
							If LCase(rs(0)) = "varchar" Then
								If Request(tag) <> "" Then
									sqlString = "UPDATE " & tableName & " SET " & fieldName & "='" & Request(tag) & "' WHERE instr_id=" & Request("instrumentID")
								Else
									sqlString = "UPDATE " & tableName & " SET " & fieldName & "=null WHERE instr_id=" & Request("instrumentID")
								End If
							Else
								If Request(tag) <> "" Then
									sqlString = "UPDATE " & tableName & " SET " & fieldName & "=" & Request(tag) & " WHERE instr_id=" & Request("instrumentID")
								Else
									sqlString = "UPDATE " & tableName & " SET " & fieldName & "=null WHERE instr_id=" & Request("instrumentID")
								End If
							End If
							cmd.CommandText = sqlString
							cmd.Execute
							If cn.Errors.Count > 0 Then
								If cn.Errors.Item(0).Number <> 0 Then
									cn.RollbackTrans
									errorFlag = True
									errorMsg = "Transaction rolled back: " & Err.Description
									Exit For
								End If
							End If
							
							'Write the change to the audit trail table.
							sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
										"VALUES ('" & currentuser & "'," & Request("instrumentID") & ",'" & tableName & "','" & fieldName & "','" & CStr(oldValue) & "','" & Request(tag) & "','update')"
							cmd.CommandText = sqlString
							cmd.Execute
							If cn.Errors.Count > 0 Then
								If cn.Errors.Item(0).Number <> 0 Then
									cn.RollbackTrans
									errorFlag = True
									errorMsg = "Transaction rolled back: " & Err.Description
									Exit For
								End If
							End If

						Else
							errorFlag = True
							errorMsg = "Data type not found for column " & fieldName & " in table " & tableName
						End If
					Else
						errorFlag = True
						errorMsg = "Instrument ID not specified"
					End If
				End If
			Next
			
			Set rs = Nothing
			Set cmd = Nothing
			If errorFlag Then
				Response.Write "<br />"
				Response.Write "<div style='text-align:center'><table style='width:70%'><tr><td style='font-weight:bold'>"
				Response.Write "An error occurred while saving the spec: " & errorMsg
				Response.Write "</td></tr></table></div>"
				cn.Close
				Set cn = Nothing
			Else
				cn.CommitTrans
				cn.Close
				Set cn = Nothing
				Response.Redirect "editspec.asp?instrID=" & Request("instrumentID") & "&page_num=" & Request("pagenum")
			End If
		Else
			Set rs = Nothing
			Set cmd = Nothing
			cn.Close
			Set cn = Nothing
			Response.Redirect "editspec.asp?instrID=" & Request("instrumentID") & "&page_num=" & Request("pagenum")
		End If
	ElseIf InStr(request("http_referer"),"AddInstrument") > 0 Then
		'Added to force spaces to ascii character 32 instead of 160.
		Dim tagname
		tagname = Request("tagnumber")
		tagname = Replace(tagname,Chr(160),Chr(32))

		'Check that this instrument name doesn't already exist.
		sqlString = "SELECT instr_id FROM instruments " & _
					"WHERE instr_name='" & tagname & "' " & _
					"AND void_sequence=0"
		Set rs = cn.Execute(sqlString)
		If rs.BOF Then

			rs.Close
			
			'Start the database transaction.
			cn.BeginTrans
		
			'Insert the new record into the "instruments" table.
			Dim equipid
			If Request("equipment") = "" Then
				equipid = "null"
			Else
				equipid = Request("equipment")
			End If
			Dim loopid
			If Request("loop") = "" Then
				loopid = "null"
			Else
				loopid = Request("loop")
			End If
			Dim procfuncid
			sqlString = "SELECT proc_func_id FROM instrument_function_types " & _
						"WHERE instr_func_type_id=" & Request("instrumenttype")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					procfuncid = rs(0)
				Else
					procfuncid = "null"
				End If
			Else
				procfuncid = "null"
			End If
			rs.Close
			sqlString = "INSERT INTO instruments (plant_area_id,instr_name,instr_num," & _
						"instr_suffix,instr_func_type_id,equip_id,loop_id," & _
						"tag_trans_name,remarks,proc_func_id) VALUES (" & _
						Request("plantarea") & ",'" & _
						tagname & "','" & _
						Request("instrnumber") & "','" & _
						Request("suffix") & "'," & _
						Request("instrumenttype") & "," & _
						equipid & "," & loopid & ",'" & _
						Request("tagnumber") & "','" & _
						Request("remarks") & "'," & _
						procfuncid & ")"
			cmd.CommandText = sqlString
			cmd.Execute
			If cn.Errors.Count > 0 Then
				If cn.Errors.Item(0).Number <> 0 Then
					cn.RollbackTrans
					errorFlag = True
					errorMsg = "Instruments transaction rolled back: " & Err.Description
				End If
			End If
		
			If Not errorFlag Then
				'Get the instr_id that was assigned to the new record.
				Dim instr_id
				sqlString = "SELECT LAST_INSERT_ID()"
				Set rs = cn.Execute(sqlString)
				If Not rs.BOF Then
					instr_id = rs(0)
				Else
					errorFlag = True
					errorMsg = "Unable to determine the new instrument ID."
				End If
				rs.Close
				
				'Write the change to the audit trail table.
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','instr_id',null,null,'insert')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','plant_area_id','null','" & Request("plantarea") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','instr_name','null','" & Request("tagnumber") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','instr_num','null','" & Request("instrnumber") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','instr_suffix','null','" & Request("suffix") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','instr_func_type_id','null','" & Request("instrumenttype") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','equip_id','null','" & equipid & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','loop_id','null','" & loopid & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','remarks','null','" & Request("remarks") & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If
				sqlString = "INSERT INTO audit_trail (change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
							"VALUES ('" & currentuser & "'," & instr_id & ",'instruments','proc_func_id','null','" & procfuncid & "','update')"
				cmd.CommandText = sqlString
				cmd.Execute
				If cn.Errors.Count > 0 Then
					If cn.Errors.Item(0).Number <> 0 Then
						cn.RollbackTrans
						errorFlag = True
						errorMsg = "audit transaction rolled back: " & Err.Description
					End If
				End If

				'Add records for this instrument in the custom spec tables.
				If Not errorFlag Then
					sqlString = "INSERT INTO pd_general (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "pd_general transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO spec_sheet_data (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "spec_sheet_data transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec1 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec1 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec2 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec2 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec3 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec3 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec4 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec4 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec5 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec5 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec6 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec6 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec7 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec7 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO add_spec8 (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "add_spec8 transaction rolled back: " & Err.Description
						End If
					End If
				End If
				If Not errorFlag Then
					sqlString = "INSERT INTO udf_instruments (instr_id) VALUES (" & instr_id & ")"
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "udf_instruments transaction rolled back: " & Err.Description
						End If
					End If
				End If
				
				'Add a record for this instrument to type-specific tables based on
				'the process function type.
				If Not errorFlag Then
					Select Case procfuncid
						Case 1
							sqlString = "INSERT INTO flow_meters (instr_id) VALUES (" & instr_id & ")"
						Case 2
							sqlString = "INSERT INTO level_meters (instr_id) VALUES (" & instr_id & ")"
						Case 3
							sqlString = "INSERT INTO pressure_meters (instr_id) VALUES (" & instr_id & ")"
						Case 4
							sqlString = "INSERT INTO temperature_meters (instr_id) VALUES (" & instr_id & ")"
						Case 5
							sqlString = "INSERT INTO analyzers (instr_id) VALUES (" & instr_id & ")"
						Case 6
							sqlString = "INSERT INTO control_valves (instr_id) VALUES (" & instr_id & ")"
						Case 7
							sqlString = "INSERT INTO relief_valves (instr_id) VALUES (" & instr_id & ")"
						Case Else
							sqlString = ""
					End Select

					If sqlString <> "" Then
						cmd.CommandText = sqlString
						cmd.Execute
						If cn.Errors.Count > 0 Then
							If cn.Errors.Item(0).Number <> 0 Then
								cn.RollbackTrans
								errorFlag = True
								errorMsg = "analyzers transaction rolled back: " & Err.Description
							End If
						End If
					End If
				End If
			End If

			Set rs = Nothing
			Set cmd = Nothing
			If errorFlag Then
				Response.Write "<br />"
				Response.Write "<div style='text-align:center'><table style='width:70%'><tr><td style='font-weight:bold'>"
				Response.Write "An error occurred while adding the instrument: " & errorMsg
				Response.Write "</td></tr></table></div>"
				cn.Close
				Set cn = Nothing
			Else
				cn.CommitTrans
				cn.Close
				Set cn = Nothing
				If Request("copy") <> "" Then
				'Open the form to allow the user to select an instrument spec to copy.
					Response.Redirect "copyinstrument.asp?instrID=" & instr_id
				Else
					'Open the form to allow the user to select a spec form for the new
					'instrument.
					Response.Redirect "selectform.asp?instrID=" & instr_id & "&page=editspec"
				End If
			End If
		Else
			rs.Close
			Set rs = Nothing
			Set cmd = Nothing
			cn.Close
			Set cn = Nothing
			Response.Write "<br />"
			Response.Write "<div style='text-align:center'><table style='width:70%'><tr><td style='font-weight:bold'>"
			Response.Write "An error occurred while adding the instrument: The specified instrument name already exists."
			Response.Write "</td></tr><tr><td>Use your browser's 'Back' button to return to the previous page to correct the problem."
			Response.Write "</td></tr></table></div>"
		End If

	ElseIf InStr(request("http_referer"),"CopyInstrument") > 0 Then
		'Read the spec configuration for the instrument to copy.
		If Request("copyid") <> "" And Request("newid") <> "" Then
			'First get the spec form id for the copy instrument.
			sqlString = "SELECT spec_id FROM instruments WHERE instr_id=" & Request("copyid")
			Set rs = cn.Execute(sqlString)
			If Not rs.BOF Then
				rs.MoveFirst
				If Not IsNull(rs(0)) Then
					idNum = rs(0)
					rs.Close
					cn.BeginTrans
					'Update the new instrument with the spec id.
					sqlString = "UPDATE instruments SET spec_id=" & idNum & " WHERE instr_id=" & Request("newid")
					cmd.CommandText = sqlString
					cmd.Execute
					If cn.Errors.Count > 0 Then
						If cn.Errors.Item(0).Number <> 0 Then
							cn.RollbackTrans
							errorFlag = True
							errorMsg = "Copy transaction rolled back: " & Err.Description
						End If
					End If
					
					'Next get the page ids for this spec form.
					sqlString = "SELECT spec_page_id FROM spec_form_pages WHERE spec_form_id=" & idNum & " ORDER BY spec_page_seq"
					Set rs = cn.Execute(sqlString)
					If Not rs.BOF Then
						rs.MoveFirst
						Dim noteCopied
						noteCopied = False
						Dim copyOk
						'Loop through the pages for this spec and copy the fields
						'specified in the spec format from the copy instrument to
						'the new one.
						Do While Not rs.EOF
							If Not IsNull(rs(0)) Then
								sqlString = "SELECT * FROM spec_formats WHERE spec_page_id=" & rs(0)
								Dim rs3
								Dim rs4
								set rs3 = CreateObject("adodb.recordset")
								set rs4 = CreateObject("adodb.recordset")
								Set rs3 = cn.Execute(sqlString)
								If Not rs3.BOF Then
									rs3.MoveFirst
									Do While Not rs3.EOF
										'Copy label database fields.
										If Not IsNull(rs3("label_table")) And Not IsNull(rs3("label_field")) Then
										'	sqlString = "UPDATE " & rs3("label_table") & " SET " & rs3("label_field") & "=" & _
										'		"(SELECT selected_value FROM " & _
										'		"(SELECT " & rs3("label_field") & " AS selected_value FROM " & rs3("label_table") & _
										'		" WHERE instr_id=" & Request("copyid") & ") AS sub_selected_value) " & _
										'		"WHERE instr_id=" & Request("newid")
											'Get the value to copy.
											sqlString = "SELECT " & rs3("label_field") & _
														" FROM " & rs3("label_table") & _
														" WHERE instr_id=" & Request("copyid")
											Set rs4 = cn.Execute(sqlString)
											If Not rs4.BOF Then
												rs4.MoveFirst
												If Not IsNull(rs4(0)) Then
													newValue = rs4(0)
												Else
													newValue = "null"
												End If
											Else
												newValue = "null"
											End If
											rs4.Close
											
											'Do not copy the field if the value is null.
											If newValue <> "null" Then
												'Find out the data type for the field and create the update string.
												sqlString = "SELECT data_type FROM information_schema.columns WHERE table_schema='instruments' AND table_name='" & rs3("label_table") & "' AND column_name='" & rs3("label_field") & "'"
												Set rs4 = cn.Execute(sqlString)
												If Not rs.BOF Then
													rs4.MoveFirst
													If LCase(rs4(0)) = "varchar" And newValue <> "null" Then
														sqlString = "UPDATE " & rs3("label_table") & " SET " & rs3("label_field") & "='" & newValue & "' WHERE instr_id=" & Request("newid")
													Else
														sqlString = "UPDATE " & rs3("label_table") & " SET " & rs3("label_field") & "=" & newValue & " WHERE instr_id=" & Request("newid")
													End If
												Else
													sqlString = "UPDATE " & rs3("label_table") & " SET " & rs3("label_field") & "=" & newValue & " WHERE instr_id=" & Request("newid")
												End If
												rs4.close
												cmd.CommandText = sqlString
												cmd.Execute
												If cn.Errors.Count > 0 Then
													If cn.Errors.Item(0).Number <> 0 Then
														cn.RollbackTrans
														errorFlag = True
														errorMsg = "Copy transaction rolled back: " & Err.Description
														Exit Do
													End If
												End If
											
												'Write the change to the audit trail table.
												sqlString = "INSERT INTO audit_trail " & _
															"(change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
															"VALUES ('" & currentuser & "'," & Request("newid") & ",'" & rs3("label_table") & "','" & rs3("label_field") & "','null','" & newValue & "','update')"
												cmd.CommandText = sqlString
												cmd.Execute
												If cn.Errors.Count > 0 Then
													If cn.Errors.Item(0).Number <> 0 Then
														cn.RollbackTrans
														errorFlag = True
														errorMsg = "Audit transaction rolled back: " & Err.Description
														Exit Do
													End If
												End If
											End If

										End If
										'Copy field database fields.
										If Not IsNull(rs3("field_table")) And Not IsNull(rs3("field_field")) And rs3("field_type") <> "Label" Then
										'	sqlString = "UPDATE " & rs3("field_table") & " SET " & rs3("field_field") & "=" & _
										'		"(SELECT selected_value FROM " & _
										'		"(SELECT " & rs3("field_field") & " AS selected_value FROM " & rs3("field_table") & _
										'		" WHERE instr_id=" & Request("copyid") & ") AS sub_selected_value) " & _
										'		"WHERE instr_id=" & Request("newid")
											'Get the value to copy.
											sqlString = "SELECT " & rs3("field_field") & _
														" FROM " & rs3("field_table") & _
														" WHERE instr_id=" & Request("copyid")
											Set rs4 = cn.Execute(sqlString)
											If Not rs4.BOF Then
												rs4.MoveFirst
												If Not IsNull(rs4(0)) Then
													newValue = rs4(0)
												Else
													newValue = "null"
												End If
											Else
												newValue = "null"
											End If
											rs4.Close
											
											'Do not copy the field if the value is null.
											If newValue <> "null" Then
												'Make sure that the spec_note is only copied once
												'even though it is on multiple pages.
												If rs3("field_field") = "spec_note" Then
													If noteCopied Then
														copyOk = False
													Else
														noteCopied = True
														copyOk = True
													End If
												Else
													copyOk = True
												End If
												If copyOk Then
													'Find out the data type for the field and create the update string.
													sqlString = "SELECT data_type FROM information_schema.columns WHERE table_schema='instruments' AND table_name='" & rs3("field_table") & "' AND column_name='" & rs3("field_field") & "'"
													Set rs4 = cn.Execute(sqlString)
													If Not rs.BOF Then
														rs4.MoveFirst
														If LCase(rs4(0)) = "varchar" And newValue <> "null" Then
															sqlString = "UPDATE " & rs3("field_table") & " SET " & rs3("field_field") & "='" & newValue & "' WHERE instr_id=" & Request("newid")
														Else
															sqlString = "UPDATE " & rs3("field_table") & " SET " & rs3("field_field") & "=" & newValue & " WHERE instr_id=" & Request("newid")
														End If
													Else
														sqlString = "UPDATE " & rs3("field_table") & " SET " & rs3("field_field") & "=" & newValue & " WHERE instr_id=" & Request("newid")
													End If
													rs4.close
													cmd.CommandText = sqlString
													cmd.Execute
													If cn.Errors.Count > 0 Then
														If cn.Errors.Item(0).Number <> 0 Then
															cn.RollbackTrans
															errorFlag = True
															errorMsg = "Copy transaction rolled back: " & Err.Description
															Exit Do
														End If
													End If

													'Write the change to the audit trail table.
													sqlString = "INSERT INTO audit_trail " & _
																"(change_user,change_instr_id,change_table,change_field,old_value,new_value,change_type) " & _
																"VALUES ('" & currentuser & "'," & Request("newid") & ",'" & rs3("field_table") & "','" & rs3("field_field") & "','null','" & newValue & "','update')"
													cmd.CommandText = sqlString
													cmd.Execute
													If cn.Errors.Count > 0 Then
														If cn.Errors.Item(0).Number <> 0 Then
															cn.RollbackTrans
															errorFlag = True
															errorMsg = "Audit transaction rolled back: " & Err.Description
															Exit Do
														End If
													End If
												End If
											End If

										End If										
											
										rs3.MoveNext
									Loop
								Else
									errorFlag = True
									errorMsg = "The spec format for the selected spec page was not found."
								End If
								rs3.Close
								Set rs3 = Nothing
								Set rs4 = Nothing
							Else
								errorFlag = True
								errorMsg = "The spec page id for the specified spec form is null."
							End If
							rs.MoveNext
						Loop
						rs.Close
					Else
						rs.Close
						errorFlag = True
						errorMsg = "Spec pages were not found for the selected spec form."
					End If
				Else
					rs.Close
					errorFlag = True
					errorMsg = "The selected 'copy from' instrument must have an assigned specification form."
				End If
			Else
				rs.Close
				errorFlag = True
				errorMsg = "The selected 'copy from' instrument was not found."
			End If
		End If
		Set rs = Nothing
		Set cmd = Nothing
		If errorFlag Then
			Response.Write "<br />"
			Response.Write "<div style='text-align:center'><table style='width:70%'><tr><td style='font-weight:bold'>"
			Response.Write "An error occurred while copying the instrument: " & errorMsg
			Response.Write "</td></tr></table></div>"
			cn.Close
			Set cn = Nothing
		Else
			cn.CommitTrans
			cn.Close
			Set cn = Nothing
			Response.Redirect "editspec.asp?instrID=" & Request("newid") & "&page_num=1"
		End If
	End If
Else
	Response.Write "<br />"
	Response.Write "<div style='text-align:center'><table style='width:70%'><tr><td style='font-weight:bold'>"
	Response.Write "An error occurred while performing the database action: " & Err.Description
	Response.Write "</td></tr></table></div>"
	Set cn = Nothing
End If

%>

<body>
</body>
</html>
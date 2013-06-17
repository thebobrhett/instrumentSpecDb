<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<!--#include file="..\Functions\HitCounter.asp"-->
<html>
<head>
<script language="javascript">
function openhelp() {
 window.open("Instrument Spec Database Administrators Guide.doc","userguide");
}
</script>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Instrument Spec Database Administration</title>
<!--#include file="../Functions/AppSecurity.asp"-->
<!--#include file="EquipmentFunctions.asp"-->
<link rel=STYLESHEET href='equipmentstyle.css' type='text/css'>
<style type="text/css">
  div {font-family:verdana;}
  input {font-family:verdana;}
  select {font-family:verdana;
		width:99%;}
  textarea {font-family:verdana;}
  @media print { .noprint {display: none; } }
</style> 
</head>
<body>
<%
'*************
' Revision History
' 
' Keith Brooks - Tuesday, April 20, 2009
'   Creation
'*************

dim HitCounts
Dim currentuser
Dim access
Dim access2

'Set/get hit counts.
HitCounts = HitCounter("equipment_admin")

'Get the current user.
currentuser = Request.ServerVariables("LOGON_USER")
currentuser = Right(currentuser,Len(currentuser)-InStr(currentuser,"\"))

'Use the function to read the allowed users for this page from the database.
'If none are specified, all users are allowed.
'access = UserAllowed(currentuser, "masterbatchentry")
access = UserAccess("equipment", "adminmenu", currentuser)
If access <> "none" Then

%>
	<div style="text-align:center">
		<table width="100%">
			<tr>
				<td style="text-align:left;vertical-align:top;width:25%"><a href="default.asp">Home</a></td>
				<td style="text-align:center;width:50%"><h1 />Instrument Spec Database Administration</td>
				<td style="text-align:right;vertical-align:top;width:25%"><a href="" onclick="openhelp();return false;" title="Open the User Guide">Help</a></td>
			</tr>
			<tr>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:100%">
						<tr>
							<td style="text-align:center;font-weight:bold;font-size:14pt">Tools</td>
						</tr>
<%
		access2 = UserAccess("equipment","audittrail",currentuser)
		If access2 <> "none" Then
%>
						<tr>
							<td style="text-align:center"><input type="button" name="audittrail" value="Instrument Spec&#10;Audit Trail" title="Open the instrument audit trail query form" style="font-size:10pt;width:140px;height:50px"  onclick="window.location='audittrail.asp'" /></td>
						</tr>
<%
		Else
%>
						<tr>
							<td>&nbsp;</td>
						</tr>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","adminaudittrail",currentuser)
		If access2 <> "none" Then
%>
						<tr>
							<td style="text-align:center"><input type="button" name="adminaudittrail" value="Administration&#10;Audit Trail" title="Open the admin audit trail query form" style="font-size:10pt;width:140px;height:50px"  onclick="window.location='adminaudittrail.asp'" /></td>
						</tr>
<%
		Else
%>
						<tr>
							<td>&nbsp;</td>
						</tr>
<%
		End If
%>
					</table>
				</td>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:80%">
						<tr>
							<td colspan="2" style="text-align:center;font-weight:bold;font-size:14pt">Lookup Tables</td>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","classifications",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="classifications" value="Classifications" title="Maintain the classifications lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='classifications.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","instrmodels",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="instrmodels" value="Models" title="Maintain the instrument_models lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='instrmodels.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","cvtypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="cvtypes" value="Control Valve Types" title="Maintain the control_valve_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='cvtypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","linetypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="linetypes" value="Line Types" title="Maintain the line_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='linetypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","drawings",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="drawings" value="Drawings" title="Maintain the drawings lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='drawings.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","lines",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="lines" value="Lines" title="Maintain the process_lines lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='lines.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","drawingtypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="drawingtypes" value="Drawing Types" title="Maintain the drawing_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='drawingtypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","loops",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="loops" value="Loops" title="Maintain the loops lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='loops.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","equipment",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="equipment" value="Equipment" title="Maintain the equipment lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='equipment.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","loopfunctions",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="loopfunctions" value="Loop Functions" title="Maintain the loop_functions lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='loopfunctions.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","equipmenttype",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="equipmenttype" value="Equipment Types" title="Maintain the equipment_type lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='equipmenttype.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","loopprocesses",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="loopprocesses" value="Loop Processes" title="Maintain the loop_processes lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='loopprocesses.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","fmsubtypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="flowmetersubtypes" value="Flow Meter Subtypes" title="Maintain the flow_meter_subtypes lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='fmsubtypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","looptypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="looptypes" value="Loop Types" title="Maintain the loop_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='looptypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","fmtypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="flowmetertypes" value="Flow Meter Types" title="Maintain the flow_meter_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='fmtypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","pipeclasses",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="pipeclasses" value="Pipe Classes" title="Maintain the pipe_classes lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='pipeclasses.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","fluidphases",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="fluidphases" value="Fluid Phases" title="Maintain the fluid_phases lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='fluidphases.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","pipematerials",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="pipematerials" value="Pipe Materials" title="Maintain the pipe_materials lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='pipematerials.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","instrfunctiontypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="instrfunctiontypes" value="Instr Func Types" title="Maintain the instrument_function_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='instrfunctiontypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","plantareas",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="plantareas" value="Plant Areas" title="Maintain the plant_areas lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='plantareas.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","instrlocations",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="instrlocations" value="Locations" title="Maintain the instrument_locations lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='instrlocations.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","processfunctions",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="processfunctions" value="Process Functions" title="Maintain the process_functions lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='processfunctions.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","instrmanufacturers",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="instrmanufacturers" value="Manufacturers" title="Maintain the instrument_manufacturers lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='instrmanufacturers.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
<%
		access2 = UserAccess("equipment","unitsofmeasure",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="unitsofmeasure" value="Units of Measure" title="Maintain the units_of_measure lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='unitsofmeasure.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
							<td>&nbsp;</td>
<%
		access2 = UserAccess("equipment","unitofmeasuretypes",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="unitofmeasuretypes" value="UOM Types" title="Maintain the unit_of_measure_types lookup table" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='unitofmeasuretypes.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
					</table>
				</td>
				<td style="text-align:center;vertical-align:top;border-style:solid;border-color:darkblue;border-width:2px">
					<table style="width:100%">
						<tr>
							<td style="text-align:center;font-weight:bold;font-size:14pt">Security</td>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","rolemembers",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="rolemembers" value="Role Members" title="Assign users to security roles" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='rolemembers.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
						<tr>
<%
		access2 = UserAccess("equipment","useraccess",currentuser)
		If access2 <> "none" Then
%>
							<td style="text-align:center"><input type="button" name="useraccess" value="User Access" title="Assign user privileges for application forms" style="font-size:10pt;width:140px;height:25px"  onclick="window.location='useraccess.asp'" /></td>
<%
		Else
%>
							<td>&nbsp;</td>
<%
		End If
%>
						</tr>
					</table>
				</td>
			</tr>
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

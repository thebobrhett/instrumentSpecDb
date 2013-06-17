<%
'*************************************************************************************
' Constants.
'*************************************************************************************
Dim strSitePath
Dim DBString
'strSitePath = request.servervariables("PATH_TRANSLATED")
'strSitePath = left(strSitePath, len(strSitePath) - (len(strSitePath) - inStrRev(strSitePath, "\")))
'DBString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strSitePath & "\SpecDB.mdb"
DBString = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;user=instrumentsuser;password=easy;DATABASE=instruments;"

'*************************************************************************************
' Function to convert an input datetime string to a MySQL datetime format:
'   YYYY-MM-DD HH24:MI:SS
'
' Keith Brooks - Friday, June 19, 2009
'   Creation.
'*************************************************************************************
Function FormatMySQLDateTime(inputvalue)

	'Format Date for MySQL
	If IsDate(inputvalue) Then
		FormatMySQLDateTime = cstr(year(inputvalue)) & "-" & cstr(month(inputvalue)) & "-" & cstr(day(inputvalue)) & " " & CStr(Hour(inputvalue)) & ":" & CStr(Minute(inputvalue)) & ":" & CStr(Second(inputvalue))
	Else
		FormatMySQLDateTime = "0000-01-01 00:00:00"
	End If

End Function

'*************************************************************************************
' Function to process an input string to allow it to be written to a MySQL text field.
'
' Keith Brooks - Monday, June 15, 2009
'   Creation.
'*************************************************************************************
Function FixString(inputvalue)

	Dim outputvalue
	outputvalue = Replace(inputvalue,"\","\\")
	outputvalue = Replace(outputvalue,"'","\'")
	FixString = outputvalue

End Function

'*************************************************************************************
' Function to determine if a record has been entered for the specified day and the
' specified table.  If so, a True is returned; otherwise, a False is returned.
' This is used to specify whether or not to display the "Add New Record" link on
' certain forms.
'
' Keith Brooks - Wednesday, August 12, 2009
'   Creation.
'*************************************************************************************

Function RecordExists(recordDate, tableName)

	Dim objConnRecord
	dim objRSRecord
	Dim strSQL
	Dim beginDate
	Dim endDate

	If IsDate(recordDate) Then
		beginDate = cstr(year(recordDate)) & "-" & cstr(month(recordDate)) & "-" & cstr(day(recordDate)) & " 00:00:00"
		endDate = cstr(year(recordDate)) & "-" & cstr(month(recordDate)) & "-" & cstr(day(recordDate)) & " 23:59:59"
	
		Set objConnRecord = CreateObject("adodb.connection")
		objConnRecord.Open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=mogsb6.aksa.local;User=polylog;password=easy;DATABASE=polylog;"
		Set objRSRecord = CreateObject("adodb.recordset")

		'Get the count of records for the specified table that exist for the specified date.
		strSQL = "SELECT COUNT(*) FROM " & tableName & " WHERE record_date>='" & beginDate & "' AND record_date<='" & endDate & "'"
		objRSRecord.Open strSQL, objConnRecord, 3
		RecordExists = False
		If Not objRSRecord.BOF Then
			objRSRecord.MoveFirst
			If objRSRecord(0) > 0 Then
				RecordExists = True
			End If
		End If
		objRSRecord.Close
		objConnRecord.Close
		Set objRSRecord = Nothing
		Set objConnRecord = Nothing
	Else
		RecordExists = False
	End If

End Function

'*************************************************************************************
' Function to pad a string with spaces to the right to a fixed length.
'*************************************************************************************

Function PadRight(sStr, nWidth)
	If Len(sStr) < nWidth Then
		Dim count
		Dim temp
		temp = sStr
		For count = Len(sStr) + 1 To nWidth
			temp = temp & Chr(32)
		Next
		PadRight = temp
'		PadRight = sStr & String(nWidth - Len(sStr), " ")
	Else
		PadRight = sStr
	End If
End Function

%>
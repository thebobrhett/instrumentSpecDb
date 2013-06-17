<%@ LANGUAGE="VBSCRIPT" %>
<%Option Explicit%>
<%server.ScriptTimeout=3600%>
<html>
<head>
<meta http-equiv='Expires' http-equiv='Content-Type' content='text/html; charset=windows-1252'>
<title>Threadline Action</title>
</head>
<body>
<%
'*************
'Bob Rhett - Thursday, September 24, 2009
'  Created
'*************

'on error resume next

dim objTLdb
dim objTLrs
dim objHrs
dim current_ts
dim submit_ts
dim strSQL

dim objXdb
dim objXrs


set objTLdb = CreateObject("adodb.connection")
objTLdb.open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;user=rootb;password=spandex;DATABASE=threadline;"
set objTLrs = CreateObject("adodb.recordset")
set objHrs = CreateObject("adodb.recordset")

set objXdb = CreateObject("adodb.connection")
objXdb.open = "driver={MySQL ODBC 3.51 Driver};option=16387;server=richmond.aksa.local;user=rootb;password=spandex;DATABASE=xsection;"
set objXrs = CreateObject("adodb.recordset")

current_ts = cstr(year(now)) & "-" & cstr(month(now)) & "-" & cstr(day(now)) & " " & cstr(hour(now)) & ":" & cstr(minute(now)) & ":" & cstr(second(now))

if request("submit") = "Set Status" then
  strSQL = "select * from status where sm=" & request("sm") & " and duct=" & request("d") & " and position=" & request("p")
  objTLrs.open strSQL, objTLdb
  if objTLrs.eof then
    'If this position does not have a status record...
    objTLrs.close
    '... and we're setting to something other than "normal", create the record.
    if request("reason") <> "Normal" then
      strSQL = "insert into status (sm, duct, position, tl, merge, seq, reason, submit_ts) values (" & request("sm") & ", " & request("d") & ", " & request("p") & ", " & request("tl") & ", '" & request("m") & "', '" & request("s") & "', '" & request("reason") & "', '" & current_ts & "')"
      objTLrs.open strSQL, objTLdb
    end if
  else
    'If this position already has a status record...
    if request("reason") <> objTLrs("reason") then
      '... and we're setting to a different status...
      select case request("action")
        case "(select from list)"
          session("action_err") = "no_action"
          session("status") = request("reason")
        case "other"
          if request("comment") = "" then
            session("action_err") = "no_comment"
            session("action") = request("action")
            session("status") = request("reason")
          else
            session("action_err") = "ok"
          end if
        case else
          session("action_err") = "ok"
      end select
      if session("action_err") = "ok" then
        submit_ts = objTLrs("submit_ts")
        submit_ts = cstr(year(submit_ts)) & "-" & cstr(month(submit_ts)) & "-" & cstr(day(submit_ts)) & " " & cstr(hour(submit_ts)) & ":" & cstr(minute(submit_ts)) & ":" & cstr(second(submit_ts))
        'Move the previous status record to history.
        strSQL = "insert into history (sm, duct, position, tl, merge, seq, reason, submit_ts, action, comment, action_ts) values (" & objTLrs("sm") & ", " & objTLrs("duct") & ", " & objTLrs("position") & ", " & objTLrs("tl") & ", '" & objTLrs("merge") & "', '" & objTLrs("seq") & "', '" & objTLrs("reason") & "', '" & submit_ts & "', '" & request("action") & "', '" & request("comment") & "', '" & current_ts & "')"
        objHrs.open strSQL, objTLdb
        objTLrs.close
        strSQL = "delete from status where sm=" & request("sm") & " and duct=" & request("d") & " and position=" & request("p")
        objTLrs.open strSQL, objTLdb
        if request("reason") <> "Normal" then
          strSQL = "insert into status (sm, duct, position, tl, merge, seq, reason, submit_ts) values (" & request("sm") & ", " & request("d") & ", " & request("p") & ", " & request("tl") & ", '" & request("m") & "', '" & request("s") & "', '" & request("reason") & "', '" & current_ts & "')"
          objTLrs.open strSQL, objTLdb
        end if
        Session.Contents.RemoveAll()
      end if
    end if
  end if
  response.redirect "duct.asp?sm=" & request("sm") & "&d=" & request("d") & "&p=" & request("p")
end if
    
set objTLrs = nothing
objTLdb.close
set objTLdb = nothing
%>
</body>
</html>

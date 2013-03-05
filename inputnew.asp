<%@ Language = VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file="handicap_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
<%
Dim objAdminConn1
OpenDB "D:\database\oakmontjuniors.com\testDB.mdb", objAdminConn1

'query the database and write the results to a recordset
Dim SQL1, rst
Set rst = Server.CreateObject("ADODB.Recordset")
SQL1 = "SELECT * FROM testHCP"
rst.Open SQL1, objAdminConn1, 3, 2
if StrComp(request.form("confirm"),"yes") = 0 then
	Response.write("adding player to database")
	rst.AddNew 
		rst("Name") = CStr(request.form("PlayerName"))
		'rst("Points") = CInt(request.form("Points"))
		rst("Entries") = 1
		'rst("FieldAverage") = CDbl(request.form("AveragePoints"))
		'rst("PlayerAverage") = CDbl(request.form("Points"))
	rst.Update
	response.write("player added")
else
	response.write("not addint to database")
end if

'close recordset
rst.Close
Set rst = Nothing
'close database
CloseDB objAdminConn1
%>
<!-- #include file=inc_botbar.asp -->
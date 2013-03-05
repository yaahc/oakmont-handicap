<%@ Language = VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file="handicap_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
<%
'make database connection
Dim objAdminConn1
OpenDB "D:\database\oakmontjuniors.com\testDB.mdb", objAdminConn1

'query the database and write the results to a recordset
Dim SQL1, rst
Set rst = Server.CreateObject("ADODB.Recordset")
SQL1 = "SELECT * FROM testHCP WHERE NAME='" & request.form("name") & "'"
rst.Open SQL1, objAdminConn1, 3, 2

'send score and connection info into database input subroutine
EnterScore request.form("name"), request.form("points"), request.form("avgpoints"), rst
'response.write(request.form("name"))

'close recordset
rst.Close
Set rst = Nothing
'close database
CloseDB objAdminConn1
%>
<!-- #include file=inc_botbar.asp -->

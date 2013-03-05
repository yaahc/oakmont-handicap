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
SQL1 = "SELECT * FROM testHCP"
rst.Open SQL1, objAdminConn1, 3, 2

'send score and connection info into database input subroutine
GetAdjustment rst, request.form("prob")

'close recordset and database
rst.Close
Set rst = Nothing
CloseDB objAdminConn1
%>
<!-- #include file=inc_botbar.asp -->

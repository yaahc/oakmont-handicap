<%@ Language = VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file="handicap_functions.asp" -->
<!-- #include file=inc_topbar.asp -->

<br>
<center><font size=4 face="calibri,arial"><b>Player Results</font></b></center>
<br><br>
<%
'make database connection
Dim objAdminConn1
OpenDB "D:\database\oakmontjuniors.com\testDB.mdb", objAdminConn1

'query the database and write the results to a recordset
Dim SQL1, rst
Set rst = Server.CreateObject("ADODB.Recordset")
if not IsNull(request.form("name")) then
	SQL1 = "SELECT * FROM testHCP WHERE Name LIKE '%" & request.form("name") & "%' ORDER BY " & request.form("sortBy") & " " & request.form("order")
	'and PlayerLastName='" & request.form("lname") & "'"
else
	SQL1 = "SELECT * FROM testHCP"
end if

rst.Open SQL1, objAdminConn1, 3, 2
%>
<table border=1 bordercolor="#000000">
	<tr>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="Name" />
			<input type="hidden" name="order" value="ASC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="Name" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="Entries" />
			<input type="hidden" name="order" value="DESC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="Entries" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="Points" />
			<input type="hidden" name="order" value="DESC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="Points" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="PlayerAverage" />
			<input type="hidden" name="order" value="DESC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="PlayerAverage" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="FieldAverage" />
			<input type="hidden" name="order" value="DESC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="FieldAverage" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="PlayerAverage" />
			<input type="hidden" name="order" value="ASC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="CummProbWorse" />
			</form>
		</td>
		<td valign=top>
			<font size=3 color="#000000" face="calibri,arial">
			<form method="post" action="report.asp">
			<input type="hidden" name="sortBy" value="PlayerAverage" />
			<input type="hidden" name="order" value="DESC" />
			<input type="hidden" name="name" value="<%=request.form("name")%>" />
			<input type="submit" value="CummProbBetterEqual" />
			</form>
		</td>
	</tr>
<%

'send score and connection info into database input subroutine
GetAdjustment rst, 0

'close recordset
rst.Close
Set rst = Nothing
'close database
CloseDB objAdminConn1
%>
<!-- #include file=inc_botbar.asp -->

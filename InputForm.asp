<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
			<p align=center class=maintitle><b>Input Player Score, exact entries only</b></p>
			<form method="post" action="input.asp">
			Name: <input type="text" name="name" list="names"><br />
			<datalist id="names">
			<%
			'make database connection
			Dim objAdminConn1
			OpenDB "D:\database\oakmontjuniors.com\testDB.mdb", objAdminConn1

			'query the database and write the results to a recordset
			Dim SQL1, rst
			Set rst = Server.CreateObject("ADODB.Recordset")
			SQL1 = "SELECT Name FROM testHCP"
			rst.Open SQL1, objAdminConn1, 3, 2
			
			While not rst.EOF
				response.write("<option value=" & rst("Name") & ">")
				rst.MoveNext
			Wend
			
			'close recordset and database
			rst.Close
			Set rst = Nothing
			CloseDB objAdminConn1
			%>
				<!--option value="Homer Simpson">
				<option value="Bart">
				<option value="Fred Flinstone"-->
			</datalist>
			Points: <input type="number" name="points" /><br />
			Field Average Points: <input type="number" name="avgpoints" /><br /><br />
			<input type="submit" value="Submit" />
			</form>
<!-- #include file=inc_botbar.asp -->

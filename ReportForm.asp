<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
			<p align=center class=maintitle><b>Search By Partial Name</b></p>
			<form method="post" action="report.asp">
			Enter partial first or last name or leave black to get a list of all members<br /><br />
			Name: <input type="text" name="name" list="names"/><br />
			<input type="hidden" name="sortBy" value="Name" />
			<input type="hidden" name="order" value="ASC" />
			<datalist id="names">
				<select name="playerNames">
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
				response.write("<option>" & rst("Name"))
				rst.MoveNext
			Wend
			
			'close recordset and database
			rst.Close
			Set rst = Nothing
			CloseDB objAdminConn1
			%>
				</select>
			</datalist>
			<input type="submit" value="Submit" />
			</form>
			
			<p align=center class=maintitle><b>Search By Probability</b></p>
			<form method="post" action="cummreport.asp">
			Enter a min cummulative probability better or = to search by<br /><br />
			prob: <input type="number" name="prob" /><br />
			<input type="hidden" name="sortBy" value="Name" />
			<input type="submit" value="Submit" />
			</form>
			
<!-- #include file=inc_botbar.asp -->
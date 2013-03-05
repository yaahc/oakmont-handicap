<%@ Language=VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
			<p align=center class=maintitle><b>Update Database</b></p>
			<form method="post" action="fileInput.asp">
			Percentile for Points (Ex. 10 = top 10% get points): <input type="number" name="points" /><br />
			<input type="submit" value="Submit" />
			</form>
<!-- #include file=inc_botbar.asp -->
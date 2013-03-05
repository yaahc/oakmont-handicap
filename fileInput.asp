<%@ Language = VBScript %>
<%Option Explicit%>
<!-- #include file="database_functions.asp" -->
<!-- #include file="handicap_functions.asp" -->
<!-- #include file=inc_topbar.asp -->
<%
'set up file system access for tourney scores
dim fs,fo,x,tfile,tstream,fpath,entry,y
set fs=Server.CreateObject("Scripting.FileSystemObject")
set fo=fs.GetFolder("D:\database\oakmontjuniors.com\TourneyScores\")

'make database connection
Dim objAdminConn1,objCommand
OpenDB "D:\database\oakmontjuniors.com\testDB.mdb", objAdminConn1
Set objCommand=Server.CreateObject("ADODB.command")
objCommand.ActiveConnection=objAdminConn1

'Run Delete Query to empty old data from database
Dim SQL2
objCommand.CommandText="DELETE from testHCP"
objCommand.Execute

'query the database and write the results to a recordset
Dim SQL1, rst
Set rst = Server.CreateObject("ADODB.Recordset")

Response.write("Test 4 <br>")

'Read the tourney scores and enter them into the database
Dim EntryNum, Header
for each x in fo.files
	'Print the name of all files in the test folder
	set tstream = x.OpenAsTextStream(1, -2)
	EntryNum=0
	Header = Split(tstream.ReadLine,",")
	while not tstream.AtEndOfStream
		entry = Split(tstream.ReadLine,",")
		response.write(entry(1) & " ")
		SQL1 = "SELECT * FROM testHCP WHERE NAME='" & entry(1) & "'"
		rst.Open SQL1, objAdminConn1, 3, 2
		If EntryNum/Header(2) < request.form("points")/100 Then
			EnterScore entry(1), 1, request.form("points")/100, rst
		Else
			EnterScore entry(1), 0, request.form("points")/100, rst
		End if
		rst.Close
		EntryNum=EntryNum+1
	wend
next

set fo=nothing
set fs=nothing
%>
<!-- #include file=inc_botbar.asp -->

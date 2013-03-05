<%
'==============================================================================
' The functions in this file are standard routines used to support opening and closing
' database connections
'
'
'Index			Name				Revision		 Date		   Date
'Number												Created		Last Changed
'
' 1				OpenDB				Basic			10/30/02	N/A
' 2				CloseDB				Basic			10/30/02	N/A

'==============================================================================



'1.   SUBROUTINE - OpenDB
'============================================================================================
' Purpose:		Open a connection to the database specified in the parameter.
' Pre:			None.
' Post:			None.
' Return:		None. 
'============================================================================================
' Date Created:	10/30/02
' Revisions:	None 
'============================================================================================
' Known Bugs:	None
'============================================================================================
' Parameters:	DataBase		- STRING: the complete pathname of the database to be opened.
'				ObjectID		- STRING: the variable name of the connection object to be created
'============================================================================================
Sub OpenDB(DataBase,ObjectID)
	Set ObjectID = Server.CreateObject("ADODB.Connection")
	ObjectID.ConnectionString = ""
	ObjectID.ConnectionString = ObjectID.ConnectionString & "Provider=Microsoft.Jet.OLEDB.4.0;"
	ObjectID.ConnectionString = ObjectID.ConnectionString & "Data Source=" & DataBase & ";"
	ObjectID.ConnectionString = ObjectID.ConnectionString & "User ID=;"
	ObjectID.ConnectionString = ObjectID.ConnectionString & "Password=;"
	ObjectID.Open
End Sub






'2.   SUBROUTINE - CloseDB
'============================================================================================
' Purpose:		Closes the database connection specified by the passed paameter
' Pre:			None.
' Post:			None.
' Return:		None.
'============================================================================================
' Date Created:	10/30/02
' Revisions:	None 
'============================================================================================
' Known Bugs:	None
'============================================================================================
' Parameters:	DBConn	- OBJECT: The connection object to be closed
'============================================================================================
Sub CloseDB(ObjectID)
	ObjectID.Close
	Set ObjectID = Nothing
End Sub




%>
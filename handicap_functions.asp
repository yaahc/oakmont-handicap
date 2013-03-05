<!-- #include file="math_functions.asp" -->
<%

'Enter score function, takes a database connection and the score information and queries the database for the player
'If the player is in the database it edits their existing info to add the new score
'if the player is not in the database it adds a new entry for the player

Dim probworse

Sub EnterScore(Name, Points, AveragePoints, rst)
	'if the player was found add new score to existing stats
	If not rst.EOF then
		'table header
		Response.Write("<table border=" & 1 & "><tr><td></td>")
		Response.Write("<td>Name</td>")
		Response.Write("<td>Entries</td>")
		Response.Write("<td>Points</td>")
		Response.Write("<td>PlayerAverage</td>")
		Response.Write("<td>FieldAverage</td>")
		Response.write("</tr>")
		
		'Write old values to table
		Response.Write("<td>Old Entry</td>")
		Response.Write("<td>" & rst("Name") & "</td>")
		Response.Write("<td>" & rst("Entries") & "</td>")
		Response.Write("<td>" & rst("Points") & "</td>")
		Response.Write("<td>" & rst("PlayerAverage") & "</td>")
		Response.Write("<td>" & rst("FieldAverage") & "</td>")
		Response.write("</tr>")
		
		'calculate and write new values to recordset
		rst("Points") = Points + rst("Points")
		rst("Entries") =  1 + rst("Entries")
		rst("FieldAverage") = (AveragePoints + ((rst("Entries")-1) * rst("FieldAverage"))) / rst("Entries")
		rst("PlayerAverage") = rst("Points")/rst("Entries")
		
		'write out new values to table
		Response.Write("<td>New Entry</td>")
		Response.Write("<td>" & rst("Name") & "</td>")
		Response.Write("<td>" & rst("Entries") & "</td>")
		Response.Write("<td>" & rst("Points") & "</td>")
		Response.Write("<td>" & rst("PlayerAverage") & "</td>")
		Response.Write("<td>" & rst("FieldAverage") & "</td>")
		Response.write("</tr></table>")
		
		'write values back to database
		rst.Update
		response.write(" entry updated <br>")
	'if the player was not found make a new entry for them
	Else
		'I was trying to add a confirmation page so that if they were trying to enter a new score they would be shown exactly what would be entered into the 
		'database, asked if they are sure they want to make a new entry, and then if they say yes it would add the new entry, but i had to many issues trying
		'to interupt it to get the confirmation, i would have liked to have some sort of blocking input or something and have an if statement around the
		'add portion that checks the input in the confirmation, but because i didnt know how to do something like that i tried to have an html form that
		'passed the variables to and then added there but the html form messed up the variables and it didnt work well and after a couple of hours i decided
		'to just stop, i figure they will know when they're adding a new entry because it wont be on the autocomplete.
	
		'response.write("Player not in database, are you sure this information is correct?")
		'Response.Write("<table border=" & 1 & "><tr>")
		'Response.Write("<td>Name</td>")
		'Response.Write("<td>Entries</td>")
		'Response.Write("<td>Points</td>")
		'Response.Write("<td>PlayerAverage</td>")
		'Response.Write("<td>FieldAverage</td>")
		'Response.write("</tr><tr>")
		'Response.Write("<td>" & Name & "</td>")
		'Response.Write("<td>" & 1 & "</td>")
		'Response.Write("<td>" & Points & "</td>")
		'Response.Write("<td>" & Points & "</td>")
		'Response.Write("<td>" & AveragePoints & "</td>")
		'Response.write("</tr></table>")
		
		'response.write("<form method="&"post"&" action="&"inputnew.asp"&">")
		'response.write("<input type="&"radio"&" name="&"confirm"&" value="&"yes"&" /> yes<br />")
		'response.write("<input type="&"radio"&" name="&"confirm"&" value="&"no"&" /> yo<br/>")
		'response.write("<input type="&"hidden"&" name="&"PlayerName"&" value=" & Name & "/>")
		'response.write("<input type="&"hidden"&" name="&"Points"&" value="&Points&"/>")
		'response.write("<input type="&"hidden"&" name="&"AveragePoints"&" value="&AveragePoints&"/>")
		'response.write("<input type="&"hidden"&" name="&"rst"&" value="&rst&"/>")
		'response.write("<input type="&"submit"&" value="&"Submit"&" />")
		'response.write("</form>")
		
		rst.AddNew 
			rst("Name") = Name
			rst("Points") = Points
			rst("Entries") = 1
			rst("FieldAverage") = AveragePoints
			rst("PlayerAverage") = Points
		rst.Update
		response.write(" added to database <br>")
	End if
End Sub

Sub GetAdjustment(rst, minprob)
	if rst.EOF then
		response.write("no players found matching search")
	end if
%>	
	
<%
	'Response.Write("<table border=" & 1 & "><tr>")
	'Response.Write("<td>Name</td>")
	'Response.Write("<td>Entries</td>")
	'Response.Write("<td>Points</td>")
	'Response.Write("<td>PlayerAverage</td>")
	'Response.Write("<td>FieldAverage</td>")
	'Response.Write("<td>CummProbWorse</td>")
	'Response.Write("<td>CummProbBetterEqual</td>")
	'Response.write("</tr>")
	
	While not rst.EOF
		Dim prob, rowColor, textColor
		prob = CPMF(rst("Points"), rst("Entries"), rst("FieldAverage"))
		probworse = 1 - prob
		if not prob > 0 then
			prob = 0
		end if
		if rst("Points") > 3 then
			rowColor = "#FF0000"
			textColor = "#FFFFFF"
		else
			rowColor = "#FFFFFF"
			textColor = "#000000"
		end if
		if prob >= CDbl(minprob) then
%>
	<tr bgcolor=<%=rowColor%>>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=rst("Name")%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=rst("Entries")%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=rst("Points")%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=rst("PlayerAverage")%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=rst("FieldAverage")%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=probworse%>
		</td>
		<td valign=top>
			<font size=3 color=<%=textColor%> face="calibri,arial">
			<%=prob%>
		</td>
	</tr>
<%			
			'Response.Write("<tr>")
			'Response.Write("<td>" & rst("Name") & "</td>")
			'Response.Write("<td>" & rst("Entries") & "</td>")
			'Response.Write("<td>" & rst("Points") & "</td>")
			'Response.Write("<td>" & rst("PlayerAverage") & "</td>")
			'Response.Write("<td>" & rst("FieldAverage") & "</td>")
			'Response.Write("<td>" & (1-prob) & "</td>")
			'Response.Write("<td>" & prob & "</td>")
			'Response.write("</tr>")
		end if
		rst.MoveNext
	Wend
%>
</table>
<%
End Sub
%>

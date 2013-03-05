
<%Sub MouseOver(URL,ImgName,ImgFilename,ImgOverFileName,ImgState,ImgAlt)%>

	<a href="<%=URL%>" onMouseOver="switchimage('<%=ImgName%>','<%=ImgOverFileName%>'); return true;" onMouseOut="switchimage('<%=ImgName%>','<%=ImgFileName%>'); return true;"><img src=<%=ImgFileName%> name=<%=ImgName%> alt="<%=ImgAlt%>" border=0></a><br>

<%End Sub%>



<%Sub FormField(FieldType, FieldName, FieldValue, FieldWidth, FieldMax)

	Select Case FieldType
	
		Case "text"%>
		
			<input type="<%=FieldType%>" name="<%=FieldName%>" value="<%=FieldValue%>" style="font:11px Verdana,Helvetica,sans-serif; width:<%=FieldWidth%>px" maxLength="<%=FieldMax%>">
				
		<%Case "checkbox"%>
		
	
	<%End Select
	
	
End Sub

Sub ShowSecondaryLink(strFileName,intLevel,strLinkText,strURL)
	Dim strTarget
	Dim intLineSpacing
	intLineSpacing = 10
	
	' To open a file or external link in a new window send a # as the first argument
	If strFileName = "#" Or strFileName = "pdf" Then
		strTarget = "_blank"
	Else
		strTarget = "_self"
	End If
	
	If intLevel = 2 Then
		If strPageLevel2 = strFileName Then%>
			<font class=SecondaryMenuFont><font class=SecondaryMenuFontSelected><img src=menu_arrow_over.gif><%=strLinkText%></font></font><br>
		<%Else%>
			<font class=SecondaryMenuFont><a href="<%=strURL%>" class=SecondaryMenuLink onMouseOver="switchimage('arrow_<%=strFileName%>','menu_arrow_over.gif'); return true;" onMouseOut="switchimage('arrow_<%=strFileName%>','menu_arrow.gif'); return true;" target=<%=strTarget%>><img src=menu_arrow.gif border=0 name="arrow_<%=strFileName%>"><%=strLinkText%></a></font><br>
		<%End If%>
		<img src=spacer.gif height="<%=intLineSpacing%>" width=1 border=0><br>
	<%End If
	If intLevel = 3 Then
		If strPageLevel3 = strFileName Then%>
			<font class=TertiaryMenuFont><font class=TertiaryMenuFontSelected><%=strLinkText%></font></font><br>
		<%Else%>
			<font class=TertiaryMenuFont><a href="<%=strURL%>" class=TertiaryMenuLink target="<%=strTarget%>"><%=strLinkText%></a></font><br>
		<%End If%>
		<img src=spacer.gif height="<%=intLineSpacing%>" width=1 border=0><br>
	<%End If

End Sub%>














<%'=========================================================================%>
<%'|	Display all the states in a drop-down box.	|%>
<%'=========================================================================%>
<%' This function takes in three parameters:
  '		selectName:		The name of the select field we are going to create.
  '		aState:			Used to automatically select the given state.
  '		defaultState:	The state to use as a default selected state if aState doesn't exist.
Sub DisplayStates(selectName, aState, defaultState)
	' Declare the variables that we will need.
	Dim statesArray													' Hold all the states' full name
	Dim statenamesArray												' Hold all the states' abbreviations
	Dim numStates													' Hold the number of states
	statesArray		= Array("Alabama", "Alaska", "Arizona",	"Arkansas",	"California", "Colorado", "Connecticut", "Delaware", "District of Columbia", "Florida",	"Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas", "Kentucky", "Louisiana",	"Maine", "Maryland", "Massachusetts", "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana", "Nebraska", "Nevada", "New Hampshire",	"New Jersey", "New Mexico", "New York",	"North Carolina", "North Dakota", "Ohio", "Oklohoma", "Oregon",	"Pennsylvania",	"Rhode Island",	"South Carolina", "South Dakota", "Tennessee", "Texas",	"Utah", "Vermont", "Virginia", "Washington", "West Virginia", "Wisconsin", "Wyoming")
	statenamesArray = Array("AL",	   "AK",	 "AZ",		"AR",		"CA",		  "CO",		  "CT",			 "DE",		 "DC",					 "FL",		"GA",	   "HI",	 "ID",	  "IL",		  "IN",		 "IA",	 "KS",	   "KY",	   "LA",		"ME",	 "MD",		 "MA",			  "MI",		  "MN",		   "MS",		  "MO",		  "MT",		 "NE",		 "NV",	   "NH",			"NJ",		  "NM",			"NY",		"NC",			  "ND",			  "OH",	  "OK",		  "OR",		"PA",			"RI",			"SC",			  "SD",			  "TN",		   "TX",	"UT",	"VT",	   "VA",	   "WA",		 "WV",			  "WI",		   "WY")
	numStates = CInt(50)

	If IsNull(aState) Then
		aState = defaultState
	ElseIf Replace(aState," ","") = "" Then
		aState = defaultState
	End If
	%>
	<SELECT name="<%=selectName%>" style="font:11px Verdana,Helvetica,sans-serif; width:50">
		<OPTION value=" ">
		<%
		Dim curState
		For curState = 1 to numStates Step 1
			%>
			<OPTION value="<%=statenamesArray(curState)%>" <%If LCase(statenamesArray(curState)) = LCase(aState) Then Response.Write "SELECTED" End If%>><%=UCase(statenamesArray(curState))%>
			<%
		Next
		%>
	</SELECT>
	<%
End Sub
%>


<%
'==============================================================================
' Purpose:	Return true or false based on whether the input value to be tested passes the requirements
'			for a valid Email address.  Requirements include one "@" in the string and a "." in the string after the "@".
' Pre:		None.
' Post:		None.
' Return:   True/false corresponded to whether the input parameter meets the legality requirements. 
'==============================================================================
' Revisions:
'==============================================================================
' Known Bugs:
'==============================================================================
' Parameters:
Function ValidEmail(strInput,AllowNull)
	Dim intCount, strInputChar, blnOK, intAtLocation, intDotLocation 
	
	intAtLocation = 0
	intDotLocation = 0
	blnOK = true 'Start by assuming everthing is OK and look for anything that is not correct
	If len(strInput) = 0 then
		If AllowNull = false then
			blnOK = false
		End if
	Else		
		For intCount = 1 to len(strInput)
			strInputChar = (mid(strInput,intCount,1))
			If strInputChar = "@" then
				If intAtLocation <> 0 then
					blnOK = false
				else
					intAtLocation = intCount
				End if
			End if
			If strInputChar = "." then
				intDotLocation = intCount
			End if
		Next
		If intAtLocation = 0 or intAtLocation > intDotLocation then
			blnOK = false
		End if
	End if			
	
	ValidEmail = blnOK
End Function

Function DisplayTime(DateValue)
	Dim strTime
	
	strTime = ""
	If Hour(DateValue) > 12 Then
		strTime = cStr(Hour(DateValue) - 12)
	Else
		strTime = cStr(Hour(DateValue))
	End If
	strTime = strTime & ":"
	If Minute(DateValue) < 10 Then strTime = strTime & "0"
	strTime = strTime & cStr(Minute(DateValue))
	If Hour(DateValue) > 11 Then
		strTime = strTime & " PM"
	Else
		strTime = strTime & " AM"
	End If
	
	DisplayTime = strTime
	
End Function

Function DisplayDate(DateValue,Format)
	Dim strDate
		
	Select Case LCase(Format)
		Case 1
			strDate = MonthName(Month(DateValue)) & " " & Day(DateValue) & ", " & Year(DateValue)
		Case 2
			strDate = Month(DateValue) & "/" & Day(DateValue) & "/" & Year(DateValue)
		Case 3
			strDate = Month(DateValue) & "/" & Day(DateValue) & "/" & Right(Year(DateValue),2)
	End Select
	
	DisplayDate = strDate

End Function
%>
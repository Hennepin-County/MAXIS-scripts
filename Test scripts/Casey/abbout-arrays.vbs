EMConnect ""
'About Arrays

'This is all about arrays.

'OUR_TEAM = Array("Ilse", "Casey", "Mark", "Megan")
' team_persons = InputBox("Who is on your team?" & vbCr & "Separate names by commas")

' OUR_TEAM = split(team_persons, ",")
Dim OUR_TEAM()
ReDim OUR_TEAM(0)
' ReDim OUR_TEAM(0)


person_count = 0
Do
	' MsgBox "my counter (person_count) is at " & person_count
	team_person = InputBox("Enter a team member name." & " (when all have been entered, type done.)")
	If team_person <> "done" Then
		ReDim Preserve OUR_TEAM(person_count)
		OUR_TEAM(person_count) = team_person
		person_count = person_count + 1
	End If
Loop until team_person = "done"
' MsgBox "person_count - " & person_count & vbCr & "UBOUND - " & UBound(OUR_TEAM)
' person_count = ""
team_person = ""

our_team_name = "Automation and Integration Team"



Do
	Dialog1 = ""
	BeginDialog Dialog1, 0, 0, 191, 200, "Dialog"
		Text 10, 10, 170, 10, "Our team is the " & our_team_name
		Text 10, 20, 150, 10, "On our team we have:"
		y_pos = 35
		For each person in OUR_TEAM
			Text 15, y_pos, 20, 10, person
			y_pos = y_pos + 10
		Next
		ButtonGroup ButtonPressed
			OkButton 130, y_pos+5, 50, 15
	EndDialog


	dialog Dialog1

	' If ButtonPressed = add_another_form_button Then
	' 	'increment your array of forms - YOU MUST KEEP A COUNTER THAT HAS THE RIGHT INFORMATION or
	' 	next_index = UBOUND(OUR_TEAM) + 1
	' End If

Loop until err_msg = ""
'MsgBox "Our team is the " & our_team_name & " and we have:" & vbCr & join(OUR_TEAM, ", ")

' For each person in OUR_TEAM
' 	MsgBox person
' Next

' For each word in our_team_name
' 	MsgBox word
' Next

'MsgBox "This is person at the 1th instance - " & OUR_TEAM(0)

' For each person in OUR_TEAM
' 	person = trim(person)
' 	MsgBox person
' Next

' For pers_index = 0 to UBound(OUR_TEAM)
' 	MsgBox pers_index & vbCr & OUR_TEAM(pers_index)
' Next

' pers_index = 0
' Do
' 	MsgBox pers_index & vbCr & OUR_TEAM(pers_index)
' 	pers_index = pers_index + 1
' Loop until pers_index > UBound(OUR_TEAM)


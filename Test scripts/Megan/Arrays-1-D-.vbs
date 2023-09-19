EMConnect "" 

'~~~~~~~~~~1 DIMENSIONAL ARRAYS~~~~~~~~~~ 
    'Arrays always start at 0 for the element count
    'If you don't put a # in the () then the system only thinhks of it as numbers (0, 1, 2, 3, 4, ...)
    'You cannot resize the first element in an array
    'Keep the last constanct always last. If you need to add more then add it above the last one. 
    'Splits/Joins are often used for 1D arrays, not multi D arrays. If no deliminator is stated, it assumes the deliminator is a space. 
    'ReDim is more common for multi D arrays. 

' ~~~'Example 1: STATIC ARRAY- array with set elements that are msgboxed. We don't often define arrays with strings becuase we don't know the elements ahead of time, usually we are collecting them.
'     'For Each Structure: 
'         'For Each [word chosen for element] in [Array_Name}
'             '[Insert Action: msgbox, write, etc.]
'         'Next
'         'More actions if desired 
'     OUR_Team = Array("Ilse", "Casey", "Mark", "Megan")          'Array_Name = Array (" ", " ", " ", ...)
'     our_team_name = "Automation and Integration Team" 

'     For each person in OUR_Team                                 'In this example, person is a word chosen to represent the element being referenced
'         Msgbox person
'     Next
    
'     msgbox "This is person at the 1st instance - " & OUR_Team(0)       


' '~~~Example 2: FOR-NEXT If you know how many elements you have use For Next. If not, look at example 3 for Do Loop. These accomplish similar things, however do loops are used if you don't know how many elements you have.
'     'For Next Structure: 
'         'For [word chosen for element] = 0 to UBound*(Array_Name) 
'             'MsgBox [word chosen for element] & vbCr & Array_Name([word chosen for element])
'         'Next

'     OUR_Team = Array("Ilse", "Casey", "Mark", "Megan")          'Array_Name = Array (" ", " ", " ", ...)
'     our_team_name = "Automation and Integration Team" 

'     For pers_index = 0 to Ubound(OUR_Team)
'         MsgBox pers_index & vbCr & OUR_Team(pers_index)
'     Next


' '~~~'Example 3: DO-LOOP If you don't know how many elements you have use Do Loop instead of For Next. 
' 'Do Loop Structure: 
'     '[word chosen for element] = 0 
'     'Do
'         'MsgBox [word chosen for element] & vbCr & Array_Name([word chosen for element])
'         '[word chosen for element] = [word chosen for element] + 1 
'     'Loop until [word chosen for element] > UBound (Array_Name)

' OUR_Team = Array("Ilse", "Casey", "Mark", "Megan")          'Array_Name = Array (" ", " ", " ", ...)
' our_team_name = "Automation and Integration Team" 

' pers_index = 0
' Do   
'     MsgBox pers_index & vbCr & OUR_Team(pers_index)           'pers_index = number, OUR_Team(pers_index) = actual name
'     pers_index = pers_index + 1                               'this keeps us from overwritting ourselves 
' Loop Until Ubound(OUR_Team) > UBound(OUR_Team)


' '~~~'Example 4: SPLIT- Collecting elements via Input box. One input box for all names. If no deliminator is entered it will use any space to separate elements. 

'     team_persons = InputBox("Who is on your team?")             'Input box to collect elements for array

'     OUR_Team = split(team_persons, ", ")
'     our_team_name = "Automation and Integration Team" 

'     For each person in OUR_Team                                 'In this example, person is a word chosen to represent the element being referenced
'         Msgbox person                                           'Msgbox the name of each person in separate msgboxes based on "," separation
'     Next
    


' '~~~'Example 5: Dim and ReDim. ReDim tells the system we are going to resize the array. Preserve keeps previous entries, otherwise it would overwrite it/delete it.

' Dim OUR_team()                  'Defining array
' ReDim OUR_Team(0)               'Redefining array so we can resize it

' person_count = 0                'Variable filled with an integer TODO??
' Do
'     team_person = InputBox("Who is on your team?")             'Input box to collect elements for array. team_person is a variable filled with an integar 
'     If team_person <> "done" Then                               'This gives us a way to know when we are done entering names
'         ReDim Preserve OUR_Team(person_count)                   'ReDim Preserve will keep all entries
'         OUR_Team(person_count) = team_person                    'Adds team members to array
'         person_count = person_count + 1                         'Increments to the next line/# so we can add another entry without overwritting the previous
'     End If
' Loop Until team_person = "done"

' our_team_name = "Automation and Integration Team" 

' MsgBox "Our team is the " & our_team_name & "and we have:"
' For each person in OUR_Team
'     MsgBox person 
' Next


'~~~'Example 6: DO-LOOP with FOR EACH-NEXT - Allows you to collect inputs from end user and then display them in a dialog without overwriting the previous one. 

Dim OUR_team()                  'Defining array
ReDim OUR_team(0)               'Redefining array so we can resize it

person_count = 0                'Variable filled with an integer TODO??
Do
    team_person = InputBox("Who is on your team?")             'Input box to collect elements for array. team_person is a variable filled with an integar 
    If team_person <> "done" Then                               'This gives us a way to know when we are done entering names
        ReDim Preserve OUR_Team(person_count)                   'ReDim Preserve will keep all entries
        OUR_Team(person_count) = team_person                    'Adds team members to array
        person_count = person_count + 1                         'Increments to the next line/# so we can add another entry without overwritting the previous
    End If
Loop Until team_person = "done"

our_team_name = "Automation and Integration Team" 
    
Do                                                                 'Dialog must be in a do-loop in order for it to update with all of the entries
    Dialog1 = ""                                                   'Dialog to display entries by end user
    BeginDialog Dialog1, 0, 0, 191, 200, "Dialog"
    Text 10, 10, 170, 10, "Our team is the" & our_team_name
    Text 10, 20, 150, 10, "On our team we have:"
    y_pos = 35                                                      'set y position so we dymanically add to it
    For each person in OUR_team                                     'for each person entered via input box which is stored in the array OUR_team
        Text 15, y_pos, 150, 10, person                              'write name in set position
        y_pos = y_pos + 10                                          'move down 10 in y position for next entry
    Next 
    ButtonGroup assess_button_pressed
        OkButton 130, y_pos + 5, 50, 15
    EndDialog

    dialog Dialog1
Loop until err_msg = ""







'~~~~~~~~~~MULTI DIMENSIONAL ARRAYS~~~~~~~~~~ 
    'ReDim is more common for multi D arrays. 
    'You will have more than 1 number in the array for a multidimensional i.e. ARRAY_NAME (worker_const, recert_cases)
    'The entries inside the () for the array are always numbers, BUT we name them as constants so that 1)we know what they stand for as we use them in our code 2) so that we can't use them as a variable somewhere else


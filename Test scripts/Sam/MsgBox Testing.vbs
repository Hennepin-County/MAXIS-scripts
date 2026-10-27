




Function search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)
    'Check on the right screen
    EMReadScreen WORKERS_check, 7, 2, 26
    ' The (DAIL) is located in the same place on both the DAIL/"" submenu and the DAIL/DAIL page. The word "WORKERS" is in a different location on both pages. This will keep the DAIL script from getting stuck on the submenu
    If WORKERS_check <> "WORKERS" then script_end_procedure("You are not in DAIL for the final deletion. The script will now end")

    'last_first_name = replace(last_first_name, " ", "")

    'Setting the row/col markers to start on 5,4 coordinates to avoid the non-DAIL information triggering accidental matches
    row = 5
    col = 4
    EMSearch last_first_name, row, col
    'If last_first_name cannot be found, loop through DAIL/DAIL with PF8. If the last_page_check triggers, end the script; otherwise, look for the name again and keep looping until the script finds it.
    If row = 0 and col = 0 then
        Do
            PF8
            EMReadScreen last_page_check, 21, 24, 2
                If last_page_check = "THIS IS THE LAST PAGE" then script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
            EMSearch last_first_name, row, col
        Loop until row <> 0 and col <> 0
    End if

    'On the page the last_first_name was found, look for the DAIL message. 
    EMSearch Original_DAIL, row, col 
    'Failsafe for if the person is on one page and their DAIL is on the next. Their name carries over to the next page - so the script will check for the name and then the message and keep looping as long as the name appears
      If row = 0 and col = 0 then
        Do
          PF8
          EMSearch last_first_name, row, col
            If row = 0 and col = 0 then script_end_procedure("The original DAIL could not be found associated with that client. The script will now end.")
          EMSearch Original_DAIL, row, col
        Loop until row <> 0 and col <> 0
      End If
    
'Final deletion logic - if the row and col parameters are not 0, that means that the Original_DAIL has been found, which means "D" can delete the DAIL using the row coordinate. Otherwise, end the script
    If row <> 0 and col <> 0 then
      EMWriteScreen "D", row, 3
      transmit
      MsgBox "The original DAIL was found and deleted"
    Else
        script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
    End If
End Function













function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function

EMConnect ""



MEMB_name_to_save = "MILLER-FRENCH,KELA"








'Compare name to panel that appeared - if a match doesn't happen, then find the correct name in the MEMB list on the left
EMReadScreen name_comparison_on_panel, 30, 4, 37
    name_comparison_on_panel = trim(name_comparison_on_panel)
    name_comparison_on_panel = Replace(name_comparison_on_panel, ", ", ",")
    MsgBox name_comparison_on_panel
If name_comparison_on_panel = MEMB_name_to_save then
    name_comparison_on_panel = MEMB_name_to_save
Else 
    commaPOS = InStr(MEMB_name_to_save, ",")
        If commaPOS > 0 then
            MEMB_name_to_save = Mid(MEMB_name_to_save, commaPOS + 1)
            MEMB_name_to_save = trim(MEMB_name_to_save)
        Else
            MsgBox "put in an end message"
        End If
    'create MAXIS friendly length for substring searching
    substring_name_comparison = Left(MEMB_name_to_save, 6)
    row = 1
    col = 12
    EMSearch substring_name_comparison, row, col
    MsgBox "Row: " & row & "Col: " & col & "Name: " & substring_name_comparison
    If row <> 0 and col = 12 then
        EMReadScreen MEMB_number_from_comparison, 2, row, 3
        EMWriteScreen MEMB_number_from_comparison, 20, 76
        transmit
    Else
        MsgBox "tee hee"
    End if

End If





Msgbox name_comparison_on_panel





last_first_name = "MILLER-FRENCH,KELA"
Original_DAIL = "SSN HAS NOT BEEN VERIFIED"

Dim row, col

function PF8()
'--- This function sends or hits the PF8 key.
 '===== Keywords: MAXIS, MMIS, PRISM, PF8
  EMSendKey "<PF8>"
  EMWaitReady 0, 0
end function



Function search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)
    'Check on the right screen
    EMReadScreen DAIL_check, 7, 2, 26
    ' The (DAIL) is located on both the DAIL/"" page and the DAIL/DAIL page. The word "WORKERS" is in a different location on both pages. This will keep the DAIL script from getting stuck on the submenu
    If DAIL_check <> "WORKERS" then script_end_procedure("You are not in DAIL for the final deletion. The script will now end")

    row = 5
    col = 4
    
    EMSearch last_first_name, row, col
    
    If row = 0 and col = 0 then
        Do
            PF8
            EMReadScreen last_page_check, 21, 24, 2
                If last_page_check = "THIS IS THE LAST PAGE" then script_end_procedure("The original DAIL could not be found for deletion. The script will now end.")
                Exit Do
            EMSearch last_first_name, row, col
        Loop until row <> 0 and col <> 0
    End if

    EMSearch Original_DAIL, row, col 
    'Failsafe for if the person is on one page and their DAIL is on the next 
      If row = 0 and col = 0 then
        Do
          PF8
          EMSearch last_first_name, row, col
            If row = 0 and col = 0 then script_end_procedure("The original DAIL could not be found associated with that client. The script will now end.")
            Exit Do
          EMSearch Original_DAIL, row, col
        Loop until row <> 0 and col <> 0
      End If
    
      EMWriteScreen "D", row, 3
      transmit
      MsgBox "The original DAIL was found and deleted"

End Function

Call search_for_correct_DAIL_and_delete(last_first_name, Original_DAIL)

MsgBox Original_DAIL & " found at row: " & row & " found at col: " & col



'Remove this one and keep in a testing file - just here to view arrays in Msgboxes as needed
Sub DisplayArray(arr)
    Dim dims, output, i, j

    ' Determine number of dimensions
    On Error Resume Next
    dims = 1
    Do While True
        i = LBound(arr, dims)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Do
        End If
        dims = dims + 1
    Loop
    dims = dims - 1
    On Error GoTo 0

    output = "Array Contents:" & vbCrLf & vbCrLf

    If dims = 1 Then
        ' Handle 1D array
        For i = LBound(arr) To UBound(arr)
            output = output & "Index " & i & ": " & arr(i) & vbCrLf
        Next

    ElseIf dims = 2 Then
        ' Handle 2D array
        For i = LBound(arr, 1) To UBound(arr, 1)
            For j = LBound(arr, 2) To UBound(arr, 2)
                output = output & "Row " & i & ", Col " & j & ": " & arr(i, j) & vbCrLf
            Next
        Next
    Else
        output = "Unsupported array dimensions: " & dims
    End If

    MsgBox output, vbOKOnly, "Array Debug Output"
End Sub
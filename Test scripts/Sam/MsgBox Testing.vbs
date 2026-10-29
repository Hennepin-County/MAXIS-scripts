


function transmit()
'--- This function sends or hits the transmit key.
 '===== Keywords: MAXIS, MMIS, PRISM, transmit
  EMSendKey "<enter>"
  EMWaitReady 0, 0
end function


Dim count, row, MEMB_name_to_save

EMConnect ""



MEMB_name_to_save = "MILLER-FRENCH,KELA"



Function MEMB_Name_DAIL_matching(MEMB_name_to_save)

    Dim MEMB_loop_array()
    ReDim MEMB_loop_array(0)

    count = 0

    row = 5

    Do
        EMReadScreen MEMB_number_from_comparison, 2, row, 3

        MembNo_value = Trim(MEMB_number_from_comparison)

        If MembNo_value = "" then
            Exit Do
        End If

        If IsNumeric(MembNo_value) then
            ReDim Preserve MEMB_loop_array(count)
            MEMB_loop_array(count) = MembNo_value
            count = count + 1
        End If

        Row = row + 1
    Loop

    For i = 0 To UBound(MEMB_loop_array)
        mna = MEMB_loop_array(i)

        Do
            EMWriteScreen mna, 20, 76
            transmit
            EMReadScreen panel_exist_check, 14, 24, 13
                If panel_exist_check = "DOES NOT EXIST" then
                    Exit Do
                End If
                
                EMReadScreen name_comparison_on_panel, 30, 4, 37
                    
                name_comparison_on_panel = replace(name_comparison_on_panel, ", ", ",")
                    If name_comparison_on_panel = MEMB_name_to_save then
                        exit function
                    Else
                        Exit Do
                    End If 
                
        Loop

    Next
    
End Function

call MEMB_Name_DAIL_matching(MEMB_name_to_save)




































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
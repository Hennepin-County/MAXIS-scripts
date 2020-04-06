'this is a function - only a function

function pause_at_certificate_of_understanding()
    region_known = FALSE        'setting this to start
    Do
        EMReadScreen check_for_cert_of_understanding, 28, 2, 28
        If check_for_cert_of_understanding = "Certificate Of Understanding" Then
            'go to training region because that is where this thing happens
            attn            'getting to the primary menu
            Do
                EMReadScreen MAI_check, 3, 1, 33
                If MAI_check <> "MAI" then EMWaitReady 1, 1
            Loop until MAI_check = "MAI"

            If region_known = FALSE Then                        'We only want to look for the region one time - otherwise it would always be Training
                region_known = TRUE
                EMReadScreen production_status, 7, 6, 15        'looking to see which session was opened
                EMReadScreen inquiry_status, 7, 7, 15
                EMReadScreen training_status, 7, 8, 15
                If production_status = "RUNNING" Then           'Setting a boolean to know which one was opened originally so we can go back to it.
                    use_prod = TRUE
                    EMWriteScreen "C", 6, 2     'here we close because otherwise the agreement stays up
                    transmit
                ElseIf inquiry_status = "RUNNING" Then
                    use_inq = TRUE
                    EMWriteScreen "C", 7, 2     'here we close because otherwise the agreement stays up
                    transmit
                ElseIf training_status = "RUNNING" Then
                    use_trn = TRUE
                End If
            End If

            EMWriteScreen "3", 2, 15                        'actually going into training region'
            transmit

            'Now we stop the script with a dialog so that the user can still interact with MAXIS
            Dialog1 = ""
            BeginDialog Dialog1, 0, 0, 211, 155, "MAXIS Certificate of Understanding"
              ButtonGroup ButtonPressed
                OkButton 155, 135, 50, 15
              Text 5, 5, 135, 15, "It appears it is time for you to review your MAXIS agreement to maintain access."
              Text 5, 25, 125, 25, "This annual agreement details of using this system in line with privacy and confidentiality requirements."
              Text 5, 60, 200, 10, "*** YOU MUST READ AND REVIEW THIS INFORMATION ***"
              GroupBox 5, 75, 200, 55, "Instructions"
              Text 15, 90, 175, 35, "Leave this dialog up and read the MAXIS screen currently displayed. Enter your agreement selection. Once this is completed, press 'OK' on this dialog and the script will continue. "
            EndDialog

            Dialog Dialog1                                  'showing the dialog here
            cancel_without_confirmation
            'If ButtonPressed = 0 Then stopscript
        End If
    Loop until check_for_cert_of_understanding <> "Certificate Of Understanding"    'we keep showing the dialog until this is done
    If region_known = TRUE Then
        'Now we are going back to the region we started in.
        attn
        Do
            EMReadScreen MAI_check, 3, 1, 33
            If MAI_check <> "MAI" then EMWaitReady 1, 1
        Loop until MAI_check = "MAI"
        EMWriteScreen "C", 8, 2
        transmit

        If use_prod = TRUE Then EMWriteScreen "1", 2, 15
        If use_inq = TRUE Then EMWriteScreen "2", 2, 15
        If use_trn = TRUE Then EMWriteScreen "3", 2, 15
        transmit
    End If
end function

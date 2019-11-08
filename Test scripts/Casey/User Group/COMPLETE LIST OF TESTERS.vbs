class tester_detail

    public tester_full_name                 'this is the full name, first and last names
    public tester_first_name
    public tester_email
    public tester_id_number
    public tester_x_number
    public tester_supervisor_name
    public tester_supervisor_email
    public tester_population
    public tester_region
    public tester_groups
    public tester_scripts

    'use this to email testers from within a script.
    public property send_tester_email(include_supervisor, body_text)
        If include_supervisor = TRUE Then
            cc_email = tester_supervisor_email
        Else
            cc_email = ""
        End If
        body_text = "Hello " & tester_first_name & ", " & vbCr & body_text & vbCr & "Thank you for all you do for us!" & vbCr & "BlueZone Script Team"
        Call create_outlook_email(tester_email, cc_email, "Testing Response - " & name_of_script, body_text, "", TRUE)
    end property

    'use this to display the message that testing will start.
    public Property display_testing_message(selection_reason)
        'use the variable 'test_reason' for what we are specifically testing in this script
        'selection_reason' should only be 'GROUP', 'REGION', 'POPULATION' or 'SCRIPT' or blank - nothing else will add information to the message.
        selection_reason = UCase(selection_reason)
        reason_text = ""
        If selection_reason = "GROUP" Then reason_text = "You have been selected because you are a part of the " & selected_group & " and we need feedback specifically in this group of testers."
        If selection_reason = "REGION" Then reason_text = "You have been selected because you are in " & tester_region & " and we need feedback specifically from this region."
        If selection_reason = "POPULATION" Then reason_text = "You have been selected because you work in " & tester_population & " and we need feedback for work in this population."
        If selection_reason = "SCRIPT" Then reason_text = "You have been selected because we believe you can offer feedback specifically on this script."

        message_text = ""
        message_text = "Hello " & tester_first_name & "!" & vbNewLine & vbNewLine & "You have been selected to test thes cript - " & name_of_script & "."
        If test_reason <> "" Then message_text = message_text & vbNewLine & "This script is being tested for: " & test_reason & ". "
        If reason_text <> "" Then message_text = message_text & vbNewLine & reason_text
        message_text = message_text & vbNewLine & "At the end of the script the Automated In-Script Error Reporting will pop-up for submitting feedback about the script run. Even 'everything was good' reports are helpful to us."
        message_text = message_text & vbNewLine & vbNewLine & "Thank you for all your hard work and assistance to us."
        message_text = message_text & vbNewLine & "                                          - The BlueZone Script Team"

        MsgBox message_text
    end property

end class

tester_num = 0
ReDim Preserve tester_array(tester_num)
tester_array(tester_num).tester_full_name           = "Casey Love"
tester_array(tester_num).tester_first_name          = "Casey"
tester_array(tester_num).tester_email               = "casey.love@hennepin.us"
tester_array(tester_num).tester_id_number           = "CALO001"
tester_array(tester_num).tester_x_number            = "x127L1S"
tester_array(tester_num).tester_supervisor_name     = "Ilse Ferris"
tester_array(tester_num).tester_supervisor_email    = "ilse.ferris@hennepin.us"
tester_array(tester_num).tester_population          = "QI"
tester_array(tester_num).tester_region              = "South"
tester_array(tester_num).tester_groups              = array("")
tester_array(tester_num).tester_scripts             = array("")

tester_num = tester_num + 1
ReDim Preserve tester_array(tester_num)
tester_array(tester_num).tester_full_name           = ""
tester_array(tester_num).tester_first_name          = ""
tester_array(tester_num).tester_email               = ""
tester_array(tester_num).tester_id_number           = ""
tester_array(tester_num).tester_x_number            = ""
tester_array(tester_num).tester_supervisor_name     = ""
tester_array(tester_num).tester_supervisor_email    = ""
tester_array(tester_num).tester_population          = ""
tester_array(tester_num).tester_region              = ""
tester_array(tester_num).tester_groups              = array("")
tester_array(tester_num).tester_scripts             = array("")

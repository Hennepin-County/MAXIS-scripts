actionable_dail_count = 0 'Setting up incrementor for counting actionable DAIL messages 
dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.

EMReadScreen DAIL_case_number, 8, dail_row - 1, 73
DAIL_case_number = trim(DAIL_case_number)
If DAIL_case_number = MAXIS_case_number then 
    DO
        'Determining if there is a new case number...
        EMReadScreen new_case, 8, dail_row, 63
        new_case = trim(new_case)
        IF new_case <> "CASE NBR" THEN '...if there is NOT a new case number, the script will read the DAIL type, month, year, and message...
            Call write_value_and_transmit("T", dail_row, 3)
            dail_row = 6
        ELSEIF new_case = "CASE NBR" THEN
            '...if the script does find that there is a new case number (indicated by "CASE NBR"), it will write a "T" in the next row and transmit, bringing that case number to the top of your DAIL
            Call write_value_and_transmit("T", dail_row + 1, 3)
            dail_row = 6
        End if
    
        'Reading the DAIL Information
        EMReadScreen DAIL_case_number, 8, dail_row - 1, 73
        DAIL_case_number = trim(DAIL_case_number)
        If DAIL_case_number <> MAXIS_case_number then exit do
        
        EMReadScreen dail_type, 4, dail_row, 6

        EMReadScreen dail_msg, 61, dail_row, 20
        dail_msg = trim(dail_msg)

        EMReadScreen dail_month, 8, dail_row, 11
        dail_month = trim(dail_month)

        Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
        IF actionable_dail = True then actionable_dail_count = actionable_dail_count + 1
        
        dail_row = dail_row + 1
    LOOP
End if 

'output actionable_dail_count into array 
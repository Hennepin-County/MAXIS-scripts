EMReadScreen number_of_dails, 1, 3, 67		'Reads where the count of DAILs is listed

DO
    If number_of_dails = " " Then exit do		'if this space is blank the rest of the DAIL reading is skipped
    dail_row = 6			'Because the script brings each new case to the top of the page, dail_row starts at 6.
    DO
        dail_type = ""
        dail_msg = ""

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
        EMReadScreen MAXIS_case_number, 8, dail_row - 1, 73
        MAXIS_case_number = trim(MAXIS_case_number)
        
        EMReadScreen dail_type, 4, dail_row, 6

        EMReadScreen dail_msg, 61, dail_row, 20
        dail_msg = trim(dail_msg)

        EMReadScreen dail_month, 8, dail_row, 11
        dail_month = trim(dail_month)

        stats_counter = stats_counter + 1   'I increment thee
        Call non_actionable_dails(actionable_dail)   'Function to evaluate the DAIL messages
        
        IF actionable_dail = True then      'actionable_dail = True will NOT be deleted and will be captured and reported out as actionable.  
            If len(dail_month) = 5 then
                output_year = ("20" & right(dail_month, 2))
                output_month = left(dail_month, 2)
                output_day = "01"
                dail_month = output_year & "-" & output_month & "-" & output_day
            elseif trim(dail_month) <> "" then
                'Adjusting data for output to SQL
                output_year     = DatePart("yyyy",dail_month)   'YYYY-MM-DD format
                output_month    = right("0" & DatePart("m", dail_month), 2)
                output_day      = DatePart("d", dail_month)
                dail_month = output_year & "-" & output_month & "-" & output_day
            End if
            
            dail_string = worker & " " & MAXIS_case_number & " " & dail_type & " " & dail_month & " " & dail_msg
            'If the case number is found in the string of case numbers, it's not added again. 
            If instr(all_dail_array, "*" & dail_string & "*") then
                If dail_type = "HIRE" then
                    add_to_array = True 
                Else 
                    add_to_array = False
                End if 
            else 
                add_to_array = True 
            End if 
            
            If add_to_array = True then          
                ReDim Preserve DAIL_array(4, DAIL_count)	'This resizes the array based on the number of rows in the Excel File'
                DAIL_array(worker_const,	           DAIL_count) = worker
                DAIL_array(maxis_case_number_const,    DAIL_count) = right("00000000" & MAXIS_case_number, 8) 'outputs in 8 digits format
                DAIL_array(dail_type_const, 	       DAIL_count) = dail_type
                DAIL_array(dail_month_const, 		   DAIL_count) = dail_month
                DAIL_array(dail_msg_const, 		       DAIL_count) = dail_msg
                Dail_count = DAIL_count + 1
                all_dail_array = trim(all_dail_array & dail_string & "*") 'Adding MAXIS case number to case number string
                dail_string = ""
            elseif add_to_array = False then 
                false_count = false_count + 1
            End if 
        End if

        dail_row = dail_row + 1
        '...going to the next page if necessary
        EMReadScreen next_dail_check, 4, dail_row, 4
        If trim(next_dail_check) = "" then
            PF8
            EMReadScreen last_page_check, 21, 24, 2
            If last_page_check = "THIS IS THE LAST PAGE" then
                all_done = true
                exit do
            Else
                dail_row = 6
            End if
        End if
    LOOP
    IF all_done = true THEN exit do
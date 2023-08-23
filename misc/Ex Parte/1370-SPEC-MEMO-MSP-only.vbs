'OneSource Policy: https://www.dhs.state.mn.us/main/idcplg?IdcService=GET_DYNAMIC_CONVERSION&RevisionSelectionMethod=LatestReleased&dDocName=ONESOURCE-15010315'
Call start_a_new_spec_memo(memo_opened, True, forms_to_arep, forms_to_swkr, send_to_other, other_name, other_street, other_city, other_state, other_zip, True)  ' start the memo writing process

Call write_variable_in_SPEC_MEMO(resident_name & "'s health care coverage has been automatically renewed effective " & first_day_of_elig_period & " for the following Medicare Savings Program:")
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO(msp_program)
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("Your Income was verified using electronic sources.")
Call write_variable_in_SPEC_MEMO("Household size: " & hh_size)
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("---Counted Income (All Amounts are Per Month)---")
For each i = 0 to Ubound(income_array)
    Call write_variable_in_SPEC_MEMO("    * " & income_source & ": " & income_amount & ".")
Next
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("(42 CFR 435.916, MN Statutes 256B.056 & 256B.057)")
Call write_variable_in_SPEC_MEMO("")
Call write_variable_in_SPEC_MEMO("If any of the information on this notice is wrong, please contact the county at the phone number listed on the notice.")
Call write_variable_in_SPEC_MEMO("Visit www.mn.gov/dhs/abdautorenew for more information about your automatic renewal.")
PF4 'Exits the MEMO
script_end_procedure("")

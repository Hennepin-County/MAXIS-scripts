* SPACE IS A CHARACTER
mystandard_array = LIST
Array(Chris, Casey, Aurelia, Ronin) - to see it
join(mystandard_array)
ReDim = I am going to change this
Preserve = keep all of that other stuff
always remember to increment the incrementor
single demention array
* you cannot change the constant - CONST just means the number, not a spot holder, not a variable does not mean something else
(kind_of_info, for_what_peeps)
incrementor = incrementor + 1 'this is a number not a person'
FOR uc_membs = 0 to uBound(uc_members_array, 2) ' uc_membs = incrementor '

FMCAMAM1 = Case Notes (NOTE)
FMCAMAM2 = Case Notes (NOTE) inside case note
FMKDLAM6 = DAIL report
FMCADAM1 =  Edit Summary (SUMM)
FMACAAM1 = Case Status Display  (CASE)
FMBDAAM9 = Household Member (MEMB)
FMCAHAP1 = Stat Panel Personal Summary (PNLP)
FMCCIAM3 = Program Action (PACT)
FMBDGAM7 = Address (ADDR)
FMBDAAM9 = Household Member (MEMB)
FMCDAAM7 = Additional Member Info (MEMI)
FMCDJAME = Work Registration (WREG)
FMCCKDMD = Case Reviews (REVW)
FMLVCAM2 = Income Verifications To Do (IEVC)
FMLWIAM3 = Verification Log Update (IULA)
FMKHTAM1 = Current Statuses Display  (CURR) 
IF STAT_note_check = "NOTE" THEN
    EMReadScreen screen_code 8, 1, 72
    IF screen_code = "FMCAMAM2" THEN 'FMCAMAM2 = Case Notes (NOTE) inside case note
        EMReadScreen mode_check 1, 2, 09
        IF mode_check = "D" THEN PF3
    ELSEIF
        screen_code = "FMCAMAM1" THEN already_at_the_correct_screen = TRUE  'FMCAMAM1 = Case Notes (NOTE)
    END IF
END IF

Add  = "A"
Edit = "E"
Display = "D"

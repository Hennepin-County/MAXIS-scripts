'sort dates script
dates_array = array(#12/23/2013#, #5/3/2010#, #6/28/2018#, #2/17/2015#, #12/23/2013#)
MsgBox "start"

function sort_dates(dates_array)

    dim ordered_dates ()
    redim ordered_dates(0)

    days =  0
    do

        prev_date = ""
        for each thing in dates_array
            check_this_date = TRUE
            For each known_date in ordered_dates
                if known_date = thing Then check_this_date = FALSE
                'MsgBox "known dates is " & known_date & vbNewLine & "thing is " & thing & vbNewLine & "match - " & check_this_date
            next
            if check_this_date = TRUE Then
                if prev_date = "" Then
                    prev_date = thing
                Else
                    if DateDiff("d", prev_date, thing) <0 then
                        prev_date = thing
                    end if
                end if
            end if
        next
        if prev_date <> "" Then
            redim preserve ordered_dates(days)
            ordered_dates(days) = prev_date
            days = days + 1
        end if
    loop until days > UBOUND(dates_array)

    dates_array = ordered_dates
end function

for each thing in dates_array
    MsgBox thing
next

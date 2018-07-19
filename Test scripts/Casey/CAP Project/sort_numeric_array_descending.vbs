'Function for sorting numeric array descending (biggest to smallest)
FUNCTION sort_numeric_array_descending(values_array, separate_character, output_array)
	'trimming and splitting the array
	values_array = trim(values_array)
	values_array = split(values_array, separate_character)
	num_of_values = ubound(values_array)

	REDIM placeholder_array(num_of_values, 1)
		' position 0 is the number, position 1 is if the number has been put in the output array

	'assigning the number values to the multi-dimensional placeholder array AND whether the specific value has been used for comparison yet (position 1)
	array_position = 0
	FOR EACH num_char IN values_array
		IF num_char <> "" THEN
			num_char = cdbl(num_char)
			placeholder_array(array_position, 0) = num_char
			placeholder_array(array_position, 1) = FALSE
			array_position = array_position + 1
		END IF
	NEXT

	'reseting array_position for the generation of the output array
	array_position = 0
	i = 0
	all_sorted = FALSE
	DO
		'stating that the number has not yet been put into the sorted array
		highest_value = FALSE
		value_to_watch = placeholder_array(i, 0)
		IF placeholder_array(i, 1) = FALSE THEN
			FOR j = 0 TO num_of_values
				'If the value is not blank AND if we still have not assigned this value to the output array. We need
				' to avoid a list of only the lowest values, which is what happens what you remove the placeholder_array(j, 1) bit
				IF placeholder_array(j, 0) <> ""  AND placeholder_array(j, 1) = FALSE THEN
					IF value_to_watch >= placeholder_array(j, 0) THEN
						highest_value = TRUE
					ELSE
						'If the function finds a value LOWER than the current one, it stops comparison and exits the FOR NEXT
						highest_value = FALSE
						EXIT FOR
					END IF
				END IF
			NEXT
		END IF

		'If we confirm that this is the highest value...
		IF highest_value = TRUE THEN
			'...then we assign position 1 as TRUE (so we will not use this value for comparison in the future)
			placeholder_array(i, 1) = TRUE
			'...we assign it to the output array...
			output_array = output_array & value_to_watch & ","
			'...and we move on to the next position in the array...
			array_position = array_position + 1
			'...until we find that we have hit the ubound for the original array. Then we stop assigning.
			IF array_position = num_of_values THEN all_sorted = TRUE
		END IF
		'If we get through this specific number and find that it does not go next on the sorted list,
		' we need to get to the next number. If we find that we have got through all the numbers and the list
		' is not complete, we need to reset this value, and start back at the beginning of the original list.
		' This way, we avoid skipping numbers that should be showing up on the list.
		i = i + 1
		IF i = num_of_values AND all_sorted = FALSE THEN i = 0
	LOOP UNTIL all_sorted = TRUE

	output_array = trim(output_array)
	output_array = split(output_array, ",")
END FUNCTION

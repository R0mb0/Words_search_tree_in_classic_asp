<%
    Class words_search_tree

    'Fields'
    Dim characters_array '-> array where store the special characters
    Dim letters_array '-> array where store letters
    Dim numbers_array '-> array where store numbers
    Dim terminator '-> terminatore character of the string
    Dim base_array '-> the array where save the infos
    Dim array_index '-> array index of last searched element 
    Dim case_sensitive '-> variable to set case sensitive
    Dim remove_special_chars '-> varaible to remove special characters from the text
    Dim remove_letters '-> variable to remove single letters in the text
    Dim remove_numbers '-> variable to remove single numbers
    Dim remove_all_numbers '-> variable to remove all numbers

    ' Initialization and destruction'
    Sub class_initialize()
        characters_array = Array(".", ",", ":", ";", "`", "/", "\", "|", "_", "-", "~", "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "+", "=", "{", "[", "}", "]", "'", "<", ">")
        letters_array = Array("q", "w", "e", "r", "t", "y", "u", "i", "o", "p", "a", "s", "d", "f", "g", "h", "j", "k", "l", "z", "x", "c", "v", "b", "n", "m", "è", "é", "ì", "í", "ò", "ó", "à", "á", "ù", "ú")
        numbers_array = Array("1", "2", "3", "4", "5", "6", "7", "8", "9", "0")
        terminator = null
        base_array = Array()
        case_sensitive = false
        remove_special_chars = false
        remove_letters = false 
        remove_numbers = false 
        remove_all_numbers = false
    End Sub
        
    Sub class_terminate()
        characters_array = nothing
        letters_array = nothing
        terminator = nothing
        base_array = nothing
        array_index = nothing
        case_sensitive = nothing
        remove_special_chars = nothing
        remove_letters = nothing 
        remove_numbers = nothing 
        remove_all_numbers = nothing 
    End Sub

    'Function to initialize the class with the terminator 
    Public Function initialize(ByVal termin, ByVal case_sens, ByVal remove_special_char, ByVal remove_lett, ByVal remove_numb, ByVal remove_all_numb)
        'Check if the terminator is a special character
        If Not(is_special_character(termin))Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "initialize - The terminator: " & termin & " is not a special character for thi reason is not valid")
        End If
        'Check if the two params are not true simultaneously
        If remove_numb and remove_all_numb Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "initialize - remove_numbers and remove_all_numbers params could not be true simultaneously!")
        End If 
        terminator = termin
        case_sensitive = case_sens
        remove_special_chars = remove_special_char
        remove_letters = remove_lett 
        remove_numbers = remove_numb 
        remove_all_numbers = remove_all_numb  
    End Function

    'Function to print a debug message 
    Private Function dp(message)
        Response.write "<br><h3> Debug print: " & message & " </h3><br>"
    End Function 

    'Function to check if a character is a special character.
    Private Function is_special_character(ByVal character)
        Select Case character
            Case "."
                is_special_character = true
                Exit Function 
            Case ","
                is_special_character = true
                Exit Function
            Case ":"
                is_special_character = true
                Exit Function
            Case ";"
                is_special_character = true
                Exit Function
            Case "`"
                is_special_character = true
                Exit Function
            Case "/"
                is_special_character = true
                Exit Function
            Case "\"
                is_special_character = true
                Exit Function
            Case "|"
                is_special_character = true
                Exit Function
            Case "_"
                is_special_character = true
                Exit Function
            Case "-"
                is_special_character = true
                Exit Function
            Case "~"
                is_special_character = true
                Exit Function
            Case "!"
                is_special_character = true
                Exit Function
            Case "@"
                is_special_character = true
                Exit Function
            Case "#"
                is_special_character = true
                Exit Function
            Case "$"
                is_special_character = true
                Exit Function
            Case "%"
                is_special_character = true
                Exit Function
            Case "^"
                is_special_character = true
                Exit Function
            Case "&"
                is_special_character = true
                Exit Function
            Case "*"
                is_special_character = true
                Exit Function
            Case "("
                is_special_character = true
                Exit Function
            Case ")"
                is_special_character = true
                Exit Function
            Case "+"
                is_special_character = true
                Exit Function
            Case "="
                is_special_character = true
                Exit Function
            Case "{"
                is_special_character = true
                Exit Function
            Case "["
                is_special_character = true
                Exit Function
            Case "}"
                is_special_character = true
                Exit Function
            Case "]"
                is_special_character = true
                Exit Function
            Case "'"
                is_special_character = true
                Exit Function
            Case "<"
                is_special_character = true
                Exit Function
            Case ">"
                is_special_character = true
                Exit Function
            Case else
                is_special_character = false
                Exit Function
        End Select
    End Function 

    'Function to check if a string contains a special character.
    Private Function recognize_special_character(ByVal my_string)
        Dim temp 
        For Each temp In characters_array
            If InStr(my_string, temp) <> 0 Then 
                recognize_special_character = temp
                Exit Function 
            End if
        Next
        recognize_special_character = null
    End Function 

    'Function to check if a string contains a special character.
    Private Function remove_special_characters(ByVal my_string)
        Dim temp_string
        temp_string = my_string
        Dim temp 
        For Each temp In characters_array
            If InStr(my_string, temp) <> 0 Then 
                temp_string = Replace(temp_string, temp, "")
            End if
        Next
        temp_string = Trim(temp_string)
        remove_special_characters = temp_string
    End Function 

    'Function to remove single letters from a text
    Private Function remove_single_letters_from_text(ByVal my_string)
        Dim temp_string
        temp_string = my_string
        Dim temp
        For Each temp In letters_array
            If InStr(temp_string, " " & temp & " ") <> 0 Then 
                temp_string = Replace(temp_string, " " & temp & " ", " ")
            End If 
        Next
        temp_string = Trim(temp_string)
        remove_single_letters_from_text = temp_string
    End Function 

    'Function to remove single numbers from a text
    Private Function remove_single_numbers_from_text(ByVal my_string)
        Dim temp_string
        temp_string = my_string
        Dim temp
        For Each temp In Split(my_string, " ")
            If IsNumeric(temp) Then 
                temp_string = Replace(temp_string, temp & " ", "")
            End If 
        Next
        temp_string = Trim(temp_string)
        remove_single_numbers_from_text = temp_string
    End Function 

    'Function to remove double spaces from text
    Private Function remove_double_spaces(ByVal my_string)   
        Dim temp_string
        temp_string = my_string
        Do While InStr(1, temp_string, "  ")
            temp_string= Replace(temp_string, "  ", " ")
        Loop
        temp_string = Trim(temp_string)
        remove_double_spaces = temp_string
    End Function

    'Function to remove all numbers from a text
    Private Function remove_all_numbers_from_text(ByVal my_string)
        Dim temp_string
        temp_string = my_string
        Dim temp
        For Each temp In numbers_array
            If InStr(temp_string, temp) <> 0 Then 
                temp_string = Replace (temp_string, temp, "")
            End If 
        Next
        remove_all_numbers_from_text = remove_double_spaces(temp_string)
    End Function 

    'Function to check if in the text there's a number
    Private Function check_number_in_text(ByVal my_string)
        Dim temp 
        For Each temp In numbers_array
            If InStr(my_string, temp) Then 
                check_number_in_text = true 
                Exit Function 
            End If  
        Next
        check_number_in_text = false 
    End Function

    'Function to check if is presente a special character inside a string and return the funded one. 
    Private Function recognize_terminator(ByVal my_string)
        If InStr(my_string, terminator) <> 0 Then 
            recognize_terminator = true
            Exit Function 
        End if
        recognize_terminator = false
    End Function 

    'Function to convert a string into an array
    Private Function string_to_array(ByVal text)
        Dim length
        length = Len(text)
        Dim outArray() 
        Dim index 
        For index = 0 to length - 1
            Redim preserve outArray(length)
            outArray(index) = Left(Right(text,(length - index)), (1))
        Next 
        Redim preserve outArray(length - 1)
        string_to_array = outArray
    End Function

    'Ad an element in the array head
    Private Function add_base_element(ByVal element, ByRef array)
        Dim temp 
        temp =  UBound(array) + 1
        Redim Preserve array(temp)
        array(temp) = element
    End Function 

    'Return true (and save the index) if find the element, else false 
    Private Function search_base_element(ByVal element, ByRef array)
        Dim temp 
        Dim my_index 
        my_index = 0
        For Each temp In array 
            If IsArray(temp) Then 
                If temp(0) = element Then 
                    array_index = my_index
                    search_base_element = true
                    Exit Function 
                End If 
            End If 
            my_index = my_index + 1
        Next 
        search_base_element = false
    End Function 

    'This function ad an element 
    Private Function node(ByVal value)
        Dim temp_array(0)
        temp_array(0) = value
        node = temp_array
    End Function 

    'Function to add a word inside the array
    Private Function adding_word(ByVal word, ByVal index, ByRef array)
        If Not(index <= UBound(word)) Then 'If index is not valid 
            Exit Function
        End If
        If UBound(array) = "-1" Then 'If the array is empty
            Redim preserve array(0)
            array(0) = node(word(index))
            adding_word word, index + 1, array(0)
            Exit Function
        End If 
        If search_base_element(word(index), array) Then 
            adding_word word, index + 1, array(array_index)
            Exit Function
        Else
            add_base_element node(word(index)) ,array 
            adding_word word, index + 1, array(UBound(array))
            Exit Function
        End If 
    End Function 

    'Function to throw the errors 
    Private Function create_allert(ByVal string_function, ByVal data)
        'Check if the class has been initializated
        If IsNull(terminator) Then 
             Call Err.Raise(vbObjectError + 10, "words_search_tree.class", string_function & " - The class has not been initalizated")
        End If 
        'Check if the word is a letter
        If remove_letters and Not(Len(data) > 1) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", string_function & " - The character: " & data & " is not a word")
        End If 
        'Check if in the word is present the terminator character 
        If recognize_terminator(data) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", string_function & " - The word contains the terminator")
        End If 
        'Check if the word is a number
        If remove_numbers and IsNumeric(data) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", string_function & " - The word: " & data & " is a number")
        End If
        'Check if a word is built with numbers or is a number
        If remove_all_numbers and check_number_in_text(data) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", string_function & " - The word: " & data & " contains numbers")
        End If
    End Function 

    'The public function to add a word in the search tree, in this case the word must be a string
    Public Function add_word(ByVal word)
        'If necessary throw error 
        create_allert "add_word", word
        Dim my_word 
        my_word = word
        'Check if a word is built with numbers 
        If remove_all_numbers and Not(IsNumeric(my_word)) Then 
            my_word = remove_all_numbers_from_text(my_word)
        End If 
        'If is case sentive 
        If Not case_sensitive Then 
            my_word = LCase(my_word)
        End If 
        my_word = string_to_array(my_word)
        add_base_element terminator, my_word  
        adding_word my_word, 0, base_array 
    End Function 

    'Private function to add word in more efficient way if the function "add_text" has been invoked
    Private Function private_add_word(ByVal word)
        Dim my_word 
        my_word = word
        'If is case sentive 
        If Not case_sensitive Then 
            my_word = LCase(my_word)
        End If 
        my_word = string_to_array(my_word)
        add_base_element terminator, my_word  
        adding_word my_word, 0, base_array 
    End Function

    'Public function to add the words of a text 
    Public Function add_text(ByVal text)
        'Check if the class has been initializated
        If IsNull(terminator) Then 
             Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "add_text - The class has not been initalizated")
        End If 
        'Check if the text is a text
        If Not(InStr(text, " ") <> 0) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "add_text - " & text & " is not a text")
        End If
        Dim temp_text 
        temp_text = text
        'Remove special characters from text if necessary
        If remove_special_chars Then 
            temp_text = remove_special_characters(temp_text)
        End If 
        'Remove single letters from text if necessary
        If remove_letters Then 
            temp_text = remove_single_letters_from_text(temp_text)
        End If 
        'Remove single numbers from text if necessary
        If remove_numbers Then 
            temp_text = remove_single_numbers_from_text(temp_text)
        End If 
        'Remove all numbers from text if necessary
        If remove_all_numbers Then 
            temp_text = remove_all_numbers_from_text(temp_text)
        End If 
        Dim temp 
        For Each temp In Split(temp_text, " ")
            private_add_word(temp)
        Next
    End Function

    'Function to print an array line 
    Private Function write_array(ByRef array, ByVal flag)
        Dim index 
        index = 0
        Dim temp 
        For Each temp In array
            If IsArray(temp) Then 
                If UBound(temp) > 1 Then 
                    write_array array(index), true
                Else
                    write_array array(index), false
                End If 
            Else 
                If temp = terminator Then 
                    Response.write "; "
                Else
                    If flag Then 
                        Response.write(temp & "-<br>")
                    Else
                        Response.write(temp)
                    End If  
                End If 
            End If 
            index = index + 1
        Next 
    End Function 

    'Function to print all the elements inside the search tree 
    Public Function Write_all_elements()
        If IsNull(terminator) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "Write_all_elements - The class has not been initalizated")
        End If 
        Dim temp 
        For Each temp In base_array
            If UBound(temp) > 1 Then
                write_array temp, true
            Else
                write_array temp, false
            End If 
            Response.write "<br>"
        Next 
    End Function 

    'Private function to find a word inside the tree
    Private Function find_word(ByVal word, ByVal index, ByRef array)
        If search_base_element(word(index), array) Then 
            If word(index) = terminator Then 
                find_word = true
                Exit Function 
            Else 
                find_word = find_word(word, index + 1, array(array_index))
                Exit Function
            End IF 
        End If 
        find_word = false 
    End Function 

    'Private Function to check if a word is in the memory
    Private Function private_is_present(ByVal word)
        Dim my_word
        my_word = word
        'If is case sentive 
        If Not case_sensitive Then 
            my_word = LCase(my_word)
        End If 
        my_word = string_to_array(word)
        add_base_element terminator, my_word 
        Dim temp 
        Dim index
        For Each temp In base_array
            If temp(0) = my_word(0) Then 
               is_present = find_word(my_word, 1, base_array(index))
               Exit Function 
            End If 
            index = index + 1
        Next
        is_present = false 
    End Function 

    'Private Function to check if a word is in the memory
    Public Function is_present(ByVal word)
        'If necessary throw error 
        create_allert "is_present", word
        ' now launch the original function
        is_present = private_is_present(word)
    End Function 

    Private Function print_words(ByVal word, ByVal index, ByRef array, ByVal flag)
        Dim temp 
        'If the words is spent
        If flag Then 
            '------------------------------------------------------------------------------------------------------------------
        Else
            If search_base_element(word(index), array) Then 
                If index = UBound(word) Then 
                    print_words = word(index) & print_words(word, index + 1, array(array_index), true)
                    Exit Function 
                End If 
                print_words = word(index) & print_words(word, index + 1, array(array_index), false)
                Exit Function 
            End If 
        End If 
    End Function 

    'Function to search a word inside the memory
    Public Function search_word(ByVal word)
        'In case of null argument 
        If word = " " and (Len(word) = 0) Then 
            add_base_element(" ", words)
            search_word = words
            Exit Function 
        End If 
        'If necessary throw error 
        create_allert "search_word", word
        'If the word is present then exit 
        If private_is_present(word) Then 
            add_base_element(" ", words)
            search_word = words
            Exit Function 
        End If 
        Dim my_word
        my_word = word
        'If is case sentive 
        If Not case_sensitive Then 
            my_word = LCase(my_word)
        End If 
        my_word = string_to_array(word)
        
    End Function 

    End Class 
%>
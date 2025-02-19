<%
    Class words_search_tree

    'Fields'
    Dim terminator '-> terminatore character of the string
    Dim base_array '-> the array where save the infos
    Dim array_index '-> array index of last searched element 

    Dim count 

    ' Initialization and destruction'
    Sub class_initialize()
        terminator = null
        base_array = Array()

        count = 0
    End Sub
        
    Sub class_terminate()
        terminator = nothing
        base_array = nothing
        array_index = nothing
    End Sub

    'Function to initialize the class with the terminator 
    Public Function initialize(ByVal termin)
        If Not(is_special_character(termin))Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "class_initialize - The terminator: " & termin & " is not a special character for thi reason is not valid")
        End If
        terminator = termin
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
    Function recognize_special_character(my_string)
        If InStr(my_string, ".") <> 0 Then 
            recognize_special_character = "."
            Exit Function 
        End if
        If InStr(my_string, ",") <> 0 Then 
            recognize_special_character = ","
            Exit Function
        End if
        If InStr(my_string, ":") <> 0 Then 
            recognize_special_character = ":"
            Exit Function
        End if
        If InStr(my_string, ";") <> 0 Then 
            recognize_special_character = ";"
            Exit Function
        End if
        If InStr(my_string, "`") <> 0 Then 
            recognize_special_character = "`"
            Exit Function
        End if
        If InStr(my_string, "/") <> 0 Then 
            recognize_special_character = "/"
            Exit Function
        End if
        If InStr(my_string, "\") <> 0 Then 
            recognize_special_character = "\"
            Exit Function
        End if
        If InStr(my_string, "|") <> 0 Then 
            recognize_special_character = "|"
            Exit Function
        End if
        If InStr(my_string, "_") <> 0 Then 
            recognize_special_character = "_"
            Exit Function
        End if
        If InStr(my_string, "-") <> 0 Then 
            recognize_special_character = "-"
            Exit Function
        End if
        If InStr(my_string, "~") <> 0 Then 
            recognize_special_character = "~"
            Exit Function
        End if
        If InStr(my_string, "!") <> 0 Then 
            recognize_special_character = "!"
            Exit Function
        End if 
        If InStr(my_string, "@") <> 0 Then 
            recognize_special_character = "@"
            Exit Function
        End if
        If InStr(my_string, "#") <> 0 Then 
            recognize_special_character = "#"
            Exit Function
        End if
        If InStr(my_string, "$") <> 0 Then 
            recognize_special_character = "$"
            Exit Function
        End if
        If InStr(my_string, "%") <> 0 Then 
            recognize_special_character = "%"
            Exit Function
        End if
        If InStr(my_string, "^") <> 0 Then 
            recognize_special_character = "^"
            Exit Function
        End if
        If InStr(my_string, "&") <> 0 Then 
            recognize_special_character = "&"
            Exit Function
        End if
        If InStr(my_string, "*") <> 0 Then 
            recognize_special_character = "*"
            Exit Function
        End if
        If InStr(my_string, "(") <> 0 Then 
            recognize_special_character = "("
            Exit Function
        End if
        If InStr(my_string, ")") <> 0 Then 
            recognize_special_character = ")"
            Exit Function
        End if
        If InStr(my_string, "+") <> 0 Then 
            recognize_special_character = "+"
            Exit Function
        End if
        If InStr(my_string, "=") <> 0 Then 
            recognize_special_character = "="
            Exit Function
        End if
        If InStr(my_string, "{") <> 0 Then 
            recognize_special_character = "{"
            Exit Function
        End if
        If InStr(my_string, "[") <> 0 Then 
            recognize_special_character = "["
            Exit Function
        End if
        If InStr(my_string, "}") <> 0 Then 
            recognize_special_character = "}"
            Exit Function
        End if
        If InStr(my_string, "]") <> 0 Then 
            recognize_special_character = "]"
            Exit Function
        End if
        If InStr(my_string, "'") <> 0 Then 
            recognize_special_character = "'"
            Exit Function
        End if
        If InStr(my_string, "<") <> 0 Then 
            recognize_special_character = "<"
            Exit Function
        End if
        If InStr(my_string, ">") <> 0 Then 
            recognize_special_character = ">"
            Exit Function
        End if
        recognize_special_character = null
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
    Private Function node(value)
        Dim temp_array(0)
        temp_array(0) = value
        node = temp_array
    End Function 

    'Function to add a word inside the array
    Private Function adding_word(ByVal word, ByVal index, ByRef array)
        If Not(index >= 0 and index <= UBound(word)) Then 'If index is not valid 
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

    'The public function to add a word in the search tree, in this case the word must be a string
    Public Function add_word(ByVal word)
        If IsNull(terminator) Then 
             Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "The class has not been initalizated")
        End If 
        If Not(Len(word) > 1) Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "add_word - The character: " & word & " is not a word")
        End If 
        If Not recognize_terminator(word) Then 
            Dim my_word
            my_word = string_to_array(word)
            add_base_element terminator, my_word 
            Dim temp 
            adding_word my_word, 0, base_array 
        Else 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "add_word - The word contains the terminator")
            End If 
    End Function 

    'Public function to add the words of a text 
    Public Function add_text(ByVal text)
        Dim temp 
        Dim character
        character = null
        For Each temp In Split(text, " ")
            character = recognize_special_character(temp)
            If IsNull(character) Then 
                add_word(temp)
            Else 
                add_word(Replace(temp, character, ""))
            End If 
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
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "The class has not been initalizated")
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

    End Class 
%>
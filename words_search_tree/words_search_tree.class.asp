<%
    Class words_search_tree

    'Fields'
    Dim terminator '-> terminatore character of the string
    Dim base_array '-> the array where save the infos
    Dim array_index '-> array index of last searched element 

    ' Initialization and destruction'
    Sub class_initialize(termin)
        If Not(is_special_character(termin))Then 
            Call Err.Raise(vbObjectError + 10, "words_search_tree.class", "The terminator: " & termin & " is not a special character for thi reason is not valid")
        End If
        terminator = termin
        base_array = Array()
    End Sub
        
    Sub class_terminate()
        terminator = nothing
        base_array = nothing
        array_index = nothing
    End Sub

    'Function to check if a character is a special character in Classic ASP.
    Private Function is_special_character(character)
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

    'Function to check if is presente a special character inside a string and return the funded one. 
    Private Function recognize_special_character(my_string)
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

    'Function to convert a string into an array
    Private Function stringToArray(text)
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
    Private Function add_base_element(element, array)
        Dim temp 
        temp =  UBound(array) + 1
        Redim Preserve array(temp)
        array(temp) = element
    End Function 

    'Return true (and save the index) if find the element, else false 
    Private Function search_base_element(element, array) 
        Dim temp 
        Dim my_index 
        my_index = 0
        For Each temp In array 
            If temp = element Then 
                search_base_element = true
                array_index = my_index
                Exit Function 
            End If 
            my_index = my_index + 1
        Next 
        search_base_element = false
    End Function 

    'Function to add a word inside the array
    Private Function adding_word(word, index, array)
        If index >= 0 and index <= UBound(word)
            If search_base_element(word(index), word)
                adding_word(word, index + 1, array(array_index))
            Else 
                Dim temp_array(1)
                temp_array(0) = word(index)
                add_base_element(temp_array, array) '-> adding a new array in the head of prevoius array
                adding_word(word, index + 1, array(UBound(array))) '-> Now change array
            End If
        End If 
    End Function 

    'The public function to add a word in the search tree, in this case the word must be a string
    Public Function add_word(word) 

    End Function 

    End Class 
%>
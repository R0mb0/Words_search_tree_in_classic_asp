<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="words_search_tree.class.asp"-->
<%
    Dim tree
    Set tree = new words_search_tree
    tree.initialize("-")
    tree.add_word("nana")
    'tree.Write_all_elements()
    tree.add_word("babu")
    'tree.Write_all_elements()
    tree.add_word("bana")
    tree.Write_all_elements()

%> 
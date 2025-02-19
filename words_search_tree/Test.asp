<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="words_search_tree.class.asp"-->
<%
    Dim tree
    Set tree = new words_search_tree
    tree.initialize("-")
    tree.add_text("Nel mezzo del cammin di nostra vita mi ritrovai per una selva oscura, che la diritta via era smarrita.")
    'tree.add_word("nana")
    'tree.Write_all_elements()
    'tree.add_word("babu")
    'tree.Write_all_elements()
    'tree.add_word("bana")
    tree.Write_all_elements()

%> 
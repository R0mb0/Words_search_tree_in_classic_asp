<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="words_search_tree.class.asp"-->
<% 

    Dim tree
    Set tree = new words_search_tree
    tree.initialize "-", true, true, true, false, true
    'tree.add_text("123 Nel mezzo del cammin di nostra45 vita mi ritrovai per una selva oscura, che la diritta via era smarrita.")
    'tree.add_text("Ahi quanto a dir 890 qual era è cosa dura esta selva22 selvaggia e aspra e forte che nel pensier rinova la paura!")
    tree.add_text("bana balena banana")
    tree.Write_all_elements()

    'Response.write "<br><h3> Posiedo la parola: di -> " & tree.is_present("di") & " </h3><br>"
    'Response.write "<br><h3> Posiedo la parola: dia -> " & tree.is_present("dia") & " </h3><br>"
    'Response.write "<br><h3> Posiedo la parola: a -> " & tree.is_present("a") & " </h3><br>"
%> 
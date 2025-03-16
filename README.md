# Words search tree in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/0170adcbf22345cf90855a8f90957a19)](https://app.codacy.com/gh/R0mb0/Words_search_tree_in_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Words_search_tree_in_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Words_search_tree_in_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## Tree structure

- **Words**
  - nana 
  - baba
  - banana

### Chart

![FlowChart](https://github.com/R0mb0/Words_search_tree_in_classic_asp/blob/main/Images/Search_tree.png)

> #### ⚠️ Commenting `Private Function write_array(...)`
> **In case of writing all tree's elements, the tree is explored in iterative way that produce a linear output**

## `words_search_tree.class.asp`'s avaible functions

- Initialize the class -> `Public Function initialize(ByVal termin, ByVal case_sens, ByVal remove_special_char, ByVal remove_lett, ByVal remove_numb, ByVal remove_all_numb)`
  >
  > **Where the params are:**
  >  - `termin` -> The terminator character (must be a special character as "-").
  >  - `case_sens` -> case sensitive option, could be "true" or "false", if "false", all text will be transformed in lower case.
  >  - `remove_special_char` -> remove special characters, could be "true" or "false", if true, the special chacters will be removed from the text.
  >  - `remove_lett` -> remove single letters, could be "true" or "false", if "true", the sigle letters will be removed, for example " a ".
  >  - `remove_numb` -> remove number, could be "true" or "false", if "true" the numbers will be removed, for example " 123 "
  >  - `remove_all_numb` -> remove all numbers from text, cuold be "true" or "false", if true, all numbers will be removed, for example: "Hi123" became "Hi"
  >
  > **⚠️ "remove_numb" and "remove_all_numb" params could not be true simultaneously!**
- Add a word to the tree -> `Public Function add_word(ByVal word)`
- Add all text words in the tree -> `Public Function add_text(ByVal text)`
- Print all elements inside the tree -> `Public Function Write_all_elements()`
- Check if a word is in the tree -> `Public Function is_present(ByVal word)`
- Search a word inside the memory -> `Public Function search_word(ByVal word, ByVal is_array)`
  >
  > **This function is usefull to search a word inside the tree, for example if `word` = "hom" the function will return: "home", "homo" and "hometown"**
  > - `word` is the word to search, it could be a part of a word, if an entire word is passed to the function, the function will return null.
  > - `is_array` change the output of the function, if `true` the function will return an array with all results, else, will be returned a string
- Function to save to file the tree state -> `Public Function save_tree(ByVal path)`
  > **Where `path` is the string with the location with the file to save location**
- Function to load the saved state tree in a file -> `Public Function load_tree(ByVal path)`
  > **Where `path` is the string with the location with the file to load location**

## How to use 

> From: `Test.asp`

1. Initialize the class
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="words_search_tree.class.asp"-->
   <% 
      Dim tree
      Set tree = new words_search_tree
      tree.initialize "-", true, true, true, false, true
   ```
2. Add values to tree   
   Possibilities:  
   - Load tree from file
     ```
     tree.load_tree("path")
     ```
   - Add words
     ```
     tree.add_word("nana")
     tree.add_word("baba")
     tree.add_word("banana")
     ```
   - Add text
     ```
     tree.add_text("Nel mezzo del cammin di nostra vita mi ritrovai per una selva oscura, che la diritta via era smarrita.")
     tree.add_text("Ahi quanto a dir qual era è cosa dura esta selva selvaggia e aspra e forte che nel pensier rinova la paura!")
     ```
3. Save the state of tree
   ```
   tree.save_tree("path")
   ```  
4. Interrogate the tree   
   Possibilities:
   - Check if a word is in the tree
     ```
     tree.is_present("banana")
     ```
   - Search a word inside the tree
     ```
       tree.search_word("bana", false)
     %>
     ```
    

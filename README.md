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

## `words_search_tree.class.asp`'s avaible functions

- Initialize the class -> `Public Function initialize(ByVal termin, ByVal case_sens, ByVal remove_lett, ByVal remove_numb, ByVal remove_all_numb)`
  >
  > **Where the params are:**
  >  - `termin` -> The terminator character (must be a special character as "-").
  >  - `case_sens` -> case sensitive option, could be "true" or "false", if "true", all text will be transformed in lower case.
  >  - `remove_lett` -> remove single letters, could be "true" or "false", if "true", the sigle letters will be removed, for example " a ".
  >  - `remove_numb` -> remove number, could be "true" or "false", if "true" the numbers will be removed, for example " 123 "
  >  - `remove_all_numb` -> remove all number from text, cuold be "true" or "false", if true, all numbers will be removed, for example: "Hi123" became "Hi"

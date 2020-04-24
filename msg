add in($a,$b) sub-function : is $b in $a?

200424--------------
- add in sub-function
- IFEQUAL("+<+in($modulename,tidl)+>+","")

200401--------------
you can run with  csv file instead of xlsx file.
First of all , you convert the file into csv file when you have error of perl package for excel.
if you use MAC OSX , your csv file will have as one line with ^M. so you can replace ^M with \r\n

--------------
support sldd & %+<+$API+>+ and restore $g_y policy (plus)

v1.0.4
- support sldd 
- support ```ITERATE %+<+$API+>+ ``` syntax
    - add ```$iterate_var_name = replace_var_with_value($iterate_var_name);```
- restore $g_y policy (plus)
    - replace_var_with_value does not break becasue break contents to improve speed  is impossible to keep the order when we count with plus().


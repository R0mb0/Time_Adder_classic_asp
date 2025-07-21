# Time Adder Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/77ab5c2aec3d48d39a70860ddf120673)](https://app.codacy.com/gh/R0mb0/Time_Adder_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Time_Adder_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Time_Adder_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## `time_adder.class.asp`'s avaible function

- Function to add two times -> `Public Function sum_times(ByVal first_time, ByVal second_time, ByVal checks, ByVal time_indicator, ByVal separation_character)`
  >
  > - `first_time`: The first time value to add (as a string).
  > - `second_time`: The second time value to add (as a string).
  > - `checks`: A boolean parameter. If `true`, the function verifies that the provided times are valid.
  > - `time_indicator`: The unit of time, which can be:
  >   - `h` -> hours
  >   - `m` -> minutes
  >   - `s` -> seconds
  > - `separation_character`: The character used to separate the components of the time. For example, in the time "12:10:24", the `separation_character` is `":"`.

## How to use 

> From `Test.asp`

1. Initialize the class
   ```asp
     <%@LANGUAGE="VBSCRIPT"%>
     <!--#include file="time_adder.class.asp"-->
     <%
          Dim timeAdder
          Set timeAdder = new time_adder
   ```
   
2. Use the class
   ```asp
    Dim time1
    Dim time2 

    time1 = "12:15:30"
    time2 = "13:45:15"
   
    Response.write timeAdder.sum_times(time1, time2, True, "h", ":") & "<br>"
   %>
   ```

   <picture>
    <source media="(prefers-color-scheme: dark)"srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDark.svg">
    <source media="(prefers-color-scheme: light)"srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAILight.svg">
    <img alt="Not made by AI" src="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDefault.svg">
  </picture>

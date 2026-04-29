# List out dates between a range of dates in Classic ASP

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/122c5874b8a24300b7b0d8a957761557)](https://app.codacy.com/gh/R0mb0/List_out_dates_between_a_range_of_dates_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/https://github.com/R0mb0/List_out_dates_between_a_range_of_dates_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/List_out_dates_between_a_range_of_dates_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

---

<div align="center">
  <a href="http://paypal.me/R0mb0">
    <picture>
      <source media="(prefers-color-scheme: dark)" srcset="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Dark.svg">
      <source media="(prefers-color-scheme: light)" srcset="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Light.svg">
      <img alt="Saved you time? Support the dev" src="https://github.com/R0mb0/Support_the_dev_badge/blob/main/Badge/SVG/Support_the_dev_badge_Default.svg">
    </picture>
  </a>
</div>


## `List_out_dates.class.asp`'s avaible function

- List out dates -> `Public Function extractDates(start_date, end_date, selector, separator, month_name, abbreviate)` - The function returns an array with all dates.
  >
  > - <ins>Where the selector could be:</ins>
  >   - "y" for Years
  >   - "m" for Months
  >   - "d" for Days
  >
  > - <ins>Where the separator could be:</ins>
  >   - An arbitrary symbol to separate date elements
  >
  > - <ins>Where the month_name coud be:</ins>
  >   - True for use MonthName function
  >   - False for don't use MonthName function
  >
  > - <ins>Where abbreviate could be:</ins>
  >   - True for abbreviate
  >   - False for don't abbreviate

## How to use 

> From `Test.asp`

1. Initialize the class
   ```asp
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="list_out_dates.class.asp"-->
    <%
      Dim dates
      Set dates = New  listOutDates
   ```

2. Create a start date and a end date
   ```asp
    Dim start_date
    start_date = "07/02/2025 11:26:46" 
    Dim end_date
    end_date = "02/04/2027 15:06:30" 
   ```
3. List out all dates from range
   ```asp
    Dim temp 
    For Each temp In dates.extractDates(start_date, end_date, "d", "/", True, False)
        Response.write(temp & "<br>")
    Next
   %>
   ```

<a href="https://github.com/R0mb0/Not_made_by_AI">
  <picture>
    <source media="(prefers-color-scheme: dark)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDark.svg">
    <source media="(prefers-color-scheme: light)" srcset="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAILight.svg">
    <img alt="Not made by AI" src="https://github.com/R0mb0/Not_made_by_AI/blob/main/Badge/SVG/NotMadeByAIDefault.svg">
  </picture>
</a>

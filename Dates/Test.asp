<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="List_out_dates.asp"-->
<%
    Dim dates
    Set dates = New  listOutDates

    Dim start_date
    start_date = "07/02/2025 11:26:46" 
    Dim end_date
    end_date = "02/04/2027 15:06:30" 

    Dim temp 
    For Each temp In dates.extractDates(start_date, end_date, "m")
        Response.write(temp & "<br>")
    Next
%>
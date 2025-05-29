<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="time_adder.class.asp"-->
<%
    'Initializating the class 
    Dim timeAdder
    Set timeAdder = new time_adder

    'Preparing test
    Dim time1
    Dim time2 

    time1 = "12:15:30"
    time2 = "13:45:15"

    Response.write "<h3> Test 1 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, False, "h", ":") & "<br>"

    Response.write "<h3> Test 2 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, True, "h", ":") & "<br>"

    time1 = "12:15"
    time2 = "13:45"

    Response.write "<h3> Test 3 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, False, "h", ":") & "<br>"

    time1 = "12"
    time2 = "13"

    Response.write "<h3> Test 4 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, True, "h", ":") & "<br>"

    time1 = "12:15"
    time2 = "13:45"

    Response.write "<h3> Test 5 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, False, "m", ":") & "<br>"

    time1 = "12"
    time2 = "13"

    Response.write "<h3> Test 6 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, True, "m", ":") & "<br>"

    Response.write "<h3> Test 7 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, True, "s", "") & "<br>"

    time1 = "12:15"
    time2 = "13:45"

    Response.write "<h3> Test 8 </h3><br>"
    Response.write timeAdder.sum_times(time1, time2, True, "s", ":") & "<br>"
%>

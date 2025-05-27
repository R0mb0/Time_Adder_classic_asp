<%
    Class time_adder

        'Initialization and destruction'
        Sub class_initialize()
        End Sub 

        Sub class_terminate()
        End Sub 

        'Private function to check if a string could be a time 
        Private Function is_time(ByVal time)
            Dim index 
            Dim char 
            Dim ascii 
            For index = 1 To Len(time) 
                char = Mid(time, index, 1)
                ascii = Asc(char)
                'Checking if there are capital letters or lowercase letters
                If(ascii >= 65 And ascii <= 90) Or (ascii >= 97 And ascii <= 122) Then
                    'Return
                    is_time = False
                    Exit Function 
                End If 
            Next
            'Return
            is_time = True
            Exit Function 
        End Function 

        'Private function to count how many time a character is in the time string 
        Private Function count_character_presence(ByVal time, ByVal separation_character)
            Dim index 
            Dim char 
            Dim count 
            count = 0
            For index = 1 To Len(time)
                char = Mid(time, i, 1)
                If(char = separation_character) Then 
                    count = count + 1 
                End If 
            Next 
            'return 
            count_character_presence = conta
        End Function 

        'Private function to check the integrity of a time 
        Private Function check_integrity(ByVal time, ByVal separation_character)
            'Check if in the time string is present the separaion character
            If(InStr(time, separation_character) > 0) Then 
                Select Case count_character_presence(time, separation_character)
                    Case 1
                        If(Len(time) > 2 And Len(time) < 6) Then 
                            'Return
                            check_integrity = True
                            Exit Function 
                        Else
                            'Return
                            check_integrity = False
                            Exit Function
                        End If 
                    Case 2
                        If(Len(time) > 4 And Len(time) < 9) Then 
                            'Return
                            check_integrity = True
                            Exit Function 
                        Else
                            'Return
                            check_integrity = False
                            Exit Function
                        End If 
                    Case Else 
                        'Return
                        check_integrity = False
                        Exit Function 
                End Select
            Else
                If(Len(time) > 2) Then 
                    'Return
                    check_integrity = False
                    Exit Function 
                Else 
                    'Return
                    check_integrity = True
                    Exit Function 
                End If 
            End If 
        End Function

        'Public function to convert a time in seconds
        Private Function tranforms_into_seconds(ByVal time, ByVal time_indicator, ByVal separation_character)
        Dim temp
            Select Case time_indicator
                Case "s"
                    tranforms_into_seconds = CInt(time)
                    Exit Function 
                Case "m"
                    temp = Split(time, separation_character)
                    Select Case UBound(temp)
                        Case 0
                            tranforms_into_seconds = (CInt(time) * 60)
                            Exit Function 
                        Case 1
                            tranforms_into_seconds = (CInt(temp(0)) * 60) + CInt(temp(1))
                            Exit Function
                        Case Else 
                            Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Wrong time indicator for the time passed")
                    End Select
                Case "h"
                    temp = Split(time, separation_character)
                    Select Case UBound(temp)
                        Case 0
                            tranforms_into_seconds = (CInt(time) * 3600)
                            Exit Function 
                        Case 1
                            tranforms_into_seconds = (CInt(temp(0)) * 3600) + CInt(temp(1) * 60)
                            Exit Function
                        Case 2 
                            tranforms_into_seconds = (CInt(temp(0)) * 3600) + CInt(temp(1) * 60) +  CInt(temp(2))
                            Exit Function
                        case Else 
                            Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Wrong time indicator for the time passed")
                            Exit Function
                    End Select 
                Case Else 
                    Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Wrong time indicator")
                    Exit Function
            End Select 
        End Function 

        'Public function to return the sum 
        Public Function sum_times(ByVal first_time, ByVal second_time, ByVal checks, ByVal time_indicator, ByVal separation_character)
            'if true checking if times are different
            If(checks) Then 
                'checking if the time strings could be a time
                If(Not(is_time(first_time)) Or Not(is_time(second_time))) Then 
                    Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Time is not correct")
                End If 

                'checking if times are good
                If(Not(check_integrity(first_time, separation_character)) Or Not(check_integrity(second_time, separation_character))) Then 
                    Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Time is not good")
                End If 

            End If 
            'Transforming time in seconds 
            Dim totalTime
            totalTime = tranforms_into_seconds(first_time) + tranforms_into_seconds(second_time)
            Dim hours
            Dim min
            Dim sec
            Select Case 
                Case "s"
                    sum_times = totalTime
                    Exit Function 
                Case "m"

                Case "h"
                Case Else 
                    Call Err.Raise(vbObjectError + 10, "time_adder.class","sum_times - Wrong time indicator")
                    Exit Function
            End Select
        End Function 

    End Class 
%> 

<% 
class listOutDates

    ' Initialization and destruction'
	Sub class_initialize()
	End Sub
	
	Sub class_terminate()
	End Sub

    'Selector type'
    '- y = Year'
    '- m = Month'
    '- d = Day'

    'Separator'
    '- An arbitrary symbol to separate date elements'

    'Month_name'
    '- True = use MonthName function
    '- False = don't use MonthName function

    'Abbreviate'
    '- True = abbreviate
    '- False = don't abbreviate

    'Example of date: 07/01/2025 11:26:46 

    Public Function extractDates(start_date, end_date, selector, separator, month_name, abbreviate) 'Selector could be one of the listed

        'Check dates order'
        If cdate(end_date) < cdate(start_date) Then 
            Call Err.Raise(vbObjectError + 10, "List_out_dates.class - extractDates", "The dates are in wrong order")
        End If 

        Dim arr 
        arr = Array()
        Dim index 
        index = 0
        Dim temp_date
        Dim A
        Dim B
        Dim C
        'Cases'
        Select Case selector
            'Extract Years'
            Case "y"
                Select Case DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    Case 0
                        Redim Preserve arr(index)
                        arr(index) = Year(cdate(start_date)) + A
                        index = index + 1 
                    Case else
                        For A = 0 To  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                            Redim Preserve arr(index)
                            arr(index) = Year(cdate(start_date)) + A
                            index = index + 1 
                        Next
                End Select 
            'Extract Months'
            Case "m"
                Select Case DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    Case 0
                        temp_date = start_date
                        For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                Redim Preserve arr(index)
                                If month_name Then 
                                    arr(index) = MonthName((Month(cdate(temp_date)) + B), abbreviate) & separator & (Year(cdate(temp_date)))
                                Else
                                    arr(index) = (Month(cdate(temp_date)) + B) & separator & (Year(cdate(temp_date)))
                                End If
                                index = index + 1 
                            Next 
                    Case else 
                        temp_date = start_date
                        For A = 0 to  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                            'Check if the date is at start or end'
                            Select Case Year(cdate(temp_date))
                                Case Year(cdate(start_date))
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                        Redim Preserve arr(index)
                                        If month_name Then 
                                            arr(index) = MonthName((Month(cdate(temp_date)) + B), abbreviate) & separator & (Year(cdate(temp_date)))
                                        Else
                                            arr(index) = (Month(cdate(temp_date)) + B) & separator & (Year(cdate(temp_date)))
                                        End If
                                        index = index + 1 
                                    Next 
                                    temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                                Case Year(cdate(end_date))
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                        Redim Preserve arr(index)
                                        If month_name Then 
                                            arr(index) = MonthName((Month(cdate(temp_date)) + B), abbreviate) & separator & (Year(cdate(temp_date)))
                                        Else
                                            arr(index) = (Month(cdate(temp_date)) + B) & separator & (Year(cdate(temp_date)))
                                        End If
                                        index = index + 1 
                                    Next 
                                    'At this point the loop should be concluded'
                                    Exit For 
                                Case else
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                        Redim Preserve arr(index)
                                        If month_name Then 
                                            arr(index) = MonthName((Month(cdate(temp_date)) + B), abbreviate) & separator & (Year(cdate(temp_date)))
                                        Else
                                            arr(index) = (Month(cdate(temp_date)) + B) & separator & (Year(cdate(temp_date)))
                                        End If
                                        index = index + 1 
                                    Next
                                temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                            End Select 
                        Next
                End Select     
            'Extract Days'
            Case "d"
                Select Case DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    Case 0
                        temp_date = start_date
                        For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                Select Case  Month(cdate(temp_date))
                                    Case Month(cdate(end_date))
                                        For C = 0 To DateDiff("d", cdate(temp_date), cdate(end_date))
                                            Redim Preserve arr(index)
                                            If month_name Then 
                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                            Else
                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                            End If
                                            index = index + 1
                                        Next
                                        'At this point the loop should be concluded'
                                        'Exit For
                                    Case else
                                        'In the last year but not in the same month'
                                        Select Case Month(cdate(temp_date))
                                            Case 12
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("31/" & Month(cdate(temp_date)) & separator & Year(cdate(temp_date)) & " 00:00:00"))
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                            Case else 
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")) -1
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                        End Select
                                        temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")
                                End Select
                            Next 
                    Case else 
                        temp_date = start_date
                        For A = 0 to  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                            'Check if the date is at start or end'
                            Select Case Year(cdate(temp_date))
                                Case Year(cdate(start_date))
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                        'Privileged case, starting from first day'
                                        Select Case Month(cdate(temp_date))
                                            Case 12
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("31/" & Month(cdate(temp_date)) & separator & Year(cdate(temp_date)) & " 00:00:00"))
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                            Case else 
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")) -1
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                        End Select
                                     temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")
                                    Next 
                                    temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                                Case Year(cdate(end_date))
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                        Select Case  Month(cdate(temp_date))
                                            Case Month(cdate(end_date))
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate(end_date))
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                                'At this point the loop should be concluded'
                                                Exit For
                                            Case else
                                                'In the last year but not in the same month'
                                                Select Case Month(cdate(temp_date))
                                                    Case 12
                                                        For C = 0 To DateDiff("d", cdate(temp_date), cdate("31/" & Month(cdate(temp_date)) & separator & Year(cdate(temp_date)) & " 00:00:00"))
                                                            Redim Preserve arr(index)
                                                            If month_name Then 
                                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                            Else
                                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                            End If
                                                            index = index + 1
                                                        Next
                                                    Case else 
                                                        For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")) -1
                                                            Redim Preserve arr(index)
                                                            If month_name Then 
                                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                            Else
                                                                arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                            End If
                                                            index = index + 1
                                                        Next
                                                End Select
                                                temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")
                                        End Select
                                    Next 
                                    'At this point the loop should be concluded'
                                    Exit For 
                                Case else
                                    'Standard case, is the same of privileged case'
                                    For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                        Select Case Month(cdate(temp_date))
                                            Case 12
                                                 For C = 0 To DateDiff("d", cdate(temp_date), cdate("31/" & Month(cdate(temp_date)) & separator & Year(cdate(temp_date)) & " 00:00:00"))
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                            Case else 
                                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")) -1
                                                    Redim Preserve arr(index)
                                                    If month_name Then 
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & MonthName(Month(cdate(temp_date)), abbreviate) & separator & (Year(cdate(temp_date)))
                                                    Else
                                                        arr(index) = (Day(cdate(temp_date)) + C) & separator & (Month(cdate(temp_date))) & separator & (Year(cdate(temp_date)))
                                                    End If
                                                    index = index + 1
                                                Next
                                        End Select
                                        temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & separator & Year(cdate(temp_date)) & " 00:00:00")
                                    Next
                                    temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                            End Select 
                        Next
                End Select
        End Select
        'Return statement'
        extractDates = arr 
    End Function
End Class 
%> 

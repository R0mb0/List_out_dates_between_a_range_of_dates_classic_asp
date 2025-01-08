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

    'Example of date: 07/01/2025 11:26:46 

    Public Function extractDates(start_date, end_date, selector) 'Selector could be one of the listed

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
                For A = 0 To  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    Redim Preserve arr(index)
                    arr(index) = Year(cdate(start_date)) + A
                    index = index + 1 
                Next
            'Extract Months'
            Case "m"
                temp_date = start_date
                For A = 0 to  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    'Check if the date is at start or end'
                    Select Case Year(cdate(temp_date))
                        Case Year(cdate(start_date))
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                Redim Preserve arr(index)
                                arr(index) = (Month(cdate(temp_date)) + B) & "/" & (Year(cdate(temp_date)))
                                index = index + 1 
                            Next 
                            temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                        Case Year(cdate(end_date))
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                Redim Preserve arr(index)
                                arr(index) = (Month(cdate(temp_date)) + B) & "/" & (Year(cdate(temp_date)))
                                index = index + 1 
                            Next 
                            'At this point the loop should be concluded'
                            Exit For 
                        Case else
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                Redim Preserve arr(index)
                                arr(index) = (Month(cdate(temp_date)) + B) & "/" & (Year(cdate(temp_date)))
                                index = index + 1 
                            Next
                            temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                    End Select 
                Next
            'Extract Days'
            Case "d"
                temp_date = start_date
                For A = 0 to  DateDiff("yyyy",cdate(start_date),cdate(end_date))
                    'Check if the date is at start or end'
                    Select Case Year(cdate(temp_date))
                        Case Year(cdate(start_date))
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                'Privileged case, starting from first day'
                                For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")) -1
                                    Redim Preserve arr(index)
                                    arr(index) = (Day(cdate(temp_date)) + C) & "/" & (Month(cdate(temp_date))) & "/" & (Year(cdate(temp_date)))
                                    index = index + 1
                                Next
                                temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")
                            Next 
                            temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                        Case Year(cdate(end_date))
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate(end_date))
                                Select Case  Month(cdate(temp_date))
                                    Case Month(cdate(end_date))
                                        For C = 0 To DateDiff("d", cdate(temp_date), cdate(end_date))
                                            Redim Preserve arr(index)
                                            arr(index) = (Day(cdate(temp_date)) + C) & "/" & (Month(cdate(temp_date))) & "/" & (Year(cdate(temp_date)))
                                            index = index + 1
                                        Next
                                        'At this point the loop should be concluded'
                                        Exit For
                                    Case else
                                        'In the last year but not in the same month'
                                        For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")) -1
                                            Redim Preserve arr(index)
                                            arr(index) = (Day(cdate(temp_date)) + C) & "/" & (Month(cdate(temp_date))) & "/" & (Year(cdate(temp_date)))
                                            index = index + 1
                                        Next
                                    temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")
                                End Select
                            Next 
                            'At this point the loop should be concluded'
                            Exit For 
                        Case else
                            'Standard case, is the same of privileged case'
                            For B = 0 To DateDiff("m",cdate(temp_date),cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")) -1 
                                 For C = 0 To DateDiff("d", cdate(temp_date), cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")) -1
                                    Redim Preserve arr(index)
                                    arr(index) = (Day(cdate(temp_date)) + C) & "/" & (Month(cdate(temp_date))) & "/" & (Year(cdate(temp_date)))
                                    index = index + 1
                                Next
                                temp_date = cdate("01/" & Month(cdate(temp_date)) + 1 & "/" & Year(cdate(temp_date)) & " 00:00:00")
                            Next
                            temp_date = cdate("01/01/" & Year(cdate(temp_date))+1 & " 00:00:00")
                    End Select 
                Next
        End Select
        'Return statement'
        extractDates = arr 
    End Function

End Class 
%> 
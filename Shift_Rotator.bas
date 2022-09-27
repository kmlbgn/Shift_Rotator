Attribute VB_Name = "Module1"

Option Explicit

Dim RotationArray() As String
Dim NamesArray()
Dim reset             As Integer
Dim NumberOfJobs      As Integer
Dim NumberOfEmployees As Integer
Dim N                 As Integer
Dim m                 As Integer
Dim k                 As Integer
Dim counter           As Integer

' For Archipelagos
' March 2020
' Kamil Begny
' kamil.begny@gmail.com

Sub RandomizeSchedule()
    Dim TempArray   As Variant
    Dim j           As Integer
    Dim random_pick As Integer
    

    NumberOfEmployees = Application.WorksheetFunction.CountA( _
                        Sheets("Names").Range("A:A")) - 1
    
    NumberOfJobs = Application.WorksheetFunction.CountA(Sheets("OFFICE"). _
                   Range("1:1"))
    
    ReDim RotationArray(1 To 7, 1 To NumberOfJobs) '7 days * number of jobs
    
    ' Copy previous weeks to 2 weeks ago first
    Sheets("OFFICE").Range(Cells(20, 3), Cells(1 + 2 * 2 + 3 * UBound(RotationArray, 1), 2 + UBound(RotationArray, 2))) = Sheets("OFFICE").Range(Cells(13, 3), Cells(1 + 2 * 2 + 2 * UBound(RotationArray, 1), 2 + UBound(RotationArray, 2))).Value
    
    ' Then move this week to previous week
    Sheets("OFFICE").Range(Cells(13, 3), Cells(1 + 2 * 2 + 2 * UBound(RotationArray, 1), 2 + UBound(RotationArray, 2))) = Sheets("OFFICE").Range(Cells(3, 3), Cells(2 + UBound(RotationArray, 1), 2 + UBound(RotationArray, 2))).Value
    
    ' Iterating the schedule (final array) cell by cell
    For k = 1 To 7
        For N = 1 To NumberOfJobs
                CountPreviousShifts
            
                ' Select person for shift
                SelectEmployee
                
                'We randomly pick within the first quarter of the remaining people in the array (For a bit of randomness and mixing)
                Randomize
                random_pick = Int((UBound(NamesArray) \ 4) * Rnd) + 1
                            
                ' Put his name on final array
                RotationArray(k, N) = NamesArray(random_pick)
                
        Next N
    Next k
    
    ' Write final array to desired schedule sheet
    Sheets("OFFICE").Range(Cells(3, 3), Cells(2 + UBound(RotationArray, 1), 2 + UBound(RotationArray, 2))) = RotationArray()
End Sub

Function SelectEmployee()
    Dim Itemk       As Variant
    Dim Itemn       As Variant
    Dim res         As Variant
    Dim TempArray   As Variant
    Dim h           As Integer
    Dim i           As Integer
 
    ' Get rid of people who already worked that shift this week, except first row to save time
    If Not (k = 1) Then
    
        ' Dinner
        If (k > 1 And N >= 6 And N <= 8) Then
            For i = 6 To 8
                For Each Itemn In Application.Index(RotationArray, 0, i)
                    res = Application.Match(Itemn, NamesArray, False)
                    If Not (IsError(res)) Then
        
                        NamesArray(Application.Match(Itemn, NamesArray, False)) = " "
                        
                    End If
                Next Itemn
            Next
        
        ' Lunch
        ElseIf (k > 1 And N >= 4 And N <= 5) Then
            For i = 4 To 5
                For Each Itemn In Application.Index(RotationArray, 0, i)
                    res = Application.Match(Itemn, NamesArray, False)
                    If Not (IsError(res)) Then
        
                        NamesArray(Application.Match(Itemn, NamesArray, False)) = " "
                        
                    End If
                Next Itemn
            Next
            
        ' Breakfast
        ElseIf (k > 1 And N >= 2 And N <= 3) Then
            For i = 2 To 3
                For Each Itemn In Application.Index(RotationArray, 0, i)
                    res = Application.Match(Itemn, NamesArray, False)
                    If Not (IsError(res)) Then
        
                        NamesArray(Application.Match(Itemn, NamesArray, False)) = " "
                        
                    End If
                Next Itemn
            Next
            
        ' Water
        ElseIf (k > 1 And N >= 12 And N <= 13) Then
            For i = 12 To 13
                For Each Itemn In Application.Index(RotationArray, 0, i)
                    res = Application.Match(Itemn, NamesArray, False)
                    If Not (IsError(res)) Then
        
                        NamesArray(Application.Match(Itemn, NamesArray, False)) = " "
                        
                    End If
                Next Itemn
            Next
        
        ' Any other single shift
        Else
            For Each Itemn In Application.Index(RotationArray, 0, N)
                res = Application.Match(Itemn, NamesArray, False)
                If Not (IsError(res)) Then
    
                    NamesArray(Application.Match(Itemn, NamesArray, False)) = " "
                    
                End If
            Next Itemn
        End If
    End If
    
    ' Get rid of people who are already working today, except first column
    If Not (N = 1) Then
        For Each Itemk In Application.Index(RotationArray, k, 0)
            res = Application.Match(Itemk, NamesArray, False)
            If Not (IsError(res)) Then
            
                NamesArray(Application.Match(Itemk, NamesArray, False)) = " "
                
            End If
        Next Itemk
    End If

    ' Recreate array and get rid of empty " " (deleted people)
    TempArray = Split(Application.WorksheetFunction.Trim(Join(NamesArray, " ")))
 
    ' If array is empty, there is no solution so we force someone
    If UBound(TempArray) = -1 Then
            ForceSelect
    Else

    ' Copy temp array on main array, first person to figure is selected
        ReDim NamesArray(LBound(NamesArray) To UBound(TempArray) + 1)
        For h = LBound(NamesArray) To UBound(NamesArray)
            NamesArray(h) = TempArray(h - 1)
        Next h
    End If
    
End Function

' Unfortunately depending on the # of people, some may have to do 2x the same shift this week,
' but we still avoid two shifts the same day.
Function ForceSelect()
    Dim TempArray   As Variant
    Dim Itemk       As Variant
    Dim Itemn       As Variant
    Dim res         As Variant
    Dim h           As Variant
    
    ' Reconstruct array with ALL names.
    Erase NamesArray
    CountPreviousShifts
    
    ' Here we remove people who already work that day to avoid double shift
    For Each Itemk In Application.Index(RotationArray, k, 0)
        res = Application.Match(Itemk, NamesArray, False)
        If Not (IsError(res)) Then
        
            NamesArray(Application.Match(Itemk, NamesArray, False)) = " "
            
        End If
    Next Itemk
    
    TempArray = Split(Application.WorksheetFunction.Trim(Join(NamesArray, " ")))

    ' Copy temp array on main array, first person is selected
    ReDim NamesArray(LBound(NamesArray) To UBound(TempArray) + 1)
    For h = LBound(NamesArray) To UBound(NamesArray)
        NamesArray(h) = TempArray(h - 1)
    Next h
    
    Debug.Print ("Double shift needed this week for : " & NamesArray(1))
End Function

Sub CountPreviousShifts()
    Dim Arr() As Variant
    Dim ArrTemp As Variant
    Dim Temp1 As Variant
    Dim Temp2 As Variant
    Dim pastShift As Variant
    Dim i As Long
    Dim j As Long
    Dim ii As Long

    ReDim NamesArray(1 To NumberOfEmployees)
    
    NamesArray() = Application.Transpose(Sheets("Names").Range("A2:A" & (1 + NumberOfEmployees)).Value)
    
    ReDim Arr(1 To NumberOfEmployees, 1 To 2)
    
     If (k > 1 And N >= 2 And N <= 3) Then
        pastShift = Join(Application.Transpose(Range("D13:D26")), " ") & " " & Join(Application.Transpose(Range("E13:E26")), " ")
        ArrTemp = Split(Join(Application.Transpose(Application.Index(RotationArray, 0, 2)), " ") & " " & Join(Application.Transpose(Application.Index(RotationArray, 0, 3)), " ") & " " & pastShift, " ")
        
        
    ElseIf (k > 1 And N >= 4 And N <= 5) Then
        pastShift = Join(Application.Transpose(Range("F13:F26")), " ") & " " & Join(Application.Transpose(Range("G13:G26")), " ")
        ArrTemp = Split(Join(Application.Transpose(Application.Index(RotationArray, 0, 4)), " ") & " " & Join(Application.Transpose(Application.Index(RotationArray, 0, 5)), " ") & " " & pastShift, " ")
        
    ElseIf (k > 1 And N >= 6 And N <= 8) Then
        pastShift = Join(Application.Transpose(Range("H13:H26")), " ") & " " & Join(Application.Transpose(Range("I13:I26")), " ") & " " & Join(Application.Transpose(Range("J13:J26")), " ")
        ArrTemp = Split(Join(Application.Transpose(Application.Index(RotationArray, 0, 6)), " ") & " " & Join(Application.Transpose(Application.Index(RotationArray, 0, 7)), " ") & " " & Join(Application.Transpose(Application.Index(RotationArray, 0, 8)), " ") & " " & pastShift, " ")

    ElseIf (k > 1 And N >= 12 And N <= 13) Then
        pastShift = Join(Application.Transpose(Range("N13:N26")), " ") & " " & Join(Application.Transpose(Range("O13:O26")), " ")
        ArrTemp = Split(Join(Application.Transpose(Application.Index(RotationArray, 0, 12)), " ") & " " & Join(Application.Transpose(Application.Index(RotationArray, 0, 13)), " ") & " " & pastShift, " ")
        
    Else
        pastShift = Join(Application.Transpose(Range(Cells(13, N + 2), Cells(26, N + 2))), " ")
        ArrTemp = Split(Join(Application.Transpose(Application.Index(RotationArray, 0, N)), " ") & " " & pastShift, " ")
    End If

    For ii = 1 To NumberOfEmployees
        Arr(ii, 1) = NamesArray(ii) ' names
        Arr(ii, 2) = Application.Count(Application.Match(ArrTemp, Array(NamesArray(ii)), 0)) ' shift count
    Next ii
    
    ' Sort the array using the bubble sort method
    For i = LBound(Arr, 1) To UBound(Arr, 1) - 1
        For j = i + 1 To UBound(Arr, 1)
            If Arr(i, 2) > Arr(j, 2) Then
                Temp1 = Arr(j, 1)
                Temp2 = Arr(j, 2)
                Arr(j, 1) = Arr(i, 1)
                Arr(j, 2) = Arr(i, 2)
                Arr(i, 1) = Temp1
                Arr(i, 2) = Temp2
            End If
        Next j
    Next i
    
    NamesArray = Application.Transpose(Application.Index(Arr, 0, 1)) ' Transpose otherwise Array is vertical and not compatible with NamesArray

End Sub






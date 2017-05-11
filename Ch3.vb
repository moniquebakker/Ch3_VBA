Option Explicit

Sub SpecialistResourceAllocation()   ' this macro should run on every Friday PM. This is also when we check annual leave arrAL(Specialist) for 6 weeks ahead, and block it.
    Const J As Integer = 6    ' 6 specialiActivity
    Const TotalWeeks As Integer = 52
    Dim arrActivityTotalCountToPSstart() As Variant
    Dim arrActivityTotalCountToPSend() As Variant
    Dim arrActivityEmpiricalTotals() As Variant
    Dim arrActivityEmpiricalNonALWeekAverage() As Variant
    Dim arrEmpWeeklyActivityCount() As Variant
    Dim arrActivity() As Variant  ' array that holds the candidate to-be-scheduled Activity
    Dim arrActivityWeeklyCountToPSend() As Variant
    Dim arrMeansAfterAllocationTillPSend() As Variant, arrMeansAfterAllocationTillPSendMin1() As Variant, arrMeansAfterAllocationTillPSendIfIs0() As Variant
    Dim arrStDevsAfterAllocationTillPSend() As Variant, arrStDevsAfterAllocationTillPSendMin1() As Variant, arrStDevsAfterAllocationTillPSendIfIs0() As Variant
    Dim arrCVsAfterAllocationTillPSend() As Variant, arrCVsAfterAllocationTillPSendMin1() As Variant, arrCVsAfterAllocationTillPSendIfIs0() As Variant
    Dim arrDeltaCV() As Variant
    Dim arrPSTempFreq() As Integer
    Dim arrPSDefFreq() As Integer
    Dim arrCouldNotAllocateFreq() As Variant
    Dim MyWorkbook As Workbook
    Dim wsActivitychoose As Worksheet
    Dim wsarrEmpCVs As Worksheet
    Dim wsarrPerfSchCVs As Worksheet
    Dim OriginalCalendarSheet As Worksheet
    Dim CVrange, DeltaCVrange As Range
    Dim Rng1, Rng2, RngU, RangeTarget As Range
    Dim Dimension1, Activity, ActivityIndex, Specialist, WeeklyCount, Instance As Integer
    Dim PSend, PSstart As Integer
    Dim StartCount As Integer
    Dim Week As Integer
    Dim PriorityCol As Integer, PrefHalfDay, PrefHalfDay2 As Integer, Startcell As Integer
    Dim TempFreqSum() As Variant
    Dim TempFreqSum1 As Integer
    Dim DefFreqSum() As Variant
    Dim DefFreqSum1 As Integer
    Dim UniqueLargestFreqColNum As Integer
    Dim RangeRecommCalc As Range
    Dim arrStDevsRecomm() As Variant
    Dim arrMeansRecomm() As Variant
    Dim arrCVsRecomm() As Variant
    Dim arrStDevsZeros() As Variant
    Dim arrMeansZeros() As Variant
    Dim arrCVsZeros() As Variant
    Dim CVsrng As Range
    Dim Counter As Integer
    Dim i As Integer
    Dim g As Integer
    Dim BiggestDegradationColNum As Integer
    Dim DegradationRng As Range
    Dim LargestValue As Integer
    Dim startrow As Integer
    Dim NatHolidaysRange As Range
    Dim HolidayRange As Range
    Dim rc As Range
    Dim RowNo As Integer
    Dim HowManyHalfDays As Integer
    Dim ALWeekNo As Integer
    Dim CurrentActiveHalfDayCellsCount() As Variant
    Dim wCell As Range
    Dim DOTW As Integer
    Dim CountAlgoInstances As Integer
    Dim RowLeftOff() As Variant
    Dim InitialScheduleLength As Integer
    Dim RowNum As Integer
    Dim Sheet1, SheetSDO As Worksheet
    Dim SDO1(), SDO2() As Variant
    Dim rCell As Range
    Dim WhatWeShouldHave() As Variant
    Dim WhatWeNowHave() As Variant
    Dim AnnualLeaveFinished() As Boolean
    Dim WhatWeShouldHaveTotal() As Variant
    Dim WhatWeNowHaveTotal() As Variant
    Dim AvailSpec3 As Integer
    Dim AvailSpec6 As Integer
    Dim DefAllocRange As Range
    Dim arrEmpiricalActiveHalfDayAverage(1 To 8, 1 To J) As Double ' using type Double for the decimal values should be ok.
    Dim arrEmpiricalActiveHalfDays(1 To J) As Integer
    Dim exists, EmpTotals2011exists, EmpTotalsCyclicexists As Boolean
    Dim NatHolidayReady As Boolean
    Dim ALReady As Boolean
    Dim Ranking() As Variant
    Dim w As Integer
    Dim arrActivityPhantom() As Variant
    Dim Phantomdata() As Variant
    Dim UseActualStdevs As Boolean
    Dim CopyAnnualLeave As Boolean
    Dim Equalizers As Variant
    Dim arrActivityMinMax As Variant
    Dim iterationcount As Integer
    Dim enforcemin, enforcemax As Boolean
    Dim Activity1PrefHalfDay1, Activity1PrefHalfDay2 As String
    Dim needtochange99sto1s As Boolean
    Dim starttime As Date
    Dim EmpTotalsCyclic  As Boolean
    Dim Twos, Threes As Integer
    Dim arrActivityIndex As Variant
    Dim Z As Integer
    Dim NHALLOW, NHALHIGH As Boolean
    Dim EqualizersSheet As String
    Dim DemandDrivenEmpTotals As Boolean
    
    starttime = Now
    Debug.Print Now & " start of run"
    
    ReDim AnnualLeaveFinished(1 To J)
    For Specialist = 1 To J
        AnnualLeaveFinished(Specialist) = False
    Next Specialist
    
    ThisWorkbook.Activate
    Set wsActivitychoose = ThisWorkbook.Worksheets("Activitychoose")
    Set MyWorkbook = ThisWorkbook
    
    ThisWorkbook.Worksheets("Activitychoose").Activate
    arrActivity = wsActivitychoose.Range("A1", Range("A1").End(xlDown).End(xlToRight)) ' Activity type specification array    Dimension1 = UBound(Activity, 1)  ' dimension 1 of Activity is the number of rows, i.e. 9 Activity types (rows)
    Dimension1 = UBound(arrActivity, 1)  ' dimension 1 of Activity is the number of rows, i.e. 9 Activity types (rows)
    
    ' Select the correct 'Friday PM' cell before continuing
    startrow = 80
    
    PSstart = startrow + 5
    PSend = startrow + 5 + 13
    InitialScheduleLength = (PSstart - 1) / 2 / 7 ' Initial schedule length in weeks
    NatHolidayReady = True ' indicates whether the empty sheet1 already has the national holidays installed
    ALReady = True   ' indicates whether the empty sheet1 already has the annual leave installed
    CopyAnnualLeave = True   ' True means AL is already in the empty schedule
    needtochange99sto1s = False
    UseActualStdevs = False ' in stead we use the lowest possible stdevs for the first x weeks
    ThisWorkbook.Worksheets("Sheet1").Range("N31").Value = UseActualStdevs
    enforcemin = True
    ThisWorkbook.Worksheets("Sheet1").Range("N32").Value = enforcemin
    enforcemax = True
    ThisWorkbook.Worksheets("Sheet1").Range("N33").Value = enforcemax
    EmpTotalsCyclic = False ' True means use the pure cyclic calendar's empirical totals are used
    DemandDrivenEmpTotals = True ' means we use emtotals calculated as 2011 demand - to #sessions conversion with some multiplier
    ThisWorkbook.Worksheets("Sheet1").Range("N34").Value = EmpTotalsCyclic
    NHALLOW = True
    ThisWorkbook.Worksheets("Sheet1").Range("N35").Value = ""
    If NHALLOW = True Then
        ThisWorkbook.Worksheets("Sheet1").Range("N35").Value = NHALLOW
    Else: ThisWorkbook.Worksheets("Sheet1").Range("N35").Value = NHALHIGH
    End If
    Application.ScreenUpdating = False
    
    '###################### PART A1 populate ActivityMinMax to ensure weelky minimums and maximums
    ThisWorkbook.Worksheets("ActivityMinMax").Activate
    ActiveSheet.Range("A2:C9").Select
    arrActivityMinMax = ThisWorkbook.Worksheets("ActivityMinMax").Range("A2", Range("C2").End(xlDown))
    
    If NHALLOW = True And DemandDrivenEmpTotals = False Then
        EqualizersSheet = "EqualizersLow"
    ElseIf NHALLOW = False And DemandDrivenEmpTotals = False Then
        EqualizersSheet = "EqualizersHigh"
    ElseIf DemandDrivenEmpTotals = True Then
        EqualizersSheet = "DemandDrivenEqualizers"
    End If
        ThisWorkbook.Worksheets(EqualizersSheet).Activate
    Equalizers = ThisWorkbook.Worksheets(EqualizersSheet).Range("A1", Range("A1").End(xlDown).End(xlToRight)).Value
    
     '###################### PART A - COUNT EMPIRICAL TOTALS AND AVERAGES FOR EACH Activity TYPE ###############################
    ' just make an array that hold each specialist's total amount of Annual Leave and active half days over 1 year
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("Sheet1 orig").Activate
        arrEmpiricalActiveHalfDays(Specialist) = _
        Application.WorksheetFunction.CountIfs(ThisWorkbook.Worksheets("Sheet1 orig").Range(Cells(1, Specialist), _
        Cells(52 * 7 * 2, Specialist)), "<10", ThisWorkbook.Worksheets("Sheet1 orig").Range(Cells(1, Specialist), _
        Cells(52 * 7 * 2, Specialist)), ">-1")
    Next Specialist
    
    ReDim arrActivityEmpiricalTotals(1 To Dimension1, 1 To J)
    ReDim arrActivityEmpiricalNonALWeekAverage(1 To Dimension1, 1 To J)
    ThisWorkbook.Worksheets("Sheet1 orig").Activate

    EmpTotals2011exists = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "EmpTotals2011" Then
        EmpTotals2011exists = True
        End If
    Next i
    If EmpTotals2011exists = True Then
        Worksheets("EmpTotals2011").Delete
        EmpTotals2011exists = False
    End If
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "EmpTotalsCyclic" Then
        EmpTotalsCyclicexists = True
        End If
    Next i
    If EmpTotalsCyclicexists = True Then
        Worksheets("EmpTotalsCyclic").Delete
        EmpTotalsCyclicexists = False
    End If
    
    If EmpTotalsCyclic = False And DemandDrivenEmpTotals = False Then
         If Not EmpTotals2011exists Then
             ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "EmpTotals2011"
             ThisWorkbook.Worksheets("Sheet1 Real2011").Activate ' Sheet1 orig does not hold Real2011 emptotals, Sheet1 Real2011 does
             For Specialist = 1 To J ' calculate the EMPIRICAL total number of minutes per session per specialist, for the period from time 0 to one year
                 For Activity = 1 To Dimension1 ' we add the non-pra (Activity 1) a bit later --> always 2 sessions p/w
                     arrActivityEmpiricalTotals(Activity, Specialist) = _
                     Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("Sheet1 Real2011").Range(Cells(1, Specialist), Cells(52 * 7 * 2, Specialist)), arrActivity(Activity, 1))
                 Next Activity
             Next Specialist
             ThisWorkbook.Worksheets("EmpTotals2011").Activate
             Worksheets("EmpTotals2011").Range(Cells(1, 1), Cells(Dimension1, J)).Value = arrActivityEmpiricalTotals
         End If
    ElseIf EmpTotalsCyclic = True Then
        If Not EmpTotalsCyclicexists Then
             ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "EmpTotalsCyclic"
             ThisWorkbook.Worksheets("CyclicNoNHNoAL").Activate ' this is emptotals without NH or AL, so the multipliers need to be much lower than for the Real2011 emptotals case
             For Specialist = 1 To J ' calculate the EMPIRICAL total number of minutes per session per specialist, for the period from time 0 to one year
                 For Activity = 1 To Dimension1 ' we add the non-pra (Activity 1) a bit later --> always 2 sessions p/w
                     arrActivityEmpiricalTotals(Activity, Specialist) = _
                     Application.WorksheetFunction.CountIf(ThisWorkbook.Worksheets("CyclicNoNHNoAL").Range(Cells(1, Specialist), Cells(52 * 7 * 2, Specialist)), arrActivity(Activity, 1))
                 Next Activity
             Next Specialist
             ThisWorkbook.Worksheets("EmpTotalsCyclic").Activate
             Worksheets("EmpTotalsCyclic").Range(Cells(1, 1), Cells(Dimension1, J)).Value = arrActivityEmpiricalTotals
             'Worksheets("EmpTotals2011").Range(Cells(10, 1), Cells(10 + Dimension1, j)).Value = arrActivityEmpiricalNonALWeekAverage
         End If
    End If
    If DemandDrivenEmpTotals = True Then
        ThisWorkbook.Worksheets("DemandDrivenEmpTotals").Activate
        arrActivityEmpiricalTotals = Range("A1", Range("A1").End(xlDown).End(xlToRight))
    End If
      

    '###################### PART B - BLOCK NATIONAL HOLIDAYS FOR 4 yrs ENTIRE SIMULATION PERIOD ###############################
    Dim nhc As Range
    Dim Column As Integer
    Dim h As Integer
    
    If NatHolidayReady = False Then
        Dim arrNatHolidays() As Variant
          ' before we start scheduling we check whether there are any National holidays that need to be blocked in PS week and onwards, till the end of the simulation run
        ThisWorkbook.Worksheets("NatHolidays").Activate
        Set NatHolidaysRange = ThisWorkbook.Worksheets("NatHolidays").Range("F3", Range("G3").End(xlDown))
        ReDim arrNatHolidays(1 To NatHolidaysRange.Count)
        h = 1
        For Each nhc In NatHolidaysRange.Cells
            arrNatHolidays(h) = nhc
            h = h + 1
        Next nhc
        ThisWorkbook.Worksheets("Sheet1").Activate
        For i = 1 To UBound(arrNatHolidays)
            For Column = 1 To 6
                ThisWorkbook.Worksheets("Sheet1").Range(Cells(arrNatHolidays(i), Column), Cells(arrNatHolidays(i), Column)).Value = 10
            Next Column
        Next i
    End If
        

    ' ########################### read EmpTotals2 in while block A and B are in quotes ##########################
    ReDim arrActivityEmpiricalTotals(1 To Dimension1, 1 To J) ' when debugging and reading in from sheet
    ReDim arrActivityEmpiricalNonALWeekAverage(1 To Dimension1, 1 To J)
    
    If EmpTotalsCyclic = False And DemandDrivenEmpTotals = False Then
        ThisWorkbook.Worksheets("EmpTotals2011").Activate
        arrActivityEmpiricalTotals = Range("A1", Range("A1").End(xlDown).End(xlToRight))
    ElseIf EmpTotalsCyclic = True And DemandDrivenEmpTotals = False Then
        ThisWorkbook.Worksheets("EmpTotalsCyclic").Activate
        arrActivityEmpiricalTotals = Range("A1", Range("A1").End(xlDown).End(xlToRight))
    ElseIf DemandDrivenEmpTotals = True Then
        ThisWorkbook.Worksheets("DemandDrivenEmpTotals").Activate
        arrActivityEmpiricalTotals = Range("A1", Range("A1").End(xlDown).End(xlToRight))
    End If
    
    '###################### PART C - COUNT Activity WEEKLY FOR THE FIRST X WEEKS AT INITIALIZATION OF THE CODE  AND WRITE INTO SPECC SHEETS ###############################
    ReDim Preserve arrActivityWeeklyCountToPSend(1 To Dimension1, 1 To J, 1 To (PSend / 2 / 7) - 1)
    ReDim Preserve arrActivityPhantom(1 To Dimension1, 1 To J, 1 To 6)
    ThisWorkbook.Worksheets("Sheet1").Activate ' Sheet1 is the new schedule where I count the PH weeks, not the to-be-scheduled week
    
    For Specialist = 1 To J   ' this loop is for counting the UNTIL PSstart weekly number of Activitys of all Activity types and all specialiActivity for StDev calculation
        For Activity = 1 To Dimension1
            arrEmpiricalActiveHalfDayAverage(Activity, Specialist) = arrActivityEmpiricalTotals(Activity, Specialist) / arrEmpiricalActiveHalfDays(Specialist)
            StartCount = 1
            For WeeklyCount = 1 To (PSend / 2 / 7) - 1 ' the first run, this is week 1-6
                arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount) = _
                Application.WorksheetFunction.CountIf(Range(Cells(StartCount, Specialist), Cells(StartCount + 13, Specialist)), arrActivity(Activity, 1))
                StartCount = StartCount + 14
            Next WeeklyCount
        Next Activity
    Next Specialist
    
    ' write the ideal (perfect) weekly counts of week 1-6 into an array
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("PHSpecc" & Specialist).Activate
        For Activity = 1 To Dimension1
            For WeeklyCount = 1 To 6
                If EmpTotalsCyclic = False Then
                    arrActivityPhantom(Activity, Specialist, WeeklyCount) = ThisWorkbook.Worksheets("PHSpecc" & Specialist).Cells(WeeklyCount, Activity)
                ElseIf EmpTotalsCyclic = True Then
                    ' means if we're using EmpTotalsCyclic (with NH AL High) then the initial 'ideal' frequencies might be different...
                    arrActivityPhantom(Activity, Specialist, WeeklyCount) = ThisWorkbook.Worksheets("PHSpecc" & Specialist).Cells(WeeklyCount + 10, Activity)
                End If
            Next WeeklyCount
        Next Activity
    Next Specialist
                
    'then write the weekly counts into the "Specc" sheets
    For Specialist = 1 To J
        exists = False
        For i = 1 To Worksheets.Count
            If Worksheets(i).Name = "Specc" & Specialist Then
                exists = True
            End If
        Next i
        If exists = True Then
            ThisWorkbook.Worksheets("Specc" & Specialist).Activate
            For Activity = 1 To Dimension1
                For WeeklyCount = 1 To (PSend / 2 / 7) - 1 ' so these three lines don't have to run again every week, only one initial time.
                    ActiveSheet.Cells(WeeklyCount, Activity).Value = arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount)
                Next WeeklyCount
            Next Activity
            ThisWorkbook.Worksheets("PHSpecc" & Specialist).Activate
            Range("A1", Range("A1").End(xlDown).End(xlToRight)).Select
            Phantomdata = Range("A1", Range("A1").End(xlDown).End(xlToRight))
            ThisWorkbook.Worksheets("Specc" & Specialist).Activate
            ThisWorkbook.Worksheets("Specc" & Specialist).Range("K1", Range("K1").Offset(6 - 1, Dimension1 - 1)).Value = Phantomdata
        End If
        If Not exists Then
            ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "Specc" & Specialist ' we're not going to re-create these sheets every time. just add Planweek's counts to the existing sheet
            exists = True
            ThisWorkbook.Worksheets("Specc" & Specialist).Activate
            For Activity = 1 To Dimension1
                For WeeklyCount = 1 To (PSend / 2 / 7) - 1 ' so these three lines don't have to run again every week, only one initial time.
                    ActiveSheet.Cells(WeeklyCount, Activity).Value = arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount)
                Next WeeklyCount
            Next Activity
            ThisWorkbook.Worksheets("PHSpecc" & Specialist).Activate
            Range("A1", Range("A1").End(xlDown).End(xlToRight)).Select
            Phantomdata = Range("A1", Range("A1").End(xlDown).End(xlToRight))
            ThisWorkbook.Worksheets("Specc" & Specialist).Activate
            ThisWorkbook.Worksheets("Specc" & Specialist).Range("K1", Range("K1").Offset(6 - 1, Dimension1 - 1)).Value = Phantomdata
        End If
    Next Specialist
    
   '###################### PART D1 - SET RowLeftOff OUTSIDE THE WEEKLY ITERATION ###############################
    iterationcount = 0
    'still need to change the 99's on Non-PRA days to 1's
    If ALReady = True And needtochange99sto1s = True Then
        For RowNo = 1 To 5824
            For Specialist = 1 To J
                Activity1PrefHalfDay1 = ThisWorkbook.Worksheets("AnnualLeave").Cells(1, Specialist * 3 + 5).Value
                Activity1PrefHalfDay2 = ThisWorkbook.Worksheets("AnnualLeave").Cells(1, Specialist * 3 + 6).Value
                If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, Specialist).Value = 99 And _
                    ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, 9).Value = Activity1PrefHalfDay1 Then
                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, Specialist).Value = 1 ' put a 1 back in the schedule where the 99 overlapped the standard day off
                End If
                If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, Specialist).Value = 99 And _
                    ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, 9).Value = Activity1PrefHalfDay2 Then
                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo, Specialist).Value = 1 ' put a 1 back in the schedule where the 99 overlapped the standard day off
                End If
            Next Specialist
        Next RowNo
    End If
    
    
    If ALReady = False Then
        ReDim RowLeftOff(1 To J)
        For Specialist = 1 To J
            RowLeftOff(Specialist) = 4
        Next Specialist

StartWeeklyIteration:
        MyWorkbook.Worksheets("Sheet1").Activate
        MyWorkbook.Worksheets("Sheet1").Range(Cells(startrow, 1), Cells(startrow, 1)).Activate
        PSstart = ActiveCell.Row + 5
        PSend = PSstart + 13
        Range(Cells(PSstart, J + 4), Cells(PSstart, J + 4).End(xlDown)).Select
        ActiveWorkbook.Worksheets("Sheet1").Names.Add Name:="FutureDateRange", RefersTo:=Selection
        
        '###################### PART D2 - CHECK AND WRITE ANNUAL LEAVE INTO THE SCHEDULE WEEK BY WEEK ###############################
        ' now write upcoming annual leave into the calendar as 99, only on weekdays and non-national holidays
        ' alternative: copy the AL cells from 'Sheet1 orig' to 'Sheet1' in order to create a fair comparison
        
        Dim c As Range
        Dim arrCopyALCells() As Variant
        ReDim arrCopyALCells(1 To 5824, 1 To 6)
        Dim copycolumn, copyrow As Integer
        Dim AM, PM As Integer
        
        If CopyAnnualLeave = False Then
            For Specialist = 1 To J
                Activity1PrefHalfDay1 = ThisWorkbook.Worksheets("AnnualLeave").Cells(1, Specialist * 3 + 5).Value
                Activity1PrefHalfDay2 = ThisWorkbook.Worksheets("AnnualLeave").Cells(1, Specialist * 3 + 6).Value
                If AnnualLeaveFinished(Specialist) = False Then
                    ThisWorkbook.Worksheets("AnnualLeave").Activate
                    ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Select
                    If ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Value = "k" Then
                        RowLeftOff(Specialist) = RowLeftOff(Specialist) + 1
                        ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Select
                        
                    End If
                    Do While ActiveCell.Value <> "end" And ActiveCell.Value <> "" And ActiveCell.Value < ThisWorkbook.Worksheets("Sheet1").Cells(PSstart, J + 4).Value Or ActiveCell.Value = "k"
                            ActiveCell.Value = "k"
                            ActiveCell.Offset(1, 0).Activate
                            RowLeftOff(Specialist) = ActiveCell.Row
                            If ActiveCell.Value = "end" Then
                                AnnualLeaveFinished(Specialist) = True
                            End If
                        Loop
                    If RowLeftOff(Specialist) > 400 Then
                        AnnualLeaveFinished(Specialist) = True
                    End If
                    For i = 1 To 400
                        If ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Offset(i, 0).Value = "k" Then
                            Exit For
                        End If
                        If ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Offset(i, 0).Value = "end" Then
                            Exit For
                        End If
                    Next i
                    If ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Offset(i, 0).Value = "k" Then
                        Range(Cells(RowLeftOff(Specialist), Specialist), Cells(RowLeftOff(Specialist) + i - 1, Specialist)).Select
                        Set HolidayRange = Selection
                    End If
                    If ThisWorkbook.Worksheets("AnnualLeave").Cells(RowLeftOff(Specialist), Specialist).Offset(i, 0).Value = "end" Then
                        Range(Cells(RowLeftOff(Specialist), Specialist), Cells(RowLeftOff(Specialist) + i - 1, Specialist)).Select
                        Set HolidayRange = Selection
                    End If
                    If AnnualLeaveFinished(Specialist) = False Then
                        
                        HolidayRange.Select
                        If HolidayRange(1, 1).Offset(0, 25).Value = ThisWorkbook.Worksheets("Sheet1").Cells(PSstart, J + 5).Value Then ' j+5 should be week number in column K
                            For Each rc In HolidayRange.Cells
                                rc.Activate
                                If rc.Value = "end" Then
                                    AnnualLeaveFinished(Specialist) = True
                                    Exit For
                                End If
                                If rc.Value = "k" Then
                                    GoTo nextrc
                                End If
                                If rc.Value < ThisWorkbook.Worksheets("Sheet1").Range("FutureDateRange").Cells(1, 1).Value Then
                                    rc.Value = "k"
                                End If
                                If rc.Value = "k" Then
                                    GoTo nextrc
                                End If
                                If rc.Value >= ThisWorkbook.Worksheets("Sheet1").Range("FutureDateRange").Cells(1, 1).Value Then
                                    RowNo = Application.WorksheetFunction.Match(rc, ThisWorkbook.Worksheets("Sheet1").Range("FutureDateRange"), 0)
                                    AM = ThisWorkbook.Worksheets("AnnualLeave").Range(Cells(rc.Row, Specialist * 3 + 5), Cells(rc.Row, Specialist * 3 + 5)).Value ' should return 1 if that was a half day AL
                                    PM = ThisWorkbook.Worksheets("AnnualLeave").Range(Cells(rc.Row, Specialist * 3 + 6), Cells(rc.Row, Specialist * 3 + 6)).Value  ' should return 1 if that was a half day AL
                                    DOTW = Range(Cells(rc.Row, Specialist * 3 + 4), Cells(rc.Row, Specialist * 3 + 4)).Value
                                    ALWeekNo = Range(Cells(rc.Row, Specialist + 25), Cells(rc.Row, Specialist + 25)).Value
                                    If AM = 1 Then
                                        ThisWorkbook.Worksheets("AnnualLeave").Activate
                                        If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value <> 10 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value > -1 And DOTW > 2 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value = 99
                                        End If
                                    End If
                                    If PM = 1 Then
                                        ThisWorkbook.Worksheets("AnnualLeave").Activate
                                        If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value <> 10 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value > -1 And DOTW > 2 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value = 99
                                        End If
                                    End If
                                    'reinstate the 1's:
                                    If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value = 99 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, 9).Value = Activity1PrefHalfDay1 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value = 1 ' put a 1 back in the schedule where the 99 overlapped the standard day off
                                    End If
                                    If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value = 99 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, 9).Value = Activity1PrefHalfDay1 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value = 1  ' put a 1 back in the schedule where the 99 overlapped the standard day off
                                    End If
                                    If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value = 99 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, 9).Value = Activity1PrefHalfDay2 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart - 1, Specialist).Value = 1 ' put a 1 back in the schedule where the 99 overlapped the standard day off
                                    End If
                                    If ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value = 99 And _
                                        ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, 9).Value = Activity1PrefHalfDay2 Then
                                            ThisWorkbook.Worksheets("Sheet1").Cells(RowNo + PSstart, Specialist).Value = 1 ' put a 1 back in the schedule where the 99 overlapped the standard day off
                                    End If
                                End If
                                ThisWorkbook.Worksheets("AnnualLeave").Cells(rc.Row, Specialist).Value = "k"  ' replace  the date with k once it's been blocked in the schedule, so that we don't do the blocking over and over
                                RowLeftOff(Specialist) = rc.Row
nextrc:                     Next rc
                        End If
                        RowLeftOff(Specialist) = ActiveCell.Row
                    End If
                End If
            Next Specialist
        End If
    End If
  
    '###################### PART D - COUNT UP TO AND INCLUDING THIS WEEK: CurrentActiveHalfDayCellsCount(Specialist): NEW SCHEDULE'S TOTALS without this week's allocation FROM t=1 TILL END OF PLANSPAN, EXCLUDE arrAL(Specialist)=(Activity 99) ###############################
    Erase arrActivityTotalCountToPSstart
    ReDim arrActivityTotalCountToPSstart(1 To Dimension1, 1 To J)
    ReDim CurrentActiveHalfDayCellsCount(1 To J)
    ThisWorkbook.Worksheets("Sheet1").Activate
    

    For Specialist = 1 To J
        With Application.WorksheetFunction
            CurrentActiveHalfDayCellsCount(Specialist) = _
            Application.WorksheetFunction.CountIfs(ThisWorkbook.Worksheets("Sheet1").Range(Cells(1, Specialist), _
            Cells(PSend, Specialist)), "<10", ThisWorkbook.Worksheets("Sheet1").Range(Cells(1, Specialist), _
            Cells(PSend, Specialist)), ">-1")
        End With
    Next Specialist
    
    For Specialist = 1 To J
    ThisWorkbook.Worksheets("Specc" & Specialist).Activate
        For Activity = 1 To Dimension1
            arrActivityTotalCountToPSstart(Activity, Specialist) = _
            Application.WorksheetFunction.Sum(Range(Cells(1, Activity), Cells((PSstart - 1) / 7 / 2, Activity)))
        Next Activity
    Next Specialist
    
    exists = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "ActivityCountToPSstart" Then
            exists = True
        End If
    Next i

    If Not exists Then
        Worksheets.Add.Name = "ActivityCountToPSstart" ' do we erase and re-create this sheet every weekly iteration, as we should? Yes we do. --> don't think we have to re-create it every week
    End If
    With ThisWorkbook.Worksheets("ActivityCountToPSstart")
        .Move After:=Sheets(Sheets.Count)
        .Activate
        .Range("A1", Range("A1").Offset(Dimension1 - 1, J - 1)).Value = arrActivityTotalCountToPSstart
    End With

    '###################### PART E - COUNT arrINITIALHALFDAYS() AND arrHALFDAYSAVAILABLE() FOR THE PS WEEK EVERY ITERATION ##############################
    ReDim Preserve arrInitialHalfDaysAvailable(1 To J, 1 To 52 * 4)
    ReDim Preserve arrHalfDaysAvailable(1 To J, 1 To 52 * 4)
    
    ThisWorkbook.Worksheets("Sheet1").Activate
    For Specialist = 1 To J
        arrInitialHalfDaysAvailable(Specialist, PSend / 2 / 7) = Application.WorksheetFunction.CountIfs(Range(Cells(PSstart, Specialist), Cells(PSstart + 9, Specialist)), "") + _
        Application.WorksheetFunction.CountIfs(Range(Cells(PSstart, Specialist), Cells(PSstart + 9, Specialist)), 0)
        arrHalfDaysAvailable(Specialist, PSend / 2 / 7) = arrInitialHalfDaysAvailable(Specialist, PSend / 2 / 7)
    Next Specialist

'###################### PART F1 - ALLOCATE TEMPORARY FREQUENCIES OF SESSION TYPES BY COMPARING TO EMPIRICAL TOTALS ###############################
    ReDim arrActivity(1 To Dimension1)  ' this array will hold the candatadtes to-be-scheduled session types, 'freq' is how many of each session type should be scheduled
    ReDim TempFreqSum(1 To J)
    ReDim arrActivityTotalCountToPSend(1 To Dimension1, 1 To J)
    ReDim arrPSTempFreqPerj(1 To Dimension1, 1 To J)
    ReDim WhatWeShouldHave(1 To Dimension1, 1 To J)
    ReDim WhatWeNowHave(1 To Dimension1, 1 To J)
    ' over here we need to adjust the multiplier that i use: * (PSend/2/7) for the arrAL(Specialist) weeks with Activity 99
    ' It means we need to subtract the cells that hold 99
    Debug.Print Now & " start allocating tempfreq"
    For Specialist = 1 To J     ' this loop is for appending each candidate to-be-scheduled Activity to an array
        For Activity = 1 To 1 ' do Activity 1 separate because it's the non-pra and the empirical totals don't reflect reality (which should be average 2x per week per j)
            WhatWeShouldHave(Activity, Specialist) = WorksheetFunction.Min(Round(CurrentActiveHalfDayCellsCount(Specialist) * arrEmpiricalActiveHalfDayAverage(Activity, Specialist) * Equalizers(Activity, Specialist), 0), 2 * PSend / 2 / 7) ' for Activity1 we require arrEmpiricalActiveHalfDayAverage * total cells up till now, not * currentNonALCells. Because e.g. per Annual Leave week, he still needs 2* Activity1
            WhatWeNowHave(Activity, Specialist) = arrActivityTotalCountToPSstart(Activity, Specialist) + arrPSTempFreqPerj(Activity, Specialist)
            PriorityCol = (Specialist * 12) - 11 + 1 ' expanded to include t=1 to 10
            PrefHalfDay = (PSstart - 1) + Worksheets("EmpMostRep").Cells(Activity + 3, PriorityCol).Value
            PrefHalfDay2 = (PSstart - 1) + Worksheets("EmpMostRep").Cells(Activity + 3, PriorityCol + 1).Value
            If WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) <= 0 Then ' so if based on empirical data, you should have .9 or more Activityxxxx in x weeks, then do it, else 0. this is better than the 0.5 roundoff point of ROUND()
                arrPSTempFreqPerj(Activity, Specialist) = 0
                Exit For
            ElseIf WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay, Specialist).Value = 10 And _
            ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value = 10 Then
                WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                Exit For
            ElseIf WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay, Specialist).Value = 99 And _
            ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value = 99 Then
                WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                Exit For
            ElseIf WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 2 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay, Specialist).Value = 99 Then
                WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                If WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value = 99 Then
                    WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                    Exit For
                End If
            ElseIf WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay, Specialist).Value = 10 And _
            ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value = 10 Then
                WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                Exit For
            ElseIf WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 2 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay, Specialist).Value = 10 Then
                WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                If WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value = 10 Then
                    WhatWeShouldHave(Activity, Specialist) = WhatWeShouldHave(Activity, Specialist) - 1
                    Exit For
                End If
                If WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) = 1 And ThisWorkbook.Worksheets("Sheet1").Cells(PrefHalfDay2, Specialist).Value <> 10 Then
                    arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) + 1
                    Exit For
                End If
            ElseIf WhatWeShouldHave(Activity, Specialist) > WhatWeNowHave(Activity, Specialist) Then
                Do While WhatWeShouldHave(Activity, Specialist) > WhatWeNowHave(Activity, Specialist) And _
                WhatWeNowHave(Activity, Specialist) <= Application.WorksheetFunction.Round(((arrActivityEmpiricalTotals(Activity, Specialist) / 52) * (PSend / 2 / 7)), 0) - 0.1 And _
                arrPSTempFreqPerj(Activity, Specialist) <= 1
                    arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) + 1
                    WhatWeNowHave(Activity, Specialist) = WhatWeNowHave(Activity, Specialist) + 1 ' note that we need to erase and recount arrActivityTotalCountToPSend each weekly iteration. Which we do.
                Loop
            End If
        Next Activity
        
        For Activity = 2 To Dimension1
            WhatWeShouldHave(Activity, Specialist) = Round(CurrentActiveHalfDayCellsCount(Specialist) * arrEmpiricalActiveHalfDayAverage(Activity, Specialist) * Equalizers(Activity, Specialist), 0)
            arrActivityTotalCountToPSend(Activity, Specialist) = arrActivityTotalCountToPSstart(Activity, Specialist) + arrPSTempFreqPerj(Activity, Specialist)
            WhatWeNowHave(Activity, Specialist) = arrActivityTotalCountToPSstart(Activity, Specialist) + arrPSTempFreqPerj(Activity, Specialist)
            If WhatWeShouldHave(Activity, Specialist) - WhatWeNowHave(Activity, Specialist) <= 0 Then ' so if based on empirical data, you should have .9 or more Activityxxxx in x weeks, then do it, else 0. this is better than the 0.5 roundoff point of ROUND()
                arrPSTempFreqPerj(Activity, Specialist) = 0
            ElseIf WhatWeShouldHave(Activity, Specialist) > WhatWeNowHave(Activity, Specialist) Then
                Do While WhatWeShouldHave(Activity, Specialist) > WhatWeNowHave(Activity, Specialist) ' don't do anything to arrHalfDaysAvailable here yet, only after definitve allocation
                    arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) + 1
                    WhatWeNowHave(Activity, Specialist) = WhatWeNowHave(Activity, Specialist) + 1
                Loop
            End If
        Next Activity
        With Application.WorksheetFunction
            TempFreqSum(Specialist) = .Sum(.Index(arrPSTempFreqPerj, 0, Specialist)) ' sums all Activity frequencies for each specialist
        End With
    Next Specialist
     
    'write the PlanSpan week's RECOMMENDED/TEMPORARY/IDEAL frequencies of Activitys arrPSTempFreq into separate sheets sheets "Specc 1" per specialist
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("Specc" & Specialist).Activate ' we're not going to re-create these sheets every time. just add Planweek's counts to the existing sheet
        For Activity = 1 To Dimension1
            For WeeklyCount = PSend / 2 / 7 To PSend / 2 / 7
                ActiveSheet.Cells(PSend / 2 / 7, Activity).Value = arrPSTempFreqPerj(Activity, Specialist)
                ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity).Value = arrPSTempFreqPerj(Activity, Specialist) ' this is a copy of the temp recommended frequencies, but the ones we will calculate with
                ActiveSheet.Cells(PSend / 2 / 7, Activity + 10).Value = arrPSTempFreqPerj(Activity, Specialist) ' then copy it again to have it under the Phantom set
                ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity + 10).Value = arrPSTempFreqPerj(Activity, Specialist)
            Next WeeklyCount
        Next Activity
    Next Specialist
 
    '###################### PART F2 - DETERMINE WHAT WE SHOULD HAVE TOTAL OF Activity TYPES BY COMPARING TO EMPIRICAL TOTALS ###############################
        
    ReDim arrPSTempFreqTotal(1 To Dimension1) ' this array will hold the TOTAL candidates to-be-scheduled Activity types, 'freq' is how many of each Activity type should be scheduled
    ReDim WhatWeShouldHaveTotal(1 To Dimension1)
    
    For Activity = 1 To Dimension1
        With Application.WorksheetFunction
            ' Total for all specialiActivity together:
            WhatWeShouldHaveTotal(Activity) = Round((.Sum(.Index(arrActivityEmpiricalNonALWeekAverage, Activity, 0)) * .Sum(CurrentActiveHalfDayCellsCount) / 2 / 7), 0)
        End With
    Next Activity
         
   '###################### PART G1 - CHECK CERTAIN ACTIVITY TYPES WEEKLY MINIMUMS AND CORRECT ############################### dd 13/1/2016
    Debug.Print Now & "check activity minmax's and correct"
    ReDim arrPSDefFreq(1 To Dimension1, 1 To J)
    ReDim WhatWeShouldHaveUnRounded(1 To Dimension1, 1 To J) As Double
    Dim place As Integer
   
    ' Activity 4 needs to take place at least 1x per week --> FLEB done by Spec 3 and 6, preference for Spec3
    ' only in the event that Spec 3 and Spec6 have arrHalfDaysAvailable = 0 then no Activity 4 is assigned to arrPSDefFreq(Activity, Specialist)
    ' same goes for Activity 5 and 6. 5 is done by all specialists but very rarely by specialist 3. Activity 6 is only done by spec 1, 2 and 4
    
    ReDim arrNeedItSoonest(1 To J)
    ReDim SpecialistWillNeedItSoonest(1 To J)
    ReDim Ranking(1 To J)
    
    'Activity 4: only compare 2 specialists so no need to use arrWillNeedItSoonest()
    Dim sumsession4 As Long
    
    With Application.WorksheetFunction
    If enforcemin = True Then
        sumsession4 = Application.WorksheetFunction.Sum(.Index(arrPSTempFreqPerj, 4, 0))
            If sumsession4 < 1 Then
                AvailSpec3 = arrHalfDaysAvailable(3, PSend / 2 / 7) - .Sum(.Index(arrPSTempFreqPerj, 0, 1))
                AvailSpec6 = arrHalfDaysAvailable(6, PSend / 2 / 7) - .Sum(.Index(arrPSTempFreqPerj, 0, 3))
                If AvailSpec3 >= AvailSpec6 And arrHalfDaysAvailable(3, PSend / 2 / 7) > 0 Then
                    arrPSDefFreq(4, 3) = arrPSDefFreq(4, 3) + 1
                    arrHalfDaysAvailable(3, PSend / 2 / 7) = arrHalfDaysAvailable(3, PSend / 2 / 7) - 1
                ElseIf arrHalfDaysAvailable(6, PSend / 2 / 7) > 0 Then
                    arrPSDefFreq(4, 6) = arrPSDefFreq(4, 6) + 1
                    arrHalfDaysAvailable(6, PSend / 2 / 7) = arrHalfDaysAvailable(6, PSend / 2 / 7) - 1
                End If
            End If
    End If
    
    ' Activity 5 & 6 enforce maximum of 2: first remove activity 5/6 from the one that has it but arrPSTempFreq > arrInitialHalfDaysAvailable
    If enforcemax = True Then
        For Specialist = 1 To J
            For Activity = 5 To 6
                If Activity = 4 Then GoTo nextactivity
                If arrPSTempFreqPerj(Activity, Specialist) > 0 And arrHalfDaysAvailable(Specialist, PSend / 2 / 7) <= 0 Then ' not available this week, remove all your 5's
                    arrPSTempFreqPerj(Activity, Specialist) = 0
                End If
nextactivity: Next Activity
        Next Specialist
    
        For Activity = 5 To 6
            If Activity = 4 Then GoTo nextactivity2
            Do While .Sum(.Index(arrPSTempFreqPerj, Activity, 0)) > 2
                ReDim arrHDAforCalc(1 To J)
                ReDim SpecialistBusiest(1 To J)
                ReDim arrNeedsItLeast(1 To Dimension1, 1 To J)
                
                For Specialist = 1 To J
                    arrHDAforCalc(Specialist) = arrHalfDaysAvailable(Specialist, PSend / 2 / 7) - .Sum(.Index(arrPSTempFreqPerj, 0, Specialist))
                Next Specialist
            
                For place = 1 To J
                    SpecialistBusiest(place) = FindMin(arrHDAforCalc())
                    arrHDAforCalc(SpecialistBusiest(place)) = 9999
                Next place
                Erase arrHDAforCalc()
                
                For i = 1 To J
                    Specialist = SpecialistBusiest(i)
                    If arrPSTempFreqPerj(Activity, Specialist) > 0 And arrHalfDaysAvailable(Specialist, PSend / 2 / 7) > 0 Then
                        arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) - 1
                        If .Sum(.Index(arrPSTempFreqPerj, Activity, 0)) <= 2 Then
                            Exit Do
                        End If
                    End If
                Next i
            Loop
nextactivity2: Next Activity
    End If
    
        'Activity 5 & 6 ensure minimum:we want to assign it to the specialist that will be needing it the soonest, and has a half day available
        ' find the specialist that needs activity 5 most based on non-rounded empirical averages and current count
        If enforcemin = True Then
        For Activity = 5 To 6
            If Activity = 4 Then GoTo nextactivity3
            If .Sum(.Index(arrPSTempFreqPerj, Activity, 0)) + .Sum(.Index(arrPSDefFreq, Activity, 0)) < 1 Then
                ReDim arrNeedItSoonest(1 To J)
                For Specialist = 1 To J
                        WhatWeShouldHaveUnRounded(Activity, Specialist) = CurrentActiveHalfDayCellsCount(Specialist) * arrEmpiricalActiveHalfDayAverage(Activity, Specialist)
                        arrNeedItSoonest(Specialist) = WhatWeShouldHaveUnRounded(Activity, Specialist) - WhatWeNowHave(Activity, Specialist)
                Next Specialist
                 
                ReDim SpecialistWillNeedItSoonest(1 To J)
                For place = LBound(arrNeedItSoonest(), 1) To UBound(arrNeedItSoonest(), 1)
                    SpecialistWillNeedItSoonest(place) = FindMax(arrNeedItSoonest()) ' SpecialistWillNeedItSoonest(place) represents specialist
                    arrNeedItSoonest(SpecialistWillNeedItSoonest(place)) = -9999
                Next place
               
                ' assign activity 5 to the  specialist that ranks highest for needing it soonest if his arrInitialHalfDaysAvailable - sum arrPSTempFreq() is largest. else try 2nd max specialist
                ' need to give a rankleastbusy to start from --> maybe not!
                Dim arrHDAminTempFreq(1 To J)
                Dim LeastBusySpecialist(1 To J)
                Dim rankleastbusy As Integer
               
                Do While .Sum(.Index(arrPSTempFreqPerj, 5, 0)) + .Sum(.Index(arrPSDefFreq, 5, 0)) < 1   '(5th row, all columns)
                    For place = 1 To J
                        If arrInitialHalfDaysAvailable(SpecialistWillNeedItSoonest(place), PSend / 2 / 7) > 0 Then
                            arrPSDefFreq(5, SpecialistWillNeedItSoonest(place)) = arrPSDefFreq(5, SpecialistWillNeedItSoonest(place)) + 1
                            Exit For
                        End If
                    Next place
                Loop
                
                Erase SpecialistWillNeedItSoonest()
                Erase arrNeedItSoonest()
                Erase Ranking()
                Erase LeastBusySpecialist()
                Dim arrSumTempFreqActivity() As Variant
             End If
nextactivity3: Next Activity
         
         For Activity = 5 To 6
             If Activity = 4 Then GoTo nextactivity4
             If .Sum(.Index(arrPSTempFreqPerj, Activity, 0)) = 1 Then
                 For Specialist = 1 To J
                     If arrPSTempFreqPerj(Activity, Specialist) = 1 Then
                         arrPSDefFreq(Activity, Specialist) = arrPSDefFreq(Activity, Specialist) + 1
                         arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) - 1
                         Exit For
                     End If
                 Next Specialist
             End If
nextactivity4: Next Activity
    End If
    End With
    'Re-write the PlanSpan week's RECOMMENDED/TEMPORARY/IDEAL frequencies of Activitys arrPSTempFreq into separate sheets sheets "Specc 1" per specialist
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("Specc" & Specialist).Activate ' we're not going to re-create these sheets every time. just add Planweek's counts to the existing sheet
        For Activity = 1 To Dimension1
            For WeeklyCount = PSend / 2 / 7 To PSend / 2 / 7
                ActiveSheet.Cells(PSend / 2 / 7, Activity).Value = arrPSTempFreqPerj(Activity, Specialist)
                ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity).Value = arrPSTempFreqPerj(Activity, Specialist) ' this is a copy of the temp recommended frequencies, but the ones we will calculate with
                ActiveSheet.Cells(PSend / 2 / 7, Activity + 10).Value = arrPSTempFreqPerj(Activity, Specialist) ' then copy it again to have it under the Phantom set
                ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity + 10).Value = arrPSTempFreqPerj(Activity, Specialist)
            Next WeeklyCount
        Next Activity
    Next Specialist
    Debug.Print Now & " end of minmax check"
    Debug.Print Now & " end allocating tempfreq"
     
    '####################################################################################################################################################
     '###################### PART G2 - PRIORITIZE BASED ON CV DETERIORATION IF TEMPFREQUENCIES EXCEED AVAILABLE HALF DAYS ###############################
    '####################################################################################################################################################
    ' see if we need to prioritize and if yes, start prioritizing with the individual max in the recommended row (i.e. the to-be-scheduled week PSend/2/7)
    ReDim arrStDevsRecomm(1 To Dimension1, 1 To J)
    ReDim arrMeansRecomm(1 To Dimension1, 1 To J)
    ReDim arrCVsRecomm(1 To Dimension1, 1 To J)
    ReDim arrStDevsZeros(1 To Dimension1, 1 To J)
    ReDim arrMeansZeros(1 To Dimension1, 1 To J)
    ReDim arrCVsZeros(1 To Dimension1, 1 To J)
    ReDim arrDifferenceCVRecommAllocated(1 To Dimension1, 1 To J)
    ReDim DefFreqSum(1 To J)
    

     Debug.Print Now & " start prioritization based on CV"
     'first write the definitive allocations so far into the Specc sheets (they start at zero)
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("Specc" & Specialist).Activate
        Set DefAllocRange = ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 3), 1), Cells(((PSend / 2 / 7) + 3), Dimension1))
        For Activity = 1 To Dimension1
            DefAllocRange(Activity).Value = arrPSDefFreq(Activity, Specialist)
        Next Activity
        
        With Application.WorksheetFunction
            TempFreqSum(Specialist) = .Sum(.Index(arrPSTempFreqPerj, 0, Specialist)) ' sums all Activity temporary frequencies for each specialist
            DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
        End With
                                                                                        ' sum these because we've allocated to deffreq in the previous part
        If arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2) < TempFreqSum(Specialist) + DefFreqSum(Specialist) _
        And arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2) > 0 Then ' means we only run the whole prioritization code IF we have fewer halfdays available than PSTempFreq
            CountAlgoInstances = CountAlgoInstances + 1
            ThisWorkbook.Worksheets("Specc" & Specialist).Activate
            ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 2), 1), Cells(((PSend / 2 / 7) + 2), Dimension1)).Select
            ActiveWorkbook.Worksheets("Specc" & Specialist).Names.Add Name:="RangeRecommCalc", RefersTo:=Selection
            If UseActualStdevs = True Then
                Do While WorksheetFunction.Large(Range("RangeRecommCalc"), 1) > WorksheetFunction.Large(Range("RangeRecommCalc"), 2) And DefFreqSum(Specialist) < arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2) ' Means we have a single highest frequency to allocate first
                    UniqueLargestFreqColNum = _
                    Application.WorksheetFunction.Match(Application.WorksheetFunction.Large(Range("RangeRecommCalc"), 1), Range("RangeRecommCalc"), 0)
                    arrPSTempFreqPerj(UniqueLargestFreqColNum, Specialist) = arrPSTempFreqPerj(UniqueLargestFreqColNum, Specialist) - 1
                    arrPSDefFreq(UniqueLargestFreqColNum, Specialist) = arrPSDefFreq(UniqueLargestFreqColNum, Specialist) + 1
                    Cells((PSend / 2 / 7) + 2, UniqueLargestFreqColNum).Value = Cells((PSend / 2 / 7) + 2, UniqueLargestFreqColNum).Value - 1
                    Cells((PSend / 2 / 7) + 3, UniqueLargestFreqColNum).Value = Cells((PSend / 2 / 7) + 3, UniqueLargestFreqColNum).Value + 1
                    With Application.WorksheetFunction
                        DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                    End With
                Loop
                With Application.WorksheetFunction
                    DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                End With
                Do While WorksheetFunction.Large(Range("RangeRecommCalc"), 1) = _
                WorksheetFunction.Large(Range("RangeRecommCalc"), 2) And WorksheetFunction.Large(Range("RangeRecommCalc"), 1) > 0 And DefFreqSum(Specialist) < arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2)
                    For Activity = 1 To Dimension1
                        If ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity).Value = WorksheetFunction.Large(Range("RangeRecommCalc"), 1) Then
                            Range(Cells(1, Activity), Cells(PSend / 2 / 7, Activity)).Select ' CV for recommended:
                            arrStDevsRecomm(Activity, Specialist) = _
                            Application.WorksheetFunction.StDev(Range(Cells(1, Activity), Cells(PSend / 2 / 7, Activity)))
                            arrMeansRecomm(Activity, Specialist) = _
                            Application.WorksheetFunction.Average(Range(Cells(1, Activity), Cells(PSend / 2 / 7, Activity))) + 0.0000000000002 ' we add a very small amoutn so that in case of a tie, we go with the bare minimum (which is the recommended number of Activitys minus 1)and to divide by 'almost 0' which is ok because it will be 0 divided by 0.000000001 which is still 0
                            If arrMeansRecomm(Activity, Specialist) > 0 And Cells((PSend / 2 / 7), Activity) > 0 Then
                                arrCVsRecomm(Activity, Specialist) = _
                                arrStDevsRecomm(Activity, Specialist) / arrMeansRecomm(Activity, Specialist)
                                Cells((PSend / 2 / 7) + 5, Activity).Value = arrCVsRecomm(Activity, Specialist)
                            End If
                            Set Rng1 = Range(Cells(1, Activity), Cells((PSend / 2 / 7) - 1, Activity))
                            Rng1.Select
                            Set Rng2 = Range(Cells((PSend / 2 / 7) + 3, Activity), Cells((PSend / 2 / 7) + 3, Activity))
                            Rng2.Select
                            Set RngU = Union(Rng1, Rng2)
                            RngU.Select
                            arrStDevsZeros(Activity, Specialist) = _
                            Application.WorksheetFunction.StDev(RngU)
                            arrMeansZeros(Activity, Specialist) = _
                            Application.WorksheetFunction.Average(RngU) + 0.0000000000002 ' we add a very small amoutn so that in case of a tie, we go with the bare minimum (which is the recommended number of Activitys minus 1)and to divide by 'almost 0' which is ok because it will be 0 divided by 0.000000001 which is still 0
                            If arrMeansZeros(Activity, Specialist) > 0 And Cells((PSend / 2 / 7), Activity) > 0 Then
                                arrCVsZeros(Activity, Specialist) = _
                                arrStDevsZeros(Activity, Specialist) / arrMeansZeros(Activity, Specialist)
                                Cells((PSend / 2 / 7) + 6, Activity).Value = arrCVsZeros(Activity, Specialist)
                            End If
                            Cells((PSend / 2 / 7) + 8, Activity).Value = arrCVsRecomm(Activity, Specialist) - arrCVsZeros(Activity, Specialist)
                            arrDifferenceCVRecommAllocated(Activity, Specialist) = _
                            arrCVsRecomm(Activity, Specialist) - arrCVsZeros(Activity, Specialist)
                        End If
                    Next Activity
                    Set DegradationRng = ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 8), 1), Cells(((PSend / 2 / 7) + 8), Dimension1))
                    DegradationRng.Select
                    ActiveWorkbook.Worksheets("Specc" & Specialist).Names.Add Name:="DegradationRng", RefersTo:=Selection
                    LargestValue = WorksheetFunction.Large(Range("RangeRecommCalc"), 1)
                    
CheckAllCols:       For Activity = Dimension1 To 1 Step -1 ' by reversing, the algorithm will give priority to Activity types 8,7,6, ... before the more general ones 3,2,1 in case of a tie in degradation
                        If Cells((PSend / 2 / 7) + 2, Activity).Value = LargestValue And _
                        Cells((PSend / 2 / 7) + 8, Activity).Value = Application.WorksheetFunction.Small(Range("DegradationRng"), 1) Then
                            arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) - 1
                            arrPSDefFreq(Activity, Specialist) = arrPSDefFreq(Activity, Specialist) + 1
                            arrHalfDaysAvailable(Specialist, PSend / 2 / 7) = arrHalfDaysAvailable(Specialist, PSend / 2 / 7) - 1
                            ThisWorkbook.Worksheets("Specc" & Specialist).Cells((PSend / 2 / 7) + 2, Activity).Select
                            Cells((PSend / 2 / 7) + 2, Activity).Value = Cells((PSend / 2 / 7) + 2, Activity).Value - 1
                            Cells((PSend / 2 / 7) + 3, Activity).Value = Cells((PSend / 2 / 7) + 3, Activity).Value + 1
                            Cells((PSend / 2 / 7) + 8, Activity).Value = ""
                            With Application.WorksheetFunction
                                DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                            End With
                            If DefFreqSum(Specialist) = arrInitialHalfDaysAvailable(Specialist, PSend / 2 / 7) Then
                                GoTo NextSpecialist ' as soon as we reach our quota we move on to the next specialist
                            End If
                            If WorksheetFunction.CountA(Range("DegradationRng")) > 0 Then
                                GoTo CheckAllCols
                            Else: GoTo loopit
                            End If
                        End If
                    Next Activity
                    
loopit:         Loop ' this should keep looping until the definitive allocation row sum is equal to initial half days available
            
            ElseIf UseActualStdevs = False Then
                Do While WorksheetFunction.Large(Range("RangeRecommCalc"), 1) > WorksheetFunction.Large(Range("RangeRecommCalc"), 2) And DefFreqSum(Specialist) < arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2) ' Means we have a single highest frequency to allocate first
                    UniqueLargestFreqColNum = _
                    Application.WorksheetFunction.Match(Application.WorksheetFunction.Large(Range("RangeRecommCalc"), 1), Range("RangeRecommCalc"), 0)
                    arrPSTempFreqPerj(UniqueLargestFreqColNum, Specialist) = arrPSTempFreqPerj(UniqueLargestFreqColNum, Specialist) - 1
                    arrPSDefFreq(UniqueLargestFreqColNum, Specialist) = arrPSDefFreq(UniqueLargestFreqColNum, Specialist) + 1
                    Cells((PSend / 2 / 7) + 2, UniqueLargestFreqColNum).Value = Cells((PSend / 2 / 7) + 2, UniqueLargestFreqColNum).Value - 1
                    Cells((PSend / 2 / 7) + 3, UniqueLargestFreqColNum).Value = Cells((PSend / 2 / 7) + 3, UniqueLargestFreqColNum).Value + 1
                    With Application.WorksheetFunction
                        DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                    End With
                Loop
                With Application.WorksheetFunction
                    DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                End With
                Do While WorksheetFunction.Large(Range("RangeRecommCalc"), 1) = _
                WorksheetFunction.Large(Range("RangeRecommCalc"), 2) And WorksheetFunction.Large(Range("RangeRecommCalc"), 1) > 0 And DefFreqSum(Specialist) < arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2)
                    For Activity = 1 To Dimension1
                        If ActiveSheet.Cells((PSend / 2 / 7) + 2, Activity).Value = WorksheetFunction.Large(Range("RangeRecommCalc"), 1) Then
                            
                            ' CV for recommended:
                            ' calculate the CVs for the 'ideal' set on the right in the Specc sheet --> hence column = Activity +10, then it continues on actual allocation
                            Range(Cells(1, Activity + 10), Cells(PSend / 2 / 7, Activity + 10)).Select
                            arrStDevsRecomm(Activity, Specialist) = _
                            Application.WorksheetFunction.StDev(Range(Cells(1, Activity + 10), Cells(PSend / 2 / 7, Activity + 10)))
                            'Debug.Print "arrStDevsRecomm  " & arrStDevsRecomm(Activity, Specialist)
                            
                            ' calculate the means for the 'ideal' set on the right in the Specc sheet --> hence column = Activity +10
                            arrMeansRecomm(Activity, Specialist) = _
                            Application.WorksheetFunction.Average(Range(Cells(1, Activity + 10), Cells(PSend / 2 / 7, Activity + 10))) + 0.0000000000002  ' we add a very small amoutn so that in case of a tie, we go with the bare minimum (which is the recommended number of Activitys minus 1)and to divide by 'almost 0' which is ok because it will be 0 divided by 0.000000001 which is still 0
                            'Debug.Print "arrMeansRecomm " & arrMeansRecomm(Activity, Specialist)
                            
                            If arrMeansRecomm(Activity, Specialist) > 0 And Cells((PSend / 2 / 7), Activity) > 0 Then
                                arrCVsRecomm(Activity, Specialist) = _
                                arrStDevsRecomm(Activity, Specialist) / arrMeansRecomm(Activity, Specialist)
                                Cells((PSend / 2 / 7) + 5, Activity + 10).Value = arrCVsRecomm(Activity, Specialist) ' Write these CVs to the right for easier recognition during debug
                            End If
                            
                            Set Rng1 = Range(Cells(1, Activity + 10), Cells((PSend / 2 / 7) - 1, Activity + 10)) ' first six weeks when run the first time
                            Rng1.Select
                            Set Rng2 = Range(Cells((PSend / 2 / 7) + 3, Activity), Cells((PSend / 2 / 7) + 3, Activity)) 'the def allocation range
                            Rng2.Select
                            Set RngU = Union(Rng1, Rng2)
                            RngU.Select
                            arrStDevsZeros(Activity, Specialist) = _
                            Application.WorksheetFunction.StDev(RngU)
                            arrMeansZeros(Activity, Specialist) = _
                            Application.WorksheetFunction.Average(RngU) + 0.0000000000002 ' we add a very small amoutn so that in case of a tie, we go with the bare minimum (which is the recommended number of Activitys minus 1)and to divide by 'almost 0' which is ok because it will be 0 divided by 0.000000001 which is still 0
             
                            
                            If arrMeansZeros(Activity, Specialist) > 0 And Cells((PSend / 2 / 7), Activity) > 0 Then
                                arrCVsZeros(Activity, Specialist) = _
                                arrStDevsZeros(Activity, Specialist) / arrMeansZeros(Activity, Specialist)
                                Cells((PSend / 2 / 7) + 6, Activity + 10).Value = arrCVsZeros(Activity, Specialist)
                            End If
                            
                            Cells((PSend / 2 / 7) + 8, Activity + 10).Value = arrCVsRecomm(Activity, Specialist) - arrCVsZeros(Activity, Specialist)
                            arrDifferenceCVRecommAllocated(Activity, Specialist) = _
                            arrCVsRecomm(Activity, Specialist) - arrCVsZeros(Activity, Specialist)
                        End If
                    Next Activity
                    
                    Set DegradationRng = ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 8), 1 + 10), Cells(((PSend / 2 / 7) + 8), Dimension1 + 10))
                    DegradationRng.Select
                    ActiveWorkbook.Worksheets("Specc" & Specialist).Names.Add Name:="DegradationRng", RefersTo:=Selection
                    LargestValue = WorksheetFunction.Large(Range("RangeRecommCalc"), 1)
                    
CheckAllCols2:       For Activity = Dimension1 To 1 Step -1 ' by reversing, the algorithm will give priority to Activity types 8,7,6, ... before the more general ones 3,2,1 in case of a tie in degradation
                                With ThisWorkbook.Worksheets("Specc" & Specialist)
                                    ' so the activity that gets allocated is the one with the smallest value in degradationrange
                                    If Cells((PSend / 2 / 7) + 2, Activity).Value = LargestValue And _
                                        Cells((PSend / 2 / 7) + 8, Activity + 10).Value = Application.WorksheetFunction.Small(Range("DegradationRng"), 1) Then
                                        arrPSTempFreqPerj(Activity, Specialist) = arrPSTempFreqPerj(Activity, Specialist) - 1
                                        arrPSDefFreq(Activity, Specialist) = arrPSDefFreq(Activity, Specialist) + 1
                                        arrHalfDaysAvailable(Specialist, PSend / 2 / 7) = arrHalfDaysAvailable(Specialist, PSend / 2 / 7) - 1
                                        ThisWorkbook.Worksheets("Specc" & Specialist).Cells((PSend / 2 / 7) + 2, Activity).Select
                                        Cells((PSend / 2 / 7) + 2, Activity).Value = Cells((PSend / 2 / 7) + 2, Activity).Value - 1
                                        Cells((PSend / 2 / 7) + 3, Activity).Value = Cells((PSend / 2 / 7) + 3, Activity).Value + 1
                                        Cells((PSend / 2 / 7) + 8, Activity + 10).Value = ""
                                        With Application.WorksheetFunction
                                            DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                                        End With
                                        If DefFreqSum(Specialist) = arrInitialHalfDaysAvailable(Specialist, PSend / 2 / 7) Then
                                            GoTo NextSpecialist ' as soon as we reach our quota we move on to the next specialist
                                        End If
                                        If WorksheetFunction.CountA(Range("DegradationRng")) > 0 Then
                                            GoTo CheckAllCols2
                                        Else: GoTo loopit2
                                        End If
                                    End If
                                End With
                    Next Activity
loopit2:         Loop ' this should keep looping until the definitive allocation row sum is equal to initial half days available
            End If
        
        
        
        ElseIf arrInitialHalfDaysAvailable(Specialist, PSend / 7 / 2) >= DefFreqSum(Specialist) + TempFreqSum(Specialist) Then
            For Activity = 1 To Dimension1
                arrPSDefFreq(Activity, Specialist) = arrPSDefFreq(Activity, Specialist) + arrPSTempFreqPerj(Activity, Specialist)
                With Application.WorksheetFunction
                    DefFreqSum(Specialist) = .Sum(.Index(arrPSDefFreq, 0, Specialist))
                End With
            Next Activity
        End If
NextSpecialist:
    Next Specialist
    
     Debug.Print Now & " end prioritization based on CV"
    
    '###################### PART Ia - ' write arrPSDefFreq() into a sheet so that Matlab can read it in ###############################
    
    exists = False
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "arrPSDefFreq" Then
            exists = True
            w = i
        End If
    Next i
    
    If exists = True Then
        Application.DisplayAlerts = False
        Worksheets(w).Delete
        exists = False
        Application.DisplayAlerts = True
    End If
    
    If Not exists Then
        ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "arrPSDefFreq"
        ThisWorkbook.Worksheets("arrPSDefFreq").Range(Cells(1, 1), Cells(Dimension1, J)).Value = arrPSDefFreq
    End If
    
    Specialist = 1
    For g = 0 To J - 1
        ThisWorkbook.Worksheets("arrPSDefFreq").Range(Cells(g * Dimension1 + 1, J + 2), Cells(g * Dimension1 + Dimension1, J + 2)) = ThisWorkbook.Worksheets("arrPSDefFreq").Range(Cells(1, Specialist), Cells(Dimension1, Specialist)).Value
        Specialist = Specialist + 1
    Next g
    '###################### PART Ib - WRITE DEFINITIVE Activity TYPES INTO THE SCHEDULE ###############################
    ' since applying optimization to each week's definite half-day allocation, in stead, we check same-half-day restrictions before each allocation, and choose next prefhalfday
    ' if restrictions are violated. To keep it fair, the specialist order of allocating activities is randomized every week.
    Debug.Print Now & " write deffreq into schedule"
    ' now write the definitive allocated Activity types into the schedule
    ReDim Preserve arrCouldNotAllocateFreq(1 To Dimension1, 1 To J) ' the first dimension tells us which Activity is concerned, the actual values are frequencies
    ThisWorkbook.Worksheets("Sheet1").Activate
                       
    '2/5/2016 shuffle array to start at random specialist when allocating definitive activities
    
    Randomize
    Dim arrSpecialists
    Dim N, K As Long
    Dim Temp As Variant
    
    arrSpecialists = Array(1, 2, 3, 4, 5, 6)
    Debug.Print arrSpecialists(1), arrSpecialists(2), arrSpecialists(3), arrSpecialists(4), arrSpecialists(5), arrSpecialists(6)
    For N = LBound(arrSpecialists) To UBound(arrSpecialists)
        K = CLng(((UBound(arrSpecialists) - N) * Rnd) + N)
        If N <> K Then
            Temp = arrSpecialists(N)
            arrSpecialists(N) = arrSpecialists(K)
            arrSpecialists(K) = Temp
        End If
    Next N
    Debug.Print arrSpecialists(1), arrSpecialists(2), arrSpecialists(3), arrSpecialists(4), arrSpecialists(5), arrSpecialists(6)
    
    arrActivityIndex = Array(8, 7, 6, 5, 4, 2, 3, 1) ' this is the order in which definitive frequencies should be allocated
    For i = 1 To J
        For Z = 1 To Dimension1
            If arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) > 0 Then
                For Instance = 1 To arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i))
                    PriorityCol = (arrSpecialists(i) * 12) - 11 + 1 ' expanded to include t=1 to 10
trynextprefhalfday:
                    PrefHalfDay = (PSstart - 1) + Worksheets("EmpMostRep").Cells(arrActivityIndex(Z) + 3, PriorityCol).Value ' WriteSTP must be on this line because PrefHalfDay needs to be reassigned a value accoriding to the PrefCol +1
                    If ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = "" Or ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = 0 Then
                        If arrActivityIndex(Z) = 3 Then ' only really need to restrict the number of twos that take place at the same time, but this is fairer comparison to cyclic scheule
                            Twos = Application.WorksheetFunction.CountIf(Range(Cells(PrefHalfDay, 1), Cells(PrefHalfDay, 6)), 2)
                            Threes = Application.WorksheetFunction.CountIf(Range(Cells(PrefHalfDay, 1), Cells(PrefHalfDay, 6)), 3)
                            If Twos + Threes <= 2 Then ' we don't really need to restrict the sum of twos and threes but this is fairer comparison to cyclic schedule. Could be relaxed later?
                                Cells(PrefHalfDay, arrSpecialists(i)).Value = arrActivityIndex(Z)
                                arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                            Else
                                PriorityCol = PriorityCol + 1
                                GoTo trynextprefhalfday
                            End If
                        ElseIf arrActivityIndex(Z) = 2 Then
                            Twos = Application.WorksheetFunction.CountIf(Range(Cells(PrefHalfDay, 1), Cells(PrefHalfDay, 6)), 2)
                            Threes = Application.WorksheetFunction.CountIf(Range(Cells(PrefHalfDay, 1), Cells(PrefHalfDay, 6)), 3)
                            If Twos <= 0 And Twos + Threes <= 1 Then
                                Cells(PrefHalfDay, arrSpecialists(i)).Value = arrActivityIndex(Z)
                                arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                            Else
                                PriorityCol = PriorityCol + 1
                                GoTo trynextprefhalfday
                            End If
                        ElseIf arrActivityIndex(Z) = 4 Or arrActivityIndex(Z) = 5 Or arrActivityIndex(Z) = 6 Then ' activities 4/5/6 are restricted to 1 instance per half day (each)
                            If Application.WorksheetFunction.CountIf(Range(Cells(PrefHalfDay, 1), Cells(PrefHalfDay, 6)), arrActivityIndex(Z)) <= 0 Then ' only one instance per half day
                                Cells(PrefHalfDay, arrSpecialists(i)).Value = arrActivityIndex(Z)
                                arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                            Else
                                PriorityCol = PriorityCol + 1
                                GoTo trynextprefhalfday
                            End If
                        Else
                            Cells(PrefHalfDay, arrSpecialists(i)).Value = arrActivityIndex(Z)
                            arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        End If
                    ElseIf arrActivityIndex(Z) = 1 And arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = 2 And ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = 10 Then
                        arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        PrefHalfDay2 = (PSstart - 1) + Worksheets("EmpMostRep").Cells(arrActivityIndex(Z) + 3, PriorityCol + 1).Value
                        If ActiveSheet.Cells(PrefHalfDay2, arrSpecialists(i)).Value = 10 Then
                            arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        End If
                    ElseIf arrActivityIndex(Z) = 1 And arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = 1 And ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)) = 10 Then
                        arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        If ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = 10 Then
                            arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        End If
                    ElseIf Not ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = "" Or ActiveSheet.Cells(PrefHalfDay, arrSpecialists(i)).Value = 0 Then
                        PriorityCol = PriorityCol + 1
                        If PriorityCol > (arrSpecialists(i) * 12) - 11 + 1 + 9 Then
                            arrCouldNotAllocateFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrCouldNotAllocateFreq(arrActivityIndex(Z), arrSpecialists(i)) + 1
                            arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) = arrPSDefFreq(arrActivityIndex(Z), arrSpecialists(i)) - 1
                        ElseIf PriorityCol <= (arrSpecialists(i) * 12) - 11 + 1 + 9 Then
                            GoTo trynextprefhalfday
                        End If
                    End If
                Next Instance
            End If
        Next Z
    Next i
    
    
    ' finally fill in the gaps in PS with 9's
    Startcell = PSstart
    For Specialist = 1 To J
        For Startcell = PSstart To PSstart + 13
            If ActiveSheet.Cells(Startcell, Specialist).Value = "" Or ActiveSheet.Cells(Startcell, Specialist).Value = 0 Then
                Cells(Startcell, Specialist).Value = 9 ' 9 means 'no activity allocated', where as 99 is annual leave taken up
            End If
        Next Startcell
    Next Specialist
    Debug.Print Now & " end writing deffreq into schedule"
    Application.ScreenUpdating = False
    
      '###################### PART H - ADD THE PLANSPAN WEEK'S final Activity FREQUENCIES TO arrActivityWeeklyCountToPSend ###############################
    ' this way we don't have to count again and again for the whole past period every week
    ' but count it from the RC sheet in order to have the correct number of 1s!!
    
    ReDim Preserve arrActivityWeeklyCountToPSend(1 To Dimension1, 1 To J, 1 To (PSend / 2 / 7))
    
    ThisWorkbook.Worksheets("Sheet1").Activate
    For Specialist = 1 To J
        For Activity = 1 To Dimension1
            StartCount = 1
            For WeeklyCount = PSend / 2 / 7 To PSend / 2 / 7 ' first run this means week 7
                arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount) = _
                Application.WorksheetFunction.CountIf(Range(Cells(PSstart, Specialist), Cells(PSstart + 13, Specialist)), Activity)
            Next WeeklyCount
        Next Activity
    Next Specialist
    
    ' Now overwrite the  frequencies in the specc sheets with the actual counts from the RC
     For Specialist = 1 To J
        ThisWorkbook.Worksheets("Specc" & Specialist).Activate ' we're not going to re-create these sheets every time. just add Planweek's counts to the existing sheet
        For Activity = 1 To Dimension1
            For WeeklyCount = PSend / 2 / 7 To PSend / 2 / 7
                 ActiveSheet.Cells(PSend / 2 / 7, Activity).Value = arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount)
                 ActiveSheet.Cells(PSend / 2 / 7, Activity + 10).Value = arrActivityWeeklyCountToPSend(Activity, Specialist, WeeklyCount) ' added on 28/04/2016
            Next WeeklyCount
        Next Activity
    Next Specialist
    
    ReDim Preserve arrStDevsAfterAllocationTillPSend(1 To Dimension1, 1 To J)
    ReDim Preserve arrMeansAfterAllocationTillPSend(1 To Dimension1, 1 To J)
    ReDim Preserve arrCVsAfterAllocationTillPSend(1 To Dimension1, 1 To J)
    
    ReDim Preserve arrStDevsAfterAllocationTillPSendMin1(1 To Dimension1, 1 To J)
    ReDim Preserve arrMeansAfterAllocationTillPSendMin1(1 To Dimension1, 1 To J)
    ReDim Preserve arrCVsAfterAllocationTillPSendMin1(1 To Dimension1, 1 To J)
    
    ReDim Preserve arrStDevsAfterAllocationTillPSendIfIs0(1 To Dimension1, 1 To J)
    ReDim Preserve arrMeansAfterAllocationTillPSendIfIs0(1 To Dimension1, 1 To J)
    ReDim Preserve arrCVsAfterAllocationTillPSendIfIs0(1 To Dimension1, 1 To J)
       
    Application.DisplayAlerts = False
    For Specialist = 1 To J
        ThisWorkbook.Worksheets("Specc" & Specialist).Activate
        ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 2), 1), Cells(((PSend / 2 / 7) + 8), Dimension1)).ClearContents
        ThisWorkbook.Worksheets("Specc" & Specialist).Range(Cells(((PSend / 2 / 7) + 2), 1 + 1), Cells(((PSend / 2 / 7) + 8), Dimension1 + 10)).ClearContents
    Next Specialist
    
    Worksheets("ActivityCountToPSstart").Delete
    
    Application.DisplayAlerts = True
    ThisWorkbook.Worksheets("Sheet1").Activate
 
    startrow = startrow + 14
    Debug.Print Now
    Application.ScreenUpdating = False
    For Specialist = 1 To J
        For Activity = 1 To Dimension1
            Debug.Print arrCouldNotAllocateFreq(Activity, Specialist)
        Next Activity
    Next Specialist
    Debug.Print startrow
    
    Erase arrActivityTotalCountToPSend() ' need to erase because the tempfreqs are included
    Erase arrActivity()   ' array that holds the candidate to-be-scheduled Activity ' have to erase because we re-calculate the sum of frequencies for each specialist every round
    Erase arrMeansAfterAllocationTillPSend(), arrMeansAfterAllocationTillPSendMin1()
    Erase arrMeansAfterAllocationTillPSendIfIs0()
    Erase arrStDevsAfterAllocationTillPSend(), arrStDevsAfterAllocationTillPSendMin1()
    Erase arrStDevsAfterAllocationTillPSendIfIs0()
    Erase arrCVsAfterAllocationTillPSend(), arrCVsAfterAllocationTillPSendMin1()
    Erase arrCVsAfterAllocationTillPSendIfIs0()
    Erase arrDeltaCV()
    Erase arrPSDefFreq() ' have to erase because we re-calculate the sum of frequencies for each specialist every round
    Erase arrPSTempFreqPerj()
    Erase arrStDevsRecomm()
    Erase arrMeansRecomm()
    Erase arrCVsRecomm()
    Erase arrStDevsZeros()
    Erase arrMeansZeros()
    Erase arrCVsZeros()
    Erase CurrentActiveHalfDayCellsCount()
    
    If startrow = 80 + 10 * 14 Then ' after 6+ 10th week
        Debug.Print Now
        Debug.Print "stcreenupdating false"
        startrow = startrow
       ' Application.ScreenUpdating = False
    End If
    If startrow >= 2908 Then
        GoTo endd
    Else
        GoTo StartWeeklyIteration
    End If
    
endd:
    Application.ScreenUpdating = True
    ThisWorkbook.Worksheets("Sheet1").Range("A1", "F2912").Copy
    ThisWorkbook.Worksheets("Sheet1").Range(Cells(2913, 1), Cells(5824, 6)).Select
    ActiveSheet.Paste
    
    For i = 1 To Worksheets.Count
        If Worksheets(i).Name = "CouldNotAllocatefreq" Then
            exists = True
        End If
    Next i

    If Not exists Then
        ThisWorkbook.Worksheets.Add(After:=Sheets(Sheets.Count)).Name = "CouldNotAllocatefreq"
    End If
    
    ThisWorkbook.Worksheets("CouldNotAllocatefreq").Activate
    Range(Cells(1, 1), Cells(Dimension1, J)).Value = arrCouldNotAllocateFreq
    
    ThisWorkbook.Worksheets("Sheet1").Range("M1").Value = CountAlgoInstances
    ThisWorkbook.Worksheets("Sheet1").Range("M1").Value = CountAlgoInstances
    
    Application.DisplayAlerts = False
    Application.DisplayAlerts = True
    
    Erase arrActivityEmpiricalTotals()
    Erase arrActivityEmpiricalNonALWeekAverage()
    Erase arrEmpWeeklyActivityCount()
    Debug.Print starttime & " start of run"
    Debug.Print Now & " end of run"
End Sub

Public Function FindMax(arr() As Variant) As Integer
  Dim myMax As Double
 
  Dim i1 As Integer
  myMax = -9999
  

    For i1 = LBound(arr, 1) To UBound(arr, 1)
      If arr(i1) > myMax Then
        myMax = arr(i1)
        FindMax = i1
      ElseIf arr(i1) = myMax And arr(i1) > -9999 Then
            If Rnd() < 0.5 Then
                myMax = i1 ' in case of tie, with 50% probability, i2 gets updated
                FindMax = i1
            End If
      End If
    Next i1


End Function
Public Function FindMin(arr() As Variant) As Integer
    Dim myMin As Double
    Dim i2 As Integer
    
    myMin = 9999
    
    For i2 = LBound(arr, 1) To UBound(arr, 1)
        If arr(i2) < myMin Then
            myMin = arr(i2)
            FindMin = i2
        ElseIf arr(i2) = myMin And arr(i2) < 9999 Then
            If Rnd() < 0.5 Then
                myMin = i2 ' in case of tie, with 50% probability, i2 gets updated
                FindMin = i2
            End If
        End If
    Next i2
    
End Function
Sub RoundM()

    Dim OriginalScheduleDurations As Range
    Dim sCell As Range
    

    ThisWorkbook.Worksheets("Sheet2 orig").Activate
    Set OriginalScheduleDurations = Range("A1", Range("A1").End(xlDown).End(xlToRight))
    For Each sCell In OriginalScheduleDurations.Cells
        If sCell.Value > 0 Then
            sCell.Value = Application.WorksheetFunction.MRound(sCell, 5)
        End If
    Next sCell
End Sub

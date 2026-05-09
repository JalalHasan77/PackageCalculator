Imports System
Imports System.Collections.Generic
Imports System.Drawing
Imports System.Linq
Imports System.Windows.Forms

' =============================================================================
' RESULT OBJECT
' =============================================================================
Public Class NextOccurrenceResult
    Public Property NextDate As Date
    Public Property EndDate As Nullable(Of Date)   ' Nothing = cancelled or no duration
    Public Property Explanation As String
    Public Sub New(ByVal nextDate As Date, ByVal explanation As String)
        Me.NextDate = nextDate
        Me.EndDate = Nothing
        Me.Explanation = explanation
    End Sub
End Class

' =============================================================================
' INPUT PARAMETERS
' =============================================================================
Public Class PackageScheduleParams
    Public Property Recurrence As String
    Public Property PackageId As String
    Public Property Units As String
    Public Property Parameters As String
    Public Property TypeOfDay As String
    Public Property TypeOfDayParameters As String
    Public Property Alteration As Integer
    ' Adjustment parameters (start date)
    Public Property FriAdjustment As String
    Public Property SatAdjustment As String
    Public Property HolidaysAdjustment As String
    ' Duration & end-date adjustment parameters
    Public Property Duration As Integer            ' 0 = no duration
    Public Property EndFriAdjustment As String     ' Fri adjustment for end date
    Public Property EndSatAdjustment As String     ' Sat adjustment for end date
    Public Property EndHolidaysAdjustment As String ' Holiday adjustment for end date
End Class

' =============================================================================
' SCHEDULER ENGINE
' =============================================================================
Public Class PackageScheduler

    Private Shared ReadOnly WeekendDaysDefault As DayOfWeek() = {DayOfWeek.Friday, DayOfWeek.Saturday}
    Private Shared ReadOnly WeekendDaysCountSat As DayOfWeek() = {DayOfWeek.Friday}

    ' -------------------------------------------------------------------------
    ' ADJUSTMENT OPTION CONSTANTS  (shared with the form for combo population)
    ' -------------------------------------------------------------------------
    Public Shared ReadOnly ADJ_THU_BEFORE As String = "Adjust to Thursday Before"
    Public Shared ReadOnly ADJ_SUN_BEFORE As String = "Adjust to Sunday Before"
    Public Shared ReadOnly ADJ_SUN_AFTER As String = "Adjust to Sunday After"
    Public Shared ReadOnly ADJ_CANCELLED As String = "Event Cancelled"
    Public Shared ReadOnly ADJ_FIRST_BIZ_AFTER As String = "Adjust to First Business Day After"
    Public Shared ReadOnly ADJ_LAST_BIZ_BEFORE As String = "Adjust to Last Business Day Before"
    Public Shared ReadOnly ADJ_NONE As String = "No Adjustment"

    ' Fri/Sat adjustment options
    Public Shared ReadOnly FriSatOptions As String() = {
        ADJ_NONE,
        ADJ_THU_BEFORE,
        ADJ_SUN_BEFORE,
        ADJ_SUN_AFTER,
        ADJ_CANCELLED
    }

    ' Holiday adjustment options (superset – includes business-day options)
    Public Shared ReadOnly HolidayOptions As String() = {
        ADJ_NONE,
        ADJ_THU_BEFORE,
        ADJ_SUN_BEFORE,
        ADJ_SUN_AFTER,
        ADJ_CANCELLED,
        ADJ_FIRST_BIZ_AFTER,
        ADJ_LAST_BIZ_BEFORE
    }

    ' -------------------------------------------------------------------------
    ' HARDCODED HOLIDAY LIST  (replace with DB call when ready)
    ' Weekends are NOT listed here – they are handled separately via Fri/Sat rules.
    ' -------------------------------------------------------------------------
    Private Shared ReadOnly HardcodedHolidays As Date() = {
        New Date(2026, 1, 1),   ' New Year's Day
        New Date(2026, 3, 20),   ' Spring Equinox / Nowruz
        New Date(2026, 4, 17),   ' Good Friday (example)
        New Date(2026, 5, 1),   ' Labour Day
        New Date(2026, 6, 15),   ' Example national holiday
        New Date(2026, 8, 15),   ' Example national holiday
        New Date(2026, 12, 25),  ' Christmas Day
        New Date(2026, 12, 31),  ' New Year's Eve
        New Date(2027, 1, 1)    ' New Year's Day 2027
    }

    Public Shared Function IsHoliday(ByVal d As Date) As Boolean
        For Each h As Date In HardcodedHolidays
            If h.Date = d.Date Then Return True
        Next
        Return False
    End Function

    ' -------------------------------------------------------------------------
    ' APPLY FRI / SAT / HOLIDAY ADJUSTMENTS
    ' Called on the raw calculated date before returning to caller.
    ' Chain: Fri/Sat check first, then holiday check (up to 7 iterations).
    ' Returns Nothing when adjustment = "Event Cancelled".
    ' -------------------------------------------------------------------------
    Public Shared Function ApplyAdjustments(ByVal rawDate As Date,
                                            ByVal p As PackageScheduleParams,
                                            ByRef adjustmentNote As String) As Nullable(Of Date)
        adjustmentNote = ""
        Dim current As Date = rawDate

        ' ── Step 1: Fri / Sat adjustment (single pass) ──────────────────────
        If current.DayOfWeek = DayOfWeek.Friday Then
            Dim adj As String = GetAdj(p.FriAdjustment)
            If adj <> ADJ_NONE Then
                Dim moved As Nullable(Of Date) = ApplySingleAdjustment(current, adj)
                If moved Is Nothing Then
                    adjustmentNote = "Event Cancelled (landed on Friday)."
                    Return Nothing
                End If
                adjustmentNote &= "Friday -> " & adj & " -> " & moved.Value.ToString("yyyy-MM-dd") & ". "
                current = moved.Value
            End If

        ElseIf current.DayOfWeek = DayOfWeek.Saturday Then
            Dim adj As String = GetAdj(p.SatAdjustment)
            If adj <> ADJ_NONE Then
                Dim moved As Nullable(Of Date) = ApplySingleAdjustment(current, adj)
                If moved Is Nothing Then
                    adjustmentNote = "Event Cancelled (landed on Saturday)."
                    Return Nothing
                End If
                adjustmentNote &= "Saturday -> " & adj & " -> " & moved.Value.ToString("yyyy-MM-dd") & ". "
                current = moved.Value
            End If
        End If

        ' ── Step 2: Holiday adjustment (chain up to 7 times) ────────────────
        Dim adj2 As String = GetAdj(p.HolidaysAdjustment)
        If adj2 <> ADJ_NONE Then
            Dim iterations As Integer = 0
            Do While IsHoliday(current) AndAlso iterations < 7
                Dim moved As Nullable(Of Date) = ApplySingleAdjustment(current, adj2)
                If moved Is Nothing Then
                    adjustmentNote &= "Event Cancelled (landed on holiday " & current.ToString("yyyy-MM-dd") & ")."
                    Return Nothing
                End If
                adjustmentNote &= "Holiday " & current.ToString("yyyy-MM-dd") &
                                  " -> " & adj2 & " -> " & moved.Value.ToString("yyyy-MM-dd") & ". "
                current = moved.Value
                iterations += 1
            Loop
        End If

        Return current
    End Function

    ' Apply a single named adjustment to a date. Returns Nothing for "Cancelled".
    Private Shared Function ApplySingleAdjustment(ByVal d As Date,
                                                  ByVal adj As String) As Nullable(Of Date)
        Select Case adj
            Case ADJ_THU_BEFORE
                ' Roll back to the nearest Thursday
                Dim candidate As Date = d.AddDays(-1)
                Do While candidate.DayOfWeek <> DayOfWeek.Thursday
                    candidate = candidate.AddDays(-1)
                Loop
                Return candidate

            Case ADJ_SUN_BEFORE
                ' Roll back to the nearest Sunday
                Dim candidate As Date = d.AddDays(-1)
                Do While candidate.DayOfWeek <> DayOfWeek.Sunday
                    candidate = candidate.AddDays(-1)
                Loop
                Return candidate

            Case ADJ_SUN_AFTER
                ' Advance to the nearest Sunday
                Dim candidate As Date = d.AddDays(1)
                Do While candidate.DayOfWeek <> DayOfWeek.Sunday
                    candidate = candidate.AddDays(1)
                Loop
                Return candidate

            Case ADJ_CANCELLED
                Return Nothing

            Case ADJ_FIRST_BIZ_AFTER
                ' First business day (Sun-Thu) strictly after d
                Dim candidate As Date = d.AddDays(1)
                Do While IsWeekendOrHolidayForBiz(candidate)
                    candidate = candidate.AddDays(1)
                Loop
                Return candidate

            Case ADJ_LAST_BIZ_BEFORE
                ' Last business day (Sun-Thu) strictly before d
                Dim candidate As Date = d.AddDays(-1)
                Do While IsWeekendOrHolidayForBiz(candidate)
                    candidate = candidate.AddDays(-1)
                Loop
                Return candidate

            Case Else
                Return d   ' ADJ_NONE or unknown
        End Select
    End Function

    ' Helper: is a date a Fri/Sat weekend OR a holiday (used for biz-day rolls)
    Private Shared Function IsWeekendOrHolidayForBiz(ByVal d As Date) As Boolean
        Return d.DayOfWeek = DayOfWeek.Friday OrElse
               d.DayOfWeek = DayOfWeek.Saturday OrElse
               IsHoliday(d)
    End Function

    ' -------------------------------------------------------------------------
    ' CALCULATE END DATE FROM START DATE + DURATION (calendar days, start=day 0)
    ' Applies the end-date-specific Fri/Sat/Holiday adjustments independently.
    ' Returns Nothing if end date is cancelled by adjustment.
    ' -------------------------------------------------------------------------
    Public Shared Function CalcEndDate(ByVal startDate As Date,
                                       ByVal p As PackageScheduleParams,
                                       ByRef endNote As String) As Nullable(Of Date)
        endNote = ""
        If p.Duration <= 0 Then Return Nothing   ' no duration set

        ' Start date = day 0, so end = startDate + Duration days
        Dim rawEnd As Date = startDate.AddDays(p.Duration)
        endNote = "Raw end: " & startDate.ToString("yyyy-MM-dd") &
                  " + " & p.Duration.ToString() & " days = " & rawEnd.ToString("yyyy-MM-dd") & ". "

        ' Build a temporary params object carrying the END adjustments
        Dim endAdj As New PackageScheduleParams
        endAdj.FriAdjustment = GetAdj(p.EndFriAdjustment)
        endAdj.SatAdjustment = GetAdj(p.EndSatAdjustment)
        endAdj.HolidaysAdjustment = GetAdj(p.EndHolidaysAdjustment)

        Dim adjNote As String = ""
        Dim adjusted As Nullable(Of Date) = ApplyAdjustments(rawEnd, endAdj, adjNote)
        If adjNote <> "" Then endNote &= "End adjustment: " & adjNote
        Return adjusted
    End Function

    Public Shared Function GetNextOccurrence(ByVal fromDate As Date, ByVal p As PackageScheduleParams) As NextOccurrenceResult
        ' Step 1: calculate raw next date
        Dim raw As NextOccurrenceResult
        Select Case p.Recurrence.Trim().ToLower()
            Case "daily"
                raw = CalcDaily(fromDate, p)
            Case "weekly"
                raw = CalcWeekly(fromDate, p)
            Case "monthly"
                raw = CalcMonthly(fromDate, p)
            Case "quarterly"
                raw = CalcQuarterlyOrSemi(fromDate, p, "Quarterly")
            Case "semiannually"
                raw = CalcQuarterlyOrSemi(fromDate, p, "SemiAnnually")
            Case "annually"
                raw = CalcAnnually(fromDate, p)
            Case Else
                Throw New ArgumentException("Unknown Recurrence: '" & p.Recurrence & "'")
        End Select

        ' Step 2: apply Fri / Sat / Holiday adjustments
        Dim adjNote As String = ""
        Dim adjusted As Nullable(Of Date) = ApplyAdjustments(raw.NextDate, p, adjNote)

        If adjusted Is Nothing Then
            ' Event cancelled – return a result with a sentinel date and clear explanation
            Return New NextOccurrenceResult(Date.MinValue,
                raw.Explanation & " | ADJUSTMENT: " & adjNote)
        End If

        Dim finalExplanation As String = raw.Explanation
        If adjNote <> "" Then
            finalExplanation &= " | ADJUSTMENT: " & adjNote
        End If

        Dim finalResult As New NextOccurrenceResult(adjusted.Value, finalExplanation)

        ' Step 3: calculate end date if duration is set
        If p.Duration > 0 Then
            Dim endNote As String = ""
            Dim endDate As Nullable(Of Date) = CalcEndDate(adjusted.Value, p, endNote)
            finalResult.EndDate = endDate
            If endNote <> "" Then
                finalResult.Explanation &= " | END DATE: " & endNote
            End If
        End If

        Return finalResult
    End Function

    ' -------------------------------------------------------------------------
    ' DAILY
    ' -------------------------------------------------------------------------
    Private Shared Function CalcDaily(ByVal fromDate As Date, ByVal p As PackageScheduleParams) As NextOccurrenceResult
        Dim stepVal As Integer
        If p.Alteration > 0 Then
            stepVal = p.Alteration
        Else
            stepVal = 1
        End If

        Dim weekends As DayOfWeek() = GetWeekends(p.TypeOfDayParameters)
        Dim isCalendar As Boolean = (p.TypeOfDay.Trim().ToLower() = "calendar")

        If isCalendar Then
            Dim nxt As Date = fromDate.AddDays(stepVal)
            Return New NextOccurrenceResult(nxt,
                "Daily (Calendar): advanced " & stepVal.ToString() & " day(s) from " & fromDate.ToString("yyyy-MM-dd") & ".")
        Else
            Dim counted As Integer = 0
            Dim candidate As Date = fromDate
            Do
                candidate = candidate.AddDays(1)
                If Not IsWeekend(candidate, weekends) Then
                    counted += 1
                End If
            Loop Until counted = stepVal

            Dim wd As String
            If weekends.Contains(DayOfWeek.Saturday) Then
                wd = "Fri+Sat weekend"
            Else
                wd = "Fri-only weekend"
            End If

            Return New NextOccurrenceResult(candidate,
                "Daily (Business, " & wd & "): advanced " & stepVal.ToString() & " business day(s) from " & fromDate.ToString("yyyy-MM-dd") & ".")
        End If
    End Function

    ' -------------------------------------------------------------------------
    ' WEEKLY
    ' -------------------------------------------------------------------------
    Private Shared Function CalcWeekly(ByVal fromDate As Date, ByVal p As PackageScheduleParams) As NextOccurrenceResult
        Dim stepVal As Integer
        If p.Alteration > 0 Then
            stepVal = p.Alteration
        Else
            stepVal = 1
        End If

        Dim allowedDays As New List(Of DayOfWeek)
        If Not String.IsNullOrEmpty(p.Parameters) AndAlso p.Parameters.Trim() <> "" Then
            Dim parts As String() = p.Parameters.Split(New Char() {"#"c})
            For Each part As String In parts
                Dim d As DayOfWeek = DayOfWeek.Sunday
                If TryParseDayName(part.Trim(), d) Then
                    allowedDays.Add(d)
                End If
            Next
        End If

        If allowedDays.Count = 0 Then
            Return New NextOccurrenceResult(fromDate.AddDays(stepVal * 7),
                "Weekly: advanced " & stepVal.ToString() & " week(s) from " & fromDate.ToString("yyyy-MM-dd") & ".")
        End If

        Dim weekStart As Date = StartOfWeek(fromDate.AddDays(1))
        Dim weeksChecked As Integer = 0

        Do While weeksChecked < 1000
            Dim sortedDays As List(Of DayOfWeek) = allowedDays.OrderBy(Function(x) x).ToList()
            For Each dow As DayOfWeek In sortedDays
                Dim dayInWeek As Date = weekStart.AddDays(CInt(dow))
                If dayInWeek > fromDate Then
                    Return New NextOccurrenceResult(dayInWeek,
                        "Weekly (every " & stepVal.ToString() & " week(s), days=" & p.Parameters & "): next allowed day after " & fromDate.ToString("yyyy-MM-dd") & ".")
                End If
            Next
            weeksChecked += stepVal
            weekStart = weekStart.AddDays(stepVal * 7)
        Loop

        Throw New InvalidOperationException("Could not find next weekly occurrence.")
    End Function

    ' -------------------------------------------------------------------------
    ' MONTHLY
    ' -------------------------------------------------------------------------
    Private Shared Function CalcMonthly(ByVal fromDate As Date, ByVal p As PackageScheduleParams) As NextOccurrenceResult
        Dim stepVal As Integer
        If p.Alteration > 0 Then
            stepVal = p.Alteration
        Else
            stepVal = 1
        End If

        ' Try current month first; if the resolved date is still in the future use it,
        ' otherwise advance by stepVal months and try again.
        Dim candidateMonth As New Date(fromDate.Year, fromDate.Month, 1)
        Dim result As NextOccurrenceResult = ResolveMonthlyDate(candidateMonth, p,
            "Monthly (every " & stepVal.ToString() & " month(s))", fromDate)

        If result.NextDate > fromDate Then
            Return result
        End If

        ' Resolved date has already passed – advance by stepVal months
        candidateMonth = candidateMonth.AddMonths(stepVal)
        Return ResolveMonthlyDate(candidateMonth, p,
            "Monthly (every " & stepVal.ToString() & " month(s))", fromDate)
    End Function

    ' -------------------------------------------------------------------------
    ' QUARTERLY / SEMI-ANNUALLY
    ' -------------------------------------------------------------------------
    Private Shared Function CalcQuarterlyOrSemi(ByVal fromDate As Date, ByVal p As PackageScheduleParams, ByVal mode As String) As NextOccurrenceResult
        Dim stepVal As Integer
        If p.Alteration > 0 Then
            stepVal = p.Alteration
        Else
            stepVal = 1
        End If

        Dim cycleMonths As New List(Of Integer)
        If Not String.IsNullOrEmpty(p.Parameters) AndAlso p.Parameters.Trim() <> "" Then
            Dim parts As String() = p.Parameters.Split(New Char() {","c})
            For Each part As String In parts
                Dim m As Integer = ParseMonthName(part.Trim())
                If m > 0 Then
                    cycleMonths.Add(m)
                End If
            Next
        End If

        If cycleMonths.Count = 0 Then
            Throw New ArgumentException(mode & ": PARAMETERS must contain comma-separated month names.")
        End If

        cycleMonths.Sort()

        ' Search from current month (not fromDate+1 day) so we can check if today's
        ' cycle month still has a future occurrence date.
        Dim searchYear As Integer = fromDate.Year
        Dim maxYear As Integer = searchYear + 10
        Dim occurrencesFound As Integer = 0

        Do While searchYear <= maxYear
            For Each m As Integer In cycleMonths
                ' Only consider cycle months that are >= current month/year
                If searchYear > fromDate.Year OrElse
                   (searchYear = fromDate.Year AndAlso m >= fromDate.Month) Then

                    Dim targetMonth As New Date(searchYear, m, 1)
                    Dim candidate As NextOccurrenceResult = ResolveMonthlyDate(targetMonth, p,
                        mode & " (every " & stepVal.ToString() & " cycle(s), group=" & p.Parameters & ")", fromDate)

                    ' Accept only if the resolved date is strictly after fromDate
                    If candidate.NextDate > fromDate Then
                        occurrencesFound += 1
                        If occurrencesFound = stepVal Then
                            Return candidate
                        End If
                    End If
                End If
            Next
            searchYear += 1
        Loop

        Throw New InvalidOperationException("Could not find next " & mode & " occurrence.")
    End Function

    ' -------------------------------------------------------------------------
    ' ANNUALLY
    ' -------------------------------------------------------------------------
    Private Shared Function CalcAnnually(ByVal fromDate As Date, ByVal p As PackageScheduleParams) As NextOccurrenceResult
        Dim stepVal As Integer
        If p.Alteration > 0 Then
            stepVal = p.Alteration
        Else
            stepVal = 1
        End If

        Dim targetMonthNum As Integer = 1
        If Not String.IsNullOrEmpty(p.Parameters) AndAlso p.Parameters.Trim() <> "" Then
            Dim parsed As Integer = ParseMonthName(p.Parameters.Trim())
            If parsed > 0 Then
                targetMonthNum = parsed
            End If
        End If

        ' Try this year's occurrence first; advance by stepVal years only if already passed
        Dim targetYear As Integer = fromDate.Year
        Dim targetMonth As New Date(targetYear, targetMonthNum, 1)
        Dim result As NextOccurrenceResult = ResolveMonthlyDate(targetMonth, p,
            "Annually (every " & stepVal.ToString() & " year(s), month=" & p.Parameters & ")", fromDate)

        If result.NextDate > fromDate Then
            Return result
        End If

        ' This year's date has already passed – advance by stepVal years
        targetMonth = New Date(targetYear + stepVal, targetMonthNum, 1)
        Return ResolveMonthlyDate(targetMonth, p,
            "Annually (every " & stepVal.ToString() & " year(s), month=" & p.Parameters & ")", fromDate)
    End Function

    ' -------------------------------------------------------------------------
    ' RESOLVE DATE WITHIN TARGET MONTH
    ' -------------------------------------------------------------------------
    Private Shared Function ResolveMonthlyDate(ByVal targetMonth As Date, ByVal p As PackageScheduleParams,
                                               ByVal context As String, ByVal fromDate As Date) As NextOccurrenceResult
        Dim weekends As DayOfWeek() = GetWeekends(p.TypeOfDayParameters)
        Dim isBusiness As Boolean = (p.TypeOfDay.Trim().ToLower() = "business")

        Select Case p.TypeOfDay.Trim().ToLower()
            Case "dateofday"
                Dim dayNum As Integer = 1
                Integer.TryParse(p.TypeOfDayParameters, dayNum)
                Dim resolved As Date = ClampToMonth(targetMonth.Year, targetMonth.Month, dayNum)
                Return New NextOccurrenceResult(resolved,
                    context & ": fixed day " & dayNum.ToString() & " of " &
                    targetMonth.ToString("MMMM yyyy") & " -> " & resolved.ToString("yyyy-MM-dd") & ".")

            Case "locationofday"
                Dim location As String = p.TypeOfDayParameters.Trim().ToUpper()
                Dim resolved As Date = ResolveLocationOfDay(targetMonth, location, isBusiness, weekends)
                Return New NextOccurrenceResult(resolved,
                    context & ": LocationOfDay=" & location & " in " &
                    targetMonth.ToString("MMMM yyyy") & " -> " & resolved.ToString("yyyy-MM-dd") & ".")

            Case Else
                Return New NextOccurrenceResult(targetMonth,
                    context & ": defaulting to 1st of " & targetMonth.ToString("MMMM yyyy") & ".")
        End Select
    End Function

    Private Shared Function ResolveLocationOfDay(ByVal targetMonth As Date, ByVal location As String,
                                                 ByVal isBusiness As Boolean, ByVal weekends As DayOfWeek()) As Date
        Dim yr As Integer = targetMonth.Year
        Dim mo As Integer = targetMonth.Month
        Dim daysInMonth As Integer = Date.DaysInMonth(yr, mo)
        Dim resolved As Date

        ' Parse combined format: BASE[+|-]OFFSET  e.g. "EOM-5", "BOM+2", "MOM+0"
        ' Also accept plain "BOM", "MOM", "EOM" with no offset (treated as +0)
        Dim baseStr As String = ""
        Dim sign As Integer = 1
        Dim offset As Integer = 0

        ' Find sign position
        Dim signPos As Integer = -1
        Dim i As Integer
        For i = 1 To location.Length - 1   ' start at 1 to skip any leading char
            If location(i) = "+"c OrElse location(i) = "-"c Then
                signPos = i
                Exit For
            End If
        Next i

        If signPos > 0 Then
            baseStr = location.Substring(0, signPos).ToUpper().Trim()
            If location(signPos) = "-"c Then
                sign = -1
            Else
                sign = 1
            End If
            Integer.TryParse(location.Substring(signPos + 1), offset)
        Else
            baseStr = location.ToUpper().Trim()
            sign = 1
            offset = 0
        End If

        ' Resolve base anchor
        Select Case baseStr
            Case "BOM"
                resolved = New Date(yr, mo, 1)
            Case "MOM"
                resolved = New Date(yr, mo, CInt(Math.Ceiling(daysInMonth / 2)))
            Case "EOM"
                resolved = New Date(yr, mo, daysInMonth)
            Case Else
                resolved = New Date(yr, mo, 1)   ' fallback
        End Select

        ' Apply offset
        resolved = resolved.AddDays(sign * offset)

        ' Clamp back into the same month if offset pushed outside
        If resolved.Month <> mo OrElse resolved.Year <> yr Then
            If sign > 0 Then
                resolved = New Date(yr, mo, daysInMonth)
            Else
                resolved = New Date(yr, mo, 1)
            End If
        End If

        ' Roll back to last business day if needed
        If isBusiness Then
            Do While IsWeekend(resolved, weekends)
                resolved = resolved.AddDays(-1)
            Loop
        End If

        Return resolved
    End Function

    ' Helper: return ADJ_NONE when value is blank, otherwise return the value
    Public Shared Function GetAdj(ByVal val As String) As String
        If String.IsNullOrEmpty(val) Then
            Return ADJ_NONE
        End If
        Return val
    End Function

    ' -------------------------------------------------------------------------
    ' HELPERS
    ' -------------------------------------------------------------------------
    Private Shared Function IsWeekend(ByVal d As Date, ByVal weekends As DayOfWeek()) As Boolean
        Return weekends.Contains(d.DayOfWeek)
    End Function

    Private Shared Function GetWeekends(ByVal typeOfDayParams As String) As DayOfWeek()
        If Not String.IsNullOrEmpty(typeOfDayParams) AndAlso
           typeOfDayParams.Trim().ToUpper() = "COUNT SAT" Then
            Return WeekendDaysCountSat
        End If
        Return WeekendDaysDefault
    End Function

    Private Shared Function StartOfWeek(ByVal d As Date) As Date
        Dim diff As Integer = CInt(d.DayOfWeek) - CInt(DayOfWeek.Sunday)
        Return d.AddDays(-diff).Date
    End Function

    Private Shared Function ClampToMonth(ByVal yr As Integer, ByVal mo As Integer, ByVal day As Integer) As Date
        Return New Date(yr, mo, Math.Min(day, Date.DaysInMonth(yr, mo)))
    End Function

    Private Shared Function TryParseDayName(ByVal name As String, ByRef result As DayOfWeek) As Boolean
        Select Case name.ToLower()
            Case "sun", "sunday" : result = DayOfWeek.Sunday : Return True
            Case "mon", "monday" : result = DayOfWeek.Monday : Return True
            Case "tue", "tuesday" : result = DayOfWeek.Tuesday : Return True
            Case "wen", "wed", "wednesday" : result = DayOfWeek.Wednesday : Return True
            Case "thu", "thursday" : result = DayOfWeek.Thursday : Return True
            Case "fri", "friday" : result = DayOfWeek.Friday : Return True
            Case "sat", "saturday" : result = DayOfWeek.Saturday : Return True
            Case Else : result = DayOfWeek.Sunday : Return False
        End Select
    End Function

    Private Shared Function ParseMonthName(ByVal name As String) As Integer
        Select Case name.ToLower()
            Case "jan", "january" : Return 1
            Case "feb", "february" : Return 2
            Case "mar", "march" : Return 3
            Case "apr", "april" : Return 4
            Case "may" : Return 5
            Case "jun", "june" : Return 6
            Case "jul", "july" : Return 7
            Case "aug", "august" : Return 8
            Case "sep", "september" : Return 9
            Case "oct", "october" : Return 10
            Case "nov", "november" : Return 11
            Case "dec", "december" : Return 12
            Case Else : Return 0
        End Select
    End Function

End Class

' =============================================================================
' WINDOWS FORM
' =============================================================================
Public Class SchedulerForm
    Inherits Form

    Private cmbRecurrence As ComboBox
    Private cmbTypeOfDay As ComboBox
    Private cmbTypeOfDayParams As ComboBox
    Private cmbParameters As ComboBox
    Private clbWeekDays As CheckedListBox
    Private lblWeekDaysHint As Label
    Private dtpFromDate As DateTimePicker
    Private txtPackageId As TextBox
    Private nudAlteration As NumericUpDown
    Private lblParameters As Label
    Private lblTypeOfDayParams As Label
    Private cmbLocBase As ComboBox
    Private cmbLocSign As ComboBox
    Private nudLocOffset As NumericUpDown
    Private lblLocFormula As Label
    Private btnCalculate As Button
    Private pnlResult As Panel
    Private lblResultDate As Label
    Private lblExplanationText As Label
    Private cmbFriAdj As ComboBox
    Private cmbSatAdj As ComboBox
    Private cmbHolAdj As ComboBox
    Private nudDuration As NumericUpDown
    Private cmbEndFriAdj As ComboBox
    Private cmbEndSatAdj As ComboBox
    Private cmbEndHolAdj As ComboBox
    Private lblEndDateCaption As Label
    Private lblEndDate As Label

    ' Guard flag – prevents event handlers firing before all controls are created
    Private _initialising As Boolean = True

    ' Colour palette
    Private ReadOnly clrBackground As Color = Color.FromArgb(18, 22, 36)
    Private ReadOnly clrPanel As Color = Color.FromArgb(26, 32, 52)
    Private ReadOnly clrAccent As Color = Color.FromArgb(56, 189, 248)
    Private ReadOnly clrAccent2 As Color = Color.FromArgb(99, 102, 241)
    Private ReadOnly clrText As Color = Color.FromArgb(226, 232, 240)
    Private ReadOnly clrSubtext As Color = Color.FromArgb(100, 116, 139)
    Private ReadOnly clrInput As Color = Color.FromArgb(30, 41, 59)
    Private ReadOnly clrBorder As Color = Color.FromArgb(51, 65, 85)
    Private ReadOnly clrSuccess As Color = Color.FromArgb(52, 211, 153)

    Public Sub New()
        InitializeComponents()
        _initialising = False
        UpdateDynamicControls()
    End Sub

    Private Sub InitializeComponents()
        Me.Text = "Package Occurrence Scheduler"
        Me.Size = New Size(636, 900)
        Me.MinimumSize = New Size(636, 600)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.BackColor = clrBackground
        Me.ForeColor = clrText
        Me.Font = New Font("Segoe UI", 9)
        Me.FormBorderStyle = FormBorderStyle.FixedSingle
        Me.MaximizeBox = False

        ' Header
        Dim pnlHeader As New Panel
        pnlHeader.Bounds = New Rectangle(0, 0, 620, 70)
        pnlHeader.BackColor = clrPanel
        AddHandler pnlHeader.Paint, AddressOf PaintHeaderBorder

        Dim lblTitle As New Label
        lblTitle.Text = "Package Scheduler"
        lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
        lblTitle.ForeColor = clrAccent
        lblTitle.AutoSize = True
        lblTitle.Location = New Point(24, 12)
        pnlHeader.Controls.Add(lblTitle)

        Dim lblSubtitle As New Label
        lblSubtitle.Text = "Calculate the next occurrence of a recurring package"
        lblSubtitle.Font = New Font("Segoe UI", 8.5F)
        lblSubtitle.ForeColor = clrSubtext
        lblSubtitle.AutoSize = True
        lblSubtitle.Location = New Point(26, 42)
        pnlHeader.Controls.Add(lblSubtitle)
        Me.Controls.Add(pnlHeader)

        ' Main input panel
        Dim pnlMain As New Panel
        pnlMain.Bounds = New Rectangle(20, 84, 576, 468)
        pnlMain.BackColor = clrPanel
        AddHandler pnlMain.Paint, AddressOf PaintBorderedPanel
        Me.Controls.Add(pnlMain)

        Dim y As Integer = 20

        ' Package ID
        pnlMain.Controls.Add(MakeSectionLabel("PACKAGE ID", 20, y))
        txtPackageId = New TextBox
        txtPackageId.Bounds = New Rectangle(20, y + 20, 536, 28)
        txtPackageId.BackColor = clrInput
        txtPackageId.ForeColor = clrText
        txtPackageId.BorderStyle = BorderStyle.FixedSingle
        txtPackageId.Font = New Font("Segoe UI", 9)
        pnlMain.Controls.Add(txtPackageId)
        y += 64

        ' From Date + Alteration
        pnlMain.Controls.Add(MakeSectionLabel("FROM DATE", 20, y))
        pnlMain.Controls.Add(MakeSectionLabel("ALTERATION (multiplier)", 310, y))

        dtpFromDate = New DateTimePicker
        dtpFromDate.Bounds = New Rectangle(20, y + 20, 270, 28)
        dtpFromDate.Format = DateTimePickerFormat.Short
        dtpFromDate.Value = Date.Today
        dtpFromDate.CalendarMonthBackground = clrPanel
        dtpFromDate.CalendarForeColor = clrText
        pnlMain.Controls.Add(dtpFromDate)

        nudAlteration = New NumericUpDown
        nudAlteration.Bounds = New Rectangle(310, y + 20, 246, 28)
        nudAlteration.Minimum = 1
        nudAlteration.Maximum = 100
        nudAlteration.Value = 1
        nudAlteration.BackColor = clrInput
        nudAlteration.ForeColor = clrText
        nudAlteration.BorderStyle = BorderStyle.FixedSingle
        nudAlteration.Font = New Font("Segoe UI", 9)
        pnlMain.Controls.Add(nudAlteration)
        y += 64

        ' Recurrence
        pnlMain.Controls.Add(MakeSectionLabel("RECURRENCE", 20, y))
        cmbRecurrence = MakeCombo(20, y + 20, 536)
        cmbRecurrence.Items.AddRange(New Object() {"Daily", "Weekly", "Monthly", "Quarterly", "SemiAnnually", "Annually"})
        cmbRecurrence.SelectedIndex = 0
        AddHandler cmbRecurrence.SelectedIndexChanged, AddressOf cmbRecurrence_SelectedIndexChanged
        pnlMain.Controls.Add(cmbRecurrence)
        y += 64

        ' Parameters – combo for non-Weekly, CheckedListBox for Weekly
        lblParameters = MakeSectionLabel("PARAMETERS", 20, y)
        pnlMain.Controls.Add(lblParameters)

        ' Standard combo (used for Quarterly / SemiAnnually / Annually)
        cmbParameters = MakeCombo(20, y + 20, 536)
        pnlMain.Controls.Add(cmbParameters)

        ' Weekly day selector: CheckedListBox (Sat → Fri order, business week = Sun-Thu)
        clbWeekDays = New CheckedListBox
        clbWeekDays.Bounds = New Rectangle(20, y + 20, 536, 110)
        clbWeekDays.BackColor = clrInput
        clbWeekDays.ForeColor = clrText
        clbWeekDays.Font = New Font("Segoe UI", 9)
        clbWeekDays.BorderStyle = BorderStyle.FixedSingle
        clbWeekDays.CheckOnClick = True
        clbWeekDays.MultiColumn = True       ' show days side-by-side
        clbWeekDays.ColumnWidth = 76
        ' Add days Sat → Fri so business days (Sun-Thu) appear naturally first
        clbWeekDays.Items.Add("Sun", True)
        clbWeekDays.Items.Add("Mon", True)
        clbWeekDays.Items.Add("Tue", True)
        clbWeekDays.Items.Add("Wed", True)
        clbWeekDays.Items.Add("Thu", True)
        clbWeekDays.Items.Add("Fri", False)
        clbWeekDays.Items.Add("Sat", False)
        clbWeekDays.Visible = False
        AddHandler clbWeekDays.ItemCheck, AddressOf clbWeekDays_ItemCheck
        pnlMain.Controls.Add(clbWeekDays)

        ' Hint label showing assembled #-string below the checklist
        lblWeekDaysHint = New Label
        lblWeekDaysHint.Font = New Font("Consolas", 8.5F)
        lblWeekDaysHint.ForeColor = clrSubtext
        lblWeekDaysHint.AutoSize = True
        lblWeekDaysHint.Location = New Point(20, y + 134)
        lblWeekDaysHint.Text = ""
        lblWeekDaysHint.Visible = False
        pnlMain.Controls.Add(lblWeekDaysHint)

        y += 64

        ' TypeOfDay + TypeOfDayParams (single combo – used for Daily & DateOfDay)
        pnlMain.Controls.Add(MakeSectionLabel("TYPE OF DAY", 20, y))
        lblTypeOfDayParams = MakeSectionLabel("TYPE OF DAY PARAMETERS", 310, y)
        pnlMain.Controls.Add(lblTypeOfDayParams)
        cmbTypeOfDay = MakeCombo(20, y + 20, 270)
        AddHandler cmbTypeOfDay.SelectedIndexChanged, AddressOf cmbTypeOfDay_SelectedIndexChanged
        pnlMain.Controls.Add(cmbTypeOfDay)
        cmbTypeOfDayParams = MakeCombo(310, y + 20, 246)
        pnlMain.Controls.Add(cmbTypeOfDayParams)
        y += 64

        ' ── LocationOfDay three-part row ──────────────────────────────────
        ' Label row
        pnlMain.Controls.Add(MakeSectionLabel("BASE", 20, y))
        pnlMain.Controls.Add(MakeSectionLabel("SIGN", 170, y))
        pnlMain.Controls.Add(MakeSectionLabel("OFFSET (days)", 260, y))
        pnlMain.Controls.Add(MakeSectionLabel("FORMULA PREVIEW", 400, y))

        ' Base combo  BOM / MOM / EOM
        cmbLocBase = MakeCombo(20, y + 20, 140)
        cmbLocBase.Items.AddRange(New Object() {"BOM", "MOM", "EOM"})
        cmbLocBase.SelectedIndex = 0
        AddHandler cmbLocBase.SelectedIndexChanged, AddressOf cmbLocBase_SelectedIndexChanged
        pnlMain.Controls.Add(cmbLocBase)

        ' Sign combo  + / -
        cmbLocSign = MakeCombo(170, y + 20, 80)
        cmbLocSign.Items.AddRange(New Object() {"+", "-"})
        cmbLocSign.SelectedIndex = 0
        AddHandler cmbLocSign.SelectedIndexChanged, AddressOf cmbLocSign_SelectedIndexChanged
        pnlMain.Controls.Add(cmbLocSign)

        ' Offset spinner  0-31
        nudLocOffset = New NumericUpDown
        nudLocOffset.Bounds = New Rectangle(260, y + 20, 130, 28)
        nudLocOffset.Minimum = 0
        nudLocOffset.Maximum = 31
        nudLocOffset.Value = 0
        nudLocOffset.BackColor = clrInput
        nudLocOffset.ForeColor = clrText
        nudLocOffset.BorderStyle = BorderStyle.FixedSingle
        nudLocOffset.Font = New Font("Segoe UI", 9)
        AddHandler nudLocOffset.ValueChanged, AddressOf nudLocOffset_ValueChanged
        pnlMain.Controls.Add(nudLocOffset)

        ' Formula preview label
        lblLocFormula = New Label
        lblLocFormula.Text = "BOM+0"
        lblLocFormula.Font = New Font("Consolas", 11, FontStyle.Bold)
        lblLocFormula.ForeColor = clrAccent
        lblLocFormula.AutoSize = True
        lblLocFormula.Location = New Point(400, y + 22)
        pnlMain.Controls.Add(lblLocFormula)

        y += 64

        ' ═══════════════════════════════════════════════════════════════════
        ' TAB CONTROL  – Tab 1: Starting Date Adjustment
        '               Tab 2: Final Date Adjustment (Duration + End Adj)
        ' Placed directly on the FORM so nothing inside pnlMain can ever
        ' overlap it, and the button is a plain Me.Controls child too.
        ' ═══════════════════════════════════════════════════════════════════
        Dim tabTop As Integer = pnlMain.Top + pnlMain.Height + 10

        Dim tabCtrl As New TabControl
        tabCtrl.Bounds = New Rectangle(20, tabTop, 576, 220)
        tabCtrl.Font = New Font("Segoe UI", 9, FontStyle.Bold)
        Me.Controls.Add(tabCtrl)

        ' ── Tab 1 : Starting Date Adjustment ────────────────────────────
        Dim tabStart As New TabPage
        tabStart.Text = "  Starting Date Adjustment  "
        tabStart.BackColor = clrPanel
        tabStart.ForeColor = clrText
        tabCtrl.TabPages.Add(tabStart)

        tabStart.Controls.Add(MakeSectionLabel("FRI_ADJUSTMENT", 10, 14))
        tabStart.Controls.Add(MakeSectionLabel("SAT_ADJUSTMENT", 288, 14))

        cmbFriAdj = MakeCombo(10, 32, 265)
        For Each opt As String In PackageScheduler.FriSatOptions
            cmbFriAdj.Items.Add(opt)
        Next
        cmbFriAdj.SelectedIndex = 0
        tabStart.Controls.Add(cmbFriAdj)

        cmbSatAdj = MakeCombo(288, 32, 265)
        For Each opt As String In PackageScheduler.FriSatOptions
            cmbSatAdj.Items.Add(opt)
        Next
        cmbSatAdj.SelectedIndex = 0
        tabStart.Controls.Add(cmbSatAdj)

        tabStart.Controls.Add(MakeSectionLabel("HOLIDAYS_ADJUSTMENT", 10, 74))
        cmbHolAdj = MakeCombo(10, 92, 540)
        For Each opt As String In PackageScheduler.HolidayOptions
            cmbHolAdj.Items.Add(opt)
        Next
        cmbHolAdj.SelectedIndex = 0
        tabStart.Controls.Add(cmbHolAdj)

        ' ── Tab 2 : Final Date Adjustment ───────────────────────────────
        Dim tabEnd As New TabPage
        tabEnd.Text = "  Final Date Adjustment  "
        tabEnd.BackColor = clrPanel
        tabEnd.ForeColor = clrText
        tabCtrl.TabPages.Add(tabEnd)

        tabEnd.Controls.Add(MakeSectionLabel("DURATION  (calendar days, start = day 0  |  0 = disabled)", 10, 14))
        nudDuration = New NumericUpDown
        nudDuration.Bounds = New Rectangle(10, 32, 150, 28)
        nudDuration.Minimum = 0
        nudDuration.Maximum = 9999
        nudDuration.Value = 0
        nudDuration.BackColor = clrInput
        nudDuration.ForeColor = clrText
        nudDuration.BorderStyle = BorderStyle.FixedSingle
        nudDuration.Font = New Font("Segoe UI", 9)
        tabEnd.Controls.Add(nudDuration)

        Dim lblDurHint As New Label
        lblDurHint.Text = "e.g.  start 10-Jun + 5 days  =  end 15-Jun"
        lblDurHint.Font = New Font("Segoe UI", 8)
        lblDurHint.ForeColor = clrSubtext
        lblDurHint.AutoSize = True
        lblDurHint.Location = New Point(170, 38)
        tabEnd.Controls.Add(lblDurHint)

        tabEnd.Controls.Add(MakeSectionLabel("END FRI_ADJUSTMENT", 10, 74))
        tabEnd.Controls.Add(MakeSectionLabel("END SAT_ADJUSTMENT", 288, 74))

        cmbEndFriAdj = MakeCombo(10, 92, 265)
        For Each opt As String In PackageScheduler.FriSatOptions
            cmbEndFriAdj.Items.Add(opt)
        Next
        cmbEndFriAdj.SelectedIndex = 0
        tabEnd.Controls.Add(cmbEndFriAdj)

        cmbEndSatAdj = MakeCombo(288, 92, 265)
        For Each opt As String In PackageScheduler.FriSatOptions
            cmbEndSatAdj.Items.Add(opt)
        Next
        cmbEndSatAdj.SelectedIndex = 0
        tabEnd.Controls.Add(cmbEndSatAdj)

        tabEnd.Controls.Add(MakeSectionLabel("END HOLIDAYS_ADJUSTMENT", 10, 134))
        cmbEndHolAdj = MakeCombo(10, 152, 540)
        For Each opt As String In PackageScheduler.HolidayOptions
            cmbEndHolAdj.Items.Add(opt)
        Next
        cmbEndHolAdj.SelectedIndex = 0
        tabEnd.Controls.Add(cmbEndHolAdj)

        ' ═══════════════════════════════════════════════════════════════════
        ' CALCULATE BUTTON  – directly on form, below TabControl
        ' ═══════════════════════════════════════════════════════════════════
        Dim btnTop As Integer = tabTop + tabCtrl.Height + 10

        btnCalculate = New Button
        btnCalculate.Text = "CALCULATE NEXT OCCURRENCE"
        btnCalculate.Bounds = New Rectangle(20, btnTop, 576, 46)
        btnCalculate.BackColor = clrAccent2
        btnCalculate.ForeColor = Color.White
        btnCalculate.FlatStyle = FlatStyle.Flat
        btnCalculate.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        btnCalculate.Cursor = Cursors.Hand
        btnCalculate.FlatAppearance.BorderSize = 0
        AddHandler btnCalculate.Click, AddressOf btnCalculate_Click
        Me.Controls.Add(btnCalculate)

        ' ═══════════════════════════════════════════════════════════════════
        ' RESULT PANEL – directly on form, below button
        ' ═══════════════════════════════════════════════════════════════════
        Dim resTop As Integer = btnTop + 56

        pnlResult = New Panel
        pnlResult.Bounds = New Rectangle(20, resTop, 576, 200)
        pnlResult.BackColor = clrPanel
        pnlResult.Visible = False
        AddHandler pnlResult.Paint, AddressOf PaintBorderedPanel
        Me.Controls.Add(pnlResult)

        ' Start date
        Dim lblStartCap As New Label
        lblStartCap.Text = "NEXT OCCURRENCE  (START DATE)"
        lblStartCap.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblStartCap.ForeColor = clrAccent
        lblStartCap.AutoSize = True
        lblStartCap.Location = New Point(20, 14)
        pnlResult.Controls.Add(lblStartCap)

        lblResultDate = New Label
        lblResultDate.Text = ""
        lblResultDate.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblResultDate.ForeColor = clrSuccess
        lblResultDate.AutoSize = True
        lblResultDate.Location = New Point(18, 30)
        pnlResult.Controls.Add(lblResultDate)

        ' Final date
        lblEndDateCaption = New Label
        lblEndDateCaption.Text = "FINAL DATE"
        lblEndDateCaption.Font = New Font("Segoe UI", 8, FontStyle.Bold)
        lblEndDateCaption.ForeColor = clrAccent
        lblEndDateCaption.AutoSize = True
        lblEndDateCaption.Location = New Point(20, 96)
        lblEndDateCaption.Visible = False
        pnlResult.Controls.Add(lblEndDateCaption)

        lblEndDate = New Label
        lblEndDate.Text = ""
        lblEndDate.Font = New Font("Segoe UI", 20, FontStyle.Bold)
        lblEndDate.ForeColor = Color.FromArgb(251, 191, 36)
        lblEndDate.AutoSize = True
        lblEndDate.Location = New Point(18, 112)
        lblEndDate.Visible = False
        pnlResult.Controls.Add(lblEndDate)

        ' Explanation
        lblExplanationText = New Label
        lblExplanationText.Text = ""
        lblExplanationText.Font = New Font("Consolas", 7.5F)
        lblExplanationText.ForeColor = clrSubtext
        lblExplanationText.Location = New Point(18, 168)
        lblExplanationText.Size = New Size(538, 24)
        lblExplanationText.AutoSize = False
        pnlResult.Controls.Add(lblExplanationText)

        ' Auto-size form to fit all content
        Me.Height = resTop + pnlResult.Height + 50
    End Sub

    ' -------------------------------------------------------------------------
    ' DYNAMIC CONTROL POPULATION
    ' -------------------------------------------------------------------------
    Private Sub UpdateDynamicControls()
        If _initialising Then Exit Sub
        Dim rec As String = "Daily"
        If cmbRecurrence IsNot Nothing AndAlso cmbRecurrence.SelectedItem IsNot Nothing Then
            rec = cmbRecurrence.SelectedItem.ToString()
        End If

        ' Show/hide the right Parameters control
        cmbParameters.Items.Clear()
        clbWeekDays.Visible = False
        lblWeekDaysHint.Visible = False
        cmbParameters.Visible = True

        Select Case rec
            Case "Weekly"
                lblParameters.Text = "PARAMETERS  (select days)"
                cmbParameters.Visible = False
                clbWeekDays.Visible = True
                lblWeekDaysHint.Visible = True
                UpdateWeekDaysHint()
                ' clbWeekDays (h=110) + hint (~16) replaces combo (h=28): adds 82px extra
                ResizePanelMain(550)
            Case "Quarterly"
                lblParameters.Text = "PARAMETERS  (quarter group)"
                cmbParameters.Items.AddRange(New Object() {
                    "Jan,Apr,Jul,Oct", "Feb,May,Aug,Nov", "Mar,Jun,Sep,Dec"})
                cmbParameters.Enabled = True
                cmbParameters.SelectedIndex = 0
                ResizePanelMain(468)
            Case "SemiAnnually"
                lblParameters.Text = "PARAMETERS  (semi-annual pair)"
                cmbParameters.Items.AddRange(New Object() {
                    "Jan,Jul", "Feb,Aug", "Mar,Sep", "Apr,Oct", "May,Nov", "Jun,Dec"})
                cmbParameters.Enabled = True
                cmbParameters.SelectedIndex = 0
                ResizePanelMain(468)
            Case "Annually"
                lblParameters.Text = "PARAMETERS  (month)"
                cmbParameters.Items.AddRange(New Object() {
                    "Jan", "Feb", "Mar", "Apr", "May", "Jun",
                    "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"})
                cmbParameters.Enabled = True
                cmbParameters.SelectedIndex = 0
                ResizePanelMain(468)
            Case Else
                lblParameters.Text = "PARAMETERS  (not applicable)"
                cmbParameters.Enabled = False
                ResizePanelMain(468)
        End Select

        cmbTypeOfDay.Items.Clear()
        cmbTypeOfDayParams.Items.Clear()

        Select Case rec
            Case "Daily"
                cmbTypeOfDay.Items.AddRange(New Object() {"Calendar", "Business"})
                cmbTypeOfDay.SelectedIndex = 0
                cmbTypeOfDay.Enabled = True
                UpdateTypeOfDayParams()
            Case "Monthly", "Quarterly", "SemiAnnually", "Annually"
                cmbTypeOfDay.Items.AddRange(New Object() {"DateOfDay", "LocationOfDay"})
                cmbTypeOfDay.SelectedIndex = 0
                cmbTypeOfDay.Enabled = True
                UpdateTypeOfDayParams()
            Case "Weekly"
                cmbTypeOfDay.Items.Add("N/A")
                cmbTypeOfDay.SelectedIndex = 0
                cmbTypeOfDay.Enabled = False
                cmbTypeOfDayParams.Items.Add("N/A")
                cmbTypeOfDayParams.SelectedIndex = 0
                cmbTypeOfDayParams.Enabled = False
                cmbTypeOfDayParams.Visible = True
                cmbLocBase.Visible = False
                cmbLocSign.Visible = False
                nudLocOffset.Visible = False
                lblLocFormula.Visible = False
                lblTypeOfDayParams.Text = "TYPE OF DAY PARAMETERS"
        End Select
    End Sub

    Private Sub UpdateTypeOfDayParams()
        If _initialising Then Exit Sub
        Dim rec As String = "Daily"
        If cmbRecurrence IsNot Nothing AndAlso cmbRecurrence.SelectedItem IsNot Nothing Then
            rec = cmbRecurrence.SelectedItem.ToString()
        End If

        Dim tod As String = ""
        If cmbTypeOfDay IsNot Nothing AndAlso cmbTypeOfDay.SelectedItem IsNot Nothing Then
            tod = cmbTypeOfDay.SelectedItem.ToString()
        End If

        ' Hide three-part location controls by default
        cmbLocBase.Visible = False
        cmbLocSign.Visible = False
        nudLocOffset.Visible = False
        lblLocFormula.Visible = False

        cmbTypeOfDayParams.Items.Clear()

        Select Case rec
            Case "Daily"
                If tod = "Business" Then
                    cmbTypeOfDayParams.Items.AddRange(New Object() {"DO NOT COUNT SAT", "COUNT SAT"})
                    cmbTypeOfDayParams.SelectedIndex = 0
                    cmbTypeOfDayParams.Enabled = True
                    cmbTypeOfDayParams.Visible = True
                    lblTypeOfDayParams.Text = "WEEKEND RULE"
                Else
                    cmbTypeOfDayParams.Items.Add("N/A")
                    cmbTypeOfDayParams.SelectedIndex = 0
                    cmbTypeOfDayParams.Enabled = False
                    cmbTypeOfDayParams.Visible = True
                    lblTypeOfDayParams.Text = "TYPE OF DAY PARAMETERS"
                End If

            Case "Monthly", "Quarterly", "SemiAnnually", "Annually"
                If tod = "DateOfDay" Then
                    Dim i As Integer
                    For i = 1 To 31
                        cmbTypeOfDayParams.Items.Add(i.ToString())
                    Next i
                    cmbTypeOfDayParams.SelectedIndex = 0
                    cmbTypeOfDayParams.Enabled = True
                    cmbTypeOfDayParams.Visible = True
                    lblTypeOfDayParams.Text = "DAY NUMBER (1-31)"

                ElseIf tod = "LocationOfDay" Then
                    ' Hide the old single combo, show the three-part controls
                    cmbTypeOfDayParams.Visible = False
                    cmbTypeOfDayParams.Enabled = False
                    lblTypeOfDayParams.Text = "LOCATION  (Base / Sign / Offset)"
                    cmbLocBase.Visible = True
                    cmbLocSign.Visible = True
                    nudLocOffset.Visible = True
                    lblLocFormula.Visible = True
                    UpdateLocFormula()
                End If
        End Select
    End Sub

    ' Build the live preview label whenever any location part changes
    Private Sub UpdateLocFormula()
        If _initialising Then Exit Sub
        If lblLocFormula Is Nothing Then Exit Sub
        Dim base As String = "BOM"
        If cmbLocBase IsNot Nothing AndAlso cmbLocBase.SelectedItem IsNot Nothing Then
            base = cmbLocBase.SelectedItem.ToString()
        End If
        Dim sign As String = "+"
        If cmbLocSign IsNot Nothing AndAlso cmbLocSign.SelectedItem IsNot Nothing Then
            sign = cmbLocSign.SelectedItem.ToString()
        End If
        Dim offset As Integer = 0
        If nudLocOffset IsNot Nothing Then
            offset = CInt(nudLocOffset.Value)
        End If
        lblLocFormula.Text = base & sign & offset.ToString()
    End Sub

    ' Build the "#"-delimited day string from checked items and show it as a hint
    Private Sub UpdateWeekDaysHint()
        If _initialising Then Exit Sub
        If clbWeekDays Is Nothing OrElse lblWeekDaysHint Is Nothing Then Exit Sub
        Dim parts As New List(Of String)
        Dim i As Integer
        For i = 0 To clbWeekDays.Items.Count - 1
            If clbWeekDays.GetItemChecked(i) Then
                parts.Add(clbWeekDays.Items(i).ToString())
            End If
        Next i
        If parts.Count > 0 Then
            lblWeekDaysHint.Text = "Selected: " & String.Join("#", parts.ToArray())
            lblWeekDaysHint.ForeColor = clrAccent
        Else
            lblWeekDaysHint.Text = "No days selected – please check at least one day."
            lblWeekDaysHint.ForeColor = Color.FromArgb(248, 113, 113)   ' red warning
        End If
    End Sub

    ' ItemCheck fires BEFORE the check state changes, so we post-invoke to get the new state
    Private Sub clbWeekDays_ItemCheck(ByVal sender As Object, ByVal e As ItemCheckEventArgs)
        Me.BeginInvoke(New Action(AddressOf UpdateWeekDaysHint))
    End Sub

    ' -------------------------------------------------------------------------
    ' EVENT HANDLERS
    ' -------------------------------------------------------------------------
    Private Sub cmbRecurrence_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        If _initialising Then Exit Sub
        UpdateDynamicControls()
        pnlResult.Visible = False
    End Sub

    Private Sub cmbTypeOfDay_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        If _initialising Then Exit Sub
        UpdateTypeOfDayParams()
        pnlResult.Visible = False
    End Sub

    Private Sub cmbLocBase_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        UpdateLocFormula()
    End Sub

    Private Sub cmbLocSign_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs)
        UpdateLocFormula()
    End Sub

    Private Sub nudLocOffset_ValueChanged(ByVal sender As Object, ByVal e As EventArgs)
        UpdateLocFormula()
    End Sub

    Private Sub btnCalculate_Click(ByVal sender As Object, ByVal e As EventArgs)
        Try
            Dim rec As String = cmbRecurrence.SelectedItem.ToString()

            Dim todVal As String = "Calendar"
            If cmbTypeOfDay.Enabled AndAlso cmbTypeOfDay.SelectedItem IsNot Nothing AndAlso
               cmbTypeOfDay.SelectedItem.ToString() <> "N/A" Then
                todVal = cmbTypeOfDay.SelectedItem.ToString()
            End If

            ' Build TypeOfDayParameters
            Dim todParamVal As String = ""
            If todVal = "LocationOfDay" Then
                Dim locBase As String = "BOM"
                If cmbLocBase.SelectedItem IsNot Nothing Then locBase = cmbLocBase.SelectedItem.ToString()
                Dim locSign As String = "+"
                If cmbLocSign.SelectedItem IsNot Nothing Then locSign = cmbLocSign.SelectedItem.ToString()
                Dim locOffset As Integer = CInt(nudLocOffset.Value)
                todParamVal = locBase & locSign & locOffset.ToString()
            Else
                If cmbTypeOfDayParams.Enabled AndAlso cmbTypeOfDayParams.SelectedItem IsNot Nothing AndAlso
                   cmbTypeOfDayParams.SelectedItem.ToString() <> "N/A" Then
                    todParamVal = cmbTypeOfDayParams.SelectedItem.ToString()
                End If
            End If

            ' Build Parameters
            Dim paramVal As String = ""
            If rec = "Weekly" Then
                Dim parts As New List(Of String)
                Dim i As Integer
                For i = 0 To clbWeekDays.Items.Count - 1
                    If clbWeekDays.GetItemChecked(i) Then
                        parts.Add(clbWeekDays.Items(i).ToString())
                    End If
                Next i
                If parts.Count = 0 Then
                    MessageBox.Show("Please select at least one day for Weekly recurrence.",
                                    "No Days Selected", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    Exit Sub
                End If
                paramVal = String.Join("#", parts.ToArray())
            Else
                If cmbParameters.Enabled AndAlso cmbParameters.SelectedItem IsNot Nothing Then
                    paramVal = cmbParameters.SelectedItem.ToString()
                End If
            End If

            Dim p As New PackageScheduleParams
            p.Recurrence = rec
            p.PackageId = txtPackageId.Text.Trim()
            p.Alteration = CInt(nudAlteration.Value)
            p.Parameters = paramVal
            p.TypeOfDay = todVal
            p.TypeOfDayParameters = todParamVal

            If cmbFriAdj.SelectedItem IsNot Nothing Then
                p.FriAdjustment = cmbFriAdj.SelectedItem.ToString()
            End If
            If cmbSatAdj.SelectedItem IsNot Nothing Then
                p.SatAdjustment = cmbSatAdj.SelectedItem.ToString()
            End If
            If cmbHolAdj.SelectedItem IsNot Nothing Then
                p.HolidaysAdjustment = cmbHolAdj.SelectedItem.ToString()
            End If

            p.Duration = CInt(nudDuration.Value)

            If cmbEndFriAdj.SelectedItem IsNot Nothing Then
                p.EndFriAdjustment = cmbEndFriAdj.SelectedItem.ToString()
            End If
            If cmbEndSatAdj.SelectedItem IsNot Nothing Then
                p.EndSatAdjustment = cmbEndSatAdj.SelectedItem.ToString()
            End If
            If cmbEndHolAdj.SelectedItem IsNot Nothing Then
                p.EndHolidaysAdjustment = cmbEndHolAdj.SelectedItem.ToString()
            End If

            Dim result As NextOccurrenceResult = PackageScheduler.GetNextOccurrence(dtpFromDate.Value.Date, p)

            If result.NextDate = Date.MinValue Then
                lblResultDate.Text = "EVENT CANCELLED"
                lblResultDate.ForeColor = Color.FromArgb(248, 113, 113)
            Else
                lblResultDate.Text = result.NextDate.ToString("dd MMM yyyy") &
                                     "   (" & result.NextDate.DayOfWeek.ToString() & ")"
                lblResultDate.ForeColor = Color.FromArgb(52, 211, 153)
            End If

            If p.Duration > 0 Then
                lblEndDateCaption.Visible = True
                lblEndDate.Visible = True
                If result.EndDate Is Nothing Then
                    lblEndDate.Text = "END DATE CANCELLED"
                    lblEndDate.ForeColor = Color.FromArgb(248, 113, 113)
                Else
                    lblEndDate.Text = result.EndDate.Value.ToString("dd MMM yyyy") &
                                      "   (" & result.EndDate.Value.DayOfWeek.ToString() & ")"
                    lblEndDate.ForeColor = Color.FromArgb(251, 191, 36)
                End If
                pnlResult.Height = 170
            Else
                lblEndDateCaption.Visible = False
                lblEndDate.Visible = False
                pnlResult.Height = 100
            End If

            lblExplanationText.Text = result.Explanation
            pnlResult.Visible = True

        Catch ex As Exception
            MessageBox.Show("ERROR: " & ex.Message & vbCrLf & ex.StackTrace,
                            "Calculation Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' -------------------------------------------------------------------------
    ' RESIZE HELPER – shrinks/grows pnlMain and slides all controls below it
    ' -------------------------------------------------------------------------
    Private Sub ResizePanelMain(ByVal newHeight As Integer)
        ' Find pnlMain – it is the first Panel child of the form after the header
        Dim pnlMain As Panel = Nothing
        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is Panel AndAlso ctrl.Top = 84 Then
                pnlMain = DirectCast(ctrl, Panel)
                Exit For
            End If
        Next

        If pnlMain Is Nothing Then Exit Sub
        If pnlMain.Height = newHeight Then Exit Sub

        Dim delta As Integer = newHeight - pnlMain.Height
        pnlMain.Height = newHeight

        ' Slide every non-header control that sits below pnlMain
        For Each ctrl As Control In Me.Controls
            If ctrl.Top > pnlMain.Top Then
                ctrl.Top += delta
            End If
        Next
        Me.Height += delta
    End Sub

    ' -------------------------------------------------------------------------
    ' PAINT HANDLERS
    ' -------------------------------------------------------------------------
    Private Sub PaintBorderedPanel(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        Using pen As New Pen(clrBorder, 1)
            e.Graphics.DrawRectangle(pen, 0, 0, pnl.Width - 1, pnl.Height - 1)
        End Using
    End Sub

    Private Sub PaintHeaderBorder(ByVal sender As Object, ByVal e As PaintEventArgs)
        Dim pnl As Panel = DirectCast(sender, Panel)
        Using pen As New Pen(clrAccent, 2)
            e.Graphics.DrawLine(pen, 0, pnl.Height - 1, pnl.Width, pnl.Height - 1)
        End Using
    End Sub

    ' -------------------------------------------------------------------------
    ' FACTORY HELPERS
    ' -------------------------------------------------------------------------
    Private Function MakeSectionLabel(ByVal text As String, ByVal x As Integer, ByVal y As Integer) As Label
        Dim lbl As New Label
        lbl.Text = text
        lbl.Font = New Font("Segoe UI", 7.5F, FontStyle.Bold)
        lbl.ForeColor = clrSubtext
        lbl.AutoSize = True
        lbl.Location = New Point(x, y)
        Return lbl
    End Function

    Private Function MakeCombo(ByVal x As Integer, ByVal y As Integer, ByVal w As Integer) As ComboBox
        Dim c As New ComboBox
        c.Bounds = New Rectangle(x, y, w, 28)
        c.BackColor = clrInput
        c.ForeColor = clrText
        c.FlatStyle = FlatStyle.Flat
        c.DropDownStyle = ComboBoxStyle.DropDownList
        c.Font = New Font("Segoe UI", 9)
        Return c
    End Function

End Class

' =============================================================================
' ENTRY POINT
' =============================================================================
Module Program
    <STAThread>
    Sub Main()
        Application.EnableVisualStyles()
        Application.SetCompatibleTextRenderingDefault(False)
        Application.Run(New SchedulerForm())
    End Sub
End Module

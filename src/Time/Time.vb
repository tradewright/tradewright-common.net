#Region "License"

' The MIT License (MIT)
'
' Copyright (c) 2017 Richard L King (TradeWright Software Systems)
' 
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
' 
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
' 
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

#End Region


''' <summary>
''' Timestamp format specifiers for use with the <c>FormatTimestamp</c>
''' global method.
''' </summary>
Public Enum TimestampFormats
    ''' <summary>
    ''' The formatted Timestamp is of the form <em>hhmmss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    TimeOnly = 0

    ''' <summary>
    '''  The formatted Timestamp is of the form <em>yyyymmdd</em>.
    ''' </summary>
    DateOnly = 1

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyymmddhhmmss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    DateAndTime = 2

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>hh:mm:ss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    TimeOnlyISO8601 = 4

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyy-mm-dd</em>.
    ''' </summary>
    DateOnlyISO8601 = 5

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>yyyy-mm-dd hh:mm:ss.lll</em> where <em>lll</em>
    ''' is milliseconds.
    ''' </summary>
    DateAndTimeISO8601 = 6

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;time&gt;.lll</em>, where &lt;time&gt; is
    ''' the time in the format defined in the Control Panel's Regional and Language Options,
    ''' and <em>lll</em> is milliseconds.
    ''' </summary>
    TimeOnlyLocal = 7

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;date&gt;.lll</em>, where &lt;date&gt; is
    ''' the date in the format defined in the Control Panel's Regional and Language Options,
    ''' and <em>lll</em> is milliseconds.
    ''' </summary>
    DateOnlyLocal = 8

    ''' <summary>
    ''' The formatted Timestamp is of the form <em>&lt;date &amp; time&gt;.lll</em>, where
    ''' &lt;date &amp; time&gt; is the date and time in the format defined in the Control Panel's
    ''' Regional and Language Options, and <em>lll</em> is milliseconds.
    ''' </summary>
    DateAndTimeLocal = 9

    ''' <summary>
    ''' Used additively in conjunction with other values, prevents the milliseconds
    ''' value being included in the Timestamp.
    ''' </summary>
    NoMillisecs = &H40000000

    ''' <summary>
    ''' Used additively in conjunction with other values, results in a timezone specifier
    ''' being appended. The format of the timezone specifier is +|-hh:mm.
    ''' </summary>
    IncludeTimezone = &H80000000
End Enum

Public Module Time

    ''' <summary>
    ''' Returns the number of the week in the year containing the supplied date. Week 1
    ''' is the first complete week in the year.
    ''' </summary>
    ''' <param name="pDate">The date.</param>
    ''' <returns>The <c>Integer/c> number of the week containing the supplied date.</returns>
    Public Function GetWeekNumber(pDate As Date) As Integer
        Return CInt(Math.Floor((pDate.DayOfYear - DayOfWeekMon(pDate) + 7) / 7))
    End Function

    ''' <summary>
    ''' Returns the date of the first day of the week containing the supplied date. Note that weeks are taken to start on Monday.
    ''' </summary>
    ''' <param name="pDate">The date.</param>
    ''' <returns>A <c>Date</c> representing the first day of the week containing the supplied date.</returns>
    Public Function GetWeekStartDate(
                pDate As Date) As Date
        Return pDate.AddDays(-DayOfWeekMon(pDate) + 1)
    End Function

    ''' <summary>
    ''' Returns the date of the first day of the specified week in the year containing the supplied date.
    ''' </summary>
    ''' <param name="weekNumber">The week number. The value supplied need not be restricted to the range 1-52.</param>
    ''' <param name="baseDate">The date that identifies the relevant year.</param>
    ''' <returns>A <c>Date</c> representing the first day of the specified week in the year containing the supplied date.</returns>
    Public Function GetWeekStartDateFromWeekNumber(
                weekNumber As Integer,
                baseDate As Date) As Date
        Dim yearStart = baseDate.AddDays(1 - baseDate.DayOfYear)
        Dim dow1 = DayOfWeekMon(yearStart) ' day of week of 1st Jan of base year

        Dim week1Date As Date
        If dow1 = 1 Then
            week1Date = yearStart
        Else
            week1Date = yearStart.AddDays(8 - dow1)
        End If

        Return week1Date.AddDays(7 * (weekNumber - 1))
    End Function

    ''' <summary>
    ''' Returns the date of the specified working day number in the year containing the supplied date.
    ''' Working days are considered to be Mondays to Fridays, and holidays are not taken into account.
    ''' </summary>
    ''' <param name="dayNumber">The working day number.</param>
    ''' <param name="baseDate">The date that identifies the relevant year.</param>
    ''' <returns>A <c>Date</c> representing the specified working day in the year containing the supplied date.</returns>
    Public Function GetWorkingDayDate(
                dayNumber As Integer,
                baseDate As Date) As Date

        Dim yearStart = baseDate.AddDays(1 - baseDate.DayOfYear)

        Do While dayNumber < 0
            Dim yearEnd = yearStart.AddDays(-1)
            yearStart = yearStart.AddYears(1)
            dayNumber = dayNumber + GetWorkingDayNumber(yearEnd) + 1
        Loop

        Dim dow1 = DayOfWeekMon(yearStart)  ' day of week of 1st Jan of base year

        Dim wd1 As Integer     ' weekdays in first week (excluding weekend)
        Dim we1 As Integer     ' weekend days at start of first week
        If dow1 = 7 Then
            ' Sunday
            wd1 = 0
            we1 = 1
        ElseIf dow1 = 6 Then
            ' Saturday
            wd1 = 0
            we1 = 2
        Else
            wd1 = 5 - dow1 + 1
            we1 = 2
        End If

        Dim doy As Integer
        If dayNumber <= wd1 Then
            doy = dayNumber
        ElseIf dayNumber - wd1 <= 5 Then
            doy = we1 + dayNumber
        Else
            ' number of whole weeks after the first week
            Dim numWholeWeeks = CInt(Math.Floor((dayNumber - wd1) / 5)) - 1
            doy = wd1 + we1 + If(numWholeWeeks > 0, 7 * numWholeWeeks + 5, 5) + If(((dayNumber - wd1) Mod 5) > 0, ((dayNumber - wd1) Mod 5) + 2, 0)
        End If

        Return yearStart.AddDays(doy - 1)
    End Function

    ''' <summary>
    ''' Returns the number of working days that have elapsed in the year up to and including the supplied date.
    ''' Working days are considered to be Mondays to Fridays, and holidays are not taken into account.
    ''' </summary>
    ''' <param name="pDate">The date.</param>
    ''' <returns>The number of working days in the year up to and including the supplied date.</returns>
    Public Function GetWorkingDayNumber(
                pDate As Date) As Integer

        Dim doy = pDate.DayOfYear   ' day of year
        Dim woy = GetWeekNumber(pDate)  ' week of year
        Dim dow = DayOfWeekMon(pDate)   ' day of week of supplied date
        Dim dow1 = DayOfWeekMon(pDate.AddDays(1 - doy)) ' day of week of 1st Jan

        Dim wd1 As Integer     ' weekdays in first week (excluding weekend)
        If dow1 = 7 Then
            ' Sunday
            wd1 = 0
        ElseIf dow1 = 6 Then
            ' Saturday
            wd1 = 0
        Else
            wd1 = 5 - dow1 + 1
        End If

        Dim wdN As Integer     ' weekdays in last week
        If dow = 7 Or dow = 6 Then
            wdN = 5
        Else
            wdN = dow
        End If

        Return wd1 + 5 * (woy - 2) + wdN
    End Function

    ''' <summary>
    ''' Formats the DateTime supplied in <paramref name="timestamp"/> according to the value specified in <paramref name="formatOption"/>.
    ''' </summary>
    ''' <param name="timestamp"></param>
    ''' <param name="formatOption"></param>
    ''' <remarks>This is a convenience method that removes the need to remember or lookup the 
    ''' various formatting options in the DateTime.ToString() method.</remarks>
    ''' <returns></returns>
    Public Function FormatTimestamp(
                        timestamp As DateTime,
                        Optional formatOption As TimestampFormats = TimestampFormats.DateAndTime) As String

        Dim includeTimezone = (formatOption And TimestampFormats.IncludeTimezone) <> 0
        formatOption = formatOption And (Not TimestampFormats.IncludeTimezone)

        Dim noMillisecs = CBool(formatOption And TimestampFormats.NoMillisecs)
        formatOption = formatOption And (Not TimestampFormats.NoMillisecs)

        Dim milliseconds = timestamp.Millisecond

        Select Case formatOption
            Case TimestampFormats.TimeOnly
                If noMillisecs Then Return timestamp.ToString("HHmmss")
                Return timestamp.ToString("HHmmss.fff")
            Case TimestampFormats.DateOnly
                Return timestamp.ToString("yyyyMMdd")
            Case TimestampFormats.DateAndTime
                If noMillisecs Then Return timestamp.ToString("yyyyMMddHHmmss")
                Return timestamp.ToString("yyyyMMddHHmmss.fff")
            Case TimestampFormats.TimeOnlyISO8601
                If noMillisecs Then Return timestamp.ToString("HH:mm:ss")
                Return timestamp.ToString("HH:mm:ss.fff")
            Case TimestampFormats.DateOnlyISO8601
                Return timestamp.ToString("yyyy-MM-dd")
            Case TimestampFormats.DateAndTimeISO8601
                If noMillisecs Then Return timestamp.ToString("yyyy-MM-dd HH:mm:ss")
                Return timestamp.ToString("yyyy-MM-dd HH:mm:ss.fff")
            Case TimestampFormats.TimeOnlyLocal
                Return timestamp.ToString("D")
            Case TimestampFormats.DateOnlyLocal
                Return timestamp.ToString("d")
            Case TimestampFormats.DateAndTimeLocal
                Return timestamp.ToString("d") & " " & timestamp.ToString("D")
            Case Else
                Throw New ArgumentException("Invalid format option")
        End Select

        If includeTimezone Then Throw New NotImplementedException
    End Function

    Private Function DayOfWeekMon([date] As Date) As Integer
        Return If([date].DayOfWeek = DayOfWeek.Sunday, 7, [date].DayOfWeek)
    End Function

End Module



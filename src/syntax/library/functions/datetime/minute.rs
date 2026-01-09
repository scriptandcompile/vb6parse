//! # Minute Function
//!
//! Returns a Variant (Integer) specifying a whole number between 0 and 59, inclusive, representing the minute of the hour.
//!
//! ## Syntax
//!
//! ```vb
//! Minute(time)
//! ```
//!
//! ## Parameters
//!
//! - `time` (Required): Any Variant, numeric expression, string expression, or combination that can represent a time
//!   - Can be a Date literal, Date variable, or numeric expression
//!   - String must be recognizable as a date/time
//!   - If Null, returns Null
//!   - If invalid date/time, error 13 (Type mismatch)
//!
//! ## Return Value
//!
//! Returns a Variant (Integer):
//! - Whole number from 0 to 59
//! - Represents the minute component of the time
//! - 0 = first minute of the hour (00:00-00:00:59)
//! - 59 = last minute of the hour (XX:59:00-XX:59:59)
//! - Returns Null if input is Null
//! - Independent of date component (only extracts minute)
//!
//! ## Remarks
//!
//! The Minute function extracts the minute from a time value:
//!
//! - **Returns Integer**: Value is always between 0 and 59
//! - **Time component only**: Ignores date portion of Date values
//! - **Null propagation**: Returns Null if input is Null
//! - **Type mismatch**: Error 13 if input cannot be converted to date/time
//! - **Various formats**: Accepts Date, String, or numeric time values
//! - **24-hour time**: Works with both 12-hour and 24-hour formats
//! - **Common use**: Extract minute for time calculations, formatting, validation
//! - **Related functions**: Hour (hour component), Second (second component), `TimeSerial` (create time)
//! - **Part of suite**: Day, Month, Year for dates; Hour, Minute, Second for times
//! - **Performance**: Fast operation, optimized in VB6
//! - **Available in**: All VB versions, VBA, `VBScript`
//!
//! ## Typical Uses
//!
//! 1. **Extract Minute from Time**
//!    ```vb
//!    currentMinute = Minute(Now)
//!    ```
//!
//! 2. **Format Time Display**
//!    ```vb
//!    timeText = Hour(Now) & ":" & Format(Minute(Now), "00")
//!    ```
//!
//! 3. **Validate Time Range**
//!    ```vb
//!    If Minute(appointmentTime) < 30 Then
//!        ' First half of the hour
//!    End If
//!    ```
//!
//! 4. **Round to Nearest Hour**
//!    ```vb
//!    If Minute(timeValue) >= 30 Then
//!        roundedHour = Hour(timeValue) + 1
//!    Else
//!        roundedHour = Hour(timeValue)
//!    End If
//!    ```
//!
//! 5. **Time Calculations**
//!    ```vb
//!    minutesPastHour = Minute(Time)
//!    minutesUntilHour = 60 - Minute(Time)
//!    ```
//!
//! 6. **Validate Appointment Times**
//!    ```vb
//!    If Minute(startTime) Mod 15 <> 0 Then
//!        MsgBox "Appointments must start on 15-minute intervals"
//!    End If
//!    ```
//!
//! 7. **Build Time String**
//!    ```vb
//!    timeStr = Format(Hour(t), "00") & ":" & Format(Minute(t), "00")
//!    ```
//!
//! 8. **Calculate Duration**
//!    ```vb
//!    durationMinutes = (Hour(endTime) - Hour(startTime)) * 60 + _
//!                      (Minute(endTime) - Minute(startTime))
//!    ```
//!
//! ## Basic Examples
//!
//! ### Example 1: Basic Usage
//! ```vb
//! Dim currentMinute As Integer
//! Dim timeValue As Date
//!
//! ' Get current minute
//! currentMinute = Minute(Now)        ' Returns 0-59
//!
//! ' Extract minute from specific time
//! timeValue = #2:45:30 PM#
//! currentMinute = Minute(timeValue)  ' Returns 45
//!
//! ' Extract minute from time string
//! currentMinute = Minute("3:27 PM")  ' Returns 27
//!
//! ' Extract minute from numeric value
//! currentMinute = Minute(0.5)        ' Returns 0 (noon = 12:00)
//! currentMinute = Minute(0.75)       ' Returns 0 (6 PM = 18:00)
//! ```
//!
//! ### Example 2: Digital Clock Display
//! ```vb
//! Sub UpdateClock()
//!     Dim h As Integer
//!     Dim m As Integer
//!     Dim s As Integer
//!     Dim currentTime As Date
//!     
//!     currentTime = Now
//!     
//!     h = Hour(currentTime)
//!     m = Minute(currentTime)
//!     s = Second(currentTime)
//!     
//!     ' Display in 24-hour format
//!     lblClock.Caption = Format(h, "00") & ":" & _
//!                       Format(m, "00") & ":" & _
//!                       Format(s, "00")
//!     
//!     ' Display in 12-hour format
//!     Dim ampm As String
//!     Dim h12 As Integer
//!     
//!     If h >= 12 Then
//!         ampm = "PM"
//!         h12 = IIf(h = 12, 12, h - 12)
//!     Else
//!         ampm = "AM"
//!         h12 = IIf(h = 0, 12, h)
//!     End If
//!     
//!     lblClock12.Caption = h12 & ":" & Format(m, "00") & " " & ampm
//! End Sub
//! ```
//!
//! ### Example 3: Appointment Scheduler Validation
//! ```vb
//! Function ValidateAppointmentTime(ByVal appointmentTime As Date) As Boolean
//!     Dim m As Integer
//!     
//!     m = Minute(appointmentTime)
//!     
//!     ' Check if time is on a 15-minute interval
//!     If m Mod 15 = 0 Then
//!         ValidateAppointmentTime = True
//!     Else
//!         MsgBox "Appointments must start on 15-minute intervals" & vbCrLf & _
//!                "Valid minutes: 00, 15, 30, 45", _
//!                vbExclamation, "Invalid Time"
//!         ValidateAppointmentTime = False
//!     End If
//! End Function
//!
//! ' Usage:
//! ' If ValidateAppointmentTime(#2:15 PM#) Then  ' Valid
//! ' If ValidateAppointmentTime(#2:23 PM#) Then  ' Invalid
//! ```
//!
//! ### Example 4: Time Rounding
//! ```vb
//! Function RoundToNearestQuarterHour(ByVal timeValue As Date) As Date
//!     Dim h As Integer
//!     Dim m As Integer
//!     Dim roundedMinute As Integer
//!     
//!     h = Hour(timeValue)
//!     m = Minute(timeValue)
//!     
//!     ' Round to nearest 15 minutes
//!     Select Case m
//!         Case 0 To 7
//!             roundedMinute = 0
//!         Case 8 To 22
//!             roundedMinute = 15
//!         Case 23 To 37
//!             roundedMinute = 30
//!         Case 38 To 52
//!             roundedMinute = 45
//!         Case 53 To 59
//!             roundedMinute = 0
//!             h = h + 1
//!             If h = 24 Then h = 0
//!     End Select
//!     
//!     RoundToNearestQuarterHour = TimeSerial(h, roundedMinute, 0)
//! End Function
//!
//! ' Usage:
//! ' rounded = RoundToNearestQuarterHour(#2:23 PM#)  ' Returns #2:30 PM#
//! ' rounded = RoundToNearestQuarterHour(#2:57 PM#)  ' Returns #3:00 PM#
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `SafeMinute` (handle Null)
//! ```vb
//! Function SafeMinute(ByVal timeValue As Variant) As Integer
//!     If IsNull(timeValue) Then
//!         SafeMinute = 0
//!     ElseIf Not IsDate(timeValue) Then
//!         SafeMinute = 0
//!     Else
//!         SafeMinute = Minute(timeValue)
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 2: `IsTopOfHour`
//! ```vb
//! Function IsTopOfHour(ByVal timeValue As Date) As Boolean
//!     IsTopOfHour = (Minute(timeValue) = 0 And Second(timeValue) = 0)
//! End Function
//! ```
//!
//! ### Pattern 3: `GetMinutesPastHour`
//! ```vb
//! Function GetMinutesPastHour(ByVal timeValue As Date) As Integer
//!     GetMinutesPastHour = Minute(timeValue)
//! End Function
//! ```
//!
//! ### Pattern 4: `GetMinutesUntilNextHour`
//! ```vb
//! Function GetMinutesUntilNextHour(ByVal timeValue As Date) As Integer
//!     Dim m As Integer
//!     m = Minute(timeValue)
//!     
//!     If m = 0 And Second(timeValue) = 0 Then
//!         GetMinutesUntilNextHour = 60
//!     Else
//!         GetMinutesUntilNextHour = 60 - m
//!         If Second(timeValue) > 0 Then
//!             GetMinutesUntilNextHour = GetMinutesUntilNextHour - 1
//!         End If
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 5: `FormatMinute`
//! ```vb
//! Function FormatMinute(ByVal timeValue As Date) As String
//!     FormatMinute = Format(Minute(timeValue), "00")
//! End Function
//! ```
//!
//! ### Pattern 6: `IsQuarterHour`
//! ```vb
//! Function IsQuarterHour(ByVal timeValue As Date) As Boolean
//!     IsQuarterHour = (Minute(timeValue) Mod 15 = 0)
//! End Function
//! ```
//!
//! ### Pattern 7: `GetQuarterHourIndex`
//! ```vb
//! Function GetQuarterHourIndex(ByVal timeValue As Date) As Integer
//!     ' Returns 0-3 for which quarter hour (0=:00, 1=:15, 2=:30, 3=:45)
//!     GetQuarterHourIndex = Minute(timeValue) \ 15
//! End Function
//! ```
//!
//! ### Pattern 8: `IsHalfHour`
//! ```vb
//! Function IsHalfHour(ByVal timeValue As Date) As Boolean
//!     IsHalfHour = (Minute(timeValue) = 0 Or Minute(timeValue) = 30)
//! End Function
//! ```
//!
//! ### Pattern 9: `CompareMinutes`
//! ```vb
//! Function CompareMinutes(ByVal time1 As Date, ByVal time2 As Date) As Integer
//!     CompareMinutes = Minute(time1) - Minute(time2)
//! End Function
//! ```
//!
//! ### Pattern 10: `MinutesSinceMidnight`
//! ```vb
//! Function MinutesSinceMidnight(ByVal timeValue As Date) As Long
//!     MinutesSinceMidnight = Hour(timeValue) * 60 + Minute(timeValue)
//! End Function
//! ```
//!
//! ## Advanced Examples
//!
//! ### Example 1: Time Slot Scheduler
//! ```vb
//! ' Class: TimeSlotScheduler
//! Private m_slotDuration As Integer  ' Minutes per slot
//! Private m_slots As Collection
//!
//! Private Sub Class_Initialize()
//!     m_slotDuration = 30  ' Default 30-minute slots
//!     Set m_slots = New Collection
//! End Sub
//!
//! Public Property Let SlotDuration(ByVal minutes As Integer)
//!     If minutes > 0 And minutes <= 60 And (60 Mod minutes = 0) Then
//!         m_slotDuration = minutes
//!     Else
//!         Err.Raise 5, , "Slot duration must evenly divide 60 (5, 10, 15, 20, 30, or 60)"
//!     End If
//! End Property
//!
//! Public Function GetSlotIndex(ByVal timeValue As Date) As Integer
//!     Dim totalMinutes As Long
//!     totalMinutes = Hour(timeValue) * 60 + Minute(timeValue)
//!     GetSlotIndex = totalMinutes \ m_slotDuration
//! End Function
//!
//! Public Function GetSlotStartTime(ByVal slotIndex As Integer) As Date
//!     Dim totalMinutes As Long
//!     Dim h As Integer
//!     Dim m As Integer
//!     
//!     totalMinutes = slotIndex * m_slotDuration
//!     h = totalMinutes \ 60
//!     m = totalMinutes Mod 60
//!     
//!     GetSlotStartTime = TimeSerial(h, m, 0)
//! End Function
//!
//! Public Function IsSlotAvailable(ByVal slotIndex As Integer) As Boolean
//!     On Error Resume Next
//!     Dim temp As Variant
//!     temp = m_slots(CStr(slotIndex))
//!     IsSlotAvailable = (Err.Number <> 0)
//!     On Error GoTo 0
//! End Function
//!
//! Public Sub BookSlot(ByVal slotIndex As Integer, ByVal description As String)
//!     If IsSlotAvailable(slotIndex) Then
//!         m_slots.Add description, CStr(slotIndex)
//!     Else
//!         Err.Raise 5, , "Slot already booked"
//!     End If
//! End Sub
//!
//! Public Function GetSlotsInHour(ByVal hour As Integer) As Integer
//!     GetSlotsInHour = 60 \ m_slotDuration
//! End Function
//! ```
//!
//! ### Example 2: Timesheet Entry Validator
//! ```vb
//! ' Class: TimesheetValidator
//! Private m_roundingInterval As Integer
//!
//! Private Sub Class_Initialize()
//!     m_roundingInterval = 15  ' Default to 15-minute intervals
//! End Sub
//!
//! Public Property Let RoundingInterval(ByVal minutes As Integer)
//!     Select Case minutes
//!         Case 1, 5, 10, 15, 30
//!             m_roundingInterval = minutes
//!         Case Else
//!             Err.Raise 5, , "Invalid rounding interval"
//!     End Select
//! End Property
//!
//! Public Function ValidateTime(ByVal timeValue As Date) As Boolean
//!     Dim m As Integer
//!     m = Minute(timeValue)
//!     ValidateTime = (m Mod m_roundingInterval = 0)
//! End Function
//!
//! Public Function RoundTime(ByVal timeValue As Date, _
//!                          Optional ByVal roundUp As Boolean = False) As Date
//!     Dim h As Integer
//!     Dim m As Integer
//!     Dim s As Integer
//!     Dim roundedMinute As Integer
//!     
//!     h = Hour(timeValue)
//!     m = Minute(timeValue)
//!     s = Second(timeValue)
//!     
//!     If roundUp Then
//!         ' Round up to next interval
//!         roundedMinute = ((m \ m_roundingInterval) + 1) * m_roundingInterval
//!         If roundedMinute >= 60 Then
//!             roundedMinute = 0
//!             h = h + 1
//!             If h = 24 Then h = 0
//!         End If
//!     Else
//!         ' Round to nearest interval
//!         Dim remainder As Integer
//!         remainder = m Mod m_roundingInterval
//!         
//!         If remainder >= (m_roundingInterval \ 2) Or s > 0 Then
//!             roundedMinute = m - remainder + m_roundingInterval
//!             If roundedMinute >= 60 Then
//!                 roundedMinute = 0
//!                 h = h + 1
//!                 If h = 24 Then h = 0
//!             End If
//!         Else
//!             roundedMinute = m - remainder
//!         End If
//!     End If
//!     
//!     RoundTime = TimeSerial(h, roundedMinute, 0)
//! End Function
//!
//! Public Function GetTimesheetString(ByVal timeValue As Date) As String
//!     GetTimesheetString = Format(Hour(timeValue), "00") & ":" & _
//!                         Format(Minute(timeValue), "00")
//! End Function
//! ```
//!
//! ### Example 3: Meeting Reminder System
//! ```vb
//! ' Module: MeetingReminders
//!
//! Public Function CheckReminders() As Collection
//!     Dim reminders As New Collection
//!     Dim currentTime As Date
//!     Dim currentMinute As Integer
//!     
//!     currentTime = Now
//!     currentMinute = Minute(currentTime)
//!     
//!     ' Check for hourly reminder (at :00)
//!     If currentMinute = 0 Then
//!         reminders.Add "Top of the hour reminder"
//!     End If
//!     
//!     ' Check for quarter-hour reminders
//!     If currentMinute Mod 15 = 0 Then
//!         reminders.Add "Quarter hour: " & Format(currentTime, "h:mm AM/PM")
//!     End If
//!     
//!     ' Check for upcoming meeting (5 minutes before)
//!     If currentMinute Mod 60 = 25 Or currentMinute Mod 60 = 55 Then
//!         reminders.Add "Meeting in 5 minutes"
//!     End If
//!     
//!     Set CheckReminders = reminders
//! End Function
//!
//! Public Function GetNextQuarterHour(ByVal fromTime As Date) As Date
//!     Dim h As Integer
//!     Dim m As Integer
//!     Dim nextMinute As Integer
//!     
//!     h = Hour(fromTime)
//!     m = Minute(fromTime)
//!     
//!     ' Calculate next quarter hour
//!     nextMinute = ((m \ 15) + 1) * 15
//!     
//!     If nextMinute >= 60 Then
//!         nextMinute = 0
//!         h = h + 1
//!         If h = 24 Then h = 0
//!     End If
//!     
//!     GetNextQuarterHour = TimeSerial(h, nextMinute, 0)
//! End Function
//!
//! Public Function MinutesUntilQuarterHour(ByVal fromTime As Date) As Integer
//!     Dim m As Integer
//!     Dim remainder As Integer
//!     
//!     m = Minute(fromTime)
//!     remainder = m Mod 15
//!     
//!     If remainder = 0 Then
//!         MinutesUntilQuarterHour = 15
//!     Else
//!         MinutesUntilQuarterHour = 15 - remainder
//!     End If
//! End Function
//! ```
//!
//! ### Example 4: Bus Schedule Matcher
//! ```vb
//! ' Class: BusSchedule
//! Private m_departureMinutes() As Integer
//! Private m_routeName As String
//!
//! Public Sub Initialize(ByVal routeName As String, departureMinutes As Variant)
//!     Dim i As Long
//!     m_routeName = routeName
//!     
//!     ReDim m_departureMinutes(LBound(departureMinutes) To UBound(departureMinutes))
//!     For i = LBound(departureMinutes) To UBound(departureMinutes)
//!         m_departureMinutes(i) = departureMinutes(i)
//!     Next i
//! End Sub
//!
//! Public Function GetNextDeparture(ByVal currentTime As Date) As Date
//!     Dim currentMinuteOfDay As Long
//!     Dim i As Long
//!     Dim h As Integer
//!     Dim m As Integer
//!     
//!     currentMinuteOfDay = Hour(currentTime) * 60 + Minute(currentTime)
//!     
//!     ' Find next departure
//!     For i = LBound(m_departureMinutes) To UBound(m_departureMinutes)
//!         If m_departureMinutes(i) > currentMinuteOfDay Then
//!             h = m_departureMinutes(i) \ 60
//!             m = m_departureMinutes(i) Mod 60
//!             GetNextDeparture = TimeSerial(h, m, 0)
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     ' No more departures today, return first tomorrow
//!     h = m_departureMinutes(LBound(m_departureMinutes)) \ 60
//!     m = m_departureMinutes(LBound(m_departureMinutes)) Mod 60
//!     GetNextDeparture = DateSerial(Year(currentTime), Month(currentTime), Day(currentTime) + 1) + _
//!                        TimeSerial(h, m, 0)
//! End Function
//!
//! Public Function GetMinutesUntilNext(ByVal currentTime As Date) As Long
//!     Dim nextDeparture As Date
//!     nextDeparture = GetNextDeparture(currentTime)
//!     
//!     GetMinutesUntilNext = DateDiff("n", currentTime, nextDeparture)
//! End Function
//!
//! Public Function IsAtDeparture(ByVal currentTime As Date, _
//!                              Optional ByVal toleranceMinutes As Integer = 0) As Boolean
//!     Dim currentMinuteOfDay As Long
//!     Dim i As Long
//!     
//!     currentMinuteOfDay = Hour(currentTime) * 60 + Minute(currentTime)
//!     
//!     For i = LBound(m_departureMinutes) To UBound(m_departureMinutes)
//!         If Abs(m_departureMinutes(i) - currentMinuteOfDay) <= toleranceMinutes Then
//!             IsAtDeparture = True
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     IsAtDeparture = False
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! ' Error 13: Type mismatch
//! ' - Input cannot be converted to date/time
//!
//! ' Safe minute extraction with error handling
//! Function SafeGetMinute(ByVal timeValue As Variant) As Integer
//!     On Error GoTo ErrorHandler
//!     
//!     If IsNull(timeValue) Then
//!         SafeGetMinute = 0
//!     ElseIf Not IsDate(timeValue) Then
//!         SafeGetMinute = 0
//!     Else
//!         SafeGetMinute = Minute(timeValue)
//!     End If
//!     
//!     Exit Function
//!     
//! ErrorHandler:
//!     SafeGetMinute = 0
//! End Function
//!
//! ' Validate before extracting
//! Function GetMinuteOrDefault(ByVal timeValue As Variant, _
//!                            Optional ByVal defaultValue As Integer = 0) As Integer
//!     If IsNull(timeValue) Then
//!         GetMinuteOrDefault = defaultValue
//!     ElseIf IsDate(timeValue) Then
//!         GetMinuteOrDefault = Minute(timeValue)
//!     Else
//!         GetMinuteOrDefault = defaultValue
//!     End If
//! End Function
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: Extracting time components is highly optimized in VB6
//! - **No String Parsing**: Direct access to internal time representation
//! - **Cache for Repeated Use**: If calling multiple times on same value
//! - **Combine Extractions**: Use Hour, Minute, Second together efficiently
//!
//! ## Best Practices
//!
//! 1. **Use with Format** - Pad with leading zero: `Format(Minute(t), "00")`
//! 2. **Validate input** - Check `IsDate` before calling Minute
//! 3. **Handle Null gracefully** - Use `IsNull` check for Variant inputs
//! 4. **Combine with Hour/Second** - Extract all time components together
//! 5. **Use for validation** - Check minute intervals for appointments
//! 6. **Consider `TimeSerial`** - To reconstruct time from components
//! 7. **Document assumptions** - Clarify if using 0-59 or 1-60 range
//! 8. **Use constants** - Define meaningful minute values (`TOP_OF_HOUR` = 0)
//! 9. **Test edge cases** - Null values, invalid strings, midnight
//! 10. **Remember range** - Always 0-59, never 60 or negative
//!
//! ## Comparison with Related Functions
//!
//! | Function | Returns | Range | Use Case |
//! |----------|---------|-------|----------|
//! | **Minute** | Minute of hour | 0-59 | Extract minute component |
//! | **Hour** | Hour of day | 0-23 | Extract hour component |
//! | **Second** | Second of minute | 0-59 | Extract second component |
//! | **`TimeSerial`** | Date/Time | N/A | Create time from components |
//!
//! ## Minute vs Hour vs Second
//!
//! ```vb
//! Dim timeValue As Date
//! timeValue = #2:45:30 PM#
//!
//! ' Extract individual components
//! Debug.Print Hour(timeValue)    ' 14 (2 PM in 24-hour format)
//! Debug.Print Minute(timeValue)  ' 45
//! Debug.Print Second(timeValue)  ' 30
//!
//! ' Reconstruct time
//! Dim reconstructed As Date
//! reconstructed = TimeSerial(Hour(timeValue), Minute(timeValue), Second(timeValue))
//! ' reconstructed = #2:45:30 PM#
//!
//! ' Format as string
//! Debug.Print Format(Hour(timeValue), "00") & ":" & _
//!            Format(Minute(timeValue), "00") & ":" & _
//!            Format(Second(timeValue), "00")
//! ' Output: "14:45:30"
//! ```
//!
//! ## Minute Range (0-59)
//!
//! ```vb
//! ' Minute function always returns 0-59
//! Debug.Print Minute(#12:00:00 AM#)  ' 0 (midnight)
//! Debug.Print Minute(#12:15:00 AM#)  ' 15
//! Debug.Print Minute(#12:30:00 AM#)  ' 30
//! Debug.Print Minute(#12:45:00 AM#)  ' 45
//! Debug.Print Minute(#12:59:59 PM#)  ' 59 (last minute of noon hour)
//! Debug.Print Minute(#11:59:59 PM#)  ' 59 (last minute before midnight)
//!
//! ' Common minute values
//! Const TOP_OF_HOUR As Integer = 0
//! Const QUARTER_PAST As Integer = 15
//! Const HALF_PAST As Integer = 30
//! Const QUARTER_TO As Integer = 45
//! ```
//!
//! ## Platform Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA core library
//! - Available in `VBScript`
//! - Returns Integer (not Long)
//! - Range is always 0-59 (inclusive)
//! - Handles Null by returning Null
//! - Type mismatch error for invalid input
//! - Same behavior across all Windows versions
//! - Works with Date variables and literals
//! - Ignores date component of Date values
//!
//! ## Limitations
//!
//! - **No milliseconds**: Does not access millisecond component
//! - **Integer only**: Returns whole minutes, not fractional
//! - **No timezone**: Does not handle timezone information
//! - **Type mismatch**: Error 13 for non-date inputs
//! - **Null propagation**: Returns Null for Null input (may need handling)
//!
//! ## Related Functions
//!
//! - `Hour`: Returns hour component (0-23)
//! - `Second`: Returns second component (0-59)
//! - `Day`: Returns day of month (1-31)
//! - `Month`: Returns month (1-12)
//! - `Year`: Returns year
//! - `Now`: Returns current date and time
//! - `Time`: Returns current time
//! - `TimeSerial`: Creates time from hour, minute, second
//! - `TimeValue`: Converts string to time
//! - `Format`: Formats date/time as string

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn minute_basic() {
        let source = r"
            currentMinute = Minute(Now)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_variable() {
        let source = r"
            m = Minute(timeValue)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_time_literal() {
        let source = r"
            m = Minute(#2:45:30 PM#)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_if_statement() {
        let source = r#"
            If Minute(appointmentTime) < 30 Then
                MsgBox "First half"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_function_return() {
        let source = r"
            Function GetMinute() As Integer
                GetMinute = Minute(Time)
            End Function
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_mod_validation() {
        let source = r#"
            If Minute(startTime) Mod 15 <> 0 Then
                MsgBox "Invalid"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_debug_print() {
        let source = r"
            Debug.Print Minute(Now)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_with_statement() {
        let source = r"
            With appointmentRecord
                .StartMinute = Minute(.StartTime)
            End With
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_select_case() {
        let source = r#"
            Select Case Minute(currentTime)
                Case 0
                    MsgBox "Top of hour"
                Case 15, 30, 45
                    MsgBox "Quarter hour"
            End Select
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_elseif() {
        let source = r#"
            If m = 0 Then
                status = "Top"
            ElseIf Minute(t) = 30 Then
                status = "Half"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_parentheses() {
        let source = r"
            result = (Minute(timeValue))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_iif() {
        let source = r#"
            result = IIf(Minute(t) >= 30, "Late", "Early")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_in_class() {
        let source = r"
            Private Sub ExtractTime()
                m_minute = Minute(m_timeValue)
            End Sub
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_function_argument() {
        let source = r"
            Call ProcessMinute(Minute(Now))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_property_assignment() {
        let source = r"
            MyObject.CurrentMinute = Minute(Time)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_array_assignment() {
        let source = r"
            minutes(i) = Minute(times(i))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_for_loop() {
        let source = r"
            For i = 1 To count
                m = Minute(appointments(i))
            Next i
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_while_wend() {
        let source = r"
            While Minute(currentTime) < 30
                DoWork
            Wend
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_do_while() {
        let source = r"
            Do While i < recordCount
                minuteValue = Minute(records(i).Time)
                i = i + 1
            Loop
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_do_until() {
        let source = r"
            Do Until Minute(Now) = 0
                Wait 1000
            Loop
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_msgbox() {
        let source = r#"
            MsgBox "Minute: " & Minute(Now)
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_concatenation() {
        let source = r#"
            timeStr = Hour(t) & ":" & Format(Minute(t), "00")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_comparison() {
        let source = r#"
            If Minute(time1) = Minute(time2) Then
                MsgBox "Same"
            End If
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_with_format() {
        let source = r#"
            formatted = Format(Minute(Now), "00")
        "#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_arithmetic() {
        let source = r"
            minutesPast = Minute(currentTime)
            minutesLeft = 60 - Minute(currentTime)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_label_caption() {
        let source = r"
            lblMinute.Caption = CStr(Minute(Time))
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn minute_calculation() {
        let source = r"
            totalMinutes = Hour(t) * 60 + Minute(t)
        ";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");
        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/parsers/syntax/library/functions/datetime/minute",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}

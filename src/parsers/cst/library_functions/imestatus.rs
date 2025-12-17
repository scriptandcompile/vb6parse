//! # `IMEStatus` Function
//!
//! Returns an `Integer` indicating the current `Input Method Editor` (`IME`) mode of Microsoft Windows.
//!
//! ## Syntax
//!
//! ```vb
//! IMEStatus()
//! ```
//!
//! ## Parameters
//!
//! None
//!
//! ## Return Value
//!
//! Returns an `Integer` representing the current `IME` mode:
//!
//! | Constant | Value | Description |
//! |----------|-------|-------------|
//! | vbIMENoOp | 0 | No `IME` installed or `IME` is disabled |
//! | vbIMEOn | 1 | `IME` is on (active) |
//! | vbIMEOff | 2 | `IME` is off (inactive) |
//! | vbIMEDisable | 3 | `IME` is disabled |
//! | vbIMEHiragana | 4 | Double-byte Hiragana mode |
//! | vbIMEKatakanHalf | 5 | Single-byte Katakana mode |
//! | vbIMEKatakanaFull | 6 | Double-byte Katakana mode |
//! | vbIMEAlphaHalf | 7 | Single-byte Alphanumeric mode |
//! | vbIMEAlphaFull | 8 | Double-byte Alphanumeric mode |
//! | vbIMEHangulHalf | 9 | Single-byte Hangul mode |
//! | vbIMEHangulFull | 10 | Double-byte Hangul mode |
//!
//! ## Remarks
//!
//! The `IMEStatus` function provides information about the `Input Method Editor`:
//!
//! - Returns the current state of the `IME` for the active window
//! - `IME` is used primarily for Asian language input (Japanese, Chinese, Korean)
//! - Only meaningful on systems with `IME` support installed
//! - Returns `vbIMENoOp` (0) if no `IME` is installed or available
//! - The return value reflects the `IME` state at the moment the function is called
//! - Can be used to detect if the user is in native language input mode
//! - Useful for applications that need to work with multibyte character sets
//! - The actual modes available depend on the installed `IME` and Windows version
//!
//! ## Typical Uses
//!
//! 1. **IME Detection**: Check if an `IME` is installed and active
//! 2. **Input Mode Validation**: Verify the user is in the correct input mode
//! 3. **Localization**: Adjust application behavior based on `IME` state
//! 4. **Data Entry**: Ensure proper input mode for specific fields
//! 5. **User Guidance**: Provide instructions based on current `IME` mode
//! 6. **Form Validation**: Check input mode before processing data
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Check if IME is available
//! Sub CheckIME()
//!     If IMEStatus() = vbIMENoOp Then
//!         MsgBox "No IME is installed"
//!     Else
//!         MsgBox "IME is available"
//!     End If
//! End Sub
//!
//! ' Example 2: Check if IME is active
//! Sub CheckIMEActive()
//!     If IMEStatus() = vbIMEOn Then
//!         MsgBox "IME is currently on"
//!     Else
//!         MsgBox "IME is currently off"
//!     End If
//! End Sub
//!
//! ' Example 3: Display current IME mode
//! Sub DisplayIMEMode()
//!     Dim mode As Integer
//!     mode = IMEStatus()
//!     Debug.Print "Current IME mode: " & mode
//! End Sub
//!
//! ' Example 4: Detect Japanese input mode
//! Function IsJapaneseInput() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     IsJapaneseInput = (status = vbIMEHiragana Or _
//!                        status = vbIMEKatakanHalf Or _
//!                        status = vbIMEKatakanaFull)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Check if IME is enabled
//! Function IsIMEEnabled() As Boolean
//!     IsIMEEnabled = (IMEStatus() <> vbIMENoOp And IMEStatus() <> vbIMEDisable)
//! End Function
//!
//! ' Pattern 2: Determine if using double-byte characters
//! Function IsDoubleByte() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     
//!     Select Case status
//!         Case vbIMEHiragana, vbIMEKatakanaFull, vbIMEAlphaFull, vbIMEHangulFull
//!             IsDoubleByte = True
//!         Case Else
//!             IsDoubleByte = False
//!     End Select
//! End Function
//!
//! ' Pattern 3: Get IME mode description
//! Function GetIMEModeDescription() As String
//!     Select Case IMEStatus()
//!         Case vbIMENoOp
//!             GetIMEModeDescription = "No IME"
//!         Case vbIMEOn
//!             GetIMEModeDescription = "IME On"
//!         Case vbIMEOff
//!             GetIMEModeDescription = "IME Off"
//!         Case vbIMEDisable
//!             GetIMEModeDescription = "IME Disabled"
//!         Case vbIMEHiragana
//!             GetIMEModeDescription = "Hiragana"
//!         Case vbIMEKatakanHalf
//!             GetIMEModeDescription = "Half-width Katakana"
//!         Case vbIMEKatakanaFull
//!             GetIMEModeDescription = "Full-width Katakana"
//!         Case vbIMEAlphaHalf
//!             GetIMEModeDescription = "Half-width Alphanumeric"
//!         Case vbIMEAlphaFull
//!             GetIMEModeDescription = "Full-width Alphanumeric"
//!         Case vbIMEHangulHalf
//!             GetIMEModeDescription = "Half-width Hangul"
//!         Case vbIMEHangulFull
//!             GetIMEModeDescription = "Full-width Hangul"
//!         Case Else
//!             GetIMEModeDescription = "Unknown mode"
//!     End Select
//! End Function
//!
//! ' Pattern 4: Validate input mode for specific field
//! Function ValidateInputMode(expectedMode As Integer) As Boolean
//!     ValidateInputMode = (IMEStatus() = expectedMode)
//! End Function
//!
//! ' Pattern 5: Detect Asian language input
//! Function IsAsianLanguageInput() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     
//!     ' Check for Japanese or Korean modes
//!     IsAsianLanguageInput = (status >= vbIMEHiragana And status <= vbIMEHangulFull)
//! End Function
//!
//! ' Pattern 6: Check for alphanumeric mode
//! Function IsAlphanumericMode() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     IsAlphanumericMode = (status = vbIMEAlphaHalf Or status = vbIMEAlphaFull)
//! End Function
//!
//! ' Pattern 7: Determine input width
//! Function GetInputWidth() As String
//!     Select Case IMEStatus()
//!         Case vbIMEKatakanHalf, vbIMEAlphaHalf, vbIMEHangulHalf
//!             GetInputWidth = "Half-width"
//!         Case vbIMEHiragana, vbIMEKatakanaFull, vbIMEAlphaFull, vbIMEHangulFull
//!             GetInputWidth = "Full-width"
//!         Case Else
//!             GetInputWidth = "N/A"
//!     End Select
//! End Function
//!
//! ' Pattern 8: Check if IME is in active input mode
//! Function IsActiveInputMode() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     ' Active modes are those actively converting input
//!     IsActiveInputMode = (status >= vbIMEHiragana And status <= vbIMEHangulFull)
//! End Function
//!
//! ' Pattern 9: Detect Katakana mode
//! Function IsKatakanaMode() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     IsKatakanaMode = (status = vbIMEKatakanHalf Or status = vbIMEKatakanaFull)
//! End Function
//!
//! ' Pattern 10: Check for Korean input
//! Function IsKoreanInput() As Boolean
//!     Dim status As Integer
//!     status = IMEStatus()
//!     IsKoreanInput = (status = vbIMEHangulHalf Or status = vbIMEHangulFull)
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: IME mode monitor
//! Public Class IMEMonitor
//!     Private m_lastStatus As Integer
//!     
//!     Public Sub Initialize()
//!         m_lastStatus = IMEStatus()
//!     End Sub
//!     
//!     Public Function HasChanged() As Boolean
//!         Dim currentStatus As Integer
//!         currentStatus = IMEStatus()
//!         
//!         If currentStatus <> m_lastStatus Then
//!             m_lastStatus = currentStatus
//!             HasChanged = True
//!         Else
//!             HasChanged = False
//!         End If
//!     End Function
//!     
//!     Public Function GetCurrentMode() As String
//!         GetCurrentMode = GetIMEModeDescription()
//!     End Function
//! End Class
//!
//! ' Example 2: TextBox IME validator
//! Private Sub txtName_GotFocus()
//!     ' For name field, we want half-width alphanumeric or Katakana
//!     If IMEStatus() <> vbIMEAlphaHalf And _
//!        IMEStatus() <> vbIMEKatakanHalf And _
//!        IMEStatus() <> vbIMEOff Then
//!         MsgBox "Please switch to half-width alphanumeric or Katakana mode", _
//!                vbInformation, "Input Mode"
//!     End If
//! End Sub
//!
//! ' Example 3: Language-aware input handler
//! Function ProcessInput(userInput As String) As String
//!     Dim imeMode As Integer
//!     imeMode = IMEStatus()
//!     
//!     Select Case imeMode
//!         Case vbIMEHiragana, vbIMEKatakanaFull, vbIMEKatakanHalf
//!             ' Japanese input - special processing
//!             ProcessInput = ProcessJapaneseText(userInput)
//!             
//!         Case vbIMEHangulHalf, vbIMEHangulFull
//!             ' Korean input - special processing
//!             ProcessInput = ProcessKoreanText(userInput)
//!             
//!         Case Else
//!             ' Standard processing
//!             ProcessInput = userInput
//!     End Select
//! End Function
//!
//! ' Example 4: Form-wide IME status display
//! Private Sub tmrIMEStatus_Timer()
//!     ' Update status bar with current IME mode
//!     Dim status As Integer
//!     status = IMEStatus()
//!     
//!     StatusBar1.Panels(1).Text = "IME: " & GetIMEModeDescription()
//!     
//!     ' Change indicator color based on mode
//!     If status = vbIMENoOp Or status = vbIMEOff Then
//!         StatusBar1.Panels(1).Picture = LoadPicture(App.Path & "\imeoff.ico")
//!     Else
//!         StatusBar1.Panels(1).Picture = LoadPicture(App.Path & "\imeon.ico")
//!     End If
//! End Sub
//!
//! ' Example 5: Data validation with IME awareness
//! Function ValidateNameField(fieldValue As String) As Boolean
//!     Dim imeMode As Integer
//!     imeMode = IMEStatus()
//!     
//!     ' For Japanese systems, allow Katakana or alphanumeric
//!     If imeMode = vbIMEHiragana Then
//!         MsgBox "Please use Katakana or alphanumeric for names", vbExclamation
//!         ValidateNameField = False
//!         Exit Function
//!     End If
//!     
//!     ' Additional validation
//!     If Len(Trim$(fieldValue)) = 0 Then
//!         ValidateNameField = False
//!     Else
//!         ValidateNameField = True
//!     End If
//! End Function
//!
//! ' Example 6: IME-aware search
//! Function PerformSearch(searchTerm As String) As Collection
//!     Dim results As New Collection
//!     Dim searchMode As String
//!     
//!     ' Determine search strategy based on IME mode
//!     Select Case IMEStatus()
//!         Case vbIMEHiragana, vbIMEKatakanaFull, vbIMEKatakanHalf
//!             searchMode = "Japanese"
//!             ' Use Japanese-specific search algorithm
//!             Set results = SearchJapanese(searchTerm)
//!             
//!         Case vbIMEHangulHalf, vbIMEHangulFull
//!             searchMode = "Korean"
//!             ' Use Korean-specific search algorithm
//!             Set results = SearchKorean(searchTerm)
//!             
//!         Case Else
//!             searchMode = "Standard"
//!             ' Use standard search
//!             Set results = SearchStandard(searchTerm)
//!     End Select
//!     
//!     Set PerformSearch = results
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The `IMEStatus` function rarely raises errors:
//!
//! - Returns `vbIMENoOp` (0) on systems without `IME` support
//! - Does not raise errors if `IME` is not available
//! - Always returns a valid `Integer` value
//! - No error handling typically required
//!
//! ```vb
//! ' Safe to call without error handling
//! Dim status As Integer
//! status = IMEStatus()
//! ```
//!
//! ## Performance Considerations
//!
//! - **Fast Operation**: `IMEStatus` is a very fast system query
//! - **No Overhead**: Minimal performance impact even when called frequently
//! - **Real-time Monitoring**: Safe to call in timer events for status updates
//! - **No Caching Needed**: The function is efficient enough to call directly
//!
//! ## Best Practices
//!
//! 1. **System Compatibility**: Always check for `vbIMENoOp` before assuming `IME` functionality
//! 2. **User Guidance**: Provide clear instructions when specific `IME` modes are required
//! 3. **Non-intrusive**: Don't force `IME` mode changes; suggest them to the user
//! 4. **Status Display**: Show current `IME` mode in status bars for user awareness
//! 5. **Localization**: Use `IMEStatus` to adapt UI for different language inputs
//! 6. **Testing**: Test on both `IME`-enabled and non-`IME` systems
//! 7. **Documentation**: Document `IME` mode requirements for specific fields
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Returns meaningful values only on Windows with `IME` support
//! - `IME` modes depend on installed Windows language packs
//! - Japanese Windows: Hiragana, Katakana modes available
//! - Korean Windows: Hangul modes available
//! - Chinese Windows: May have different mode constants
//! - Western Windows without `IME`: Typically returns `vbIMENoOp`
//!
//! ## Limitations
//!
//! - Only detects `IME` state, cannot change it (use `SendKeys` or Windows API for that)
//! - Return values depend on installed `IME` and language packs
//! - Some `IME` modes may not be available on all systems
//! - Does not detect which specific `IME` software is being used
//! - Limited to Windows `IME` implementation
//! - Cannot distinguish between different Chinese `IME` modes
//! - Return value reflects system state at call time (may change immediately after)
//!
//! ## Related Functions and Properties
//!
//! - `IMEMode` property: Sets/gets the `IME` mode for controls
//! - `SendKeys`: Can be used to change `IME` mode via keyboard shortcuts
//! - Windows API functions for `IME` control (`ImmGetContext`, etc.)
//!
//! ## `IME` Mode Constants Reference
//!
//! ```vb
//! Public Const vbIMENoOp = 0         ' No IME
//! Public Const vbIMEOn = 1           ' IME On
//! Public Const vbIMEOff = 2          ' IME Off  
//! Public Const vbIMEDisable = 3      ' IME Disabled
//! Public Const vbIMEHiragana = 4     ' Hiragana
//! Public Const vbIMEKatakanHalf = 5  ' Half Katakana
//! Public Const vbIMEKatakanaFull = 6 ' Full Katakana
//! Public Const vbIMEAlphaHalf = 7    ' Half Alphanumeric
//! Public Const vbIMEAlphaFull = 8    ' Full Alphanumeric
//! Public Const vbIMEHangulHalf = 9   ' Half Hangul
//! Public Const vbIMEHangulFull = 10  ' Full Hangul
//! ```

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn imestatus_basic() {
        let source = r#"
Sub Test()
    status = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_in_function() {
        let source = r#"
Function GetIMEMode() As Integer
    GetIMEMode = IMEStatus()
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_if_statement() {
        let source = r#"
Sub Test()
    If IMEStatus() = vbIMEOn Then
        Debug.Print "IME is on"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_select_case() {
        let source = r#"
Sub Test()
    Select Case IMEStatus()
        Case vbIMENoOp
            Debug.Print "No IME"
        Case vbIMEOn
            Debug.Print "IME On"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_comparison() {
        let source = r#"
Sub Test()
    If IMEStatus() <> vbIMENoOp Then
        MsgBox "IME available"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_assignment() {
        let source = r#"
Sub Test()
    Dim mode As Integer
    mode = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "IME Status: " & IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_concatenation() {
        let source = r#"
Sub Test()
    msg = "Current mode: " & CStr(IMEStatus())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_with_parentheses() {
        let source = r#"
Sub Test()
    value = (IMEStatus())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_in_do_loop() {
        let source = r#"
Sub Test()
    Do While IMEStatus() = vbIMENoOp
        DoEvents
    Loop
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_class_member() {
        let source = r#"
Private Sub Class_Initialize()
    m_imeMode = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_multiple_comparison() {
        let source = r#"
Sub Test()
    If IMEStatus() = vbIMEHiragana Or IMEStatus() = vbIMEKatakanaFull Then
        Debug.Print "Japanese input"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_function_argument() {
        let source = r#"
Sub Test()
    Call ProcessMode(IMEStatus())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_property_assignment() {
        let source = r#"
Sub Test()
    obj.IMEMode = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_array_assignment() {
        let source = r#"
Sub Test()
    modes(0) = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_with_statement() {
        let source = r#"
Sub Test()
    With statusInfo
        .CurrentMode = IMEStatus()
    End With
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_in_for_loop() {
        let source = r#"
Sub Test()
    Dim i As Integer
    For i = 1 To 10
        If IMEStatus() <> vbIMENoOp Then Exit For
    Next i
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_nested_call() {
        let source = r#"
Sub Test()
    result = CStr(IMEStatus())
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_iif() {
        let source = r#"
Sub Test()
    msg = IIf(IMEStatus() = vbIMEOn, "On", "Off")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_range_check() {
        let source = r#"
Sub Test()
    Dim status As Integer
    status = IMEStatus()
    If status >= vbIMEHiragana And status <= vbIMEHangulFull Then
        Debug.Print "Asian language mode"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_collection_add() {
        let source = r#"
Sub Test()
    col.Add IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_return_value() {
        let source = r#"
Function CheckIME() As Integer
    CheckIME = IMEStatus()
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_type_field() {
        let source = r#"
Sub Test()
    Dim info As SystemInfo
    info.IMEMode = IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_boolean_expression() {
        let source = r#"
Sub Test()
    isEnabled = (IMEStatus() <> vbIMENoOp)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_format() {
        let source = r#"
Sub Test()
    text = "Mode: " & Format$(IMEStatus(), "0")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn imestatus_timer_event() {
        let source = r#"
Private Sub Timer1_Timer()
    lblStatus.Caption = "IME: " & IMEStatus()
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IMEStatus"));
        assert!(text.contains("Identifier"));
    }
}

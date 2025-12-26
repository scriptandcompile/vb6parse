//! # `IRR` Function
//!
//! Returns a `Double` specifying the internal rate of return for a series of periodic cash flows (payments and receipts).
//!
//! ## Syntax
//!
//! ```vb
//! IRR(values()[, guess])
//! ```
//!
//! ## Parameters
//!
//! - `values()` (Required): `Array` of `Double` specifying cash flow values. The array must contain at least one positive value (receipt) and one negative value (payment)
//! - `guess` (Optional): `Variant` specifying value you estimate will be returned by `IRR`. If omitted, guess is 0.1 (10 percent)
//!
//! ## Return Value
//!
//! Returns a `Double` representing the internal rate of return:
//! - Expressed as a decimal (0.1 = 10%)
//! - The discount rate that makes the net present value (`NPV`) of all cash flows equal to zero
//! - Used to evaluate the profitability of potential investments
//! - Higher `IRR` indicates more desirable investment
//!
//! ## Remarks
//!
//! The internal rate of return is the interest rate received for an investment consisting of payments and receipts that occur at regular intervals:
//!
//! - `IRR` uses the order of values within the array to interpret the order of cash flows
//! - Cash flows must occur at regular intervals (monthly, quarterly, annually, etc.)
//! - First element is typically a negative value (initial investment)
//! - Array must contain at least one positive and one negative value
//! - Uses an iterative technique to calculate `IRR`
//! - Begins with the value of `guess` and cycles through until result is accurate to within 0.00001 percent
//! - If `IRR` can't find a result after 20 tries, it fails with Error 5
//! - Most cases, you don't need to provide `guess`; if omitted, 10% is assumed
//! - If `IRR` returns Error 5, try different value for `guess`
//! - `IRR` is closely related to `NPV` (net present value) function
//! - `IRR` is the rate where `NPV` equals zero: `NPV(IRR(values), values) = 0` (approximately)
//!
//! ## Typical Uses
//!
//! 1. **Investment Analysis**: Evaluate profitability of potential investments
//! 2. **Project Evaluation**: Compare multiple projects to select most profitable
//! 3. **Capital Budgeting**: Assess capital expenditure decisions
//! 4. **Business Case Analysis**: Justify business investments with `ROI` calculations
//! 5. **Equipment Purchase**: Evaluate cost savings from new equipment
//! 6. **Real Estate Investment**: Analyze property investment returns
//! 7. **Lease vs Buy**: Compare financial impact of leasing versus purchasing
//! 8. **Portfolio Management**: Assess historical returns on investments
//!
//! ## Basic Usage Examples
//!
//! ```vb
//! ' Example 1: Simple investment analysis
//! Dim cashFlows(0 To 4) As Double
//! Dim returnRate As Double
//!
//! cashFlows(0) = -10000  ' Initial investment (negative = cash out)
//! cashFlows(1) = 3000    ' Year 1 return
//! cashFlows(2) = 3500    ' Year 2 return
//! cashFlows(3) = 4000    ' Year 3 return
//! cashFlows(4) = 4500    ' Year 4 return
//!
//! returnRate = IRR(cashFlows)
//! Debug.Print "Internal Rate of Return: " & Format$(returnRate * 100, "0.00") & "%"
//! ' Prints approximately: 28.09%
//!
//! ' Example 2: Equipment purchase evaluation
//! Dim equipmentCosts(0 To 5) As Double
//! equipmentCosts(0) = -50000  ' Equipment cost
//! equipmentCosts(1) = 12000   ' Year 1 savings
//! equipmentCosts(2) = 15000   ' Year 2 savings
//! equipmentCosts(3) = 18000   ' Year 3 savings
//! equipmentCosts(4) = 21000   ' Year 4 savings
//! equipmentCosts(5) = 24000   ' Year 5 savings
//!
//! returnRate = IRR(equipmentCosts)
//! If returnRate > 0.15 Then  ' 15% hurdle rate
//!     MsgBox "Equipment purchase approved - IRR: " & Format$(returnRate * 100, "0.00") & "%"
//! Else
//!     MsgBox "Equipment purchase rejected - IRR too low"
//! End If
//!
//! ' Example 3: Comparing two projects
//! Dim projectA(0 To 3) As Double
//! Dim projectB(0 To 3) As Double
//!
//! projectA(0) = -25000: projectA(1) = 10000: projectA(2) = 12000: projectA(3) = 15000
//! projectB(0) = -30000: projectB(1) = 15000: projectB(2) = 14000: projectB(3) = 13000
//!
//! Dim irrA As Double, irrB As Double
//! irrA = IRR(projectA)
//! irrB = IRR(projectB)
//!
//! Debug.Print "Project A IRR: " & Format$(irrA * 100, "0.00") & "%"
//! Debug.Print "Project B IRR: " & Format$(irrB * 100, "0.00") & "%"
//!
//! If irrA > irrB Then
//!     MsgBox "Select Project A"
//! Else
//!     MsgBox "Select Project B"
//! End If
//!
//! ' Example 4: Using guess parameter for difficult calculations
//! Dim complexFlows(0 To 6) As Double
//! complexFlows(0) = -100000
//! complexFlows(1) = -50000   ' Additional investment in year 2
//! complexFlows(2) = 20000
//! complexFlows(3) = 40000
//! complexFlows(4) = 50000
//! complexFlows(5) = 60000
//! complexFlows(6) = 70000
//!
//! ' Provide guess to help convergence
//! On Error Resume Next
//! returnRate = IRR(complexFlows, 0.2)  ' Start with 20% guess
//! If Err.Number = 0 Then
//!     Debug.Print "Complex IRR: " & Format$(returnRate * 100, "0.00") & "%"
//! Else
//!     Debug.Print "Could not calculate IRR"
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Calculate IRR for investment
//! Function CalculateInvestmentIRR(initialInvestment As Double, returns() As Double) As Double
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     
//!     ReDim cashFlows(0 To UBound(returns) + 1)
//!     cashFlows(0) = -Abs(initialInvestment)  ' Ensure negative
//!     
//!     For i = 0 To UBound(returns)
//!         cashFlows(i + 1) = returns(i)
//!     Next i
//!     
//!     CalculateInvestmentIRR = IRR(cashFlows)
//! End Function
//!
//! ' Pattern 2: IRR with hurdle rate comparison
//! Function MeetsHurdleRate(cashFlows() As Double, hurdleRate As Double) As Boolean
//!     On Error Resume Next
//!     Dim rate As Double
//!     rate = IRR(cashFlows)
//!     
//!     If Err.Number = 0 Then
//!         MeetsHurdleRate = (rate >= hurdleRate)
//!     Else
//!         MeetsHurdleRate = False
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! ' Pattern 3: Format IRR as percentage
//! Function FormatIRR(cashFlows() As Double) As String
//!     On Error Resume Next
//!     Dim rate As Double
//!     rate = IRR(cashFlows)
//!     
//!     If Err.Number = 0 Then
//!         FormatIRR = Format$(rate * 100, "0.00") & "%"
//!     Else
//!         FormatIRR = "N/A"
//!     End If
//!     On Error GoTo 0
//! End Function
//!
//! ' Pattern 4: Select best investment from multiple options
//! Function SelectBestInvestment(investments As Collection) As Integer
//!     Dim bestIRR As Double
//!     Dim bestIndex As Integer
//!     Dim currentIRR As Double
//!     Dim i As Integer
//!     
//!     bestIRR = -999999
//!     bestIndex = -1
//!     
//!     For i = 1 To investments.Count
//!         On Error Resume Next
//!         currentIRR = IRR(investments(i))
//!         
//!         If Err.Number = 0 And currentIRR > bestIRR Then
//!             bestIRR = currentIRR
//!             bestIndex = i
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     SelectBestInvestment = bestIndex
//! End Function
//!
//! ' Pattern 5: Calculate IRR with validation
//! Function SafeIRR(cashFlows() As Double, Optional guess As Double = 0.1) As Variant
//!     Dim hasPositive As Boolean
//!     Dim hasNegative As Boolean
//!     Dim i As Integer
//!     
//!     ' Validate array has both positive and negative values
//!     For i = LBound(cashFlows) To UBound(cashFlows)
//!         If cashFlows(i) > 0 Then hasPositive = True
//!         If cashFlows(i) < 0 Then hasNegative = True
//!     Next i
//!     
//!     If Not (hasPositive And hasNegative) Then
//!         SafeIRR = Null
//!         Exit Function
//!     End If
//!     
//!     On Error Resume Next
//!     SafeIRR = IRR(cashFlows, guess)
//!     If Err.Number <> 0 Then SafeIRR = Null
//!     On Error GoTo 0
//! End Function
//!
//! ' Pattern 6: Compare project IRRs
//! Sub CompareProjects(project1() As Double, project2() As Double)
//!     Dim irr1 As Double, irr2 As Double
//!     
//!     irr1 = IRR(project1)
//!     irr2 = IRR(project2)
//!     
//!     Debug.Print "Project 1 IRR: " & Format$(irr1 * 100, "0.00") & "%"
//!     Debug.Print "Project 2 IRR: " & Format$(irr2 * 100, "0.00") & "%"
//!     Debug.Print "Difference: " & Format$((irr1 - irr2) * 100, "0.00") & " percentage points"
//! End Sub
//!
//! ' Pattern 7: Calculate breakeven IRR
//! Function GetBreakevenIRR(costOfCapital As Double, cashFlows() As Double) As String
//!     Dim projectIRR As Double
//!     projectIRR = IRR(cashFlows)
//!     
//!     If projectIRR > costOfCapital Then
//!         GetBreakevenIRR = "Project exceeds cost of capital by " & _
//!                          Format$((projectIRR - costOfCapital) * 100, "0.00") & "%"
//!     ElseIf projectIRR < costOfCapital Then
//!         GetBreakevenIRR = "Project falls short of cost of capital by " & _
//!                          Format$((costOfCapital - projectIRR) * 100, "0.00") & "%"
//!     Else
//!         GetBreakevenIRR = "Project exactly meets cost of capital"
//!     End If
//! End Function
//!
//! ' Pattern 8: IRR for monthly cash flows
//! Function MonthlyIRR(monthlyCashFlows() As Double) As Double
//!     ' Returns annualized IRR from monthly cash flows
//!     Dim monthlyRate As Double
//!     monthlyRate = IRR(monthlyCashFlows)
//!     MonthlyIRR = ((1 + monthlyRate) ^ 12) - 1  ' Convert to annual rate
//! End Function
//!
//! ' Pattern 9: Try multiple guesses if IRR fails
//! Function RobustIRR(cashFlows() As Double) As Variant
//!     Dim guesses As Variant
//!     Dim i As Integer
//!     Dim result As Double
//!     
//!     guesses = Array(0.1, 0.2, 0.5, -0.1, -0.2, 0.01, 0.9)
//!     
//!     For i = 0 To UBound(guesses)
//!         On Error Resume Next
//!         result = IRR(cashFlows, guesses(i))
//!         
//!         If Err.Number = 0 Then
//!             RobustIRR = result
//!             On Error GoTo 0
//!             Exit Function
//!         End If
//!         On Error GoTo 0
//!     Next i
//!     
//!     RobustIRR = Null  ' Could not calculate
//! End Function
//!
//! ' Pattern 10: Incremental IRR analysis
//! Function IncrementalIRR(baseProject() As Double, incrementalProject() As Double) As Double
//!     Dim incrementalFlows() As Double
//!     Dim i As Integer
//!     Dim maxIndex As Integer
//!     
//!     ' Calculate incremental cash flows
//!     maxIndex = IIf(UBound(baseProject) > UBound(incrementalProject), _
//!                    UBound(baseProject), UBound(incrementalProject))
//!     
//!     ReDim incrementalFlows(0 To maxIndex)
//!     
//!     For i = 0 To maxIndex
//!         incrementalFlows(i) = 0
//!         If i <= UBound(incrementalProject) Then
//!             incrementalFlows(i) = incrementalFlows(i) + incrementalProject(i)
//!         End If
//!         If i <= UBound(baseProject) Then
//!             incrementalFlows(i) = incrementalFlows(i) - baseProject(i)
//!         End If
//!     Next i
//!     
//!     IncrementalIRR = IRR(incrementalFlows)
//! End Function
//! ```
//!
//! ## Advanced Usage Examples
//!
//! ```vb
//! ' Example 1: Investment analyzer class
//! Public Class InvestmentAnalyzer
//!     Private m_cashFlows() As Double
//!     Private m_irr As Variant
//!     Private m_calculated As Boolean
//!     
//!     Public Sub SetCashFlows(cashFlows() As Double)
//!         Dim i As Integer
//!         ReDim m_cashFlows(LBound(cashFlows) To UBound(cashFlows))
//!         
//!         For i = LBound(cashFlows) To UBound(cashFlows)
//!             m_cashFlows(i) = cashFlows(i)
//!         Next i
//!         
//!         m_calculated = False
//!     End Sub
//!     
//!     Public Function GetIRR() As Variant
//!         If Not m_calculated Then
//!             On Error Resume Next
//!             m_irr = IRR(m_cashFlows)
//!             If Err.Number <> 0 Then m_irr = Null
//!             On Error GoTo 0
//!             m_calculated = True
//!         End If
//!         GetIRR = m_irr
//!     End Function
//!     
//!     Public Function GetFormattedIRR() As String
//!         Dim rate As Variant
//!         rate = GetIRR()
//!         
//!         If IsNull(rate) Then
//!             GetFormattedIRR = "N/A"
//!         Else
//!             GetFormattedIRR = Format$(rate * 100, "0.00") & "%"
//!         End If
//!     End Function
//!     
//!     Public Function IsAcceptable(hurdleRate As Double) As Boolean
//!         Dim rate As Variant
//!         rate = GetIRR()
//!         
//!         If IsNull(rate) Then
//!             IsAcceptable = False
//!         Else
//!             IsAcceptable = (rate >= hurdleRate)
//!         End If
//!     End Function
//!     
//!     Public Function CompareToRate(targetRate As Double) As String
//!         Dim rate As Variant
//!         rate = GetIRR()
//!         
//!         If IsNull(rate) Then
//!             CompareToRate = "Unable to calculate IRR"
//!         ElseIf rate > targetRate Then
//!             CompareToRate = "Exceeds target by " & _
//!                            Format$((rate - targetRate) * 100, "0.00") & "%"
//!         ElseIf rate < targetRate Then
//!             CompareToRate = "Below target by " & _
//!                            Format$((targetRate - rate) * 100, "0.00") & "%"
//!         Else
//!             CompareToRate = "Exactly meets target"
//!         End If
//!     End Function
//! End Class
//!
//! ' Example 2: Project portfolio manager
//! Public Class ProjectPortfolio
//!     Private m_projects As Collection
//!     
//!     Private Sub Class_Initialize()
//!         Set m_projects = New Collection
//!     End Sub
//!     
//!     Public Sub AddProject(projectName As String, cashFlows() As Double)
//!         Dim projectData As Variant
//!         projectData = Array(projectName, cashFlows)
//!         m_projects.Add projectData
//!     End Sub
//!     
//!     Public Function GetBestProject() As String
//!         Dim bestIRR As Double
//!         Dim bestName As String
//!         Dim currentIRR As Double
//!         Dim i As Integer
//!         Dim projectData As Variant
//!         
//!         bestIRR = -999999
//!         bestName = ""
//!         
//!         For i = 1 To m_projects.Count
//!             projectData = m_projects(i)
//!             
//!             On Error Resume Next
//!             currentIRR = IRR(projectData(1))
//!             
//!             If Err.Number = 0 And currentIRR > bestIRR Then
//!                 bestIRR = currentIRR
//!                 bestName = projectData(0)
//!             End If
//!             On Error GoTo 0
//!         Next i
//!         
//!         GetBestProject = bestName & " (IRR: " & Format$(bestIRR * 100, "0.00") & "%)"
//!     End Function
//!     
//!     Public Function GetRankedProjects() As String
//!         Dim rankings() As Variant
//!         Dim i As Integer, j As Integer
//!         Dim temp As Variant
//!         Dim result As String
//!         Dim projectData As Variant
//!         Dim projectIRR As Double
//!         
//!         ReDim rankings(1 To m_projects.Count)
//!         
//!         ' Build array of project names and IRRs
//!         For i = 1 To m_projects.Count
//!             projectData = m_projects(i)
//!             
//!             On Error Resume Next
//!             projectIRR = IRR(projectData(1))
//!             If Err.Number <> 0 Then projectIRR = -999999
//!             On Error GoTo 0
//!             
//!             rankings(i) = Array(projectData(0), projectIRR)
//!         Next i
//!         
//!         ' Sort by IRR (descending)
//!         For i = 1 To UBound(rankings) - 1
//!             For j = i + 1 To UBound(rankings)
//!                 If rankings(j)(1) > rankings(i)(1) Then
//!                     temp = rankings(i)
//!                     rankings(i) = rankings(j)
//!                     rankings(j) = temp
//!                 End If
//!             Next j
//!         Next i
//!         
//!         ' Build result string
//!         result = "Project Rankings:" & vbCrLf
//!         For i = 1 To UBound(rankings)
//!             result = result & i & ". " & rankings(i)(0) & ": " & _
//!                      Format$(rankings(i)(1) * 100, "0.00") & "%" & vbCrLf
//!         Next i
//!         
//!         GetRankedProjects = result
//!     End Function
//! End Class
//!
//! ' Example 3: Capital budgeting calculator
//! Function EvaluateCapitalProject(initialCost As Double, annualSavings As Double, _
//!                                years As Integer, salvageValue As Double, _
//!                                hurdleRate As Double) As String
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     Dim projectIRR As Double
//!     Dim result As String
//!     
//!     ReDim cashFlows(0 To years)
//!     cashFlows(0) = -Abs(initialCost)
//!     
//!     For i = 1 To years - 1
//!         cashFlows(i) = annualSavings
//!     Next i
//!     
//!     cashFlows(years) = annualSavings + salvageValue
//!     
//!     projectIRR = IRR(cashFlows)
//!     
//!     result = "Capital Project Evaluation" & vbCrLf
//!     result = result & "Initial Cost: " & Format$(initialCost, "Currency") & vbCrLf
//!     result = result & "Annual Savings: " & Format$(annualSavings, "Currency") & vbCrLf
//!     result = result & "Project Life: " & years & " years" & vbCrLf
//!     result = result & "Salvage Value: " & Format$(salvageValue, "Currency") & vbCrLf
//!     result = result & "IRR: " & Format$(projectIRR * 100, "0.00") & "%" & vbCrLf
//!     result = result & "Hurdle Rate: " & Format$(hurdleRate * 100, "0.00") & "%" & vbCrLf
//!     
//!     If projectIRR >= hurdleRate Then
//!         result = result & "Recommendation: APPROVE"
//!     Else
//!         result = result & "Recommendation: REJECT"
//!     End If
//!     
//!     EvaluateCapitalProject = result
//! End Function
//!
//! ' Example 4: Real estate investment analyzer
//! Function AnalyzeRealEstateInvestment(purchasePrice As Double, downPayment As Double, _
//!                                      monthlyRent As Double, monthlyExpenses As Double, _
//!                                      years As Integer, appreciationRate As Double) As String
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     Dim salePrice As Double
//!     Dim annualIRR As Double
//!     Dim result As String
//!     
//!     ReDim cashFlows(0 To years)
//!     
//!     ' Initial investment (down payment)
//!     cashFlows(0) = -Abs(downPayment)
//!     
//!     ' Annual net cash flows
//!     For i = 1 To years - 1
//!         cashFlows(i) = (monthlyRent - monthlyExpenses) * 12
//!     Next i
//!     
//!     ' Final year includes sale
//!     salePrice = purchasePrice * ((1 + appreciationRate) ^ years)
//!     cashFlows(years) = (monthlyRent - monthlyExpenses) * 12 + salePrice - (purchasePrice - downPayment)
//!     
//!     annualIRR = IRR(cashFlows)
//!     
//!     result = "Real Estate Investment Analysis" & vbCrLf
//!     result = result & "Purchase Price: " & Format$(purchasePrice, "Currency") & vbCrLf
//!     result = result & "Down Payment: " & Format$(downPayment, "Currency") & vbCrLf
//!     result = result & "Monthly Rent: " & Format$(monthlyRent, "Currency") & vbCrLf
//!     result = result & "Monthly Expenses: " & Format$(monthlyExpenses, "Currency") & vbCrLf
//!     result = result & "Holding Period: " & years & " years" & vbCrLf
//!     result = result & "Annual IRR: " & Format$(annualIRR * 100, "0.00") & "%"
//!     
//!     AnalyzeRealEstateInvestment = result
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! The IRR function can raise errors:
//!
//! - **Invalid procedure call (Error 5)**: If IRR can't find a result after 20 iterations, or if array doesn't contain at least one positive and one negative value
//! - **Type Mismatch (Error 13)**: If values array is not numeric
//! - **Subscript out of range (Error 9)**: If array is invalid
//!
//! ```vb
//! On Error GoTo ErrorHandler
//! Dim cashFlows(0 To 4) As Double
//! Dim rate As Double
//!
//! cashFlows(0) = -10000
//! cashFlows(1) = 3000
//! cashFlows(2) = 3500
//! cashFlows(3) = 4000
//! cashFlows(4) = 4500
//!
//! rate = IRR(cashFlows)
//! Debug.Print "IRR: " & Format$(rate * 100, "0.00") & "%"
//! Exit Sub
//!
//! ErrorHandler:
//!     If Err.Number = 5 Then
//!         MsgBox "Unable to calculate IRR. Try a different guess value.", vbCritical
//!     Else
//!         MsgBox "Error calculating IRR: " & Err.Description, vbCritical
//!     End If
//! ```
//!
//! ## Performance Considerations
//!
//! - **Iterative Calculation**: `IRR` uses iterative algorithm that can be slow for complex cash flows
//! - **Convergence**: May require multiple iterations; providing good `guess` can improve performance
//! - **Array Size**: Larger arrays take longer to process
//! - **Caching**: Cache calculated `IRR` values rather than recalculating repeatedly
//!
//! ## Best Practices
//!
//! 1. **Validate Input**: Ensure array contains at least one positive and one negative value
//! 2. **Error Handling**: Always wrap `IRR` in error handler as it may fail to converge
//! 3. **Sign Convention**: Use negative for cash outflows (investments), positive for inflows (returns)
//! 4. **Provide Guess**: For complex cash flows or when default fails, provide appropriate guess value
//! 5. **Regular Intervals**: Ensure cash flows occur at regular, consistent intervals
//! 6. **Order Matters**: Values must be in chronological order in the array
//! 7. **Hurdle Rate**: Compare `IRR` to hurdle rate or cost of capital to make decisions
//! 8. **Multiple IRRs**: Be aware that some cash flow patterns can have multiple valid `IRR`s
//! 9. **Complement with NPV**: Use `NPV` alongside `IRR` for complete investment analysis
//! 10. **Format for Display**: Multiply by 100 and format as percentage for user display
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Return Value | Use Case |
//! |----------|---------|--------------|----------|
//! | `IRR` | Internal rate of return | Rate (`Decimal`) | Evaluate single investment profitability |
//! | `MIRR` | Modified `IRR` | Rate (`Decimal`) | Handle reinvestment assumptions |
//! | `NPV` | Net present value | `Currency` amount | Calculate dollar value at given rate |
//! | `PV` | Present value | `Currency` amount | Simple annuity present value |
//! | `FV` | Future value | `Currency` amount | Simple annuity future value |
//!
//! ## Platform and Version Notes
//!
//! - Available in all VB6 versions
//! - Part of VBA financial functions
//! - Uses `Double` precision
//! - Consistent with Excel's `IRR` function
//! - Maximum 20 iterations for convergence
//!
//! ## Limitations
//!
//! - Assumes cash flows occur at regular intervals
//! - May fail to converge for certain cash flow patterns
//! - Assumes reinvestment at the `IRR` rate (use `MIRR` for different assumption)
//! - Cannot handle irregular time periods between cash flows (use `XIRR` in Excel for that)
//! - Multiple `IRR`s possible for some cash flow patterns (multiple sign changes)
//! - Does not account for risk differences between projects
//!
//! ## Related Functions
//!
//! - `MIRR`: Modified internal rate of return with reinvestment rate
//! - `NPV`: Net present value
//! - `PV`: Present value
//! - `FV`: Future value
//! - `Rate`: Interest rate per period

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn irr_basic() {
        let source = r"
Sub Test()
    rate = IRR(cashFlows)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_with_guess() {
        let source = r"
Sub Test()
    rate = IRR(cashFlows, 0.1)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_array_literal() {
        let source = r"
Sub Test()
    Dim flows(0 To 4) As Double
    rate = IRR(flows)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_if_statement() {
        let source = r#"
Sub Test()
    If IRR(cashFlows) > 0.15 Then
        MsgBox "Acceptable investment"
    End If
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_function_return() {
        let source = r"
Function CalculateReturn() As Double
    CalculateReturn = IRR(projectCashFlows)
End Function
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_comparison() {
        let source = r"
Sub Test()
    If IRR(project1) > IRR(project2) Then
        selectedProject = 1
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_format() {
        let source = r#"
Sub Test()
    formatted = Format$(IRR(cashFlows) * 100, "0.00") & "%"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_debug_print() {
        let source = r#"
Sub Test()
    Debug.Print "IRR: " & Format$(IRR(cashFlows) * 100, "0.00") & "%"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_variable_assignment() {
        let source = r"
Sub Test()
    Dim returnRate As Double
    returnRate = IRR(investmentFlows)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_property_assignment() {
        let source = r"
Sub Test()
    investment.ReturnRate = IRR(investment.CashFlows)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_in_class() {
        let source = r"
Private Sub Class_Initialize()
    m_irr = IRR(m_cashFlows)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_with_statement() {
        let source = r"
Sub Test()
    With project
        .IRR = IRR(.CashFlows, .Guess)
    End With
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_function_argument() {
        let source = r"
Sub Test()
    Call EvaluateInvestment(IRR(cashFlows))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_select_case() {
        let source = r#"
Sub Test()
    Select Case IRR(cashFlows)
        Case Is > 0.2
            rating = "Excellent"
        Case Is > 0.1
            rating = "Good"
    End Select
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Investment return: " & Format$(IRR(cashFlows) * 100, "0.00") & "%"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_collection_add() {
        let source = r"
Sub Test()
    results.Add IRR(projectFlows(i))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_math_expression() {
        let source = r"
Sub Test()
    spread = IRR(project1) - IRR(project2)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_iif() {
        let source = r#"
Sub Test()
    decision = IIf(IRR(cashFlows) > hurdleRate, "Approve", "Reject")
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    For i = 1 To projectCount
        projectIRRs(i) = IRR(projectFlows(i))
    Next i
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_do_loop() {
        let source = r"
Sub Test()
    Do While IRR(currentFlows) < targetRate
        currentFlows(1) = currentFlows(1) + 1000
    Loop
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_boolean_expression() {
        let source = r"
Sub Test()
    isAcceptable = IRR(cashFlows) > hurdleRate And totalCost < budget
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_error_handling() {
        let source = r"
Sub Test()
    On Error Resume Next
    rate = IRR(cashFlows, 0.2)
    If Err.Number = 0 Then
        Debug.Print rate
    End If
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_multiplication() {
        let source = r"
Sub Test()
    percentageRate = IRR(cashFlows) * 100
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_parentheses() {
        let source = r"
Sub Test()
    value = (IRR(cashFlows))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_concatenation() {
        let source = r#"
Sub Test()
    result = "IRR: " & IRR(cashFlows)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_array_access() {
        let source = r"
Sub Test()
    projectRates(index) = IRR(projectCashFlows(index))
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn irr_nested_call() {
        let source = r"
Sub Test()
    percentage = CStr(IRR(cashFlows) * 100)
End Sub
";
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("IRR"));
        assert!(text.contains("Identifier"));
    }
}

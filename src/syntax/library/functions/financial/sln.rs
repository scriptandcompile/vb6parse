/// # SLN Function
///
/// Returns a Double specifying the straight-line depreciation of an asset for a single period.
///
/// ## Syntax
///
/// ```vb
/// SLN(cost, salvage, life)
/// ```
///
/// ## Parameters
///
/// - `cost` - Required. Double specifying initial cost of the asset.
/// - `salvage` - Required. Double specifying value of the asset at the end of its useful life.
/// - `life` - Required. Double specifying length of the useful life of the asset.
///
/// ## Return Value
///
/// Returns a Double representing the depreciation of an asset for a single period using the straight-line method.
///
/// The formula is:
/// ```text
/// SLN = (cost - salvage) / life
/// ```
///
/// ## Remarks
///
/// The SLN function calculates straight-line depreciation, which is the simplest and most commonly used depreciation method. It assumes that the asset depreciates by the same amount each period over its useful life.
///
/// Key characteristics:
/// - Depreciation is constant for each period
/// - Total depreciation = cost - salvage value
/// - Each period's depreciation = (cost - salvage) / life
/// - All three arguments must be positive
/// - Life represents the number of periods (years, months, etc.)
/// - Units must be consistent (if life is in years, result is yearly depreciation)
///
/// Straight-line depreciation is used when:
/// - Asset provides uniform benefit over its life
/// - Simple, easy-to-understand method is preferred
/// - Tax regulations require or allow straight-line method
/// - Asset doesn't become obsolete quickly
/// - Usage is relatively constant over time
///
/// The straight-line method is one of several depreciation methods:
/// - **SLN**: Straight-line (this function) - constant depreciation
/// - **DDB**: Double-declining balance - accelerated depreciation
/// - **SYD**: Sum-of-years digits - accelerated depreciation
///
/// ## Typical Uses
///
/// 1. **Asset Depreciation**: Calculate annual depreciation expense
/// 2. **Financial Planning**: Project future asset values
/// 3. **Tax Calculations**: Determine tax deductions
/// 4. **Budgeting**: Estimate replacement costs
/// 5. **Accounting**: Prepare financial statements
/// 6. **Asset Management**: Track asset value over time
/// 7. **Cost Analysis**: Calculate total ownership cost
/// 8. **Investment Analysis**: Evaluate asset investments
///
/// ## Basic Examples
///
/// ```vb
/// ' Example 1: Calculate yearly depreciation for equipment
/// Dim equipmentCost As Double
/// Dim salvageValue As Double
/// Dim usefulLife As Double
/// Dim yearlyDepreciation As Double
///
/// equipmentCost = 50000      ' $50,000 initial cost
/// salvageValue = 5000        ' $5,000 salvage value
/// usefulLife = 5             ' 5 years useful life
///
/// yearlyDepreciation = SLN(equipmentCost, salvageValue, usefulLife)
/// ' Returns 9000 ($9,000 per year)
/// ```
///
/// ```vb
/// ' Example 2: Calculate monthly depreciation
/// Dim cost As Double
/// Dim salvage As Double
/// Dim months As Double
/// Dim monthlyDepreciation As Double
///
/// cost = 24000
/// salvage = 2000
/// months = 36  ' 3 years = 36 months
///
/// monthlyDepreciation = SLN(cost, salvage, months)
/// ' Returns 611.11 (approximately)
/// ```
///
/// ```vb
/// ' Example 3: Calculate book value after n periods
/// Dim initialCost As Double
/// Dim salvage As Double
/// Dim life As Double
/// Dim periods As Integer
/// Dim annualDepreciation As Double
/// Dim bookValue As Double
///
/// initialCost = 100000
/// salvage = 10000
/// life = 10
/// periods = 3  ' After 3 years
///
/// annualDepreciation = SLN(initialCost, salvage, life)
/// bookValue = initialCost - (annualDepreciation * periods)
/// ' bookValue = 73000
/// ```
///
/// ```vb
/// ' Example 4: Display depreciation schedule
/// Dim cost As Double
/// Dim salvage As Double
/// Dim life As Integer
/// Dim depreciation As Double
/// Dim i As Integer
///
/// cost = 30000
/// salvage = 3000
/// life = 5
/// depreciation = SLN(cost, salvage, life)
///
/// For i = 1 To life
///     Debug.Print "Year " & i & ": $" & depreciation
/// Next i
/// ```
///
/// ## Common Patterns
///
/// ### Pattern 1: `CalculateBookValue`
/// Calculate asset book value after specified periods
/// ```vb
/// Function CalculateBookValue(cost As Double, salvage As Double, _
///                             life As Double, periodsElapsed As Integer) As Double
///     Dim annualDepreciation As Double
///     Dim totalDepreciation As Double
///     
///     annualDepreciation = SLN(cost, salvage, life)
///     totalDepreciation = annualDepreciation * periodsElapsed
///     
///     ' Ensure book value doesn't go below salvage value
///     If totalDepreciation > (cost - salvage) Then
///         CalculateBookValue = salvage
///     Else
///         CalculateBookValue = cost - totalDepreciation
///     End If
/// End Function
/// ```
///
/// ### Pattern 2: `GenerateDepreciationSchedule`
/// Create complete depreciation schedule
/// ```vb
/// Sub GenerateDepreciationSchedule(cost As Double, salvage As Double, _
///                                  life As Integer)
///     Dim depreciation As Double
///     Dim bookValue As Double
///     Dim i As Integer
///     
///     depreciation = SLN(cost, salvage, life)
///     bookValue = cost
///     
///     Debug.Print "Year" & vbTab & "Depreciation" & vbTab & "Book Value"
///     Debug.Print "0" & vbTab & "0" & vbTab & bookValue
///     
///     For i = 1 To life
///         bookValue = bookValue - depreciation
///         Debug.Print i & vbTab & depreciation & vbTab & bookValue
///     Next i
/// End Sub
/// ```
///
/// ### Pattern 3: `CompareDepreciationMethods`
/// Compare straight-line with other methods
/// ```vb
/// Sub CompareDepreciationMethods(cost As Double, salvage As Double, _
///                                life As Integer, year As Integer)
///     Dim slnDepreciation As Double
///     Dim ddbDepreciation As Double
///     
///     slnDepreciation = SLN(cost, salvage, life)
///     ddbDepreciation = DDB(cost, salvage, life, year)
///     
///     Debug.Print "Year " & year
///     Debug.Print "Straight-Line: $" & Format(slnDepreciation, "#,##0.00")
///     Debug.Print "Double-Declining: $" & Format(ddbDepreciation, "#,##0.00")
/// End Sub
/// ```
///
/// ### Pattern 4: `CalculateTotalDepreciation`
/// Calculate total depreciation over asset life
/// ```vb
/// Function CalculateTotalDepreciation(cost As Double, salvage As Double) As Double
///     CalculateTotalDepreciation = cost - salvage
/// End Function
/// ```
///
/// ### Pattern 5: `ValidateDepreciationInputs`
/// Validate inputs before calculating depreciation
/// ```vb
/// Function ValidateDepreciationInputs(cost As Double, salvage As Double, _
///                                     life As Double) As Boolean
///     ValidateDepreciationInputs = False
///     
///     If cost <= 0 Then
///         MsgBox "Cost must be positive", vbExclamation
///         Exit Function
///     End If
///     
///     If salvage < 0 Then
///         MsgBox "Salvage value cannot be negative", vbExclamation
///         Exit Function
///     End If
///     
///     If salvage >= cost Then
///         MsgBox "Salvage value must be less than cost", vbExclamation
///         Exit Function
///     End If
///     
///     If life <= 0 Then
///         MsgBox "Life must be positive", vbExclamation
///         Exit Function
///     End If
///     
///     ValidateDepreciationInputs = True
/// End Function
/// ```
///
/// ### Pattern 6: `CalculateMonthlyDepreciation`
/// Convert annual to monthly depreciation
/// ```vb
/// Function CalculateMonthlyDepreciation(cost As Double, salvage As Double, _
///                                       yearsLife As Double) As Double
///     Dim annualDepreciation As Double
///     
///     annualDepreciation = SLN(cost, salvage, yearsLife)
///     CalculateMonthlyDepreciation = annualDepreciation / 12
/// End Function
/// ```
///
/// ### Pattern 7: `CalculateDepreciationRate`
/// Calculate depreciation rate as percentage
/// ```vb
/// Function CalculateDepreciationRate(cost As Double, salvage As Double, _
///                                    life As Double) As Double
///     Dim annualDepreciation As Double
///     
///     annualDepreciation = SLN(cost, salvage, life)
///     CalculateDepreciationRate = (annualDepreciation / cost) * 100
/// End Function
/// ```
///
/// ### Pattern 8: `CalculateReplacementYear`
/// Determine when asset should be replaced
/// ```vb
/// Function CalculateReplacementYear(cost As Double, salvage As Double, _
///                                   life As Double, _
///                                   minimumValue As Double) As Integer
///     Dim depreciation As Double
///     Dim bookValue As Double
///     Dim year As Integer
///     
///     depreciation = SLN(cost, salvage, life)
///     bookValue = cost
///     
///     For year = 1 To life
///         bookValue = bookValue - depreciation
///         If bookValue <= minimumValue Then
///             CalculateReplacementYear = year
///             Exit Function
///         End If
///     Next year
///     
///     CalculateReplacementYear = life
/// End Function
/// ```
///
/// ### Pattern 9: `CalculateAccumulatedDepreciation`
/// Calculate accumulated depreciation at specific period
/// ```vb
/// Function CalculateAccumulatedDepreciation(cost As Double, salvage As Double, _
///                                           life As Double, _
///                                           period As Integer) As Double
///     Dim annualDepreciation As Double
///     
///     annualDepreciation = SLN(cost, salvage, life)
///     
///     If period >= life Then
///         CalculateAccumulatedDepreciation = cost - salvage
///     Else
///         CalculateAccumulatedDepreciation = annualDepreciation * period
///     End If
/// End Function
/// ```
///
/// ### Pattern 10: `FormatDepreciationReport`
/// Format depreciation information for display
/// ```vb
/// Function FormatDepreciationReport(cost As Double, salvage As Double, _
///                                   life As Double, year As Integer) As String
///     Dim depreciation As Double
///     Dim accumulated As Double
///     Dim bookValue As Double
///     Dim report As String
///     
///     depreciation = SLN(cost, salvage, life)
///     accumulated = depreciation * year
///     bookValue = cost - accumulated
///     
///     report = "Depreciation Report - Year " & year & vbCrLf
///     report = report & "Initial Cost: $" & Format(cost, "#,##0.00") & vbCrLf
///     report = report & "Annual Depreciation: $" & Format(depreciation, "#,##0.00") & vbCrLf
///     report = report & "Accumulated Depreciation: $" & Format(accumulated, "#,##0.00") & vbCrLf
///     report = report & "Book Value: $" & Format(bookValue, "#,##0.00") & vbCrLf
///     
///     FormatDepreciationReport = report
/// End Function
/// ```
///
/// ## Advanced Usage
///
/// ### Example 1: `AssetDepreciationTracker` Class
/// Track depreciation for multiple assets
/// ```vb
/// ' Class: AssetDepreciationTracker
/// Private Type AssetInfo
///     AssetName As String
///     Cost As Double
///     Salvage As Double
///     Life As Double
///     PurchaseDate As Date
///     AnnualDepreciation As Double
/// End Type
///
/// Private m_assets() As AssetInfo
/// Private m_assetCount As Integer
///
/// Private Sub Class_Initialize()
///     m_assetCount = 0
///     ReDim m_assets(0 To 9)
/// End Sub
///
/// Public Sub AddAsset(assetName As String, cost As Double, _
///                     salvage As Double, life As Double, _
///                     purchaseDate As Date)
///     If m_assetCount > UBound(m_assets) Then
///         ReDim Preserve m_assets(0 To m_assetCount * 2)
///     End If
///     
///     With m_assets(m_assetCount)
///         .AssetName = assetName
///         .Cost = cost
///         .Salvage = salvage
///         .Life = life
///         .PurchaseDate = purchaseDate
///         .AnnualDepreciation = SLN(cost, salvage, life)
///     End With
///     
///     m_assetCount = m_assetCount + 1
/// End Sub
///
/// Public Function GetAssetBookValue(assetIndex As Integer, _
///                                   asOfDate As Date) As Double
///     Dim yearsElapsed As Double
///     Dim totalDepreciation As Double
///     Dim bookValue As Double
///     
///     If assetIndex < 0 Or assetIndex >= m_assetCount Then
///         GetAssetBookValue = 0
///         Exit Function
///     End If
///     
///     With m_assets(assetIndex)
///         yearsElapsed = DateDiff("d", .PurchaseDate, asOfDate) / 365.25
///         
///         If yearsElapsed >= .Life Then
///             GetAssetBookValue = .Salvage
///         Else
///             totalDepreciation = .AnnualDepreciation * yearsElapsed
///             bookValue = .Cost - totalDepreciation
///             
///             If bookValue < .Salvage Then
///                 GetAssetBookValue = .Salvage
///             Else
///                 GetAssetBookValue = bookValue
///             End If
///         End If
///     End With
/// End Function
///
/// Public Function GetTotalBookValue(asOfDate As Date) As Double
///     Dim i As Integer
///     Dim total As Double
///     
///     total = 0
///     For i = 0 To m_assetCount - 1
///         total = total + GetAssetBookValue(i, asOfDate)
///     Next i
///     
///     GetTotalBookValue = total
/// End Function
///
/// Public Function GetAnnualDepreciationExpense() As Double
///     Dim i As Integer
///     Dim total As Double
///     
///     total = 0
///     For i = 0 To m_assetCount - 1
///         total = total + m_assets(i).AnnualDepreciation
///     Next i
///     
///     GetAnnualDepreciationExpense = total
/// End Function
///
/// Public Function GetAssetCount() As Integer
///     GetAssetCount = m_assetCount
/// End Function
/// ```
///
/// ### Example 2: `DepreciationScheduleGenerator` Module
/// Generate detailed depreciation schedules
/// ```vb
/// ' Module: DepreciationScheduleGenerator
///
/// Public Function GenerateScheduleArray(cost As Double, salvage As Double, _
///                                       life As Integer) As Variant
///     Dim schedule() As Variant
///     Dim depreciation As Double
///     Dim bookValue As Double
///     Dim accumulated As Double
///     Dim i As Integer
///     
///     ReDim schedule(0 To life, 0 To 4)
///     
///     ' Headers
///     schedule(0, 0) = "Year"
///     schedule(0, 1) = "Beginning Value"
///     schedule(0, 2) = "Depreciation"
///     schedule(0, 3) = "Accumulated"
///     schedule(0, 4) = "Ending Value"
///     
///     depreciation = SLN(cost, salvage, life)
///     bookValue = cost
///     accumulated = 0
///     
///     For i = 1 To life
///         schedule(i, 0) = i
///         schedule(i, 1) = bookValue
///         schedule(i, 2) = depreciation
///         accumulated = accumulated + depreciation
///         schedule(i, 3) = accumulated
///         bookValue = cost - accumulated
///         schedule(i, 4) = bookValue
///     Next i
///     
///     GenerateScheduleArray = schedule
/// End Function
///
/// Public Sub ExportScheduleToCSV(cost As Double, salvage As Double, _
///                                life As Integer, filePath As String)
///     Dim schedule As Variant
///     Dim fileNum As Integer
///     Dim i As Integer
///     Dim j As Integer
///     Dim line As String
///     
///     schedule = GenerateScheduleArray(cost, salvage, life)
///     
///     fileNum = FreeFile
///     Open filePath For Output As #fileNum
///     
///     For i = 0 To life
///         line = ""
///         For j = 0 To 4
///             If j > 0 Then line = line & ","
///             line = line & schedule(i, j)
///         Next j
///         Print #fileNum, line
///     Next i
///     
///     Close #fileNum
/// End Sub
///
/// Public Function CalculateQuarterlyDepreciation(cost As Double, salvage As Double, _
///                                                yearsLife As Integer) As Double()
///     Dim annualDepreciation As Double
///     Dim quarterlyDepreciation As Double
///     Dim quarters() As Double
///     Dim i As Integer
///     
///     annualDepreciation = SLN(cost, salvage, yearsLife)
///     quarterlyDepreciation = annualDepreciation / 4
///     
///     ReDim quarters(1 To yearsLife * 4)
///     
///     For i = 1 To yearsLife * 4
///         quarters(i) = quarterlyDepreciation
///     Next i
///     
///     CalculateQuarterlyDepreciation = quarters
/// End Function
/// ```
///
/// ### Example 3: `DepreciationComparison` Class
/// Compare different depreciation methods
/// ```vb
/// ' Class: DepreciationComparison
/// Private m_cost As Double
/// Private m_salvage As Double
/// Private m_life As Integer
///
/// Public Sub Initialize(cost As Double, salvage As Double, life As Integer)
///     m_cost = cost
///     m_salvage = salvage
///     m_life = life
/// End Sub
///
/// Public Function GetStraightLineDepreciation(year As Integer) As Double
///     GetStraightLineDepreciation = SLN(m_cost, m_salvage, m_life)
/// End Function
///
/// Public Function GetDoubleDecliningDepreciation(year As Integer) As Double
///     GetDoubleDecliningDepreciation = DDB(m_cost, m_salvage, m_life, year)
/// End Function
///
/// Public Function GetStraightLineBookValue(year As Integer) As Double
///     Dim depreciation As Double
///     depreciation = SLN(m_cost, m_salvage, m_life)
///     GetStraightLineBookValue = m_cost - (depreciation * year)
/// End Function
///
/// Public Function GetTotalDifference(year As Integer) As Double
///     Dim slnTotal As Double
///     Dim ddbTotal As Double
///     Dim i As Integer
///     
///     slnTotal = 0
///     ddbTotal = 0
///     
///     For i = 1 To year
///         slnTotal = slnTotal + GetStraightLineDepreciation(i)
///         ddbTotal = ddbTotal + GetDoubleDecliningDepreciation(i)
///     Next i
///     
///     GetTotalDifference = ddbTotal - slnTotal
/// End Function
///
/// Public Function GenerateComparisonReport(maxYears As Integer) As String
///     Dim report As String
///     Dim i As Integer
///     
///     report = "Depreciation Method Comparison" & vbCrLf
///     report = report & "Cost: $" & Format(m_cost, "#,##0.00") & vbCrLf
///     report = report & "Salvage: $" & Format(m_salvage, "#,##0.00") & vbCrLf
///     report = report & "Life: " & m_life & " years" & vbCrLf & vbCrLf
///     
///     report = report & "Year" & vbTab & "SLN" & vbTab & "DDB" & vbCrLf
///     
///     For i = 1 To maxYears
///         report = report & i & vbTab
///         report = report & Format(GetStraightLineDepreciation(i), "#,##0") & vbTab
///         report = report & Format(GetDoubleDecliningDepreciation(i), "#,##0") & vbCrLf
///     Next i
///     
///     GenerateComparisonReport = report
/// End Function
/// ```
///
/// ### Example 4: `FinancialPlanner` Module
/// Financial planning with depreciation
/// ```vb
/// ' Module: FinancialPlanner
///
/// Public Function CalculateReplacementFund(cost As Double, salvage As Double, _
///                                         life As Double) As Double
///     ' Calculate annual savings needed to replace asset
///     Dim replacementCost As Double
///     
///     ' Assume replacement cost increases by inflation
///     replacementCost = cost * 1.03 ^ life  ' 3% annual inflation
///     
///     CalculateReplacementFund = (replacementCost - salvage) / life
/// End Function
///
/// Public Function CalculateTaxSavings(cost As Double, salvage As Double, _
///                                     life As Double, taxRate As Double) As Double
///     ' Calculate annual tax savings from depreciation
///     Dim annualDepreciation As Double
///     
///     annualDepreciation = SLN(cost, salvage, life)
///     CalculateTaxSavings = annualDepreciation * taxRate
/// End Function
///
/// Public Function CalculateNetPresentValue(cost As Double, salvage As Double, _
///                                          life As Integer, _
///                                          discountRate As Double) As Double
///     ' Calculate NPV of depreciation tax shield
///     Dim annualDepreciation As Double
///     Dim taxShield As Double
///     Dim npv As Double
///     Dim i As Integer
///     Const TAX_RATE As Double = 0.3  ' 30% tax rate
///     
///     annualDepreciation = SLN(cost, salvage, life)
///     taxShield = annualDepreciation * TAX_RATE
///     npv = 0
///     
///     For i = 1 To life
///         npv = npv + taxShield / ((1 + discountRate) ^ i)
///     Next i
///     
///     CalculateNetPresentValue = npv
/// End Function
///
/// Public Function ShouldReplaceAsset(currentBookValue As Double, _
///                                    replacementCost As Double, _
///                                    yearsRemaining As Double, _
///                                    expectedSavings As Double) As Boolean
///     ' Determine if asset should be replaced now
///     Dim replacementBenefit As Double
///     
///     replacementBenefit = expectedSavings * yearsRemaining
///     ShouldReplaceAsset = replacementBenefit > replacementCost
/// End Function
/// ```
///
/// ## Error Handling
///
/// The SLN function can generate the following errors:
///
/// - **Error 5** (Invalid procedure call or argument): If life is zero or negative
/// - **Error 6** (Overflow): If result exceeds Double range
/// - **Error 13** (Type mismatch): Arguments not numeric
///
/// Always validate inputs before calling SLN:
/// ```vb
/// On Error Resume Next
/// depreciation = SLN(cost, salvage, life)
/// If Err.Number <> 0 Then
///     MsgBox "Error calculating depreciation: " & Err.Description
/// End If
/// ```
///
/// ## Performance Considerations
///
/// - SLN is a very fast calculation (simple division)
/// - No iterative calculations required
/// - Can be called repeatedly without performance concerns
/// - Consider caching result if used multiple times with same inputs
///
/// ## Best Practices
///
/// 1. **Validate Inputs**: Check cost > salvage, life > 0
/// 2. **Consistent Units**: Ensure life matches desired period (years, months)
/// 3. **Handle Edge Cases**: Check for zero life, negative values
/// 4. **Document Assumptions**: Clearly state depreciation period
/// 5. **Salvage Value**: Use realistic salvage value estimates
/// 6. **Format Output**: Use Format function for currency display
/// 7. **Complete Schedules**: Generate full schedules for planning
/// 8. **Compare Methods**: Evaluate if straight-line is most appropriate
/// 9. **Tax Compliance**: Verify method meets tax requirements
/// 10. **Regular Review**: Update assumptions periodically
///
/// ## Comparison with Related Functions
///
/// | Function | Method | Depreciation Pattern | Best For |
/// |----------|--------|---------------------|----------|
/// | SLN | Straight-line | Constant each period | Uniform usage assets |
/// | DDB | Double-declining balance | Accelerated (higher early) | Tech, vehicles |
/// | SYD | Sum-of-years digits | Accelerated (graduated) | Equipment with declining efficiency |
///
/// ## Platform Considerations
///
/// - Available in VB6, VBA (all versions)
/// - Part of financial functions library
/// - Returns Double for precision
/// - Consistent across platforms
///
/// ## Limitations
///
/// - Assumes constant depreciation (may not reflect reality)
/// - Doesn't account for accelerated wear
/// - No adjustment for partial periods
/// - Salvage value is an estimate
/// - Doesn't consider tax implications directly
/// - Can't handle mid-period asset purchases without adjustment
///
/// ## Related Functions
///
/// - `DDB`: Returns depreciation using double-declining balance method
/// - `SYD`: Returns depreciation using sum-of-years digits method
/// - `FV`: Calculates future value (inverse concept)
/// - `PV`: Calculates present value
///
#[cfg(test)]
mod tests {
    use crate::assert_tree;
    use crate::*;
    #[test]
    fn sln_basic() {
        let source = r"
Sub Test()
    Dim depreciation As Double
    depreciation = SLN(50000, 5000, 5)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("depreciation"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("depreciation"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("50000"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5000"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("5"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_with_variables() {
        let source = r"
Sub Test()
    Dim cost As Double
    Dim salvage As Double
    Dim life As Double
    Dim result As Double
    result = SLN(cost, salvage, life)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("cost"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("salvage"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("life"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("result"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("result"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_if_statement() {
        let source = r#"
Sub Test()
    If SLN(cost, salvage, life) > 1000 Then
        MsgBox "High depreciation"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("SLN"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("cost"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("salvage"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("life"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("1000"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"High depreciation\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_function_return() {
        let source = r"
Function CalculateDepreciation(c As Double, s As Double, l As Double) As Double
    CalculateDepreciation = SLN(c, s, l)
End Function
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            FunctionStatement {
                FunctionKeyword,
                Whitespace,
                Identifier ("CalculateDepreciation"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("c"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("s"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("l"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("CalculateDepreciation"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("c"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("l"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_variable_assignment() {
        let source = r"
Sub Test()
    Dim annualDepreciation As Double
    annualDepreciation = SLN(100000, 10000, 10)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("annualDepreciation"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("annualDepreciation"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("100000"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10000"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("10"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_msgbox() {
        let source = r#"
Sub Test()
    MsgBox "Annual depreciation: " & SLN(cost, salvage, life)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Annual depreciation: \""),
                        Whitespace,
                        Ampersand,
                        Whitespace,
                        Identifier ("SLN"),
                        LeftParenthesis,
                        Identifier ("cost"),
                        Comma,
                        Whitespace,
                        Identifier ("salvage"),
                        Comma,
                        Whitespace,
                        Identifier ("life"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_debug_print() {
        let source = r"
Sub Test()
    Debug.Print SLN(30000, 3000, 5)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("Debug"),
                        PeriodOperator,
                        PrintKeyword,
                        Whitespace,
                        Identifier ("SLN"),
                        LeftParenthesis,
                        IntegerLiteral ("30000"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("3000"),
                        Comma,
                        Whitespace,
                        IntegerLiteral ("5"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_select_case() {
        let source = r#"
Sub Test()
    Select Case SLN(cost, salvage, life)
        Case Is > 10000
            MsgBox "High"
        Case Is > 5000
            MsgBox "Medium"
        Case Else
            MsgBox "Low"
    End Select
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    SelectCaseStatement {
                        Whitespace,
                        SelectKeyword,
                        Whitespace,
                        CaseKeyword,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                        Whitespace,
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("10000"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"High\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseClause {
                            CaseKeyword,
                            Whitespace,
                            IsKeyword,
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IntegerLiteral ("5000"),
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Medium\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        CaseElseClause {
                            CaseKeyword,
                            Whitespace,
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Low\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        SelectKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_class_usage() {
        let source = r"
Class AssetManager
    Public Function GetDepreciation(c As Double, s As Double, l As Double) As Double
        GetDepreciation = SLN(c, s, l)
    End Function
End Class
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("AssetManager"),
                Newline,
            },
            Whitespace,
            FunctionStatement {
                PublicKeyword,
                Whitespace,
                FunctionKeyword,
                Whitespace,
                Identifier ("GetDepreciation"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("c"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("s"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    Comma,
                    Whitespace,
                    Identifier ("l"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
                StatementList {
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("GetDepreciation"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("c"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("s"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("l"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    Whitespace,
                },
                EndKeyword,
                Whitespace,
                FunctionKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
        ]);
    }

    #[test]
    fn sln_with_statement() {
        let source = r"
Sub Test()
    With Asset
        Dim dep As Double
        dep = SLN(.Cost, .Salvage, .Life)
    End With
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WithStatement {
                        Whitespace,
                        WithKeyword,
                        Whitespace,
                        Identifier ("Asset"),
                        Newline,
                        StatementList {
                            Whitespace,
                            DimStatement {
                                DimKeyword,
                                Whitespace,
                                Identifier ("dep"),
                                Whitespace,
                                AsKeyword,
                                Whitespace,
                                DoubleKeyword,
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("dep"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("SLN"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                PeriodOperator,
                                            },
                                        },
                                    },
                                },
                            },
                            CallStatement {
                                Identifier ("Cost"),
                                Comma,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("Salvage"),
                                Comma,
                                Whitespace,
                                PeriodOperator,
                                Identifier ("Life"),
                                RightParenthesis,
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        WithKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_elseif() {
        let source = r#"
Sub Test()
    Dim d As Double
    d = SLN(cost, salvage, life)
    If d > 10000 Then
        MsgBox "High"
    ElseIf d > 5000 Then
        MsgBox "Medium"
    Else
        MsgBox "Low"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("d"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("d"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("d"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("10000"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"High\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        ElseIfClause {
                            ElseIfKeyword,
                            Whitespace,
                            BinaryExpression {
                                IdentifierExpression {
                                    Identifier ("d"),
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                NumericLiteralExpression {
                                    IntegerLiteral ("5000"),
                                },
                            },
                            Whitespace,
                            ThenKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Medium\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        ElseClause {
                            ElseKeyword,
                            Newline,
                            StatementList {
                                Whitespace,
                                CallStatement {
                                    Identifier ("MsgBox"),
                                    Whitespace,
                                    StringLiteral ("\"Low\""),
                                    Newline,
                                },
                                Whitespace,
                            },
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_for_loop() {
        let source = r"
Sub Test()
    Dim i As Integer
    Dim dep As Double
    dep = SLN(cost, salvage, life)
    For i = 1 To life
        Debug.Print dep
    Next i
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("dep"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("dep"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("life"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                Identifier ("dep"),
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_do_while() {
        let source = r"
Sub Test()
    Do While bookValue > salvage
        bookValue = bookValue - SLN(cost, salvage, life)
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("bookValue"),
                            },
                            Whitespace,
                            GreaterThanOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("salvage"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("bookValue"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("bookValue"),
                                    },
                                    Whitespace,
                                    SubtractionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("SLN"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("cost"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("salvage"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("life"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_do_until() {
        let source = r"
Sub Test()
    Do Until accumulated >= totalDepreciable
        accumulated = accumulated + SLN(cost, salvage, life)
    Loop
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    DoStatement {
                        Whitespace,
                        DoKeyword,
                        Whitespace,
                        UntilKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("accumulated"),
                            },
                            Whitespace,
                            GreaterThanOrEqualOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("totalDepreciable"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("accumulated"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("accumulated"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("SLN"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("cost"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("salvage"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("life"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        LoopKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_while_wend() {
        let source = r"
Sub Test()
    While year <= life
        total = total + SLN(cost, salvage, life)
        year = year + 1
    Wend
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    WhileStatement {
                        Whitespace,
                        WhileKeyword,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("year"),
                            },
                            Whitespace,
                            LessThanOrEqualOperator,
                            Whitespace,
                            IdentifierExpression {
                                Identifier ("life"),
                            },
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("total"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("total"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    CallExpression {
                                        Identifier ("SLN"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("cost"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("salvage"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("life"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                            AssignmentStatement {
                                IdentifierExpression {
                                    Identifier ("year"),
                                },
                                Whitespace,
                                EqualityOperator,
                                Whitespace,
                                BinaryExpression {
                                    IdentifierExpression {
                                        Identifier ("year"),
                                    },
                                    Whitespace,
                                    AdditionOperator,
                                    Whitespace,
                                    NumericLiteralExpression {
                                        IntegerLiteral ("1"),
                                    },
                                },
                                Newline,
                            },
                            Whitespace,
                        },
                        WendKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_parentheses() {
        let source = r"
Sub Test()
    Dim total As Double
    total = (SLN(cost1, salvage1, life1) + SLN(cost2, salvage2, life2))
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("total"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("total"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("SLN"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("cost1"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("salvage1"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("life1"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                AdditionOperator,
                                Whitespace,
                                CallExpression {
                                    Identifier ("SLN"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("cost2"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("salvage2"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("life2"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_iif() {
        let source = r#"
Sub Test()
    Dim msg As String
    msg = IIf(SLN(cost, salvage, life) > threshold, "High", "Low")
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("msg"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("msg"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("IIf"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    BinaryExpression {
                                        CallExpression {
                                            Identifier ("SLN"),
                                            LeftParenthesis,
                                            ArgumentList {
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("cost"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("salvage"),
                                                    },
                                                },
                                                Comma,
                                                Whitespace,
                                                Argument {
                                                    IdentifierExpression {
                                                        Identifier ("life"),
                                                    },
                                                },
                                            },
                                            RightParenthesis,
                                        },
                                        Whitespace,
                                        GreaterThanOperator,
                                        Whitespace,
                                        IdentifierExpression {
                                            Identifier ("threshold"),
                                        },
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"High\""),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    StringLiteralExpression {
                                        StringLiteral ("\"Low\""),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_array_assignment() {
        let source = r"
Sub Test()
    Dim schedule(10) As Double
    schedule(0) = SLN(cost, salvage, life)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("schedule"),
                        LeftParenthesis,
                        NumericLiteralExpression {
                            IntegerLiteral ("10"),
                        },
                        RightParenthesis,
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        CallExpression {
                            Identifier ("schedule"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    NumericLiteralExpression {
                                        IntegerLiteral ("0"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_property_assignment() {
        let source = r"
Class Asset
    Public AnnualDepreciation As Double
End Class

Sub Test()
    Dim a As New Asset
    a.AnnualDepreciation = SLN(cost, salvage, life)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            Unknown,
            Whitespace,
            CallStatement {
                Identifier ("Asset"),
                Newline,
            },
            Whitespace,
            DimStatement {
                PublicKeyword,
                Whitespace,
                Identifier ("AnnualDepreciation"),
                Whitespace,
                AsKeyword,
                Whitespace,
                DoubleKeyword,
                Newline,
            },
            Unknown,
            Whitespace,
            Unknown,
            Newline,
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("a"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        NewKeyword,
                        Whitespace,
                        Identifier ("Asset"),
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        MemberAccessExpression {
                            Identifier ("a"),
                            PeriodOperator,
                            Identifier ("AnnualDepreciation"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_function_argument() {
        let source = r"
Sub ProcessDepreciation(value As Double)
End Sub

Sub Test()
    ProcessDepreciation SLN(cost, salvage, life)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("ProcessDepreciation"),
                ParameterList {
                    LeftParenthesis,
                    Identifier ("value"),
                    Whitespace,
                    AsKeyword,
                    Whitespace,
                    DoubleKeyword,
                    RightParenthesis,
                },
                Newline,
                StatementList,
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    CallStatement {
                        Identifier ("ProcessDepreciation"),
                        Whitespace,
                        Identifier ("SLN"),
                        LeftParenthesis,
                        Identifier ("cost"),
                        Comma,
                        Whitespace,
                        Identifier ("salvage"),
                        Comma,
                        Whitespace,
                        Identifier ("life"),
                        RightParenthesis,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_concatenation() {
        let source = r#"
Sub Test()
    Dim report As String
    report = "Depreciation: $" & SLN(cost, salvage, life)
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("report"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        StringKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("report"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            StringLiteralExpression {
                                StringLiteral ("\"Depreciation: $\""),
                            },
                            Whitespace,
                            Ampersand,
                            Whitespace,
                            CallExpression {
                                Identifier ("SLN"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("cost"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("salvage"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("life"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_comparison() {
        let source = r"
Sub Test()
    Dim needsAttention As Boolean
    needsAttention = (SLN(cost, salvage, life) > budget)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("needsAttention"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        BooleanKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("needsAttention"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        ParenthesizedExpression {
                            LeftParenthesis,
                            BinaryExpression {
                                CallExpression {
                                    Identifier ("SLN"),
                                    LeftParenthesis,
                                    ArgumentList {
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("cost"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("salvage"),
                                            },
                                        },
                                        Comma,
                                        Whitespace,
                                        Argument {
                                            IdentifierExpression {
                                                Identifier ("life"),
                                            },
                                        },
                                    },
                                    RightParenthesis,
                                },
                                Whitespace,
                                GreaterThanOperator,
                                Whitespace,
                                IdentifierExpression {
                                    Identifier ("budget"),
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_arithmetic() {
        let source = r"
Sub Test()
    Dim bookValue As Double
    bookValue = initialCost - (SLN(cost, salvage, life) * years)
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("bookValue"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("bookValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            IdentifierExpression {
                                Identifier ("initialCost"),
                            },
                            Whitespace,
                            SubtractionOperator,
                            Whitespace,
                            ParenthesizedExpression {
                                LeftParenthesis,
                                BinaryExpression {
                                    CallExpression {
                                        Identifier ("SLN"),
                                        LeftParenthesis,
                                        ArgumentList {
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("cost"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("salvage"),
                                                },
                                            },
                                            Comma,
                                            Whitespace,
                                            Argument {
                                                IdentifierExpression {
                                                    Identifier ("life"),
                                                },
                                            },
                                        },
                                        RightParenthesis,
                                    },
                                    Whitespace,
                                    MultiplicationOperator,
                                    Whitespace,
                                    IdentifierExpression {
                                        Identifier ("years"),
                                    },
                                },
                                RightParenthesis,
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_monthly_calculation() {
        let source = r"
Sub Test()
    Dim monthlyDep As Double
    monthlyDep = SLN(cost, salvage, life) / 12
End Sub
";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("monthlyDep"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("monthlyDep"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        BinaryExpression {
                            CallExpression {
                                Identifier ("SLN"),
                                LeftParenthesis,
                                ArgumentList {
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("cost"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("salvage"),
                                        },
                                    },
                                    Comma,
                                    Whitespace,
                                    Argument {
                                        IdentifierExpression {
                                            Identifier ("life"),
                                        },
                                    },
                                },
                                RightParenthesis,
                            },
                            Whitespace,
                            DivisionOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("12"),
                            },
                        },
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_error_handling() {
        let source = r#"
Sub Test()
    On Error Resume Next
    Dim d As Double
    d = SLN(cost, salvage, life)
    If Err.Number <> 0 Then
        MsgBox "Error"
    End If
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        ResumeKeyword,
                        Whitespace,
                        NextKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("d"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("d"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    IfStatement {
                        Whitespace,
                        IfKeyword,
                        Whitespace,
                        BinaryExpression {
                            MemberAccessExpression {
                                Identifier ("Err"),
                                PeriodOperator,
                                Identifier ("Number"),
                            },
                            Whitespace,
                            InequalityOperator,
                            Whitespace,
                            NumericLiteralExpression {
                                IntegerLiteral ("0"),
                            },
                        },
                        Whitespace,
                        ThenKeyword,
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("MsgBox"),
                                Whitespace,
                                StringLiteral ("\"Error\""),
                                Newline,
                            },
                            Whitespace,
                        },
                        EndKeyword,
                        Whitespace,
                        IfKeyword,
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_on_error_goto() {
        let source = r#"
Sub Test()
    On Error GoTo ErrorHandler
    Dim depValue As Double
    depValue = SLN(assetCost, assetSalvage, assetLife)
    Exit Sub
ErrorHandler:
    MsgBox "Error calculating depreciation"
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    OnErrorStatement {
                        Whitespace,
                        OnKeyword,
                        Whitespace,
                        ErrorKeyword,
                        Whitespace,
                        GotoKeyword,
                        Whitespace,
                        Identifier ("ErrorHandler"),
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("depValue"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("depValue"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("assetCost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("assetSalvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("assetLife"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    ExitStatement {
                        Whitespace,
                        ExitKeyword,
                        Whitespace,
                        SubKeyword,
                        Newline,
                    },
                    LabelStatement {
                        Identifier ("ErrorHandler"),
                        ColonOperator,
                        Newline,
                    },
                    Whitespace,
                    CallStatement {
                        Identifier ("MsgBox"),
                        Whitespace,
                        StringLiteral ("\"Error calculating depreciation\""),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }

    #[test]
    fn sln_schedule_generation() {
        let source = r#"
Sub Test()
    Dim i As Integer
    Dim annualDep As Double
    annualDep = SLN(cost, salvage, life)
    For i = 1 To life
        Debug.Print "Year " & i & ": " & annualDep
    Next i
End Sub
"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        assert_tree!(cst, [
            Newline,
            SubStatement {
                SubKeyword,
                Whitespace,
                Identifier ("Test"),
                ParameterList {
                    LeftParenthesis,
                    RightParenthesis,
                },
                Newline,
                StatementList {
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        IntegerKeyword,
                        Newline,
                    },
                    Whitespace,
                    DimStatement {
                        DimKeyword,
                        Whitespace,
                        Identifier ("annualDep"),
                        Whitespace,
                        AsKeyword,
                        Whitespace,
                        DoubleKeyword,
                        Newline,
                    },
                    Whitespace,
                    AssignmentStatement {
                        IdentifierExpression {
                            Identifier ("annualDep"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        CallExpression {
                            Identifier ("SLN"),
                            LeftParenthesis,
                            ArgumentList {
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("cost"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("salvage"),
                                    },
                                },
                                Comma,
                                Whitespace,
                                Argument {
                                    IdentifierExpression {
                                        Identifier ("life"),
                                    },
                                },
                            },
                            RightParenthesis,
                        },
                        Newline,
                    },
                    ForStatement {
                        Whitespace,
                        ForKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("i"),
                        },
                        Whitespace,
                        EqualityOperator,
                        Whitespace,
                        NumericLiteralExpression {
                            IntegerLiteral ("1"),
                        },
                        Whitespace,
                        ToKeyword,
                        Whitespace,
                        IdentifierExpression {
                            Identifier ("life"),
                        },
                        Newline,
                        StatementList {
                            Whitespace,
                            CallStatement {
                                Identifier ("Debug"),
                                PeriodOperator,
                                PrintKeyword,
                                Whitespace,
                                StringLiteral ("\"Year \""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("i"),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                StringLiteral ("\": \""),
                                Whitespace,
                                Ampersand,
                                Whitespace,
                                Identifier ("annualDep"),
                                Newline,
                            },
                            Whitespace,
                        },
                        NextKeyword,
                        Whitespace,
                        Identifier ("i"),
                        Newline,
                    },
                },
                EndKeyword,
                Whitespace,
                SubKeyword,
                Newline,
            },
        ]);
    }
}

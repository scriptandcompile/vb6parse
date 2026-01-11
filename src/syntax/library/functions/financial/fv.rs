//! `Fv` Function
//!
//! Returns a Double specifying the future value of an annuity based on periodic, fixed payments and a fixed interest rate.
//!
//! # Syntax
//!
//! ```vb
//! Fv(rate, nper, pmt[, pv[, type]])
//! ```
//!
//! # Parameters
//!
//! - `rate` - Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
//! - `nper` - Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a four-year car loan, your loan has a total of 4 * 12 (or 48) payment periods.
//! - `pmt` - Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
//! - `pv` - Optional. Variant specifying present value (or lump sum) of a series of future payments. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make. If omitted, 0 is assumed.
//! - `type` - Optional. Variant specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.
//!
//! # Return Value
//!
//! Returns a Double representing the future value of an annuity based on periodic, fixed payments and a fixed interest rate.
//!
//! # Remarks
//!
//! - An annuity is a series of fixed cash payments made over a period of time.
//! - An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
//! - The rate and nper arguments must be calculated using payment periods expressed in the same units.
//! - For example, if rate is calculated using months, nper must also be calculated using months.
//! - For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.
//! - `Fv` is related to the `Pv` function. `Fv` calculates what a series of payments will be worth in the future, while `Pv` calculates what a series of future payments is worth now.
//!
//! # Typical Uses
//!
//! - Calculating savings account balance after regular deposits
//! - Determining investment portfolio value after periodic contributions
//! - Computing retirement fund balance
//! - Estimating college savings fund growth
//! - Analyzing long-term investment returns
//! - Planning for future financial goals
//!
//! # Basic Usage Examples
//!
//! ```vb
//! ' Calculate future value of monthly savings
//! Dim monthlyDeposit As Double
//! Dim annualRate As Double
//! Dim years As Integer
//! Dim futureValue As Double
//!
//! monthlyDeposit = -100  ' $100 per month (negative because it's paid out)
//! annualRate = 0.06      ' 6% annual interest
//! years = 10
//!
//! ' Calculate future value
//! futureValue = Fv(annualRate / 12, years * 12, monthlyDeposit)
//! ' Returns approximately $16,387.93
//!
//! ' Calculate future value with initial deposit
//! Dim initialDeposit As Double
//! initialDeposit = -1000  ' $1000 initial deposit
//!
//! futureValue = Fv(annualRate / 12, years * 12, monthlyDeposit, initialDeposit)
//! ' Returns approximately $18,193.97
//!
//! ' Calculate with payments at beginning of period
//! futureValue = Fv(annualRate / 12, years * 12, monthlyDeposit, initialDeposit, 1)
//! ' Returns approximately $18,284.68
//!
//! ' Calculate investment growth with no regular payments
//! Dim lumpSum As Double
//! lumpSum = -5000  ' $5000 one-time investment
//!
//! futureValue = Fv(0.08 / 12, 5 * 12, 0, lumpSum)
//! ' Returns approximately $7,449.23 (compound interest on lump sum)
//! ```
//!
//! # Common Patterns
//!
//! ## 1. Monthly Savings Calculator
//!
//! ```vb
//! Function CalculateSavings(monthlyAmount As Double, years As Integer, _
//!                           annualRate As Double) As Double
//!     Dim monthlyRate As Double
//!     Dim periods As Integer
//!     
//!     monthlyRate = annualRate / 12
//!     periods = years * 12
//!     
//!     ' Negative because it's money paid out
//!     CalculateSavings = Fv(monthlyRate, periods, -monthlyAmount)
//! End Function
//!
//! ' Usage
//! Dim savings As Double
//! savings = CalculateSavings(200, 20, 0.07)  ' $200/month, 20 years, 7%
//! MsgBox "Future value: " & FormatCurrency(savings)
//! ```
//!
//! ## 2. Retirement Planning
//!
//! ```vb
//! Function CalculateRetirementFund(monthlyContribution As Double, _
//!                                  currentAge As Integer, _
//!                                  retirementAge As Integer, _
//!                                  currentBalance As Double, _
//!                                  expectedReturn As Double) As Double
//!     Dim years As Integer
//!     Dim periods As Integer
//!     Dim monthlyRate As Double
//!     
//!     years = retirementAge - currentAge
//!     periods = years * 12
//!     monthlyRate = expectedReturn / 12
//!     
//!     CalculateRetirementFund = Fv(monthlyRate, periods, _
//!                                  -monthlyContribution, _
//!                                  -currentBalance)
//! End Function
//!
//! ' Usage
//! Dim retirementValue As Double
//! retirementValue = CalculateRetirementFund(500, 30, 65, 10000, 0.08)
//! Debug.Print "Retirement fund at 65: " & FormatCurrency(retirementValue)
//! ```
//!
//! ## 3. College Savings Plan
//!
//! ```vb
//! Function CollegeSavingsPlan(yearsUntilCollege As Integer, _
//!                             monthlyDeposit As Double, _
//!                             initialAmount As Double, _
//!                             expectedRate As Double) As Double
//!     Dim monthlyRate As Double
//!     Dim periods As Integer
//!     
//!     monthlyRate = expectedRate / 12
//!     periods = yearsUntilCollege * 12
//!     
//!     CollegeSavingsPlan = Fv(monthlyRate, periods, _
//!                             -monthlyDeposit, _
//!                             -initialAmount)
//! End Function
//!
//! ' Usage
//! Dim collegeFund As Double
//! collegeFund = CollegeSavingsPlan(18, 250, 5000, 0.06)
//! MsgBox "College fund in 18 years: " & FormatCurrency(collegeFund)
//! ```
//!
//! ## 4. Investment Comparison
//!
//! ```vb
//! Sub CompareInvestments()
//!     Dim option1 As Double
//!     Dim option2 As Double
//!     Dim years As Integer
//!     
//!     years = 10
//!     
//!     ' Option 1: $100/month at 6%
//!     option1 = Fv(0.06 / 12, years * 12, -100)
//!     
//!     ' Option 2: $50/month at 8%
//!     option2 = Fv(0.08 / 12, years * 12, -50)
//!     
//!     Debug.Print "Option 1 (6%): " & FormatCurrency(option1)
//!     Debug.Print "Option 2 (8%): " & FormatCurrency(option2)
//!     
//!     If option1 > option2 Then
//!         Debug.Print "Option 1 is better"
//!     Else
//!         Debug.Print "Option 2 is better"
//!     End If
//! End Sub
//! ```
//!
//! ## 5. Compound Interest Calculator
//!
//! ```vb
//! Function CompoundInterest(principal As Double, rate As Double, _
//!                          years As Integer, _
//!                          Optional compoundFrequency As Integer = 12) As Double
//!     Dim periods As Integer
//!     Dim periodRate As Double
//!     
//!     periods = years * compoundFrequency
//!     periodRate = rate / compoundFrequency
//!     
//!     ' No periodic payment, just compound the principal
//!     CompoundInterest = Fv(periodRate, periods, 0, -principal)
//! End Function
//!
//! ' Usage
//! Dim finalAmount As Double
//! finalAmount = CompoundInterest(10000, 0.05, 10, 12)  ' Monthly compounding
//! Debug.Print "Principal grows to: " & FormatCurrency(finalAmount)
//! ```
//!
//! ## 6. Savings Goal Calculator
//!
//! ```vb
//! Function MonthlyDepositNeeded(targetAmount As Double, _
//!                               years As Integer, _
//!                               rate As Double, _
//!                               Optional startingBalance As Double = 0) As Double
//!     ' This is the inverse - given FV, find PMT
//!     ' Using trial and error or formula
//!     Dim monthlyRate As Double
//!     Dim periods As Integer
//!     
//!     monthlyRate = rate / 12
//!     periods = years * 12
//!     
//!     ' Use Pmt function instead for accurate calculation
//!     ' This example shows how Fv relates to the goal
//!     Dim testPayment As Double
//!     testPayment = 100
//!     
//!     Do While Fv(monthlyRate, periods, -testPayment, -startingBalance) < targetAmount
//!         testPayment = testPayment + 10
//!     Loop
//!     
//!     MonthlyDepositNeeded = testPayment
//! End Function
//! ```
//!
//! ## 7. Annuity Future Value
//!
//! ```vb
//! Function AnnuityFutureValue(payment As Double, rate As Double, _
//!                             years As Integer, _
//!                             paymentTiming As Integer) As Double
//!     ' paymentTiming: 0 = end of period, 1 = beginning of period
//!     Dim periods As Integer
//!     
//!     periods = years
//!     AnnuityFutureValue = Fv(rate, periods, -payment, 0, paymentTiming)
//! End Function
//!
//! ' Usage
//! Dim fvOrdinary As Double
//! Dim fvDue As Double
//!
//! fvOrdinary = AnnuityFutureValue(1000, 0.05, 10, 0)  ' Ordinary annuity
//! fvDue = AnnuityFutureValue(1000, 0.05, 10, 1)       ' Annuity due
//!
//! Debug.Print "Ordinary annuity FV: " & FormatCurrency(fvOrdinary)
//! Debug.Print "Annuity due FV: " & FormatCurrency(fvDue)
//! ```
//!
//! ## 8. Investment Portfolio Projection
//!
//! ```vb
//! Type PortfolioProjection
//!     Years As Integer
//!     FutureValue As Double
//! End Type
//!
//! Function ProjectPortfolio(monthlyDeposit As Double, _
//!                          startingBalance As Double, _
//!                          rate As Double, _
//!                          maxYears As Integer) As PortfolioProjection()
//!     Dim projections() As PortfolioProjection
//!     Dim i As Integer
//!     Dim monthlyRate As Double
//!     
//!     ReDim projections(1 To maxYears)
//!     monthlyRate = rate / 12
//!     
//!     For i = 1 To maxYears
//!         projections(i).Years = i
//!         projections(i).FutureValue = Fv(monthlyRate, i * 12, _
//!                                         -monthlyDeposit, _
//!                                         -startingBalance)
//!     Next i
//!     
//!     ProjectPortfolio = projections
//! End Function
//! ```
//!
//! ## 9. Loan Payoff Calculator (Inverse Use)
//!
//! ```vb
//! Sub AnalyzeLoanPayoff()
//!     Dim loanAmount As Double
//!     Dim monthlyPayment As Double
//!     Dim annualRate As Double
//!     Dim years As Integer
//!     Dim remainingBalance As Double
//!     
//!     loanAmount = 200000     ' Initial loan
//!     monthlyPayment = 1200   ' Monthly payment
//!     annualRate = 0.045      ' 4.5% APR
//!     years = 5               ' After 5 years
//!     
//!     ' Future value will be negative (debt remaining)
//!     remainingBalance = -Fv(annualRate / 12, years * 12, _
//!                           monthlyPayment, -loanAmount)
//!     
//!     Debug.Print "Remaining balance after " & years & " years: " & _
//!                 FormatCurrency(remainingBalance)
//! End Sub
//! ```
//!
//! ## 10. Recurring Deposit Calculator
//!
//! ```vb
//! Sub RecurringDepositCalculator()
//!     Dim deposit As Double
//!     Dim rate As Double
//!     Dim quarters As Integer
//!     Dim maturityValue As Double
//!     
//!     deposit = 500           ' Quarterly deposit
//!     rate = 0.06 / 4        ' Quarterly rate (6% annual)
//!     quarters = 20          ' 5 years
//!     
//!     maturityValue = Fv(rate, quarters, -deposit)
//!     
//!     Debug.Print "Maturity value: " & FormatCurrency(maturityValue)
//!     Debug.Print "Total deposits: " & FormatCurrency(deposit * quarters)
//!     Debug.Print "Interest earned: " & _
//!                 FormatCurrency(maturityValue - (deposit * quarters))
//! End Sub
//! ```
//!
//! # Advanced Usage
//!
//! ## 1. Flexible Savings Calculator with UI
//!
//! ```vb
//! Sub CalculateAndDisplay()
//!     Dim monthlyDeposit As Double
//!     Dim years As Integer
//!     Dim annualRate As Double
//!     Dim initialBalance As Double
//!     Dim paymentType As Integer
//!     Dim futureValue As Double
//!     
//!     ' Get inputs from form controls
//!     monthlyDeposit = CDbl(txtMonthlyDeposit.Text)
//!     years = CInt(txtYears.Text)
//!     annualRate = CDbl(txtRate.Text) / 100
//!     initialBalance = CDbl(txtInitialBalance.Text)
//!     
//!     ' Check if payments at beginning or end
//!     paymentType = IIf(chkBeginning.Value = 1, 1, 0)
//!     
//!     ' Calculate
//!     futureValue = Fv(annualRate / 12, years * 12, _
//!                     -monthlyDeposit, _
//!                     -initialBalance, _
//!                     paymentType)
//!     
//!     ' Display result
//!     lblResult.Caption = "Future Value: " & FormatCurrency(futureValue, 2)
//!     
//!     ' Calculate total contributions
//!     Dim totalContributions As Double
//!     totalContributions = initialBalance + (monthlyDeposit * years * 12)
//!     
//!     ' Calculate interest earned
//!     Dim interestEarned As Double
//!     interestEarned = futureValue - totalContributions
//!     
//!     lblTotalDeposits.Caption = "Total Deposits: " & _
//!                                FormatCurrency(totalContributions, 2)
//!     lblInterest.Caption = "Interest Earned: " & _
//!                          FormatCurrency(interestEarned, 2)
//! End Sub
//! ```
//!
//! ## 2. Scenario Analysis
//!
//! ```vb
//! Sub AnalyzeScenarios()
//!     Dim rates() As Double
//!     Dim deposit As Double
//!     Dim years As Integer
//!     Dim i As Integer
//!     
//!     deposit = 300
//!     years = 15
//!     rates = Array(0.04, 0.06, 0.08, 0.10)  ' Different return scenarios
//!     
//!     Debug.Print "Scenario Analysis for $" & deposit & "/month over " & years & " years:"
//!     Debug.Print String(60, "=")
//!     
//!     For i = LBound(rates) To UBound(rates)
//!         Dim fv As Double
//!         fv = Fv(rates(i) / 12, years * 12, -deposit)
//!         
//!         Debug.Print "At " & FormatPercent(rates(i), 0) & ": " & _
//!                    FormatCurrency(fv, 2) & " (gain: " & _
//!                    FormatCurrency(fv - (deposit * years * 12), 2) & ")"
//!     Next i
//! End Sub
//! ```
//!
//! ## 3. Goal-Based Planning
//!
//! ```vb
//! Function YearsToReachGoal(targetAmount As Double, _
//!                           monthlyDeposit As Double, _
//!                           startingBalance As Double, _
//!                           annualRate As Double) As Double
//!     Dim years As Double
//!     Dim fv As Double
//!     Dim monthlyRate As Double
//!     
//!     monthlyRate = annualRate / 12
//!     years = 1
//!     
//!     Do While years <= 50  ' Max 50 years
//!         fv = Fv(monthlyRate, years * 12, -monthlyDeposit, -startingBalance)
//!         
//!         If fv >= targetAmount Then
//!             YearsToReachGoal = years
//!             Exit Function
//!         End If
//!         
//!         years = years + 0.25  ' Check quarterly
//!     Loop
//!     
//!     YearsToReachGoal = -1  ' Goal not reachable
//! End Function
//!
//! ' Usage
//! Dim yearsNeeded As Double
//! yearsNeeded = YearsToReachGoal(500000, 1000, 50000, 0.07)
//!
//! If yearsNeeded > 0 Then
//!     MsgBox "You will reach your goal in " & Format(yearsNeeded, "0.0") & " years"
//! Else
//!     MsgBox "Goal not reachable with current parameters"
//! End If
//! ```
//!
//! ## 4. Monte Carlo Simulation
//!
//! ```vb
//! Function SimulateFutureValue(deposit As Double, years As Integer, _
//!                              avgRate As Double, volatility As Double, _
//!                              simulations As Integer) As Variant
//!     Dim results() As Double
//!     Dim i As Integer
//!     Dim simulatedRate As Double
//!     
//!     ReDim results(1 To simulations)
//!     Randomize
//!     
//!     For i = 1 To simulations
//!         ' Simple random variation around average rate
//!         simulatedRate = avgRate + ((Rnd() - 0.5) * 2 * volatility)
//!         
//!         ' Ensure rate doesn't go negative
//!         If simulatedRate < 0 Then simulatedRate = 0
//!         
//!         results(i) = Fv(simulatedRate / 12, years * 12, -deposit)
//!     Next i
//!     
//!     SimulateFutureValue = results
//! End Function
//!
//! ' Analyze results
//! Sub AnalyzeSimulation()
//!     Dim results As Variant
//!     Dim avg As Double, minVal As Double, maxVal As Double
//!     Dim i As Integer
//!     
//!     results = SimulateFutureValue(500, 20, 0.07, 0.02, 1000)
//!     
//!     avg = 0
//!     minVal = 1E+308
//!     maxVal = -1E+308
//!     
//!     For i = 1 To UBound(results)
//!         avg = avg + results(i)
//!         If results(i) < minVal Then minVal = results(i)
//!         If results(i) > maxVal Then maxVal = results(i)
//!     Next i
//!     
//!     avg = avg / UBound(results)
//!     
//!     Debug.Print "Average FV: " & FormatCurrency(avg, 2)
//!     Debug.Print "Min FV: " & FormatCurrency(minVal, 2)
//!     Debug.Print "Max FV: " & FormatCurrency(maxVal, 2)
//! End Sub
//! ```
//!
//! ## 5. Tax-Advantaged Account Calculator
//!
//! ```vb
//! Function TaxAdvantaged401k(salary As Double, contributionPct As Double, _
//!                            employerMatch As Double, years As Integer, _
//!                            currentBalance As Double, rate As Double) As Double
//!     Dim monthlyContribution As Double
//!     Dim monthlyEmployerMatch As Double
//!     Dim totalMonthlyDeposit As Double
//!     
//!     ' Calculate monthly contributions
//!     monthlyContribution = (salary * contributionPct) / 12
//!     monthlyEmployerMatch = (salary * employerMatch) / 12
//!     totalMonthlyDeposit = monthlyContribution + monthlyEmployerMatch
//!     
//!     TaxAdvantaged401k = Fv(rate / 12, years * 12, _
//!                            -totalMonthlyDeposit, _
//!                            -currentBalance)
//! End Function
//!
//! ' Usage
//! Dim retirement401k As Double
//! retirement401k = TaxAdvantaged401k(75000, 0.06, 0.03, 30, 25000, 0.08)
//! Debug.Print "401(k) at retirement: " & FormatCurrency(retirement401k, 2)
//! ```
//!
//! ## 6. Education Savings with Increasing Contributions
//!
//! ```vb
//! Function EducationSavingsWithIncrease(initialMonthly As Double, _
//!                                       annualIncrease As Double, _
//!                                       years As Integer, _
//!                                       rate As Double) As Double
//!     Dim yearlyFV As Double
//!     Dim currentMonthly As Double
//!     Dim i As Integer
//!     
//!     yearlyFV = 0
//!     currentMonthly = initialMonthly
//!     
//!     For i = 1 To years
//!         ' Calculate FV for this year's contributions
//!         Dim yearContribution As Double
//!         yearContribution = Fv(rate / 12, (years - i + 1) * 12, _
//!                              -currentMonthly)
//!         
//!         yearlyFV = yearlyFV + yearContribution
//!         
//!         ' Increase for next year
//!         currentMonthly = currentMonthly * (1 + annualIncrease)
//!     Next i
//!     
//!     EducationSavingsWithIncrease = yearlyFV
//! End Function
//! ```
//!
//! # Error Handling
//!
//! ```vb
//! Function SafeFv(rate As Double, nper As Integer, pmt As Double, _
//!                 Optional pv As Variant, Optional pType As Variant) As Variant
//!     On Error GoTo ErrorHandler
//!     
//!     ' Validate inputs
//!     If nper <= 0 Then
//!         SafeFv = "Error: Number of periods must be positive"
//!         Exit Function
//!     End If
//!     
//!     If rate <= -1 Then
//!         SafeFv = "Error: Rate must be greater than -100%"
//!         Exit Function
//!     End If
//!     
//!     ' Set defaults
//!     If IsMissing(pv) Then pv = 0
//!     If IsMissing(pType) Then pType = 0
//!     
//!     ' Calculate
//!     SafeFv = Fv(rate, nper, pmt, pv, pType)
//!     Exit Function
//!     
//! ErrorHandler:
//!     Select Case Err.Number
//!         Case 5  ' Invalid procedure call
//!             SafeFv = "Error: Invalid arguments"
//!         Case 6  ' Overflow
//!             SafeFv = "Error: Result too large"
//!         Case 13  ' Type mismatch
//!             SafeFv = "Error: Invalid data types"
//!         Case Else
//!             SafeFv = "Error: " & Err.Description
//!     End Select
//! End Function
//! ```
//!
//! Common errors:
//! - **Error 5 (Invalid procedure call)**: Invalid argument values (e.g., negative periods).
//! - **Error 6 (Overflow)**: Result is too large to fit in a Double.
//! - **Error 13 (Type mismatch)**: Arguments are not numeric.
//!
//! # Performance Considerations
//!
//! - `Fv` is a mathematical calculation, very fast
//! - No I/O or external dependencies
//! - Safe to call repeatedly in loops for scenario analysis
//! - Consider caching results if using same parameters multiple times
//! - For large-scale simulations, consider batch calculations
//!
//! # Best Practices
//!
//! 1. **Use negative values for cash outflows** (payments, deposits)
//! 2. **Use positive values for cash inflows** (receipts, withdrawals)
//! 3. **Match time units** - if rate is monthly, nper should be in months
//! 4. **Validate inputs** - check for reasonable ranges
//! 5. **Handle edge cases** - zero rate, very long periods
//! 6. **Document assumptions** - especially for rate projections
//! 7. **Consider inflation** - future value in today's dollars may differ
//!
//! # Comparison with Other Functions
//!
//! ## `Fv` vs `Pv`
//!
//! ```vb
//! ' Fv: Future value of an investment
//! futureValue = Fv(0.06 / 12, 10 * 12, -100)  ' What will I have?
//!
//! ' Pv: Present value of an investment
//! presentValue = Pv(0.06 / 12, 10 * 12, -100)  ' What is it worth today?
//! ```
//!
//! ## `Fv` vs `NPer`
//!
//! ```vb
//! ' Fv: Calculate future value given payments
//! fv = Fv(0.05 / 12, 120, -100)
//!
//! ' NPer: Calculate periods needed to reach a goal
//! periods = NPer(0.05 / 12, -100, 0, 16000)  ' How long to reach $16,000?
//! ```
//!
//! ## Fv vs Pmt
//!
//! ```vb
//! ' Fv: Calculate future value given payment amount
//! fv = Fv(0.06 / 12, 120, -200)
//!
//! ' Pmt: Calculate payment needed to reach future value
//! payment = Pmt(0.06 / 12, 120, 0, -30000)  ' How much to save for $30,000?
//! ```
//!
//! # Limitations
//!
//! - Assumes constant interest rate (real-world rates vary)
//! - Assumes regular, fixed payments (life is rarely this predictable)
//! - Does not account for taxes, fees, or inflation
//! - Does not consider compounding frequency variations
//! - Limited to Double precision (very large values may overflow)
//! - No built-in risk or uncertainty modeling
//!
//! # Mathematical Formula
//!
//! The future value calculation uses the formula:
//!
//! ```text
//! For type = 0 (payments at end of period):
//! FV = -PV * (1 + rate)^nper - PMT * [((1 + rate)^nper - 1) / rate]
//!
//! For type = 1 (payments at beginning of period):
//! FV = -PV * (1 + rate)^nper - PMT * [((1 + rate)^nper - 1) / rate] * (1 + rate)
//!
//! Special case when rate = 0:
//! FV = -PV - PMT * nper
//! ```
//!
//! # Related Functions
//!
//! - `Pv` - Returns the present value of an annuity
//! - `Pmt` - Returns the payment for an annuity
//! - `PPmt` - Returns the principal payment for a specific period
//! - `IPmt` - Returns the interest payment for a specific period
//! - `NPer` - Returns the number of periods for an annuity
//! - `Rate` - Returns the interest rate per period
//! - `DDB` - Returns depreciation using double-declining balance
//! - `SLN` - Returns straight-line depreciation
//! - `SYD` - Returns sum-of-years' digits depreciation

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn fv_basic() {
        let source = r"result = Fv(0.06 / 12, 120, -100)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings
            .set_snapshot_path("../../../../../snapshots/syntax/library/functions/financial/fv");
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_with_pv() {
        let source = r"futureValue = Fv(rate / 12, years * 12, monthlyDeposit, initialDeposit)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_all_parameters() {
        let source =
            r"futureValue = Fv(annualRate / 12, years * 12, monthlyDeposit, initialDeposit, 1)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_no_payment() {
        let source = r"futureValue = Fv(0.08 / 12, 5 * 12, 0, lumpSum)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_in_function() {
        let source = r"Function CalculateSavings() As Double
    CalculateSavings = Fv(monthlyRate, periods, -monthlyAmount)
End Function";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_retirement() {
        let source =
            r"retirementValue = Fv(monthlyRate, periods, -monthlyContribution, -currentBalance)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_debug_print() {
        let source = r#"Debug.Print "Future value: " & Fv(0.05 / 12, 10 * 12, -200)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_comparison() {
        let source =
            r#"If Fv(0.06 / 12, years * 12, -100) > targetAmount Then MsgBox "Goal reached""#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_assignment() {
        let source = r"Dim fv As Double
fv = Fv(rate, periods, payment)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_in_loop() {
        let source = r"For i = 1 To maxYears
    projections(i).FutureValue = Fv(monthlyRate, i * 12, -monthlyDeposit, -startingBalance)
Next i";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_negative_result() {
        let source = r"balance = Fv(rate / 12, years * 12, payment, -principal)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_formatcurrency() {
        let source = r#"lblResult.Caption = "Future Value: " & FormatCurrency(Fv(rate, periods, payment), 2)"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_array_assignment() {
        let source = r"results(i) = Fv(simulatedRate / 12, years * 12, -deposit)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_error_handling() {
        let source = r"On Error GoTo ErrorHandler
fv = Fv(rate, nper, pmt, pv, pType)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_do_while() {
        let source = r"Do While Fv(monthlyRate, periods, -testPayment, -startingBalance) < targetAmount
    testPayment = testPayment + 10
Loop";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_select_case() {
        let source = r"Select Case investment
    Case 1
        result = Fv(0.06 / 12, years * 12, -100)
End Select";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_iif() {
        let source = r"fv = IIf(useHighRate, Fv(0.08 / 12, periods, -deposit), Fv(0.05 / 12, periods, -deposit))";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_msgbox() {
        let source =
            r#"MsgBox "Your savings will grow to " & FormatCurrency(Fv(rate, periods, -deposit))"#;
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_calculation() {
        let source = r"totalValue = Fv(rate, periods, payment) + bonus";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_type_member() {
        let source = r"investment.FutureValue = Fv(rate, periods, payment, principal)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_subtraction() {
        let source = r"interestEarned = Fv(rate, periods, payment) - totalContributions";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_division() {
        let source = r"monthlyEquivalent = Fv(rate / 12, periods * 12, payment / 12)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_with_cdbl() {
        let source = r"result = Fv(CDbl(txtRate.Text) / 100 / 12, CInt(txtYears.Text) * 12, -CDbl(txtDeposit.Text))";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_quarterly() {
        let source = r"maturityValue = Fv(rate, quarters, -deposit)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_annual() {
        let source = r"fvOrdinary = Fv(rate, periods, -payment, 0, paymentTiming)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }

    #[test]
    fn fv_compound_interest() {
        let source = r"finalAmount = Fv(periodRate, periods, 0, -principal)";
        let (cst_opt, _failures) = ConcreteSyntaxTree::from_text("test.bas", source).unpack();
        let cst = cst_opt.expect("CST should be parsed");

        let tree = cst.to_serializable();

        let mut settings = insta::Settings::clone_current();
        settings.set_snapshot_path(
            "../../../../../snapshots/syntax/syntax/library/functions/financial/fv",
        );
        settings.set_prepend_module_to_snapshot(false);
        let _guard = settings.bind_to_scope();
        insta::assert_yaml_snapshot!(tree);
    }
}

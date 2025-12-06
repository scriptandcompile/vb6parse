//! # `NPer` Function
//!
//! Returns a Double specifying the number of periods for an annuity based on periodic, fixed payments and a fixed interest rate.
//!
//! ## Syntax
//!
//! ```vb
//! NPer(rate, pmt, pv, [fv], [type])
//! ```
//!
//! ## Parameters
//!
//! - **rate** (Required) - Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10 percent and make monthly payments, the rate per period is 0.1/12, or 0.0083.
//! - **pmt** (Required) - Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
//! - **pv** (Required) - Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.
//! - **fv** (Optional) - Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.
//! - **type** (Optional) - Variant specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.
//!
//! ## Return Value
//!
//! Returns a **Double** specifying the number of payment periods in an annuity.
//!
//! ## Remarks
//!
//! The `NPer` function calculates how many periods (usually months or years) it will take to pay off a loan or reach a savings goal, given a fixed payment amount and interest rate.
//!
//! ### Key Characteristics:
//! - Returns number of periods (can be fractional)
//! - Assumes constant payment amounts
//! - Assumes constant interest rate
//! - For loans, pv is positive (amount borrowed), pmt is negative (payment out)
//! - For investments, pv is negative (deposit), fv is positive (goal)
//! - Type parameter: 0 = end of period (default), 1 = beginning of period
//! - Rate must be expressed per period (annual rate / periods per year)
//! - Commonly used with PMT, PV, FV, and Rate functions
//!
//! ### Common Use Cases:
//! - Calculate loan payoff time
//! - Determine how long to reach savings goal
//! - Plan retirement savings duration
//! - Calculate time to pay off credit card debt
//! - Determine mortgage term needed
//! - Investment planning timelines
//! - Debt consolidation analysis
//! - Education savings planning
//!
//! ## Typical Uses
//!
//! 1. **Loan Payoff Time** - Calculate how many months to pay off a loan
//! 2. **Savings Goal Timeline** - Determine time to reach a savings target
//! 3. **Mortgage Planning** - Calculate term needed for specific payments
//! 4. **Credit Card Payoff** - Estimate months to pay off credit card balance
//! 5. **Investment Duration** - Calculate time to reach investment goal
//! 6. **Retirement Planning** - Determine years needed to save for retirement
//! 7. **Debt Analysis** - Compare payoff timelines for different payment amounts
//! 8. **What-If Scenarios** - Model different payment scenarios
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: How many months to pay off a $10,000 loan at 8% APR with $200/month payments?
//! Dim months As Double
//! months = NPer(0.08 / 12, -200, 10000)
//! ' Result: approximately 62.6 months (5.2 years)
//! ```
//!
//! ```vb
//! ' Example 2: How many years to save $50,000 with $300/month at 6% annual return?
//! Dim years As Double
//! years = NPer(0.06 / 12, -300, 0, 50000) / 12
//! ' Result: approximately 10.8 years
//! ```
//!
//! ```vb
//! ' Example 3: Time to pay off credit card with minimum payments
//! Dim balance As Double
//! Dim monthlyPayment As Double
//! Dim apr As Double
//! Dim payoffMonths As Double
//!
//! balance = 5000
//! monthlyPayment = -100
//! apr = 0.1899 ' 18.99% APR
//! payoffMonths = NPer(apr / 12, monthlyPayment, balance)
//! ' Result: approximately 79 months
//! ```
//!
//! ```vb
//! ' Example 4: Retirement savings timeline
//! Dim monthlyDeposit As Double
//! Dim retirementGoal As Double
//! Dim annualReturn As Double
//! Dim yearsNeeded As Double
//!
//! monthlyDeposit = -500
//! retirementGoal = 1000000
//! annualReturn = 0.07
//! yearsNeeded = NPer(annualReturn / 12, monthlyDeposit, 0, retirementGoal) / 12
//! ' Result: approximately 38.3 years
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Loan payoff calculator
//! Function CalculatePayoffMonths(loanAmount As Double, _
//!                                monthlyPayment As Double, _
//!                                apr As Double) As Double
//!     Dim monthlyRate As Double
//!     monthlyRate = apr / 12
//!     CalculatePayoffMonths = NPer(monthlyRate, -monthlyPayment, loanAmount)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Savings goal timeline
//! Function YearsToReachGoal(monthlyDeposit As Double, _
//!                          currentBalance As Double, _
//!                          targetAmount As Double, _
//!                          annualReturn As Double) As Double
//!     Dim months As Double
//!     months = NPer(annualReturn / 12, -monthlyDeposit, -currentBalance, targetAmount)
//!     YearsToReachGoal = months / 12
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 3: Compare payment scenarios
//! Sub ComparePaymentScenarios(balance As Double, apr As Double)
//!     Dim payment As Double
//!     Dim months As Double
//!     Dim i As Integer
//!     
//!     For i = 1 To 5
//!         payment = balance * 0.02 * i ' 2%, 4%, 6%, 8%, 10% of balance
//!         months = NPer(apr / 12, -payment, balance)
//!         Debug.Print "Payment: $" & Format(payment, "0.00") & _
//!                     " - Months: " & Format(months, "0.0")
//!     Next i
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 4: Interest rate impact on duration
//! Function CalculateRateImpact(loanAmount As Double, _
//!                             monthlyPayment As Double) As String
//!     Dim result As String
//!     Dim rate As Double
//!     Dim months As Double
//!     
//!     result = "Interest Rate Impact:" & vbCrLf
//!     
//!     For rate = 0.03 To 0.15 Step 0.02
//!         months = NPer(rate / 12, -monthlyPayment, loanAmount)
//!         result = result & Format(rate * 100, "0.0") & "% APR: " & _
//!                  Format(months, "0.0") & " months" & vbCrLf
//!     Next rate
//!     
//!     CalculateRateImpact = result
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 5: Minimum payment warning
//! Function IsMinimumPaymentTooLow(balance As Double, _
//!                                 payment As Double, _
//!                                 apr As Double) As Boolean
//!     Dim months As Double
//!     Dim maxYears As Double
//!     
//!     maxYears = 10
//!     months = NPer(apr / 12, -payment, balance)
//!     
//!     IsMinimumPaymentTooLow = (months > maxYears * 12)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Early payoff calculation
//! Function MonthsSavedByExtraPayment(balance As Double, _
//!                                    regularPayment As Double, _
//!                                    extraPayment As Double, _
//!                                    apr As Double) As Double
//!     Dim regularMonths As Double
//!     Dim acceleratedMonths As Double
//!     
//!     regularMonths = NPer(apr / 12, -regularPayment, balance)
//!     acceleratedMonths = NPer(apr / 12, -(regularPayment + extraPayment), balance)
//!     
//!     MonthsSavedByExtraPayment = regularMonths - acceleratedMonths
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 7: Formatted duration display
//! Function FormatPayoffTime(months As Double) As String
//!     Dim years As Long
//!     Dim remainingMonths As Long
//!     
//!     years = Int(months / 12)
//!     remainingMonths = Int(months Mod 12)
//!     
//!     If years > 0 Then
//!         FormatPayoffTime = years & " year"
//!         If years > 1 Then FormatPayoffTime = FormatPayoffTime & "s"
//!         If remainingMonths > 0 Then
//!             FormatPayoffTime = FormatPayoffTime & ", " & remainingMonths & " month"
//!             If remainingMonths > 1 Then FormatPayoffTime = FormatPayoffTime & "s"
//!         End If
//!     Else
//!         FormatPayoffTime = remainingMonths & " month"
//!         If remainingMonths <> 1 Then FormatPayoffTime = FormatPayoffTime & "s"
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: College savings timeline
//! Function CalculateCollegeSavingsTime(currentAge As Integer, _
//!                                      currentSavings As Double, _
//!                                      monthlyDeposit As Double, _
//!                                      collegeGoal As Double, _
//!                                      annualReturn As Double) As Double
//!     Dim months As Double
//!     Dim yearsUntilCollege As Integer
//!     
//!     months = NPer(annualReturn / 12, -monthlyDeposit, -currentSavings, collegeGoal)
//!     yearsUntilCollege = 18 - currentAge
//!     
//!     CalculateCollegeSavingsTime = months / 12
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Debt consolidation comparison
//! Function CompareConsolidation(debt1 As Double, rate1 As Double, pmt1 As Double, _
//!                              debt2 As Double, rate2 As Double, pmt2 As Double, _
//!                              consolidatedRate As Double) As String
//!     Dim currentMonths As Double
//!     Dim consolidatedMonths As Double
//!     Dim totalDebt As Double
//!     Dim totalPayment As Double
//!     Dim result As String
//!     
//!     ' Calculate current payoff time
//!     currentMonths = Application.Max( _
//!         NPer(rate1 / 12, -pmt1, debt1), _
//!         NPer(rate2 / 12, -pmt2, debt2))
//!     
//!     ' Calculate consolidated payoff time
//!     totalDebt = debt1 + debt2
//!     totalPayment = pmt1 + pmt2
//!     consolidatedMonths = NPer(consolidatedRate / 12, -totalPayment, totalDebt)
//!     
//!     result = "Current: " & Format(currentMonths, "0.0") & " months" & vbCrLf
//!     result = result & "Consolidated: " & Format(consolidatedMonths, "0.0") & " months" & vbCrLf
//!     result = result & "Savings: " & Format(currentMonths - consolidatedMonths, "0.0") & " months"
//!     
//!     CompareConsolidation = result
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Payment sufficiency check
//! Function IsPaymentSufficient(balance As Double, _
//!                             payment As Double, _
//!                             apr As Double) As Boolean
//!     Dim monthlyInterest As Double
//!     
//!     ' Check if payment exceeds monthly interest
//!     monthlyInterest = balance * (apr / 12)
//!     
//!     If payment <= monthlyInterest Then
//!         IsPaymentSufficient = False ' Will never pay off
//!     Else
//!         On Error Resume Next
//!         Dim periods As Double
//!         periods = NPer(apr / 12, -payment, balance)
//!         IsPaymentSufficient = (Err.Number = 0)
//!         On Error GoTo 0
//!     End If
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Loan Analyzer Class
//!
//! ```vb
//! ' Class: LoanAnalyzer
//! ' Comprehensive loan analysis with payoff scenarios
//!
//! Option Explicit
//!
//! Private m_principal As Double
//! Private m_apr As Double
//! Private m_monthlyPayment As Double
//!
//! Public Sub Initialize(principal As Double, apr As Double, monthlyPayment As Double)
//!     m_principal = principal
//!     m_apr = apr
//!     m_monthlyPayment = monthlyPayment
//! End Sub
//!
//! Public Function GetPayoffMonths() As Double
//!     GetPayoffMonths = NPer(m_apr / 12, -m_monthlyPayment, m_principal)
//! End Function
//!
//! Public Function GetPayoffYears() As Double
//!     GetPayoffYears = GetPayoffMonths() / 12
//! End Function
//!
//! Public Function GetTotalInterest() As Double
//!     Dim months As Double
//!     Dim totalPaid As Double
//!     
//!     months = GetPayoffMonths()
//!     totalPaid = m_monthlyPayment * months
//!     GetTotalInterest = totalPaid - m_principal
//! End Function
//!
//! Public Function GetPayoffDate() As Date
//!     Dim months As Double
//!     months = GetPayoffMonths()
//!     GetPayoffDate = DateAdd("m", Int(months), Date)
//! End Function
//!
//! Public Function AnalyzeExtraPayment(extraMonthly As Double) As String
//!     Dim baseMonths As Double
//!     Dim acceleratedMonths As Double
//!     Dim monthsSaved As Double
//!     Dim interestSaved As Double
//!     Dim result As String
//!     
//!     baseMonths = NPer(m_apr / 12, -m_monthlyPayment, m_principal)
//!     acceleratedMonths = NPer(m_apr / 12, -(m_monthlyPayment + extraMonthly), m_principal)
//!     
//!     monthsSaved = baseMonths - acceleratedMonths
//!     interestSaved = (m_monthlyPayment * baseMonths - m_principal) - _
//!                     ((m_monthlyPayment + extraMonthly) * acceleratedMonths - m_principal)
//!     
//!     result = "Extra Payment Analysis:" & vbCrLf
//!     result = result & "Extra payment: $" & Format(extraMonthly, "#,##0.00") & vbCrLf
//!     result = result & "Time saved: " & Format(monthsSaved, "0.0") & " months" & vbCrLf
//!     result = result & "Interest saved: $" & Format(interestSaved, "#,##0.00")
//!     
//!     AnalyzeExtraPayment = result
//! End Function
//!
//! Public Function GenerateAmortizationSummary() As String
//!     Dim summary As String
//!     Dim months As Double
//!     Dim totalPaid As Double
//!     Dim totalInterest As Double
//!     
//!     months = GetPayoffMonths()
//!     totalPaid = m_monthlyPayment * months
//!     totalInterest = totalPaid - m_principal
//!     
//!     summary = "Loan Amortization Summary" & vbCrLf
//!     summary = summary & String(50, "-") & vbCrLf
//!     summary = summary & "Principal: $" & Format(m_principal, "#,##0.00") & vbCrLf
//!     summary = summary & "APR: " & Format(m_apr * 100, "0.00") & "%" & vbCrLf
//!     summary = summary & "Monthly Payment: $" & Format(m_monthlyPayment, "#,##0.00") & vbCrLf
//!     summary = summary & "Payoff Time: " & Format(months, "0.0") & " months (" & _
//!                        Format(months / 12, "0.0") & " years)" & vbCrLf
//!     summary = summary & "Total Paid: $" & Format(totalPaid, "#,##0.00") & vbCrLf
//!     summary = summary & "Total Interest: $" & Format(totalInterest, "#,##0.00") & vbCrLf
//!     summary = summary & "Payoff Date: " & Format(GetPayoffDate(), "mmm dd, yyyy")
//!     
//!     GenerateAmortizationSummary = summary
//! End Function
//! ```
//!
//! ### Example 2: Retirement Planner Class
//!
//! ```vb
//! ' Class: RetirementPlanner
//! ' Plans retirement savings timeline and scenarios
//!
//! Option Explicit
//!
//! Private m_currentAge As Integer
//! Private m_retirementAge As Integer
//! Private m_currentSavings As Double
//! Private m_monthlyContribution As Double
//! Private m_expectedReturn As Double
//! Private m_retirementGoal As Double
//!
//! Public Sub Initialize(currentAge As Integer, _
//!                      retirementAge As Integer, _
//!                      currentSavings As Double, _
//!                      monthlyContribution As Double, _
//!                      expectedReturn As Double, _
//!                      retirementGoal As Double)
//!     m_currentAge = currentAge
//!     m_retirementAge = retirementAge
//!     m_currentSavings = currentSavings
//!     m_monthlyContribution = monthlyContribution
//!     m_expectedReturn = expectedReturn
//!     m_retirementGoal = retirementGoal
//! End Sub
//!
//! Public Function GetYearsToGoal() As Double
//!     Dim months As Double
//!     months = NPer(m_expectedReturn / 12, -m_monthlyContribution, -m_currentSavings, m_retirementGoal)
//!     GetYearsToGoal = months / 12
//! End Function
//!
//! Public Function WillReachGoal() As Boolean
//!     Dim yearsNeeded As Double
//!     Dim yearsAvailable As Integer
//!     
//!     yearsNeeded = GetYearsToGoal()
//!     yearsAvailable = m_retirementAge - m_currentAge
//!     
//!     WillReachGoal = (yearsNeeded <= yearsAvailable)
//! End Function
//!
//! Public Function GetRequiredMonthlyContribution() As Double
//!     Dim monthsAvailable As Long
//!     monthsAvailable = (m_retirementAge - m_currentAge) * 12
//!     
//!     GetRequiredMonthlyContribution = -PMT(m_expectedReturn / 12, _
//!                                          monthsAvailable, _
//!                                          -m_currentSavings, _
//!                                          m_retirementGoal)
//! End Function
//!
//! Public Function GenerateScenarioAnalysis() As String
//!     Dim result As String
//!     Dim yearsNeeded As Double
//!     Dim yearsAvailable As Integer
//!     Dim shortfall As Double
//!     
//!     yearsNeeded = GetYearsToGoal()
//!     yearsAvailable = m_retirementAge - m_currentAge
//!     
//!     result = "Retirement Scenario Analysis" & vbCrLf
//!     result = result & String(50, "-") & vbCrLf
//!     result = result & "Current Age: " & m_currentAge & vbCrLf
//!     result = result & "Retirement Age: " & m_retirementAge & vbCrLf
//!     result = result & "Years Available: " & yearsAvailable & vbCrLf
//!     result = result & "Current Savings: $" & Format(m_currentSavings, "#,##0.00") & vbCrLf
//!     result = result & "Monthly Contribution: $" & Format(m_monthlyContribution, "#,##0.00") & vbCrLf
//!     result = result & "Expected Return: " & Format(m_expectedReturn * 100, "0.00") & "%" & vbCrLf
//!     result = result & "Retirement Goal: $" & Format(m_retirementGoal, "#,##0.00") & vbCrLf
//!     result = result & vbCrLf
//!     result = result & "Years Needed: " & Format(yearsNeeded, "0.0") & vbCrLf
//!     
//!     If WillReachGoal() Then
//!         result = result & "Status: On track! " & Format(yearsAvailable - yearsNeeded, "0.0") & _
//!                  " years ahead of schedule."
//!     Else
//!         shortfall = yearsNeeded - yearsAvailable
//!         result = result & "Status: Behind schedule by " & Format(shortfall, "0.0") & " years." & vbCrLf
//!         result = result & "Required Monthly: $" & _
//!                  Format(GetRequiredMonthlyContribution(), "#,##0.00")
//!     End If
//!     
//!     GenerateScenarioAnalysis = result
//! End Function
//! ```
//!
//! ### Example 3: Debt Payoff Optimizer Module
//!
//! ```vb
//! ' Module: DebtPayoffOptimizer
//! ' Optimizes debt payoff strategies
//!
//! Option Explicit
//!
//! Private Type DebtInfo
//!     name As String
//!     balance As Double
//!     apr As Double
//!     minimumPayment As Double
//!     payoffMonths As Double
//! End Type
//!
//! Public Function AnalyzeSnowballMethod(debts() As DebtInfo, _
//!                                       extraPayment As Double) As String
//!     Dim i As Integer
//!     Dim result As String
//!     Dim totalMonths As Double
//!     
//!     ' Sort debts by balance (smallest first)
//!     SortDebtsByBalance debts
//!     
//!     result = "Debt Snowball Method Analysis" & vbCrLf
//!     result = result & String(60, "-") & vbCrLf
//!     
//!     totalMonths = 0
//!     
//!     For i = LBound(debts) To UBound(debts)
//!         If i = LBound(debts) Then
//!             debts(i).payoffMonths = NPer(debts(i).apr / 12, _
//!                                          -(debts(i).minimumPayment + extraPayment), _
//!                                          debts(i).balance)
//!         Else
//!             ' Add freed-up payment from previous debt
//!             Dim availablePayment As Double
//!             availablePayment = debts(i).minimumPayment + extraPayment
//!             
//!             For j = LBound(debts) To i - 1
//!                 availablePayment = availablePayment + debts(j).minimumPayment
//!             Next j
//!             
//!             debts(i).payoffMonths = totalMonths + _
//!                 NPer(debts(i).apr / 12, -availablePayment, debts(i).balance)
//!         End If
//!         
//!         totalMonths = debts(i).payoffMonths
//!         
//!         result = result & debts(i).name & ": " & _
//!                  Format(debts(i).payoffMonths, "0.0") & " months" & vbCrLf
//!     Next i
//!     
//!     result = result & vbCrLf & "Total Time: " & Format(totalMonths, "0.0") & " months"
//!     
//!     AnalyzeSnowballMethod = result
//! End Function
//!
//! Public Function AnalyzeAvalancheMethod(debts() As DebtInfo, _
//!                                        extraPayment As Double) As String
//!     ' Sort debts by APR (highest first)
//!     SortDebtsByRate debts
//!     
//!     ' Use same logic as snowball but with different sort
//!     AnalyzeAvalancheMethod = AnalyzeSnowballMethod(debts, extraPayment)
//! End Function
//!
//! Private Sub SortDebtsByBalance(debts() As DebtInfo)
//!     ' Simple bubble sort
//!     Dim i As Integer, j As Integer
//!     Dim temp As DebtInfo
//!     
//!     For i = LBound(debts) To UBound(debts) - 1
//!         For j = i + 1 To UBound(debts)
//!             If debts(i).balance > debts(j).balance Then
//!                 temp = debts(i)
//!                 debts(i) = debts(j)
//!                 debts(j) = temp
//!             End If
//!         Next j
//!     Next i
//! End Sub
//!
//! Private Sub SortDebtsByRate(debts() As DebtInfo)
//!     ' Simple bubble sort
//!     Dim i As Integer, j As Integer
//!     Dim temp As DebtInfo
//!     
//!     For i = LBound(debts) To UBound(debts) - 1
//!         For j = i + 1 To UBound(debts)
//!             If debts(i).apr < debts(j).apr Then
//!                 temp = debts(i)
//!                 debts(i) = debts(j)
//!                 debts(j) = temp
//!             End If
//!         Next j
//!     Next i
//! End Sub
//! ```
//!
//! ### Example 4: Financial Goal Tracker
//!
//! ```vb
//! ' Class: FinancialGoalTracker
//! ' Tracks progress toward multiple financial goals
//!
//! Option Explicit
//!
//! Private Type Goal
//!     name As String
//!     targetAmount As Double
//!     currentAmount As Double
//!     monthlyDeposit As Double
//!     expectedReturn As Double
//!     targetDate As Date
//!     isOnTrack As Boolean
//!     monthsNeeded As Double
//! End Type
//!
//! Private m_goals As Collection
//!
//! Private Sub Class_Initialize()
//!     Set m_goals = New Collection
//! End Sub
//!
//! Public Sub AddGoal(name As String, _
//!                   targetAmount As Double, _
//!                   currentAmount As Double, _
//!                   monthlyDeposit As Double, _
//!                   expectedReturn As Double, _
//!                   targetDate As Date)
//!     Dim goal As Goal
//!     
//!     goal.name = name
//!     goal.targetAmount = targetAmount
//!     goal.currentAmount = currentAmount
//!     goal.monthlyDeposit = monthlyDeposit
//!     goal.expectedReturn = expectedReturn
//!     goal.targetDate = targetDate
//!     
//!     ' Calculate if on track
//!     goal.monthsNeeded = NPer(expectedReturn / 12, -monthlyDeposit, -currentAmount, targetAmount)
//!     
//!     Dim monthsAvailable As Long
//!     monthsAvailable = DateDiff("m", Date, targetDate)
//!     goal.isOnTrack = (goal.monthsNeeded <= monthsAvailable)
//!     
//!     m_goals.Add goal, name
//! End Sub
//!
//! Public Function GetGoalStatus(goalName As String) As String
//!     Dim goal As Goal
//!     Dim monthsAvailable As Long
//!     Dim result As String
//!     
//!     On Error Resume Next
//!     goal = m_goals(goalName)
//!     
//!     If Err.Number <> 0 Then
//!         GetGoalStatus = "Goal not found"
//!         Exit Function
//!     End If
//!     On Error GoTo 0
//!     
//!     monthsAvailable = DateDiff("m", Date, goal.targetDate)
//!     
//!     result = "Goal: " & goal.name & vbCrLf
//!     result = result & "Target: $" & Format(goal.targetAmount, "#,##0.00") & vbCrLf
//!     result = result & "Current: $" & Format(goal.currentAmount, "#,##0.00") & vbCrLf
//!     result = result & "Months Available: " & monthsAvailable & vbCrLf
//!     result = result & "Months Needed: " & Format(goal.monthsNeeded, "0.0") & vbCrLf
//!     
//!     If goal.isOnTrack Then
//!         result = result & "Status: On track!"
//!     Else
//!         Dim requiredMonthly As Double
//!         requiredMonthly = -PMT(goal.expectedReturn / 12, monthsAvailable, _
//!                               -goal.currentAmount, goal.targetAmount)
//!         result = result & "Status: Behind - Need $" & _
//!                  Format(requiredMonthly, "#,##0.00") & "/month"
//!     End If
//!     
//!     GetGoalStatus = result
//! End Function
//!
//! Public Function GenerateAllGoalsReport() As String
//!     Dim report As String
//!     Dim goal As Goal
//!     Dim i As Long
//!     
//!     report = "Financial Goals Progress Report" & vbCrLf
//!     report = report & String(60, "=") & vbCrLf & vbCrLf
//!     
//!     For i = 1 To m_goals.Count
//!         goal = m_goals(i)
//!         report = report & GetGoalStatus(goal.name) & vbCrLf & vbCrLf
//!     Next i
//!     
//!     GenerateAllGoalsReport = report
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! Dim periods As Double
//! periods = NPer(rate, pmt, pv, fv, type)
//! If Err.Number <> 0 Then
//!     MsgBox "Error calculating periods: " & Err.Description & vbCrLf & _
//!            "This may occur if payment is too small to ever pay off the balance."
//!     ' Payment must exceed interest to avoid infinite periods
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - `NPer` uses iterative calculation - reasonably fast
//! - More complex than simple arithmetic operations
//! - For many calculations, consider caching results
//! - Typically completes in microseconds
//! - Performance similar to other financial functions (PMT, PV, FV)
//!
//! ## Best Practices
//!
//! 1. **Use consistent periods** - Match rate and payment periods (monthly, yearly, etc.)
//! 2. **Sign conventions** - Money out is negative, money in is positive
//! 3. **Validate inputs** - Check that payment exceeds interest charge
//! 4. **Handle errors** - Wrap in error handling for invalid scenarios
//! 5. **Round results** - Round to appropriate precision for display
//! 6. **Document assumptions** - State whether payments are start/end of period
//! 7. **Validate reasonableness** - Check that results make sense
//! 8. **Use with other functions** - Combine with PMT, PV, FV for analysis
//! 9. **Consider type parameter** - Specify payment timing when relevant
//! 10. **Format for users** - Convert to years/months for clarity
//!
//! ## Comparison with Related Functions
//!
//! | Function | Calculates | Given |
//! |----------|-----------|-------|
//! | **`NPer`** | Number of periods | Rate, payment, present value, future value |
//! | **PMT** | Payment amount | Rate, periods, present value, future value |
//! | **PV** | Present value | Rate, periods, payment, future value |
//! | **FV** | Future value | Rate, periods, payment, present value |
//! | **Rate** | Interest rate | Periods, payment, present value, future value |
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, Word, etc.)
//! - Available in VB6
//! - **Not available in `VBScript`**
//! - Uses iterative algorithm to solve
//! - Consistent with Excel's NPER function
//! - Part of VBA financial functions library
//!
//! ## Limitations
//!
//! - Assumes constant payment amounts
//! - Assumes constant interest rate
//! - Payment must exceed periodic interest or will fail/return error
//! - Does not account for fees, taxes, or insurance
//! - Cannot handle irregular payment schedules
//! - Type parameter limited to 0 or 1
//! - May return fractional periods
//!
//! ## Related Functions
//!
//! - **PMT** - Calculates payment amount for a loan
//! - **PV** - Calculates present value of an investment
//! - **FV** - Calculates future value of an investment
//! - **Rate** - Calculates interest rate per period
//! - **`IPmt`** - Calculates interest payment for a specific period
//! - **`PPmt`** - Calculates principal payment for a specific period
//!
//! ## VB6 Parser Notes
//!
//! `NPer` is parsed as a regular function call (`CallExpression`). This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers
//! working with financial calculations, loan analysis, investment planning, and time-value-of-money
//! operations.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn nper_basic() {
        let source = r#"
Dim months As Double
months = NPer(0.08 / 12, -200, 10000)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_with_fv() {
        let source = r#"
Dim periods As Double
periods = NPer(0.06 / 12, -300, 0, 50000)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_with_type() {
        let source = r#"
Dim n As Double
n = NPer(rate, pmt, pv, fv, 1)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_if_statement() {
        let source = r#"
If NPer(apr / 12, -payment, balance) > 60 Then
    MsgBox "Payoff will take over 5 years"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_function_return() {
        let source = r#"
Function CalculatePayoffMonths(loan As Double, payment As Double, rate As Double) As Double
    CalculatePayoffMonths = NPer(rate / 12, -payment, loan)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_variable_assignment() {
        let source = r#"
Dim payoffYears As Double
payoffYears = NPer(0.05 / 12, -1000, 200000) / 12
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_msgbox() {
        let source = r#"
MsgBox "Months to payoff: " & Format(NPer(rate, pmt, pv), "0.0")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_debug_print() {
        let source = r#"
Debug.Print "Periods: " & NPer(interestRate, monthlyPmt, principal)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_select_case() {
        let source = r#"
Select Case Int(NPer(0.08 / 12, -250, 5000) / 12)
    Case Is < 2
        MsgBox "Short term"
    Case 2 To 5
        MsgBox "Medium term"
    Case Else
        MsgBox "Long term"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_class_usage() {
        let source = r#"
Private m_periods As Double

Public Sub CalculatePeriods()
    m_periods = NPer(m_rate / 12, -m_payment, m_balance)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_with_statement() {
        let source = r#"
With loanInfo
    .PayoffMonths = NPer(.Rate / 12, -.MonthlyPayment, .Principal)
End With
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_elseif() {
        let source = r#"
If x > 0 Then
    y = 1
ElseIf NPer(r, p, v) < 36 Then
    y = 2
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_for_loop() {
        let source = r#"
For payment = 100 To 500 Step 50
    months = NPer(0.1 / 12, -payment, 10000)
    Debug.Print payment, months
Next payment
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_do_while() {
        let source = r#"
Do While NPer(rate, -payment, balance) > targetMonths
    payment = payment + 10
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_do_until() {
        let source = r#"
Do Until NPer(apr / 12, -pmt, bal) <= 12
    pmt = pmt + 50
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_while_wend() {
        let source = r#"
While NPer(0.05 / 12, -amount, 50000) > 120
    amount = amount + 25
Wend
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_parentheses() {
        let source = r#"
Dim result As Double
result = (NPer(rate, payment, principal))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_iif() {
        let source = r#"
Dim term As String
term = IIf(NPer(r, p, v) < 36, "Short", "Long")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_comparison() {
        let source = r#"
If NPer(rate1, pmt, bal) < NPer(rate2, pmt, bal) Then
    MsgBox "Rate 1 is faster"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_array_assignment() {
        let source = r#"
Dim payoffTimes(10) As Double
payoffTimes(i) = NPer(rates(i) / 12, -payment, balance)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_property_assignment() {
        let source = r#"
Set obj = New LoanCalculator
obj.PayoffPeriods = NPer(obj.Rate / 12, -obj.Payment, obj.Balance)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_function_argument() {
        let source = r#"
Call DisplayPayoffSchedule(NPer(apr / 12, -monthlyPmt, loanAmount))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_format() {
        let source = r#"
Dim formatted As String
formatted = Format(NPer(0.06 / 12, -500, 25000), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_arithmetic() {
        let source = r#"
Dim years As Double
years = NPer(rate / 12, -payment, principal) / 12
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_concatenation() {
        let source = r#"
Dim msg As String
msg = "Payoff time: " & NPer(r, p, v) & " months"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_int_conversion() {
        let source = r#"
Dim wholeMonths As Integer
wholeMonths = Int(NPer(0.08 / 12, -300, 15000))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn nper_error_handling() {
        let source = r#"
On Error Resume Next
n = NPer(rate, payment, balance)
If Err.Number <> 0 Then
    MsgBox "Invalid calculation"
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPer"));
        assert!(text.contains("Identifier"));
    }
}

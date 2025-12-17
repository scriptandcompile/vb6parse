//! # Pmt Function
//!
//! Returns a Double specifying the payment for an annuity based on periodic, fixed payments and a fixed interest rate.
//!
//! ## Syntax
//!
//! ```vb
//! Pmt(rate, nper, pv, [fv], [type])
//! ```
//!
//! ## Parameters
//!
//! - `rate` - Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10% and make monthly payments, the rate per period is 0.1/12, or 0.0083.
//! - `nper` - Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a 4-year car loan, your loan has 4 * 12 (or 48) payment periods.
//! - `pv` - Required. Double specifying present value (or lump sum) that a series of payments to be paid in the future is worth now. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.
//! - `fv` - Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.
//! - `type` - Optional. Variant specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.
//!
//! ## Return Value
//!
//! Returns a `Double` specifying the payment amount per period. The payment includes principal and interest but includes no taxes, reserve payments, or fees sometimes associated with loans.
//!
//! ## Remarks
//!
//! The `Pmt` function is one of the most commonly used financial functions in VB6. It calculates the periodic payment
//! required to pay off a loan or to accumulate a certain amount in a savings plan, given a fixed interest rate and
//! fixed payment periods.
//!
//! An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage)
//! or an investment (such as a monthly savings plan).
//!
//! The `rate` and `nper` arguments must be calculated using payment periods expressed in the same units. For example, if `rate`
//! is calculated using months, `nper` must also be calculated using months.
//!
//! For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received
//! (such as dividend checks) is represented by positive numbers.
//!
//! **Important**: The payment returned by `Pmt` is typically a negative number when calculating loan payments, because
//! it represents money flowing out. When calculating savings deposits, the payment is positive if you want to know
//! how much to deposit (money flowing out).
//!
//! ## Typical Uses
//!
//! 1. **Mortgage Calculation**: Calculate monthly mortgage payments for home loans
//! 2. **Auto Loan Payments**: Determine monthly car loan payments
//! 3. **Savings Goals**: Calculate required monthly deposits to reach a savings target
//! 4. **Retirement Planning**: Determine periodic contributions needed for retirement funds
//! 5. **Lease Payments**: Calculate lease payment amounts for equipment or vehicles
//! 6. **Student Loan Amortization**: Compute monthly student loan payments
//! 7. **Investment Planning**: Determine periodic investment amounts to reach financial goals
//! 8. **Loan Comparison**: Compare different loan options by payment amount
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Loan Payment
//! ```vb
//! ' Calculate monthly payment for a $20,000 loan at 6% APR for 5 years
//! Dim monthlyPayment As Double
//! monthlyPayment = Pmt(0.06 / 12, 5 * 12, 20000)
//! ' Returns approximately -386.66 (negative because it's money paid out)
//! ```
//!
//! ### Example 2: Mortgage Payment
//! ```vb
//! ' $200,000 home loan, 30 years, 4.5% APR
//! Dim mortgagePayment As Double
//! mortgagePayment = Pmt(0.045 / 12, 30 * 12, 200000)
//! ' Returns approximately -1,013.37 per month
//! ```
//!
//! ### Example 3: Savings Plan
//! ```vb
//! ' How much to save monthly to accumulate $50,000 in 10 years at 5% annual return?
//! Dim monthlyDeposit As Double
//! monthlyDeposit = Pmt(0.05 / 12, 10 * 12, 0, -50000)
//! ' Returns approximately -322.67 per month (negative = deposit)
//! ```
//!
//! ### Example 4: Payment Due at Beginning of Period
//! ```vb
//! ' Lease payment due at start of month
//! Dim leasePayment As Double
//! leasePayment = Pmt(0.08 / 12, 36, 25000, 0, 1)
//! ' Returns slightly lower payment due to beginning-of-period timing
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `CalculateLoanPayment`
//! ```vb
//! Function CalculateLoanPayment(loanAmount As Double, _
//!                               annualRate As Double, _
//!                               years As Integer) As Double
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     
//!     monthlyRate = annualRate / 12
//!     numPayments = years * 12
//!     
//!     ' Return positive value (absolute value of payment)
//!     CalculateLoanPayment = Abs(Pmt(monthlyRate, numPayments, loanAmount))
//! End Function
//! ```
//!
//! ### Pattern 2: `CalculateTotalInterest`
//! ```vb
//! Function CalculateTotalInterest(principal As Double, _
//!                                 annualRate As Double, _
//!                                 years As Integer) As Double
//!     Dim payment As Double
//!     Dim totalPaid As Double
//!     Dim numPayments As Integer
//!     
//!     numPayments = years * 12
//!     payment = Abs(Pmt(annualRate / 12, numPayments, principal))
//!     totalPaid = payment * numPayments
//!     
//!     CalculateTotalInterest = totalPaid - principal
//! End Function
//! ```
//!
//! ### Pattern 3: `CompareLoanOptions`
//! ```vb
//! Sub CompareLoanOptions(amount As Double)
//!     Dim payment15 As Double
//!     Dim payment30 As Double
//!     Dim totalInterest15 As Double
//!     Dim totalInterest30 As Double
//!     
//!     ' 15-year loan at 3.5%
//!     payment15 = Abs(Pmt(0.035 / 12, 15 * 12, amount))
//!     totalInterest15 = (payment15 * 15 * 12) - amount
//!     
//!     ' 30-year loan at 4.0%
//!     payment30 = Abs(Pmt(0.04 / 12, 30 * 12, amount))
//!     totalInterest30 = (payment30 * 30 * 12) - amount
//!     
//!     Debug.Print "15-year: $" & Format(payment15, "0.00") & "/mo, Total Interest: $" & Format(totalInterest15, "#,##0")
//!     Debug.Print "30-year: $" & Format(payment30, "0.00") & "/mo, Total Interest: $" & Format(totalInterest30, "#,##0")
//! End Sub
//! ```
//!
//! ### Pattern 4: `CalculateAffordableAmount`
//! ```vb
//! Function CalculateAffordableAmount(monthlyPayment As Double, _
//!                                    annualRate As Double, _
//!                                    years As Integer) As Double
//!     ' Calculate how much you can borrow given a payment amount
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     Dim principal As Double
//!     
//!     monthlyRate = annualRate / 12
//!     numPayments = years * 12
//!     
//!     ' Use negative payment because Pmt returns negative for outflows
//!     principal = PV(monthlyRate, numPayments, -monthlyPayment)
//!     CalculateAffordableAmount = principal
//! End Function
//! ```
//!
//! ### Pattern 5: `SavingsCalculator`
//! ```vb
//! Function CalculateSavingsDeposit(targetAmount As Double, _
//!                                  years As Integer, _
//!                                  annualReturn As Double) As Double
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     
//!     monthlyRate = annualReturn / 12
//!     numPayments = years * 12
//!     
//!     ' Use negative FV because it's a goal (money we want)
//!     ' Return positive value (absolute value)
//!     CalculateSavingsDeposit = Abs(Pmt(monthlyRate, numPayments, 0, -targetAmount))
//! End Function
//! ```
//!
//! ### Pattern 6: `ValidatePmtParameters`
//! ```vb
//! Function ValidatePmtParameters(rate As Double, nper As Integer, _
//!                                pv As Double) As Boolean
//!     ValidatePmtParameters = False
//!     
//!     If nper <= 0 Then
//!         MsgBox "Number of periods must be positive"
//!         Exit Function
//!     End If
//!     
//!     If rate < 0 Then
//!         MsgBox "Interest rate cannot be negative"
//!         Exit Function
//!     End If
//!     
//!     If pv = 0 Then
//!         MsgBox "Present value cannot be zero for loan calculations"
//!         Exit Function
//!     End If
//!     
//!     ValidatePmtParameters = True
//! End Function
//! ```
//!
//! ### Pattern 7: `CalculateWithExtraPayment`
//! ```vb
//! Function CalculatePayoffWithExtra(principal As Double, _
//!                                   annualRate As Double, _
//!                                   years As Integer, _
//!                                   extraPayment As Double, _
//!                                   ByRef periodsToPayoff As Integer) As Double
//!     Dim regularPayment As Double
//!     Dim totalPayment As Double
//!     Dim balance As Double
//!     Dim monthlyRate As Double
//!     Dim period As Integer
//!     
//!     monthlyRate = annualRate / 12
//!     regularPayment = Abs(Pmt(monthlyRate, years * 12, principal))
//!     totalPayment = regularPayment + extraPayment
//!     
//!     balance = principal
//!     period = 0
//!     
//!     Do While balance > 0 And period < years * 12
//!         balance = balance * (1 + monthlyRate) - totalPayment
//!         period = period + 1
//!     Loop
//!     
//!     periodsToPayoff = period
//!     CalculatePayoffWithExtra = totalPayment
//! End Function
//! ```
//!
//! ### Pattern 8: `AmortizationSchedule`
//! ```vb
//! Sub CreateAmortizationSchedule(principal As Double, _
//!                                annualRate As Double, _
//!                                years As Integer)
//!     Dim payment As Double
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     Dim balance As Double
//!     Dim interestPaid As Double
//!     Dim principalPaid As Double
//!     Dim period As Integer
//!     
//!     monthlyRate = annualRate / 12
//!     numPayments = years * 12
//!     payment = Abs(Pmt(monthlyRate, numPayments, principal))
//!     balance = principal
//!     
//!     Debug.Print "Period", "Payment", "Interest", "Principal", "Balance"
//!     Debug.Print String(60, "-")
//!     
//!     For period = 1 To numPayments
//!         interestPaid = balance * monthlyRate
//!         principalPaid = payment - interestPaid
//!         balance = balance - principalPaid
//!         
//!         Debug.Print period, _
//!                     Format(payment, "0.00"), _
//!                     Format(interestPaid, "0.00"), _
//!                     Format(principalPaid, "0.00"), _
//!                     Format(balance, "0.00")
//!     Next period
//! End Sub
//! ```
//!
//! ### Pattern 9: `BiweeklyPaymentCalculator`
//! ```vb
//! Function CalculateBiweeklyPayment(principal As Double, _
//!                                   annualRate As Double, _
//!                                   years As Integer) As Double
//!     Dim monthlyPayment As Double
//!     Dim biweeklyPayment As Double
//!     
//!     ' Calculate monthly payment
//!     monthlyPayment = Abs(Pmt(annualRate / 12, years * 12, principal))
//!     
//!     ' Biweekly is half the monthly payment
//!     ' This results in 26 payments per year instead of 24, paying off faster
//!     biweeklyPayment = monthlyPayment / 2
//!     
//!     CalculateBiweeklyPayment = biweeklyPayment
//! End Function
//! ```
//!
//! ### Pattern 10: `RefinanceAnalysis`
//! ```vb
//! Function ShouldRefinance(currentBalance As Double, _
//!                          currentRate As Double, _
//!                          currentYearsLeft As Integer, _
//!                          newRate As Double, _
//!                          newYears As Integer, _
//!                          closingCosts As Double) As Boolean
//!     Dim currentPayment As Double
//!     Dim newPayment As Double
//!     Dim monthlySavings As Double
//!     Dim breakEvenMonths As Double
//!     
//!     currentPayment = Abs(Pmt(currentRate / 12, currentYearsLeft * 12, currentBalance))
//!     newPayment = Abs(Pmt(newRate / 12, newYears * 12, currentBalance))
//!     
//!     monthlySavings = currentPayment - newPayment
//!     
//!     If monthlySavings <= 0 Then
//!         ShouldRefinance = False
//!         Exit Function
//!     End If
//!     
//!     breakEvenMonths = closingCosts / monthlySavings
//!     
//!     ' Refinance if break-even is less than 3 years and you'll stay that long
//!     ShouldRefinance = (breakEvenMonths <= 36)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Comprehensive Loan Calculator Class
//! ```vb
//! ' Complete loan calculator with full amortization
//! Class LoanCalculator
//!     Private m_principal As Double
//!     Private m_annualRate As Double
//!     Private m_years As Integer
//!     Private m_paymentType As Integer  ' 0 = end of period, 1 = beginning
//!     
//!     Public Property Let Principal(value As Double)
//!         If value > 0 Then m_principal = value
//!     End Property
//!     
//!     Public Property Let AnnualRate(value As Double)
//!         If value >= 0 Then m_annualRate = value
//!     End Property
//!     
//!     Public Property Let Years(value As Integer)
//!         If value > 0 Then m_years = value
//!     End Property
//!     
//!     Public Property Let PaymentType(value As Integer)
//!         If value = 0 Or value = 1 Then m_paymentType = value
//!     End Property
//!     
//!     Public Function GetMonthlyPayment() As Double
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         
//!         monthlyRate = m_annualRate / 12
//!         numPayments = m_years * 12
//!         
//!         GetMonthlyPayment = Abs(Pmt(monthlyRate, numPayments, m_principal, 0, m_paymentType))
//!     End Function
//!     
//!     Public Function GetTotalInterest() As Double
//!         Dim payment As Double
//!         Dim totalPaid As Double
//!         
//!         payment = GetMonthlyPayment()
//!         totalPaid = payment * m_years * 12
//!         
//!         GetTotalInterest = totalPaid - m_principal
//!     End Function
//!     
//!     Public Function GetTotalPaid() As Double
//!         GetTotalPaid = GetMonthlyPayment() * m_years * 12
//!     End Function
//!     
//!     Public Function GetAmortizationSchedule() As Collection
//!         Dim schedule As Collection
//!         Dim payment As Double
//!         Dim monthlyRate As Double
//!         Dim balance As Double
//!         Dim interestPaid As Double
//!         Dim principalPaid As Double
//!         Dim period As Integer
//!         Dim entry As Object
//!         
//!         Set schedule = New Collection
//!         monthlyRate = m_annualRate / 12
//!         payment = GetMonthlyPayment()
//!         balance = m_principal
//!         
//!         For period = 1 To m_years * 12
//!             interestPaid = balance * monthlyRate
//!             principalPaid = payment - interestPaid
//!             balance = balance - principalPaid
//!             
//!             Set entry = CreateObject("Scripting.Dictionary")
//!             entry.Add "Period", period
//!             entry.Add "Payment", payment
//!             entry.Add "Interest", interestPaid
//!             entry.Add "Principal", principalPaid
//!             entry.Add "Balance", balance
//!             
//!             schedule.Add entry
//!         Next period
//!         
//!         Set GetAmortizationSchedule = schedule
//!     End Function
//!     
//!     Public Function GenerateReport() As String
//!         Dim report As String
//!         
//!         report = "Loan Analysis Report" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Loan Amount: $" & Format(m_principal, "#,##0.00") & vbCrLf
//!         report = report & "Annual Rate: " & Format(m_annualRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "Term: " & m_years & " years" & vbCrLf
//!         report = report & "Payment Type: " & IIf(m_paymentType = 0, "End of Period", "Beginning of Period") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Monthly Payment: $" & Format(GetMonthlyPayment(), "#,##0.00") & vbCrLf
//!         report = report & "Total Paid: $" & Format(GetTotalPaid(), "#,##0.00") & vbCrLf
//!         report = report & "Total Interest: $" & Format(GetTotalInterest(), "#,##0.00") & vbCrLf
//!         report = report & "Interest as % of Principal: " & _
//!                  Format((GetTotalInterest() / m_principal) * 100, "0.0") & "%"
//!         
//!         GenerateReport = report
//!     End Function
//!     
//!     Public Function CalculatePayoffDate(startDate As Date) As Date
//!         CalculatePayoffDate = DateAdd("m", m_years * 12, startDate)
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Retirement Savings Planner
//! ```vb
//! ' Plan retirement savings with multiple scenarios
//! Class RetirementPlanner
//!     Private m_currentAge As Integer
//!     Private m_retirementAge As Integer
//!     Private m_targetAmount As Double
//!     Private m_expectedReturn As Double
//!     Private m_currentSavings As Double
//!     
//!     Public Sub Initialize(currentAge As Integer, retirementAge As Integer, _
//!                          targetAmount As Double, expectedReturn As Double, _
//!                          currentSavings As Double)
//!         m_currentAge = currentAge
//!         m_retirementAge = retirementAge
//!         m_targetAmount = targetAmount
//!         m_expectedReturn = expectedReturn
//!         m_currentSavings = currentSavings
//!     End Sub
//!     
//!     Public Function GetRequiredMonthlyContribution() As Double
//!         Dim yearsToRetirement As Integer
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         Dim futureValueNeeded As Double
//!         
//!         yearsToRetirement = m_retirementAge - m_currentAge
//!         monthlyRate = m_expectedReturn / 12
//!         numPayments = yearsToRetirement * 12
//!         
//!         ' Account for current savings growing
//!         futureValueNeeded = m_targetAmount - (m_currentSavings * ((1 + monthlyRate) ^ numPayments))
//!         
//!         ' Calculate required payment
//!         GetRequiredMonthlyContribution = Abs(Pmt(monthlyRate, numPayments, 0, -futureValueNeeded))
//!     End Function
//!     
//!     Public Function ProjectSavingsGrowth() As Collection
//!         Dim projections As Collection
//!         Dim monthlyContribution As Double
//!         Dim balance As Double
//!         Dim monthlyRate As Double
//!         Dim year As Integer
//!         Dim month As Integer
//!         Dim entry As Object
//!         
//!         Set projections = New Collection
//!         monthlyContribution = GetRequiredMonthlyContribution()
//!         monthlyRate = m_expectedReturn / 12
//!         balance = m_currentSavings
//!         
//!         For year = m_currentAge To m_retirementAge - 1
//!             For month = 1 To 12
//!                 balance = balance * (1 + monthlyRate) + monthlyContribution
//!             Next month
//!             
//!             Set entry = CreateObject("Scripting.Dictionary")
//!             entry.Add "Age", year + 1
//!             entry.Add "Balance", balance
//!             entry.Add "YearlyContribution", monthlyContribution * 12
//!             
//!             projections.Add entry
//!         Next year
//!         
//!         Set ProjectSavingsGrowth = projections
//!     End Function
//!     
//!     Public Function AnalyzeReturnScenarios() As String
//!         Dim report As String
//!         Dim rates() As Double
//!         Dim i As Integer
//!         Dim payment As Double
//!         
//!         ReDim rates(0 To 4)
//!         rates(0) = 0.04
//!         rates(1) = 0.06
//!         rates(2) = 0.08
//!         rates(3) = 0.10
//!         rates(4) = 0.12
//!         
//!         report = "Retirement Contribution Analysis" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Target Amount: $" & Format(m_targetAmount, "#,##0") & vbCrLf
//!         report = report & "Years to Retirement: " & (m_retirementAge - m_currentAge) & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Return Rate    Monthly Contribution" & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         
//!         For i = 0 To 4
//!             m_expectedReturn = rates(i)
//!             payment = GetRequiredMonthlyContribution()
//!             report = report & Format(rates(i) * 100, "00.0") & "%          $" & _
//!                      Format(payment, "#,##0.00") & vbCrLf
//!         Next i
//!         
//!         AnalyzeReturnScenarios = report
//!     End Function
//!     
//!     Public Function GenerateRetirementPlan() As String
//!         Dim plan As String
//!         Dim monthlyContribution As Double
//!         Dim totalContributions As Double
//!         Dim yearsToRetirement As Integer
//!         
//!         monthlyContribution = GetRequiredMonthlyContribution()
//!         yearsToRetirement = m_retirementAge - m_currentAge
//!         totalContributions = monthlyContribution * 12 * yearsToRetirement
//!         
//!         plan = "Personalized Retirement Plan" & vbCrLf
//!         plan = plan & String(50, "=") & vbCrLf
//!         plan = plan & "Current Age: " & m_currentAge & vbCrLf
//!         plan = plan & "Retirement Age: " & m_retirementAge & vbCrLf
//!         plan = plan & "Current Savings: $" & Format(m_currentSavings, "#,##0") & vbCrLf
//!         plan = plan & "Target Amount: $" & Format(m_targetAmount, "#,##0") & vbCrLf
//!         plan = plan & "Expected Return: " & Format(m_expectedReturn * 100, "0.0") & "%" & vbCrLf
//!         plan = plan & String(50, "-") & vbCrLf
//!         plan = plan & "Required Monthly Contribution: $" & Format(monthlyContribution, "#,##0.00") & vbCrLf
//!         plan = plan & "Total You'll Contribute: $" & Format(totalContributions, "#,##0") & vbCrLf
//!         plan = plan & "Expected Investment Growth: $" & _
//!                Format(m_targetAmount - totalContributions - m_currentSavings, "#,##0")
//!         
//!         GenerateRetirementPlan = plan
//!     End Function
//! End Class
//! ```
//!
//! ### Example 3: Auto Loan Comparison Tool
//! ```vb
//! ' Compare multiple auto loan options
//! Module AutoLoanAnalyzer
//!     Private Type LoanOption
//!         LenderName As String
//!         Rate As Double
//!         Term As Integer
//!         DownPayment As Double
//!         TradeInValue As Double
//!         Fees As Double
//!     End Type
//!     
//!     Public Function AnalyzeLoanOptions(carPrice As Double, _
//!                                       options() As LoanOption) As String
//!         Dim report As String
//!         Dim i As Integer
//!         Dim loanAmount As Double
//!         Dim payment As Double
//!         Dim totalCost As Double
//!         Dim totalInterest As Double
//!         Dim bestOption As Integer
//!         Dim lowestCost As Double
//!         
//!         report = "Auto Loan Comparison" & vbCrLf
//!         report = report & "Vehicle Price: $" & Format(carPrice, "#,##0") & vbCrLf
//!         report = report & String(70, "=") & vbCrLf
//!         report = report & "Lender          Rate   Term   Payment   Total Cost   Interest" & vbCrLf
//!         report = report & String(70, "-") & vbCrLf
//!         
//!         lowestCost = 999999999
//!         
//!         For i = LBound(options) To UBound(options)
//!             loanAmount = carPrice - options(i).DownPayment - options(i).TradeInValue + options(i).Fees
//!             payment = Abs(Pmt(options(i).Rate / 12, options(i).Term, loanAmount))
//!             totalCost = payment * options(i).Term + options(i).DownPayment - options(i).TradeInValue
//!             totalInterest = (payment * options(i).Term) - loanAmount
//!             
//!             report = report & Left(options(i).LenderName & Space(15), 15) & " "
//!             report = report & Format(options(i).Rate * 100, "0.0") & "%  "
//!             report = report & Format(options(i).Term, "00") & "mo  "
//!             report = report & "$" & Format(payment, "000.00") & "  "
//!             report = report & "$" & Format(totalCost, "#,##0") & "  "
//!             report = report & "$" & Format(totalInterest, "#,##0")
//!             
//!             If totalCost < lowestCost Then
//!                 lowestCost = totalCost
//!                 bestOption = i
//!                 report = report & " *BEST*"
//!             End If
//!             
//!             report = report & vbCrLf
//!         Next i
//!         
//!         report = report & String(70, "-") & vbCrLf
//!         report = report & "Recommended: " & options(bestOption).LenderName
//!         
//!         AnalyzeLoanOptions = report
//!     End Function
//!     
//!     Public Function CalculateBreakEvenTerm(price As Double, rate1 As Double, _
//!                                           term1 As Integer, rate2 As Double, _
//!                                           term2 As Integer) As String
//!         Dim payment1 As Double
//!         Dim payment2 As Double
//!         Dim totalCost1 As Double
//!         Dim totalCost2 As Double
//!         Dim monthlySavings As Double
//!         
//!         payment1 = Abs(Pmt(rate1 / 12, term1, price))
//!         payment2 = Abs(Pmt(rate2 / 12, term2, price))
//!         totalCost1 = payment1 * term1
//!         totalCost2 = payment2 * term2
//!         monthlySavings = payment1 - payment2
//!         
//!         CalculateBreakEvenTerm = "Option 1: " & term1 & " months at " & _
//!                                 Format(rate1 * 100, "0.0") & "% = $" & _
//!                                 Format(payment1, "0.00") & "/mo (Total: $" & _
//!                                 Format(totalCost1, "#,##0") & ")" & vbCrLf & _
//!                                 "Option 2: " & term2 & " months at " & _
//!                                 Format(rate2 * 100, "0.0") & "% = $" & _
//!                                 Format(payment2, "0.00") & "/mo (Total: $" & _
//!                                 Format(totalCost2, "#,##0") & ")"
//!     End Function
//! End Module
//! ```
//!
//! ### Example 4: Mortgage Affordability Calculator
//! ```vb
//! ' Calculate home affordability based on income
//! Class MortgageAffordabilityCalculator
//!     Private m_monthlyIncome As Double
//!     Private m_monthlyDebts As Double
//!     Private m_downPaymentPercent As Double
//!     Private m_annualRate As Double
//!     Private m_loanTermYears As Integer
//!     Private m_propertyTaxRate As Double
//!     Private m_insuranceRate As Double
//!     
//!     Public Sub SetIncome(monthlyIncome As Double)
//!         m_monthlyIncome = monthlyIncome
//!     End Sub
//!     
//!     Public Sub SetDebts(monthlyDebts As Double)
//!         m_monthlyDebts = monthlyDebts
//!     End Sub
//!     
//!     Public Sub SetLoanTerms(downPaymentPercent As Double, _
//!                            annualRate As Double, years As Integer)
//!         m_downPaymentPercent = downPaymentPercent
//!         m_annualRate = annualRate
//!         m_loanTermYears = years
//!     End Sub
//!     
//!     Public Sub SetHousingCosts(propertyTaxRate As Double, insuranceRate As Double)
//!         m_propertyTaxRate = propertyTaxRate
//!         m_insuranceRate = insuranceRate
//!     End Sub
//!     
//!     Public Function GetMaxMonthlyPayment() As Double
//!         ' Use 28% front-end ratio (housing costs / gross income)
//!         ' and 36% back-end ratio (total debt / gross income)
//!         Dim maxByFrontEnd As Double
//!         Dim maxByBackEnd As Double
//!         
//!         maxByFrontEnd = m_monthlyIncome * 0.28
//!         maxByBackEnd = (m_monthlyIncome * 0.36) - m_monthlyDebts
//!         
//!         ' Use the more conservative (lower) value
//!         If maxByFrontEnd < maxByBackEnd Then
//!             GetMaxMonthlyPayment = maxByFrontEnd
//!         Else
//!             GetMaxMonthlyPayment = maxByBackEnd
//!         End If
//!     End Function
//!     
//!     Public Function GetAffordableHomePrice() As Double
//!         Dim maxPayment As Double
//!         Dim principalAndInterest As Double
//!         Dim loanAmount As Double
//!         Dim homePrice As Double
//!         
//!         maxPayment = GetMaxMonthlyPayment()
//!         
//!         ' Estimate property tax and insurance (rough approximation)
//!         ' Actual payment = P&I + taxes + insurance
//!         ' taxes ≈ (price × tax rate) / 12
//!         ' insurance ≈ (price × insurance rate) / 12
//!         ' P&I = maxPayment - (price × (taxRate + insuranceRate) / 12)
//!         
//!         ' For simplicity, allocate 80% of max payment to P&I
//!         principalAndInterest = maxPayment * 0.8
//!         
//!         ' Use PV to find loan amount from payment
//!         loanAmount = Abs(PV(m_annualRate / 12, m_loanTermYears * 12, -principalAndInterest))
//!         
//!         ' Calculate home price from loan amount
//!         homePrice = loanAmount / (1 - m_downPaymentPercent)
//!         
//!         GetAffordableHomePrice = homePrice
//!     End Function
//!     
//!     Public Function GenerateAffordabilityReport() As String
//!         Dim report As String
//!         Dim maxPayment As Double
//!         Dim affordablePrice As Double
//!         Dim downPayment As Double
//!         Dim loanAmount As Double
//!         Dim estimatedPayment As Double
//!         
//!         maxPayment = GetMaxMonthlyPayment()
//!         affordablePrice = GetAffordableHomePrice()
//!         downPayment = affordablePrice * m_downPaymentPercent
//!         loanAmount = affordablePrice - downPayment
//!         estimatedPayment = Abs(Pmt(m_annualRate / 12, m_loanTermYears * 12, loanAmount))
//!         
//!         report = "Mortgage Affordability Analysis" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Income & Debt Information:" & vbCrLf
//!         report = report & "  Monthly Gross Income: $" & Format(m_monthlyIncome, "#,##0") & vbCrLf
//!         report = report & "  Monthly Debt Payments: $" & Format(m_monthlyDebts, "#,##0") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Loan Parameters:" & vbCrLf
//!         report = report & "  Down Payment: " & Format(m_downPaymentPercent * 100, "0") & "%" & vbCrLf
//!         report = report & "  Interest Rate: " & Format(m_annualRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "  Loan Term: " & m_loanTermYears & " years" & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Affordability Results:" & vbCrLf
//!         report = report & "  Max Monthly Payment: $" & Format(maxPayment, "#,##0") & vbCrLf
//!         report = report & "  Affordable Home Price: $" & Format(affordablePrice, "#,##0") & vbCrLf
//!         report = report & "  Required Down Payment: $" & Format(downPayment, "#,##0") & vbCrLf
//!         report = report & "  Loan Amount: $" & Format(loanAmount, "#,##0") & vbCrLf
//!         report = report & "  Est. Monthly P&I: $" & Format(estimatedPayment, "#,##0")
//!         
//!         GenerateAffordabilityReport = report
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `Pmt` function can raise errors in the following situations:
//!
//! - **Invalid Procedure Call (Error 5)**: When:
//!   - `nper` is 0 or negative
//!   - `rate` is -1 (causes division by zero in the formula)
//! - **Type Mismatch (Error 13)**: When arguments cannot be converted to numeric values
//! - **Overflow (Error 6)**: When calculated payment exceeds Double range
//!
//! Always validate input parameters:
//!
//! ```vb
//! On Error Resume Next
//! payment = Pmt(rate, nper, pv, fv, type)
//! If Err.Number <> 0 Then
//!     MsgBox "Error calculating payment: " & Err.Description
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - The `Pmt` function is very fast for individual calculations
//! - For amortization schedules with hundreds of periods, pre-calculate the payment once
//! - Avoid calling `Pmt` repeatedly in tight loops with the same parameters
//! - Use variables to store intermediate calculations (monthly rate, etc.)
//!
//! ## Best Practices
//!
//! 1. **Convert Rates Properly**: Always divide annual rates by 12 for monthly payments
//! 2. **Match Time Units**: Ensure rate and nper use the same time period
//! 3. **Use Absolute Value**: Use `Abs()` to display positive payment amounts to users
//! 4. **Validate Inputs**: Check that nper > 0 and rate is reasonable before calling
//! 5. **Handle Sign Conventions**: Remember negative = outflow, positive = inflow
//! 6. **Round for Display**: Use `Format()` to display payments with 2 decimal places
//! 7. **Consider Type Parameter**: Use type=1 for beginning-of-period payments (leases, etc.)
//! 8. **Document Assumptions**: Clearly state what rate, term, and conditions are used
//! 9. **Test Edge Cases**: Verify behavior with 0% rate, very short/long terms
//! 10. **Combine with Other Functions**: Use with `IPmt`, `PPmt`, `PV`, `FV` for complete analysis
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **Pmt** | Calculate periodic payment | Double (payment amount) | Loan payments, savings deposits |
//! | **`IPmt`** | Interest portion of payment | Double (interest amount) | Amortization schedules, tax deductions |
//! | **`PPmt`** | Principal portion of payment | Double (principal amount) | Tracking loan balance reduction |
//! | **PV** | Present value | Double (current value) | Reverse calculation from payment |
//! | **FV** | Future value | Double (future value) | Investment growth projections |
//! | **`NPer`** | Number of periods | Double (period count) | How long to pay off debt |
//! | **Rate** | Interest rate | Double (rate per period) | Finding effective interest rate |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VBA and VB6
//! - Behavior is consistent across Windows platforms
//! - Uses standard annuity formulas from financial mathematics
//! - For zero interest rates, payment is simply pv / nper
//! - Maximum precision limited by Double data type
//!
//! ## Limitations
//!
//! - Assumes constant interest rate over entire period
//! - Assumes equal payment amounts (no variable payments)
//! - Does not account for taxes, fees, or insurance directly
//! - Cannot handle variable rate loans (ARMs) without recalculation
//! - Does not consider payment frequency other than what you specify
//! - Sign convention can be confusing (negative for outflows)
//!
//! ## Related Functions
//!
//! - `IPmt`: Returns the interest payment for a specific period
//! - `PPmt`: Returns the principal payment for a specific period
//! - `PV`: Returns the present value of an investment
//! - `FV`: Returns the future value of an investment
//! - `NPer`: Returns the number of periods for an investment
//! - `Rate`: Returns the interest rate per period

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn pmt_basic() {
        let source = r#"
Dim payment As Double
payment = Pmt(0.06 / 12, 60, 20000)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_with_all_parameters() {
        let source = r#"
Dim monthlyPayment As Double
monthlyPayment = Pmt(0.045 / 12, 360, 200000, 0, 0)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_if_statement() {
        let source = r#"
If Abs(Pmt(rate, nper, principal)) > maxPayment Then
    MsgBox "Payment too high"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_function_return() {
        let source = r#"
Function CalculatePayment(amount As Double, years As Integer) As Double
    CalculatePayment = Abs(Pmt(0.05 / 12, years * 12, amount))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_variable_assignment() {
        let source = r#"
Dim loanPayment As Double
loanPayment = Pmt(interestRate, periods, loanAmount)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_msgbox() {
        let source = r#"
MsgBox "Monthly payment: $" & Format(Abs(Pmt(0.06 / 12, 60, 25000)), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_debug_print() {
        let source = r#"
Debug.Print "Payment: " & Pmt(monthlyRate, numPayments, principal)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_select_case() {
        let source = r#"
Dim payment As Double
payment = Abs(Pmt(0.05 / 12, 360, loanAmount))
Select Case payment
    Case Is < 1000
        category = "Affordable"
    Case Is < 2000
        category = "Moderate"
    Case Else
        category = "High"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_class_usage() {
        let source = r#"
Private m_payment As Double

Public Sub CalculateLoan()
    m_payment = Pmt(m_rate, m_periods, m_principal)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_with_statement() {
        let source = r#"
With loanCalc
    .MonthlyPayment = Pmt(.Rate, .NumPayments, .Amount)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_elseif() {
        let source = r#"
If amount < 10000 Then
    rate = 0.05
ElseIf Pmt(0.06 / 12, 60, amount) < budget Then
    rate = 0.06
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_for_loop() {
        let source = r#"
For years = 1 To 30
    payment = Abs(Pmt(0.05 / 12, years * 12, 100000))
    Debug.Print years, payment
Next years
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_do_while() {
        let source = r#"
Do While Abs(Pmt(rate, 360, amount)) > maxPayment
    amount = amount - 1000
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_do_until() {
        let source = r#"
Do Until Abs(Pmt(r / 12, n, principal)) <= affordablePayment
    n = n + 12
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_while_wend() {
        let source = r#"
While Pmt(interestRate, periods, loanAmt) < -500
    periods = periods + 1
Wend
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_parentheses() {
        let source = r#"
Dim result As Double
result = (Pmt(rate, nper, pv))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_iif() {
        let source = r#"
Dim payment As Double
payment = IIf(useBeginning, Pmt(r, n, pv, fv, 1), Pmt(r, n, pv, fv, 0))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_comparison() {
        let source = r#"
If Abs(Pmt(rate1, term1, amt)) < Abs(Pmt(rate2, term2, amt)) Then
    MsgBox "Option 1 has lower payment"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_array_assignment() {
        let source = r#"
Dim payments(10) As Double
payments(i) = Pmt(rates(i), periods, principal)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_property_assignment() {
        let source = r#"
Set obj = New Loan
obj.Payment = Pmt(obj.Rate, obj.Term, obj.Principal)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_function_argument() {
        let source = r#"
Call DisplayLoanInfo(loanAmount, Pmt(monthlyRate, months, loanAmount))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_arithmetic() {
        let source = r#"
Dim totalCost As Double
totalCost = Abs(Pmt(rate, nper, principal)) * nper
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_concatenation() {
        let source = r#"
Dim msg As String
msg = "Your monthly payment is $" & Format(Abs(Pmt(r, n, amt)), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_abs_function() {
        let source = r#"
Dim displayPayment As Double
displayPayment = Abs(Pmt(interestRate / 12, years * 12, loanAmount))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_error_handling() {
        let source = r#"
On Error Resume Next
payment = Pmt(rate, nper, pv, fv, type)
If Err.Number <> 0 Then
    MsgBox "Error calculating payment"
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn pmt_on_error_goto() {
        let source = r#"
Sub CalculateMonthlyPayment()
    On Error GoTo ErrorHandler
    Dim pmt As Double
    pmt = Pmt(monthlyRate, numMonths, loanPrincipal)
    Exit Sub
ErrorHandler:
    MsgBox "Error in payment calculation"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Pmt"));
        assert!(text.contains("Identifier"));
    }
}

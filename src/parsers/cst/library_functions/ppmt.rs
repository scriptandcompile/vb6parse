//! # `PPmt` Function
//!
//! Returns a Double specifying the principal payment for a given period of an annuity based on periodic, fixed payments and a fixed interest rate.
//!
//! ## Syntax
//!
//! ```vb
//! PPmt(rate, per, nper, pv, [fv], [type])
//! ```
//!
//! ## Parameters
//!
//! - `rate` - Required. Double specifying interest rate per period. For example, if you get a car loan at an annual percentage rate (APR) of 10% and make monthly payments, the rate per period is 0.1/12, or 0.0083.
//! - `per` - Required. Integer specifying payment period in the range 1 through nper. This is the specific period for which you want to know the principal payment.
//! - `nper` - Required. Integer specifying total number of payment periods in the annuity. For example, if you make monthly payments on a 4-year car loan, your loan has 4 * 12 (or 48) payment periods.
//! - `pv` - Required. Double specifying present value, or lump sum, that a series of payments to be paid in the future is worth now. For example, when you borrow money to buy a car, the loan amount is the present value to the lender.
//! - `fv` - Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0. If omitted, 0 is assumed.
//! - `type` - Optional. Variant specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.
//!
//! ## Return Value
//!
//! Returns a `Double` specifying the principal payment for the given period. The principal is the portion of the payment that reduces the loan balance, excluding interest.
//!
//! ## Remarks
//!
//! The `PPmt` function is essential for creating amortization schedules and understanding how loans are paid down over time.
//! While the total payment amount stays constant (calculated by `Pmt`), the split between principal and interest changes
//! each period. Early in the loan, most of the payment goes to interest; later, most goes to principal.
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
//! The principal payment returned by `PPmt` plus the interest payment returned by `IPmt` for the same period equals
//! the total payment returned by `Pmt`.
//!
//! **Important**: The `per` parameter must be between 1 and `nper`. The function will raise an error if `per` is
//! outside this range.
//!
//! ## Typical Uses
//!
//! 1. **Amortization Schedules**: Breaking down each payment into principal and interest components
//! 2. **Loan Balance Tracking**: Calculating how much principal is paid in each period
//! 3. **Tax Deduction Analysis**: Separating deductible interest from non-deductible principal
//! 4. **Equity Building**: Tracking home equity growth through principal payments
//! 5. **Refinancing Analysis**: Comparing principal paydown between different loan options
//! 6. **Investment Analysis**: Understanding the principal contribution in annuity investments
//! 7. **Financial Planning**: Projecting debt reduction over time
//! 8. **Prepayment Scenarios**: Analyzing impact of extra principal payments
//!
//! ## Basic Examples
//!
//! ### Example 1: Simple Principal Payment
//! ```vb
//! ' Calculate principal payment in month 12 of a 60-month loan
//! Dim principalPmt As Double
//! principalPmt = PPmt(0.06 / 12, 12, 60, 20000)
//! ' Returns the principal portion of the 12th payment (negative value)
//! ```
//!
//! ### Example 2: First vs Last Payment
//! ```vb
//! ' $200,000 mortgage, 30 years, 4.5% APR
//! Dim firstPrincipal As Double
//! Dim lastPrincipal As Double
//!
//! firstPrincipal = Abs(PPmt(0.045 / 12, 1, 360, 200000))
//! lastPrincipal = Abs(PPmt(0.045 / 12, 360, 360, 200000))
//!
//! ' First payment: mostly interest, small principal
//! ' Last payment: mostly principal, small interest
//! Debug.Print "First payment principal: $" & Format(firstPrincipal, "0.00")
//! Debug.Print "Last payment principal: $" & Format(lastPrincipal, "0.00")
//! ```
//!
//! ### Example 3: Verify Payment Split
//! ```vb
//! ' Verify that PPmt + IPmt = Pmt for any period
//! Dim totalPayment As Double
//! Dim principalPart As Double
//! Dim interestPart As Double
//!
//! totalPayment = Pmt(0.05 / 12, 60, 15000)
//! principalPart = PPmt(0.05 / 12, 24, 60, 15000)
//! interestPart = IPmt(0.05 / 12, 24, 60, 15000)
//!
//! ' principalPart + interestPart should equal totalPayment
//! ```
//!
//! ### Example 4: Calculate Principal in First Year
//! ```vb
//! Function CalculateFirstYearPrincipal(loanAmount As Double, _
//!                                      annualRate As Double, _
//!                                      years As Integer) As Double
//!     Dim month As Integer
//!     Dim totalPrincipal As Double
//!     Dim monthlyRate As Double
//!     
//!     monthlyRate = annualRate / 12
//!     totalPrincipal = 0
//!     
//!     For month = 1 To 12
//!         totalPrincipal = totalPrincipal + PPmt(monthlyRate, month, years * 12, loanAmount)
//!     Next month
//!     
//!     CalculateFirstYearPrincipal = Abs(totalPrincipal)
//! End Function
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `BuildAmortizationSchedule`
//! ```vb
//! Sub BuildAmortizationSchedule(principal As Double, _
//!                               annualRate As Double, _
//!                               years As Integer)
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     Dim period As Integer
//!     Dim payment As Double
//!     Dim principalPmt As Double
//!     Dim interestPmt As Double
//!     Dim balance As Double
//!     
//!     monthlyRate = annualRate / 12
//!     numPayments = years * 12
//!     payment = Abs(Pmt(monthlyRate, numPayments, principal))
//!     balance = principal
//!     
//!     Debug.Print "Period", "Payment", "Principal", "Interest", "Balance"
//!     Debug.Print String(60, "-")
//!     
//!     For period = 1 To numPayments
//!         principalPmt = Abs(PPmt(monthlyRate, period, numPayments, principal))
//!         interestPmt = Abs(IPmt(monthlyRate, period, numPayments, principal))
//!         balance = balance - principalPmt
//!         
//!         Debug.Print period, _
//!                     Format(payment, "0.00"), _
//!                     Format(principalPmt, "0.00"), _
//!                     Format(interestPmt, "0.00"), _
//!                     Format(balance, "0.00")
//!     Next period
//! End Sub
//! ```
//!
//! ### Pattern 2: `CalculatePrincipalPaidRange`
//! ```vb
//! Function CalculatePrincipalPaidRange(rate As Double, _
//!                                      startPeriod As Integer, _
//!                                      endPeriod As Integer, _
//!                                      nper As Integer, _
//!                                      pv As Double) As Double
//!     ' Calculate total principal paid between two periods
//!     Dim period As Integer
//!     Dim totalPrincipal As Double
//!     
//!     totalPrincipal = 0
//!     For period = startPeriod To endPeriod
//!         totalPrincipal = totalPrincipal + PPmt(rate, period, nper, pv)
//!     Next period
//!     
//!     CalculatePrincipalPaidRange = totalPrincipal
//! End Function
//! ```
//!
//! ### Pattern 3: `GetRemainingBalance`
//! ```vb
//! Function GetRemainingBalance(principal As Double, _
//!                              rate As Double, _
//!                              nper As Integer, _
//!                              currentPeriod As Integer) As Double
//!     ' Calculate remaining balance after a specific period
//!     Dim period As Integer
//!     Dim principalPaid As Double
//!     
//!     principalPaid = 0
//!     For period = 1 To currentPeriod
//!         principalPaid = principalPaid + PPmt(rate, period, nper, principal)
//!     Next period
//!     
//!     ' Principal paid is negative, so subtract it (adds to get reduction)
//!     GetRemainingBalance = principal + principalPaid
//! End Function
//! ```
//!
//! ### Pattern 4: `ComparePrincipalPaydown`
//! ```vb
//! Sub ComparePrincipalPaydown(amount As Double)
//!     Dim principal15 As Double
//!     Dim principal30 As Double
//!     Dim year As Integer
//!     
//!     Debug.Print "Year", "15-yr Principal", "30-yr Principal", "Difference"
//!     Debug.Print String(60, "-")
//!     
//!     For year = 1 To 15
//!         ' Calculate total principal paid in this year
//!         principal15 = Abs(CalculatePrincipalPaidRange(0.035 / 12, _
//!                          (year - 1) * 12 + 1, year * 12, 15 * 12, amount))
//!         principal30 = Abs(CalculatePrincipalPaidRange(0.04 / 12, _
//!                          (year - 1) * 12 + 1, year * 12, 30 * 12, amount))
//!         
//!         Debug.Print year, _
//!                     Format(principal15, "#,##0"), _
//!                     Format(principal30, "#,##0"), _
//!                     Format(principal15 - principal30, "#,##0")
//!     Next year
//! End Sub
//! ```
//!
//! ### Pattern 5: `ValidatePPmtParameters`
//! ```vb
//! Function ValidatePPmtParameters(rate As Double, per As Integer, _
//!                                 nper As Integer, pv As Double) As Boolean
//!     ValidatePPmtParameters = False
//!     
//!     If per < 1 Or per > nper Then
//!         MsgBox "Period must be between 1 and " & nper
//!         Exit Function
//!     End If
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
//!     ValidatePPmtParameters = True
//! End Function
//! ```
//!
//! ### Pattern 6: `CalculateEquityGrowth`
//! ```vb
//! Function CalculateEquityGrowth(homeValue As Double, _
//!                                loanAmount As Double, _
//!                                rate As Double, _
//!                                nper As Integer, _
//!                                currentPeriod As Integer) As Double
//!     ' Calculate home equity = home value - remaining loan balance
//!     Dim remainingBalance As Double
//!     Dim equity As Double
//!     
//!     remainingBalance = GetRemainingBalance(loanAmount, rate, nper, currentPeriod)
//!     equity = homeValue - remainingBalance
//!     
//!     CalculateEquityGrowth = equity
//! End Function
//! ```
//!
//! ### Pattern 7: `ExtraPrincipalImpact`
//! ```vb
//! Function CalculateExtraPrincipalImpact(principal As Double, _
//!                                        rate As Double, _
//!                                        nper As Integer, _
//!                                        extraPayment As Double, _
//!                                        ByRef monthsSaved As Integer) As Double
//!     ' Calculate how much faster loan pays off with extra principal
//!     Dim regularPayment As Double
//!     Dim balance As Double
//!     Dim monthlyRate As Double
//!     Dim period As Integer
//!     Dim principalPmt As Double
//!     
//!     monthlyRate = rate
//!     regularPayment = Abs(Pmt(monthlyRate, nper, principal))
//!     balance = principal
//!     period = 0
//!     
//!     Do While balance > 0.01 And period < nper
//!         period = period + 1
//!         principalPmt = Abs(PPmt(monthlyRate, period, nper, principal))
//!         balance = balance - principalPmt - extraPayment
//!     Loop
//!     
//!     monthsSaved = nper - period
//!     CalculateExtraPrincipalImpact = extraPayment * period
//! End Function
//! ```
//!
//! ### Pattern 8: `GetPrincipalPercent`
//! ```vb
//! Function GetPrincipalPercent(rate As Double, _
//!                              period As Integer, _
//!                              nper As Integer, _
//!                              pv As Double) As Double
//!     ' Calculate what percentage of payment is principal
//!     Dim payment As Double
//!     Dim principalPmt As Double
//!     
//!     payment = Abs(Pmt(rate, nper, pv))
//!     principalPmt = Abs(PPmt(rate, period, nper, pv))
//!     
//!     If payment > 0 Then
//!         GetPrincipalPercent = (principalPmt / payment) * 100
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 9: `FindBreakEvenPeriod`
//! ```vb
//! Function FindBreakEvenPeriod(principal As Double, _
//!                              rate As Double, _
//!                              nper As Integer) As Integer
//!     ' Find period where principal payment exceeds interest payment
//!     Dim period As Integer
//!     Dim principalPmt As Double
//!     Dim interestPmt As Double
//!     
//!     For period = 1 To nper
//!         principalPmt = Abs(PPmt(rate, period, nper, principal))
//!         interestPmt = Abs(IPmt(rate, period, nper, principal))
//!         
//!         If principalPmt > interestPmt Then
//!             FindBreakEvenPeriod = period
//!             Exit Function
//!         End If
//!     Next period
//!     
//!     FindBreakEvenPeriod = nper ' Never crossed over
//! End Function
//! ```
//!
//! ### Pattern 10: `ProjectPrincipalPaydown`
//! ```vb
//! Sub ProjectPrincipalPaydown(loanAmount As Double, _
//!                             annualRate As Double, _
//!                             years As Integer)
//!     Dim year As Integer
//!     Dim yearlyPrincipal As Double
//!     Dim cumulativePrincipal As Double
//!     Dim remainingBalance As Double
//!     Dim monthlyRate As Double
//!     Dim numPayments As Integer
//!     
//!     monthlyRate = annualRate / 12
//!     numPayments = years * 12
//!     cumulativePrincipal = 0
//!     
//!     Debug.Print "Year", "Principal Paid", "Cumulative", "Balance"
//!     Debug.Print String(60, "-")
//!     
//!     For year = 1 To years
//!         yearlyPrincipal = Abs(CalculatePrincipalPaidRange(monthlyRate, _
//!                              (year - 1) * 12 + 1, year * 12, numPayments, loanAmount))
//!         cumulativePrincipal = cumulativePrincipal + yearlyPrincipal
//!         remainingBalance = loanAmount - cumulativePrincipal
//!         
//!         Debug.Print year, _
//!                     Format(yearlyPrincipal, "#,##0"), _
//!                     Format(cumulativePrincipal, "#,##0"), _
//!                     Format(remainingBalance, "#,##0")
//!     Next year
//! End Sub
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Complete Amortization Analyzer
//! ```vb
//! ' Comprehensive amortization analysis with principal tracking
//! Class AmortizationAnalyzer
//!     Private m_principal As Double
//!     Private m_annualRate As Double
//!     Private m_years As Integer
//!     Private m_paymentType As Integer
//!     
//!     Public Sub Initialize(principal As Double, annualRate As Double, _
//!                          years As Integer, Optional paymentType As Integer = 0)
//!         m_principal = principal
//!         m_annualRate = annualRate
//!         m_years = years
//!         m_paymentType = paymentType
//!     End Sub
//!     
//!     Public Function GetPaymentSchedule() As Collection
//!         Dim schedule As Collection
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         Dim period As Integer
//!         Dim entry As Object
//!         Dim payment As Double
//!         Dim principalPmt As Double
//!         Dim interestPmt As Double
//!         Dim balance As Double
//!         
//!         Set schedule = New Collection
//!         monthlyRate = m_annualRate / 12
//!         numPayments = m_years * 12
//!         payment = Abs(Pmt(monthlyRate, numPayments, m_principal, 0, m_paymentType))
//!         balance = m_principal
//!         
//!         For period = 1 To numPayments
//!             principalPmt = Abs(PPmt(monthlyRate, period, numPayments, m_principal, 0, m_paymentType))
//!             interestPmt = Abs(IPmt(monthlyRate, period, numPayments, m_principal, 0, m_paymentType))
//!             balance = balance - principalPmt
//!             
//!             Set entry = CreateObject("Scripting.Dictionary")
//!             entry.Add "Period", period
//!             entry.Add "Payment", payment
//!             entry.Add "Principal", principalPmt
//!             entry.Add "Interest", interestPmt
//!             entry.Add "Balance", balance
//!             entry.Add "PrincipalPercent", (principalPmt / payment) * 100
//!             
//!             schedule.Add entry
//!         Next period
//!         
//!         Set GetPaymentSchedule = schedule
//!     End Function
//!     
//!     Public Function GetYearlySummary() As Collection
//!         Dim summary As Collection
//!         Dim year As Integer
//!         Dim startPeriod As Integer
//!         Dim endPeriod As Integer
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         Dim yearlyPrincipal As Double
//!         Dim yearlyInterest As Double
//!         Dim period As Integer
//!         Dim entry As Object
//!         
//!         Set summary = New Collection
//!         monthlyRate = m_annualRate / 12
//!         numPayments = m_years * 12
//!         
//!         For year = 1 To m_years
//!             startPeriod = (year - 1) * 12 + 1
//!             endPeriod = year * 12
//!             
//!             yearlyPrincipal = 0
//!             yearlyInterest = 0
//!             
//!             For period = startPeriod To endPeriod
//!                 yearlyPrincipal = yearlyPrincipal + _
//!                     Abs(PPmt(monthlyRate, period, numPayments, m_principal, 0, m_paymentType))
//!                 yearlyInterest = yearlyInterest + _
//!                     Abs(IPmt(monthlyRate, period, numPayments, m_principal, 0, m_paymentType))
//!             Next period
//!             
//!             Set entry = CreateObject("Scripting.Dictionary")
//!             entry.Add "Year", year
//!             entry.Add "Principal", yearlyPrincipal
//!             entry.Add "Interest", yearlyInterest
//!             entry.Add "Total", yearlyPrincipal + yearlyInterest
//!             
//!             summary.Add entry
//!         Next year
//!         
//!         Set GetYearlySummary = summary
//!     End Function
//!     
//!     Public Function GetBalanceAtPeriod(period As Integer) As Double
//!         Dim p As Integer
//!         Dim principalPaid As Double
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         
//!         monthlyRate = m_annualRate / 12
//!         numPayments = m_years * 12
//!         principalPaid = 0
//!         
//!         For p = 1 To period
//!             principalPaid = principalPaid + _
//!                 Abs(PPmt(monthlyRate, p, numPayments, m_principal, 0, m_paymentType))
//!         Next p
//!         
//!         GetBalanceAtPeriod = m_principal - principalPaid
//!     End Function
//!     
//!     Public Function GenerateReport() As String
//!         Dim report As String
//!         Dim summary As Collection
//!         Dim entry As Object
//!         Dim totalPrincipal As Double
//!         Dim totalInterest As Double
//!         
//!         Set summary = GetYearlySummary()
//!         
//!         report = "Amortization Analysis Report" & vbCrLf
//!         report = report & String(70, "=") & vbCrLf
//!         report = report & "Loan Amount: $" & Format(m_principal, "#,##0.00") & vbCrLf
//!         report = report & "Annual Rate: " & Format(m_annualRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "Term: " & m_years & " years" & vbCrLf
//!         report = report & String(70, "-") & vbCrLf
//!         report = report & "Year   Principal      Interest       Total" & vbCrLf
//!         report = report & String(70, "-") & vbCrLf
//!         
//!         totalPrincipal = 0
//!         totalInterest = 0
//!         
//!         For Each entry In summary
//!             report = report & Format(entry("Year"), "00") & "   $" & _
//!                      Format(entry("Principal"), "#,##0.00") & "   $" & _
//!                      Format(entry("Interest"), "#,##0.00") & "   $" & _
//!                      Format(entry("Total"), "#,##0.00") & vbCrLf
//!             totalPrincipal = totalPrincipal + entry("Principal")
//!             totalInterest = totalInterest + entry("Interest")
//!         Next entry
//!         
//!         report = report & String(70, "-") & vbCrLf
//!         report = report & "Total Principal: $" & Format(totalPrincipal, "#,##0.00") & vbCrLf
//!         report = report & "Total Interest: $" & Format(totalInterest, "#,##0.00")
//!         
//!         GenerateReport = report
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Loan Comparison with Principal Analysis
//! ```vb
//! ' Compare multiple loan options focusing on principal paydown
//! Module LoanPrincipalComparison
//!     Private Type LoanDetails
//!         Name As String
//!         Principal As Double
//!         Rate As Double
//!         Years As Integer
//!     End Type
//!     
//!     Public Function CompareLoans(loans() As LoanDetails, _
//!                                 comparisonYear As Integer) As String
//!         Dim report As String
//!         Dim i As Integer
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         Dim yearPrincipal As Double
//!         Dim yearInterest As Double
//!         Dim balance As Double
//!         Dim period As Integer
//!         
//!         report = "Loan Comparison - Year " & comparisonYear & vbCrLf
//!         report = report & String(80, "=") & vbCrLf
//!         report = report & "Loan              Principal Paid  Interest Paid  " & _
//!                  "Balance        Principal %" & vbCrLf
//!         report = report & String(80, "-") & vbCrLf
//!         
//!         For i = LBound(loans) To UBound(loans)
//!             monthlyRate = loans(i).Rate / 12
//!             numPayments = loans(i).Years * 12
//!             
//!             ' Calculate year totals
//!             yearPrincipal = 0
//!             yearInterest = 0
//!             
//!             For period = (comparisonYear - 1) * 12 + 1 To comparisonYear * 12
//!                 If period <= numPayments Then
//!                     yearPrincipal = yearPrincipal + _
//!                         Abs(PPmt(monthlyRate, period, numPayments, loans(i).Principal))
//!                     yearInterest = yearInterest + _
//!                         Abs(IPmt(monthlyRate, period, numPayments, loans(i).Principal))
//!                 End If
//!             Next period
//!             
//!             ' Get balance at end of year
//!             balance = loans(i).Principal
//!             For period = 1 To comparisonYear * 12
//!                 If period <= numPayments Then
//!                     balance = balance - Abs(PPmt(monthlyRate, period, numPayments, loans(i).Principal))
//!                 End If
//!             Next period
//!             
//!             report = report & Left(loans(i).Name & Space(16), 16) & "  $" & _
//!                      Format(yearPrincipal, "#,##0") & "      $" & _
//!                      Format(yearInterest, "#,##0") & "      $" & _
//!                      Format(balance, "#,##0") & "      " & _
//!                      Format((yearPrincipal / (yearPrincipal + yearInterest)) * 100, "00.0") & _
//!                      "%" & vbCrLf
//!         Next i
//!         
//!         CompareLoans = report
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Equity Builder Tracker
//! ```vb
//! ' Track home equity growth over time
//! Class EquityTracker
//!     Private m_homeValue As Double
//!     Private m_loanAmount As Double
//!     Private m_annualRate As Double
//!     Private m_loanYears As Integer
//!     Private m_appreciationRate As Double
//!     
//!     Public Sub Initialize(homeValue As Double, loanAmount As Double, _
//!                          annualRate As Double, loanYears As Integer, _
//!                          appreciationRate As Double)
//!         m_homeValue = homeValue
//!         m_loanAmount = loanAmount
//!         m_annualRate = annualRate
//!         m_loanYears = loanYears
//!         m_appreciationRate = appreciationRate
//!     End Sub
//!     
//!     Public Function GetEquityAtYear(year As Integer) As Double
//!         Dim appreciatedValue As Double
//!         Dim remainingBalance As Double
//!         Dim period As Integer
//!         Dim monthlyRate As Double
//!         Dim numPayments As Integer
//!         Dim principalPaid As Double
//!         
//!         ' Calculate appreciated home value
//!         appreciatedValue = m_homeValue * ((1 + m_appreciationRate) ^ year)
//!         
//!         ' Calculate remaining loan balance
//!         monthlyRate = m_annualRate / 12
//!         numPayments = m_loanYears * 12
//!         principalPaid = 0
//!         
//!         For period = 1 To year * 12
//!             If period <= numPayments Then
//!                 principalPaid = principalPaid + _
//!                     Abs(PPmt(monthlyRate, period, numPayments, m_loanAmount))
//!             End If
//!         Next period
//!         
//!         remainingBalance = m_loanAmount - principalPaid
//!         
//!         GetEquityAtYear = appreciatedValue - remainingBalance
//!     End Function
//!     
//!     Public Function GetEquityProjection() As Collection
//!         Dim projection As Collection
//!         Dim year As Integer
//!         Dim entry As Object
//!         Dim equity As Double
//!         Dim homeValue As Double
//!         Dim loanBalance As Double
//!         
//!         Set projection = New Collection
//!         
//!         For year = 0 To m_loanYears
//!             equity = GetEquityAtYear(year)
//!             homeValue = m_homeValue * ((1 + m_appreciationRate) ^ year)
//!             loanBalance = homeValue - equity
//!             
//!             Set entry = CreateObject("Scripting.Dictionary")
//!             entry.Add "Year", year
//!             entry.Add "HomeValue", homeValue
//!             entry.Add "LoanBalance", loanBalance
//!             entry.Add "Equity", equity
//!             entry.Add "EquityPercent", (equity / homeValue) * 100
//!             
//!             projection.Add entry
//!         Next year
//!         
//!         Set GetEquityProjection = projection
//!     End Function
//!     
//!     Public Function GenerateEquityReport() As String
//!         Dim report As String
//!         Dim projection As Collection
//!         Dim entry As Object
//!         Dim year As Integer
//!         
//!         Set projection = GetEquityProjection()
//!         
//!         report = "Home Equity Growth Projection" & vbCrLf
//!         report = report & String(70, "=") & vbCrLf
//!         report = report & "Initial Home Value: $" & Format(m_homeValue, "#,##0") & vbCrLf
//!         report = report & "Loan Amount: $" & Format(m_loanAmount, "#,##0") & vbCrLf
//!         report = report & "Appreciation Rate: " & Format(m_appreciationRate * 100, "0.0") & "%" & vbCrLf
//!         report = report & String(70, "-") & vbCrLf
//!         report = report & "Year  Home Value    Loan Balance  Equity        Equity %" & vbCrLf
//!         report = report & String(70, "-") & vbCrLf
//!         
//!         For Each entry In projection
//!             year = entry("Year")
//!             If year Mod 5 = 0 Or year = 1 Then  ' Show every 5 years
//!                 report = report & Format(year, "00") & "    $" & _
//!                          Format(entry("HomeValue"), "#,##0") & "     $" & _
//!                          Format(entry("LoanBalance"), "#,##0") & "     $" & _
//!                          Format(entry("Equity"), "#,##0") & "      " & _
//!                          Format(entry("EquityPercent"), "00.0") & "%" & vbCrLf
//!             End If
//!         Next entry
//!         
//!         GenerateEquityReport = report
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Refinance Decision Tool
//! ```vb
//! ' Analyze refinancing with detailed principal comparison
//! Class RefinanceAnalyzer
//!     Private m_currentBalance As Double
//!     Private m_currentRate As Double
//!     Private m_currentYearsLeft As Integer
//!     Private m_newRate As Double
//!     Private m_newYears As Integer
//!     Private m_closingCosts As Double
//!     
//!     Public Sub SetCurrentLoan(balance As Double, rate As Double, yearsLeft As Integer)
//!         m_currentBalance = balance
//!         m_currentRate = rate
//!         m_currentYearsLeft = yearsLeft
//!     End Sub
//!     
//!     Public Sub SetNewLoan(rate As Double, years As Integer, closingCosts As Double)
//!         m_newRate = rate
//!         m_newYears = years
//!         m_closingCosts = closingCosts
//!     End Sub
//!     
//!     Public Function ComparePrincipalPaydown(years As Integer) As String
//!         Dim report As String
//!         Dim year As Integer
//!         Dim currentPrincipal As Double
//!         Dim newPrincipal As Double
//!         Dim period As Integer
//!         Dim currentMonthlyRate As Double
//!         Dim newMonthlyRate As Double
//!         Dim currentNumPayments As Integer
//!         Dim newNumPayments As Integer
//!         
//!         currentMonthlyRate = m_currentRate / 12
//!         currentNumPayments = m_currentYearsLeft * 12
//!         newMonthlyRate = m_newRate / 12
//!         newNumPayments = m_newYears * 12
//!         
//!         report = "Principal Paydown Comparison" & vbCrLf
//!         report = report & String(60, "=") & vbCrLf
//!         report = report & "Year  Current Loan  New Loan      Difference" & vbCrLf
//!         report = report & String(60, "-") & vbCrLf
//!         
//!         For year = 1 To years
//!             ' Calculate principal paid in this year for current loan
//!             currentPrincipal = 0
//!             For period = (year - 1) * 12 + 1 To year * 12
//!                 If period <= currentNumPayments Then
//!                     currentPrincipal = currentPrincipal + _
//!                         Abs(PPmt(currentMonthlyRate, period, currentNumPayments, m_currentBalance))
//!                 End If
//!             Next period
//!             
//!             ' Calculate principal paid in this year for new loan
//!             newPrincipal = 0
//!             For period = (year - 1) * 12 + 1 To year * 12
//!                 If period <= newNumPayments Then
//!                     newPrincipal = newPrincipal + _
//!                         Abs(PPmt(newMonthlyRate, period, newNumPayments, m_currentBalance + m_closingCosts))
//!                 End If
//!             Next period
//!             
//!             report = report & Format(year, "00") & "    $" & _
//!                      Format(currentPrincipal, "#,##0") & "      $" & _
//!                      Format(newPrincipal, "#,##0") & "      $" & _
//!                      Format(newPrincipal - currentPrincipal, "#,##0") & vbCrLf
//!         Next year
//!         
//!         ComparePrincipalPaydown = report
//!     End Function
//!     
//!     Public Function ShouldRefinance() As Boolean
//!         Dim currentPayment As Double
//!         Dim newPayment As Double
//!         Dim monthlySavings As Double
//!         Dim breakEvenMonths As Integer
//!         
//!         currentPayment = Abs(Pmt(m_currentRate / 12, m_currentYearsLeft * 12, m_currentBalance))
//!         newPayment = Abs(Pmt(m_newRate / 12, m_newYears * 12, m_currentBalance + m_closingCosts))
//!         
//!         monthlySavings = currentPayment - newPayment
//!         
//!         If monthlySavings > 0 Then
//!             breakEvenMonths = m_closingCosts / monthlySavings
//!             ShouldRefinance = (breakEvenMonths <= 36)  ' 3 years or less
//!         Else
//!             ShouldRefinance = False
//!         End If
//!     End Function
//!     
//!     Public Function GenerateAnalysis() As String
//!         Dim analysis As String
//!         Dim currentPayment As Double
//!         Dim newPayment As Double
//!         Dim monthlySavings As Double
//!         Dim breakEvenMonths As Integer
//!         
//!         currentPayment = Abs(Pmt(m_currentRate / 12, m_currentYearsLeft * 12, m_currentBalance))
//!         newPayment = Abs(Pmt(m_newRate / 12, m_newYears * 12, m_currentBalance + m_closingCosts))
//!         monthlySavings = currentPayment - newPayment
//!         
//!         analysis = "Refinance Analysis" & vbCrLf
//!         analysis = analysis & String(50, "=") & vbCrLf
//!         analysis = analysis & "Current Loan:" & vbCrLf
//!         analysis = analysis & "  Balance: $" & Format(m_currentBalance, "#,##0") & vbCrLf
//!         analysis = analysis & "  Rate: " & Format(m_currentRate * 100, "0.00") & "%" & vbCrLf
//!         analysis = analysis & "  Years Left: " & m_currentYearsLeft & vbCrLf
//!         analysis = analysis & "  Payment: $" & Format(currentPayment, "#,##0.00") & vbCrLf
//!         analysis = analysis & String(50, "-") & vbCrLf
//!         analysis = analysis & "New Loan:" & vbCrLf
//!         analysis = analysis & "  Amount: $" & Format(m_currentBalance + m_closingCosts, "#,##0") & vbCrLf
//!         analysis = analysis & "  Rate: " & Format(m_newRate * 100, "0.00") & "%" & vbCrLf
//!         analysis = analysis & "  Term: " & m_newYears & " years" & vbCrLf
//!         analysis = analysis & "  Payment: $" & Format(newPayment, "#,##0.00") & vbCrLf
//!         analysis = analysis & String(50, "-") & vbCrLf
//!         
//!         If monthlySavings > 0 Then
//!             breakEvenMonths = m_closingCosts / monthlySavings
//!             analysis = analysis & "Monthly Savings: $" & Format(monthlySavings, "#,##0.00") & vbCrLf
//!             analysis = analysis & "Break-even: " & breakEvenMonths & " months" & vbCrLf
//!             analysis = analysis & "Recommendation: " & _
//!                        IIf(ShouldRefinance(), "REFINANCE", "DON'T REFINANCE")
//!         Else
//!             analysis = analysis & "New payment is HIGHER - refinancing not recommended"
//!         End If
//!         
//!         GenerateAnalysis = analysis
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `PPmt` function can raise errors in the following situations:
//!
//! - **Invalid Procedure Call (Error 5)**: When:
//!   - `per` is less than 1 or greater than `nper`
//!   - `nper` is 0 or negative
//!   - `rate` is -1 (causes division by zero in the formula)
//! - **Type Mismatch (Error 13)**: When arguments cannot be converted to numeric values
//! - **Overflow (Error 6)**: When calculated values exceed Double range
//!
//! Always validate input parameters:
//!
//! ```vb
//! If per >= 1 And per <= nper Then
//!     principalPayment = PPmt(rate, per, nper, pv)
//! Else
//!     MsgBox "Period must be between 1 and " & nper
//! End If
//! ```
//!
//! ## Performance Considerations
//!
//! - The `PPmt` function is fast for individual period calculations
//! - For complete amortization schedules, calling `PPmt` for every period can be slow
//! - Consider calculating balance iteratively: balance = balance - `PPmt`(...)
//! - Pre-calculate monthly rate and other constants outside loops
//! - For large schedules (hundreds of periods), consider caching results
//!
//! ## Best Practices
//!
//! 1. **Validate Period Range**: Always check that per is between 1 and nper
//! 2. **Match Time Units**: Ensure rate and nper use the same time period
//! 3. **Use with `IPmt`**: Combine `PPmt` and `IPmt` to verify they sum to total payment
//! 4. **Use Absolute Value**: Use `Abs()` when displaying to users
//! 5. **Handle Sign Conventions**: Remember negative = outflow, positive = inflow
//! 6. **Optimize Loops**: Pre-calculate constants before looping through periods
//! 7. **Consider Rounding**: Use proper rounding for financial calculations
//! 8. **Verify Totals**: Sum of all principal payments should equal original loan amount
//! 9. **Document Calculations**: Clearly state assumptions in amortization reports
//! 10. **Test Edge Cases**: Verify behavior at period 1, final period, and 0% rate
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **`PPmt`** | Principal portion of payment | Double (principal amount) | Amortization schedules, equity tracking |
//! | **`IPmt`** | Interest portion of payment | Double (interest amount) | Tax deductions, interest expense tracking |
//! | **Pmt** | Total periodic payment | Double (payment amount) | Loan budgeting, payment calculation |
//! | **PV** | Present value | Double (current value) | Valuation, reverse calculations |
//! | **FV** | Future value | Double (future value) | Investment projections |
//! | **`NPer`** | Number of periods | Double (period count) | Loan term calculation |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VBA and VB6
//! - Behavior is consistent across Windows platforms
//! - Uses standard financial formulas for amortization
//! - For zero interest rates, principal = pv / nper for each period
//! - Maximum precision limited by Double data type
//!
//! ## Limitations
//!
//! - Assumes constant interest rate over entire period
//! - Assumes equal payment amounts (standard amortization)
//! - Does not account for fees, taxes, or insurance
//! - Cannot handle variable rate loans without recalculation
//! - Period must be an integer (no fractional periods)
//! - Does not consider prepayment or extra principal payments
//!
//! ## Related Functions
//!
//! - `IPmt`: Returns the interest payment for a specific period
//! - `Pmt`: Returns the total payment for an annuity
//! - `PV`: Returns the present value of an investment
//! - `FV`: Returns the future value of an investment
//! - `NPer`: Returns the number of periods for an investment
//! - `Rate`: Returns the interest rate per period

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn ppmt_basic() {
        let source = r#"
Dim principalPmt As Double
principalPmt = PPmt(0.06 / 12, 12, 60, 20000)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_with_all_parameters() {
        let source = r#"
Dim principal As Double
principal = PPmt(0.045 / 12, 1, 360, 200000, 0, 0)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_if_statement() {
        let source = r#"
If Abs(PPmt(rate, period, nper, principal)) > threshold Then
    MsgBox "High principal payment"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_function_return() {
        let source = r#"
Function GetPrincipalPayment(per As Integer) As Double
    GetPrincipalPayment = Abs(PPmt(0.05 / 12, per, 60, 15000))
End Function
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_variable_assignment() {
        let source = r#"
Dim principalPortion As Double
principalPortion = PPmt(monthlyRate, period, numPayments, loanAmount)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_msgbox() {
        let source = r#"
MsgBox "Principal: $" & Format(Abs(PPmt(0.06 / 12, 24, 60, 25000)), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_debug_print() {
        let source = r#"
Debug.Print "Period " & per & " Principal: " & PPmt(rate, per, nper, pv)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_select_case() {
        let source = r#"
Dim principal As Double
principal = Abs(PPmt(0.05 / 12, period, 360, loanAmount))
Select Case principal
    Case Is < 100
        category = "Low"
    Case Is < 500
        category = "Medium"
    Case Else
        category = "High"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_class_usage() {
        let source = r#"
Private m_principalPayment As Double

Public Sub CalculateForPeriod(period As Integer)
    m_principalPayment = PPmt(m_rate, period, m_nper, m_pv)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_with_statement() {
        let source = r#"
With amortization
    .PrincipalPmt = PPmt(.Rate, .Period, .NumPayments, .LoanAmount)
End With
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_elseif() {
        let source = r#"
If period < 1 Then
    principal = 0
ElseIf PPmt(rate, period, nper, pv) < -1000 Then
    principal = PPmt(rate, period, nper, pv)
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_for_loop() {
        let source = r#"
For period = 1 To 360
    principalPmt = Abs(PPmt(0.045 / 12, period, 360, 200000))
    totalPrincipal = totalPrincipal + principalPmt
Next period
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_do_while() {
        let source = r#"
Do While Abs(PPmt(rate, period, nper, balance)) < targetPrincipal
    period = period + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_do_until() {
        let source = r#"
Do Until Abs(PPmt(r / 12, p, n, principal)) > minPrincipal
    p = p + 1
Loop
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_while_wend() {
        let source = r#"
While period <= numPeriods
    balance = balance - Abs(PPmt(interestRate, period, numPeriods, loanAmt))
    period = period + 1
Wend
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_parentheses() {
        let source = r#"
Dim result As Double
result = (PPmt(rate, per, nper, pv))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_iif() {
        let source = r#"
Dim principal As Double
principal = IIf(useFV, PPmt(r, p, n, pv, fv), PPmt(r, p, n, pv))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_comparison() {
        let source = r#"
If Abs(PPmt(rate1, per, nper, amt)) > Abs(PPmt(rate2, per, nper, amt)) Then
    MsgBox "Loan 1 pays more principal"
End If
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_array_assignment() {
        let source = r#"
Dim principalPayments(360) As Double
principalPayments(i) = PPmt(rate, i, numPayments, principal)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_property_assignment() {
        let source = r#"
Set obj = New AmortizationSchedule
obj.PrincipalPayment = PPmt(obj.Rate, obj.Period, obj.Term, obj.Amount)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_function_argument() {
        let source = r#"
Call UpdateBalance(currentBalance, PPmt(monthlyRate, month, totalMonths, loanPrincipal))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_arithmetic() {
        let source = r#"
Dim newBalance As Double
newBalance = oldBalance - Abs(PPmt(rate, period, nper, originalAmount))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_concatenation() {
        let source = r#"
Dim msg As String
msg = "Principal payment: $" & Format(Abs(PPmt(r, p, n, amt)), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_abs_function() {
        let source = r#"
Dim displayPrincipal As Double
displayPrincipal = Abs(PPmt(interestRate / 12, period, years * 12, loanAmount))
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_sum_with_ipmt() {
        let source = r#"
Dim totalPayment As Double
totalPayment = PPmt(rate, per, nper, pv) + IPmt(rate, per, nper, pv)
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_error_handling() {
        let source = r#"
On Error Resume Next
principal = PPmt(rate, per, nper, pv, fv, type)
If Err.Number <> 0 Then
    MsgBox "Error calculating principal payment"
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn ppmt_on_error_goto() {
        let source = r#"
Sub CalculatePrincipal()
    On Error GoTo ErrorHandler
    Dim principalPmt As Double
    principalPmt = PPmt(monthlyRate, period, numMonths, loanPrincipal)
    Exit Sub
ErrorHandler:
    MsgBox "Error in principal calculation"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_text("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("PPmt"));
        assert!(text.contains("Identifier"));
    }
}

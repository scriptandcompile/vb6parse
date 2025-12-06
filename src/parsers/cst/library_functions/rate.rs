//! # Rate Function
//!
//! Returns a Double specifying the interest rate per period for an annuity.
//!
//! ## Syntax
//!
//! ```vb
//! Rate(nper, pmt, pv, [fv], [type], [guess])
//! ```
//!
//! ## Parameters
//!
//! - `nper` - Required. Double specifying total number of payment periods in the annuity. For example, if you make monthly payments on a 4-year car loan, your loan has 4 * 12 (or 48) payment periods.
//! - `pmt` - Required. Double specifying payment to be made each period. Payments usually contain principal and interest that doesn't change over the life of the annuity.
//! - `pv` - Required. Double specifying present value, or value today, of a series of future payments or receipts. For example, when you borrow money to buy a car, the loan amount is the present value to the lender of the monthly car payments you will make.
//! - `fv` - Optional. Variant specifying future value or cash balance you want after you've made the final payment. For example, the future value of a loan is $0 because that's its value after the final payment. However, if you want to save $50,000 over 18 years for your child's education, then $50,000 is the future value. If omitted, 0 is assumed.
//! - `type` - Optional. Variant specifying when payments are due. Use 0 if payments are due at the end of the payment period, or use 1 if payments are due at the beginning of the period. If omitted, 0 is assumed.
//! - `guess` - Optional. Variant specifying value you estimate will be returned by Rate. If omitted, guess is 0.1 (10 percent). Rate is calculated by iteration and can have zero or more than one solution.
//!
//! ## Return Value
//!
//! Returns a `Double` specifying the interest rate per period for an annuity. The rate is calculated using an iterative algorithm and is returned as a decimal (e.g., 0.08 for 8%).
//!
//! ## Remarks
//!
//! The `Rate` function calculates the interest rate per period for an annuity based on periodic, fixed payments and a fixed principal. This is useful when you know the loan amount, payment, and term, but need to determine what interest rate is being charged.
//!
//! An annuity is a series of fixed cash payments made over a period of time. An annuity can be a loan (such as a home mortgage) or an investment (such as a monthly savings plan).
//!
//! For all arguments, cash paid out (such as deposits to savings) is represented by negative numbers; cash received (such as dividend checks) is represented by positive numbers.
//!
//! **Important Notes**:
//! - The `Rate` function uses an iterative technique to calculate the interest rate
//! - If `Rate` cannot find a result after 20 iterations, it fails and returns an error
//! - Different values for `guess` can result in different solutions or no solution
//! - The rate returned is per period - multiply by periods per year for annual rate
//! - Be consistent with units: if `nper` is in months, the result is monthly rate
//!
//! **Calculation Method**:
//! The Rate function solves the present value equation for the rate:
//! ```text
//! PV + PMT * ((1 - (1 + rate)^-nper) / rate) + FV / (1 + rate)^nper = 0
//! ```
//!
//! ## Typical Uses
//!
//! 1. **Loan Analysis**: Determine the interest rate on a loan given payment and terms
//! 2. **APR Calculation**: Calculate Annual Percentage Rate from payment information
//! 3. **Investment Returns**: Find the rate of return on an investment
//! 4. **Lease Rate Discovery**: Determine implicit interest rate in a lease
//! 5. **Loan Comparison**: Compare effective rates between different loan offers
//! 6. **Reverse Engineering**: Find the rate when only payment details are known
//! 7. **Financial Planning**: Calculate required rate of return for goals
//! 8. **Credit Card Analysis**: Determine effective rate from minimum payments
//!
//! ## Basic Examples
//!
//! ### Example 1: Find Loan Interest Rate
//! ```vb
//! ' You borrowed $10,000, pay $200/month for 5 years. What's the monthly rate?
//! Dim monthlyRate As Double
//! Dim annualRate As Double
//! monthlyRate = Rate(60, -200, 10000)
//! annualRate = monthlyRate * 12
//! ' monthlyRate ≈ 0.00618 (0.618% per month)
//! ' annualRate ≈ 0.0742 (7.42% APR)
//! ```
//!
//! ### Example 2: Investment Rate of Return
//! ```vb
//! ' Invested $5,000, withdrew $100/month for 5 years, ended with $3,000. What was the rate?
//! Dim monthlyReturn As Double
//! monthlyReturn = Rate(60, 100, -5000, 3000)
//! ' Returns the monthly rate of return
//! ```
//!
//! ### Example 3: Find Rate with Guess
//! ```vb
//! ' Sometimes need to provide a guess to help convergence
//! Dim rate As Double
//! rate = Rate(48, -250, 10000, 0, 0, 0.08)  ' Guess 8% annual (0.08/12 monthly)
//! ```
//!
//! ### Example 4: Annual Rate from Monthly Terms
//! ```vb
//! ' Calculate APR from monthly payment information
//! Dim monthlyRate As Double
//! Dim apr As Double
//! monthlyRate = Rate(360, -1000, 150000)  ' 30-year mortgage
//! apr = monthlyRate * 12 * 100  ' Convert to annual percentage
//! MsgBox "APR: " & Format(apr, "0.00") & "%"
//! ```
//!
//! ## Common Patterns
//!
//! ### Pattern 1: `CalculateAPR`
//! ```vb
//! Function CalculateAPR(loanAmount As Double, monthlyPayment As Double, _
//!                       months As Integer) As Double
//!     ' Calculate Annual Percentage Rate from loan terms
//!     Dim monthlyRate As Double
//!     
//!     On Error Resume Next
//!     monthlyRate = Rate(months, -monthlyPayment, loanAmount)
//!     
//!     If Err.Number = 0 Then
//!         CalculateAPR = monthlyRate * 12  ' Convert to annual rate
//!     Else
//!         CalculateAPR = -1  ' Error indicator
//!         Err.Clear
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 2: `CompareEffectiveRates`
//! ```vb
//! Function CompareEffectiveRates(loan1PV As Double, loan1PMT As Double, loan1Nper As Integer, _
//!                                loan2PV As Double, loan2PMT As Double, loan2Nper As Integer) As String
//!     Dim rate1 As Double, rate2 As Double
//!     
//!     On Error Resume Next
//!     rate1 = Rate(loan1Nper, -loan1PMT, loan1PV) * 12
//!     rate2 = Rate(loan2Nper, -loan2PMT, loan2PV) * 12
//!     
//!     If Err.Number <> 0 Then
//!         CompareEffectiveRates = "Error calculating rates"
//!         Err.Clear
//!         Exit Function
//!     End If
//!     On Error GoTo 0
//!     
//!     If rate1 < rate2 Then
//!         CompareEffectiveRates = "Loan 1 has lower rate: " & Format(rate1 * 100, "0.00") & "%"
//!     ElseIf rate2 < rate1 Then
//!         CompareEffectiveRates = "Loan 2 has lower rate: " & Format(rate2 * 100, "0.00") & "%"
//!     Else
//!         CompareEffectiveRates = "Both loans have same rate: " & Format(rate1 * 100, "0.00") & "%"
//!     End If
//! End Function
//! ```
//!
//! ### Pattern 3: `ValidateRateParameters`
//! ```vb
//! Function ValidateRateParameters(nper As Integer, pmt As Double, pv As Double) As Boolean
//!     ValidateRateParameters = False
//!     
//!     If nper <= 0 Then
//!         MsgBox "Number of periods must be positive"
//!         Exit Function
//!     End If
//!     
//!     If pmt = 0 Then
//!         MsgBox "Payment cannot be zero"
//!         Exit Function
//!     End If
//!     
//!     If pv = 0 Then
//!         MsgBox "Present value cannot be zero"
//!         Exit Function
//!     End If
//!     
//!     ' Check if payment and PV have proper sign relationship
//!     If (pmt > 0 And pv > 0) Or (pmt < 0 And pv < 0) Then
//!         MsgBox "Payment and present value must have opposite signs"
//!         Exit Function
//!     End If
//!     
//!     ValidateRateParameters = True
//! End Function
//! ```
//!
//! ### Pattern 4: `FindRateWithRetry`
//! ```vb
//! Function FindRateWithRetry(nper As Integer, pmt As Double, pv As Double, _
//!                            Optional fv As Double = 0) As Double
//!     ' Try multiple guess values if initial attempt fails
//!     Dim guesses As Variant
//!     Dim i As Integer
//!     Dim rate As Double
//!     
//!     guesses = Array(0.1, 0.05, 0.15, 0.01, 0.2, 0.001)
//!     
//!     For i = LBound(guesses) To UBound(guesses)
//!         On Error Resume Next
//!         rate = Rate(nper, pmt, pv, fv, 0, guesses(i))
//!         
//!         If Err.Number = 0 Then
//!             FindRateWithRetry = rate
//!             Exit Function
//!         End If
//!         
//!         Err.Clear
//!     Next i
//!     
//!     On Error GoTo 0
//!     FindRateWithRetry = -999  ' Error code
//! End Function
//! ```
//!
//! ### Pattern 5: `CalculateEffectiveAPR`
//! ```vb
//! Function CalculateEffectiveAPR(loanAmount As Double, payment As Double, _
//!                                years As Integer, fees As Double) As Double
//!     ' Calculate APR including fees
//!     Dim months As Integer
//!     Dim netLoanAmount As Double
//!     Dim monthlyRate As Double
//!     
//!     months = years * 12
//!     netLoanAmount = loanAmount - fees  ' Reduce by fees paid upfront
//!     
//!     monthlyRate = Rate(months, -payment, netLoanAmount)
//!     CalculateEffectiveAPR = monthlyRate * 12
//! End Function
//! ```
//!
//! ### Pattern 6: `GetLeaseImplicitRate`
//! ```vb
//! Function GetLeaseImplicitRate(vehiclePrice As Double, monthlyPayment As Double, _
//!                               leaseTerm As Integer, residualValue As Double) As Double
//!     ' Find the implicit interest rate in a lease
//!     Dim monthlyRate As Double
//!     
//!     ' For a lease, the PV is the vehicle price, FV is the residual value
//!     monthlyRate = Rate(leaseTerm, -monthlyPayment, vehiclePrice, -residualValue)
//!     GetLeaseImplicitRate = monthlyRate * 12  ' Annual rate
//! End Function
//! ```
//!
//! ### Pattern 7: `CalculateRealRate`
//! ```vb
//! Function CalculateRealRate(nper As Integer, pmt As Double, pv As Double, _
//!                            inflationRate As Double) As Double
//!     ' Calculate real (inflation-adjusted) rate of return
//!     Dim nominalRate As Double
//!     Dim realRate As Double
//!     
//!     nominalRate = Rate(nper, pmt, pv)
//!     
//!     ' Fisher equation: (1 + nominal) = (1 + real)(1 + inflation)
//!     realRate = ((1 + nominalRate) / (1 + inflationRate / 12)) - 1
//!     
//!     CalculateRealRate = realRate
//! End Function
//! ```
//!
//! ### Pattern 8: `ConvertToAPY`
//! ```vb
//! Function ConvertToAPY(periodicRate As Double, periodsPerYear As Integer) As Double
//!     ' Convert periodic rate to Annual Percentage Yield (with compounding)
//!     ConvertToAPY = ((1 + periodicRate) ^ periodsPerYear) - 1
//! End Function
//! ```
//!
//! ### Pattern 9: `BackoutRate`
//! ```vb
//! Function BackoutRate(payment As Double, principal As Double, _
//!                      years As Integer, paymentType As Integer) As Double
//!     ' Reverse engineer the rate from payment information
//!     Dim periods As Integer
//!     Dim rate As Double
//!     
//!     periods = years * 12
//!     
//!     On Error Resume Next
//!     rate = Rate(periods, -payment, principal, 0, paymentType)
//!     
//!     If Err.Number = 0 Then
//!         BackoutRate = rate * 12  ' Annual rate
//!     Else
//!         BackoutRate = -1
//!         Err.Clear
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ### Pattern 10: `IsRateReasonable`
//! ```vb
//! Function IsRateReasonable(calculatedRate As Double) As Boolean
//!     ' Validate that calculated rate is within reasonable bounds
//!     Dim annualRate As Double
//!     
//!     annualRate = calculatedRate * 12
//!     
//!     ' Check if annual rate is between -50% and +50%
//!     IsRateReasonable = (annualRate >= -0.5 And annualRate <= 0.5)
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Comprehensive Loan Rate Analyzer
//! ```vb
//! ' Analyze and compare loan rates with detailed calculations
//! Class LoanRateAnalyzer
//!     Private m_loanAmount As Double
//!     Private m_monthlyPayment As Double
//!     Private m_numPayments As Integer
//!     Private m_fees As Double
//!     Private m_calculatedRate As Double
//!     Private m_effectiveRate As Double
//!     
//!     Public Sub Initialize(loanAmount As Double, monthlyPayment As Double, _
//!                          years As Integer, Optional fees As Double = 0)
//!         m_loanAmount = loanAmount
//!         m_monthlyPayment = monthlyPayment
//!         m_numPayments = years * 12
//!         m_fees = fees
//!     End Sub
//!     
//!     Public Function CalculateNominalRate() As Double
//!         ' Calculate the stated interest rate
//!         Dim monthlyRate As Double
//!         
//!         On Error Resume Next
//!         monthlyRate = Rate(m_numPayments, -m_monthlyPayment, m_loanAmount)
//!         
//!         If Err.Number = 0 Then
//!             m_calculatedRate = monthlyRate * 12
//!             CalculateNominalRate = m_calculatedRate
//!         Else
//!             CalculateNominalRate = -1
//!             Err.Clear
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function CalculateEffectiveRate() As Double
//!         ' Calculate APR including fees
//!         Dim netAmount As Double
//!         Dim monthlyRate As Double
//!         
//!         netAmount = m_loanAmount - m_fees
//!         
//!         On Error Resume Next
//!         monthlyRate = Rate(m_numPayments, -m_monthlyPayment, netAmount)
//!         
//!         If Err.Number = 0 Then
//!             m_effectiveRate = monthlyRate * 12
//!             CalculateEffectiveRate = m_effectiveRate
//!         Else
//!             CalculateEffectiveRate = -1
//!             Err.Clear
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function CalculateAPY() As Double
//!         ' Calculate Annual Percentage Yield (with compounding)
//!         Dim monthlyRate As Double
//!         
//!         monthlyRate = m_calculatedRate / 12
//!         CalculateAPY = ((1 + monthlyRate) ^ 12) - 1
//!     End Function
//!     
//!     Public Function GetTotalInterestPaid() As Double
//!         ' Calculate total interest over life of loan
//!         GetTotalInterestPaid = (m_monthlyPayment * m_numPayments) - m_loanAmount
//!     End Function
//!     
//!     Public Function GetInterestPercentage() As Double
//!         ' Calculate interest as percentage of principal
//!         GetInterestPercentage = GetTotalInterestPaid() / m_loanAmount
//!     End Function
//!     
//!     Public Function GenerateRateReport() As String
//!         Dim report As String
//!         Dim nominalRate As Double
//!         Dim effectiveRate As Double
//!         Dim apy As Double
//!         
//!         nominalRate = CalculateNominalRate()
//!         effectiveRate = CalculateEffectiveRate()
//!         
//!         If nominalRate < 0 Or effectiveRate < 0 Then
//!             GenerateRateReport = "Error: Could not calculate interest rate"
//!             Exit Function
//!         End If
//!         
//!         apy = CalculateAPY()
//!         
//!         report = "Loan Rate Analysis" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Loan Amount: $" & Format(m_loanAmount, "#,##0.00") & vbCrLf
//!         report = report & "Monthly Payment: $" & Format(m_monthlyPayment, "#,##0.00") & vbCrLf
//!         report = report & "Term: " & (m_numPayments / 12) & " years (" & m_numPayments & " months)" & vbCrLf
//!         report = report & "Fees: $" & Format(m_fees, "#,##0.00") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Nominal APR: " & Format(nominalRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "Effective APR (with fees): " & Format(effectiveRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "APY (with compounding): " & Format(apy * 100, "0.00") & "%" & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Total Interest Paid: $" & Format(GetTotalInterestPaid(), "#,##0.00") & vbCrLf
//!         report = report & "Interest as % of Principal: " & Format(GetInterestPercentage() * 100, "0.00") & "%"
//!         
//!         GenerateRateReport = report
//!     End Function
//! End Class
//! ```
//!
//! ### Example 2: Multi-Loan Rate Comparison Tool
//! ```vb
//! ' Compare rates across multiple loan offers
//! Module LoanRateComparison
//!     Private Type LoanOffer
//!         Name As String
//!         Principal As Double
//!         Payment As Double
//!         Months As Integer
//!         Fees As Double
//!         NominalRate As Double
//!         EffectiveRate As Double
//!     End Type
//!     
//!     Public Function CompareLoans(offers() As LoanOffer) As String
//!         Dim i As Integer
//!         Dim report As String
//!         Dim bestOffer As Integer
//!         Dim lowestRate As Double
//!         
//!         lowestRate = 999
//!         bestOffer = LBound(offers)
//!         
//!         ' Calculate rates for all offers
//!         For i = LBound(offers) To UBound(offers)
//!             With offers(i)
//!                 On Error Resume Next
//!                 .NominalRate = Rate(.Months, -.Payment, .Principal) * 12
//!                 .EffectiveRate = Rate(.Months, -.Payment, .Principal - .Fees) * 12
//!                 
//!                 If Err.Number <> 0 Then
//!                     .NominalRate = -1
//!                     .EffectiveRate = -1
//!                     Err.Clear
//!                 End If
//!                 On Error GoTo 0
//!                 
//!                 If .EffectiveRate > 0 And .EffectiveRate < lowestRate Then
//!                     lowestRate = .EffectiveRate
//!                     bestOffer = i
//!                 End If
//!             End With
//!         Next i
//!         
//!         ' Generate comparison report
//!         report = "Loan Offer Comparison" & vbCrLf
//!         report = report & String(80, "=") & vbCrLf
//!         report = report & "Offer                Principal    Payment   Term   Fees      APR      Eff.APR" & vbCrLf
//!         report = report & String(80, "-") & vbCrLf
//!         
//!         For i = LBound(offers) To UBound(offers)
//!             With offers(i)
//!                 report = report & Left(.Name & Space(20), 20)
//!                 report = report & " $" & Right(Space(9) & Format(.Principal, "#,##0"), 9)
//!                 report = report & "  $" & Right(Space(7) & Format(.Payment, "#,##0"), 7)
//!                 report = report & Right(Space(5) & (.Months / 12), 5) & "y"
//!                 report = report & " $" & Right(Space(6) & Format(.Fees, "#,##0"), 6)
//!                 
//!                 If .NominalRate >= 0 Then
//!                     report = report & Right(Space(6) & Format(.NominalRate * 100, "0.00"), 6) & "%"
//!                     report = report & Right(Space(7) & Format(.EffectiveRate * 100, "0.00"), 7) & "%"
//!                 Else
//!                     report = report & "  Error   Error"
//!                 End If
//!                 
//!                 If i = bestOffer Then report = report & " *BEST*"
//!                 report = report & vbCrLf
//!             End With
//!         Next i
//!         
//!         report = report & String(80, "-") & vbCrLf
//!         report = report & "Best Offer: " & offers(bestOffer).Name & _
//!                  " (Effective APR: " & Format(offers(bestOffer).EffectiveRate * 100, "0.00") & "%)"
//!         
//!         CompareLoans = report
//!     End Function
//!     
//!     Public Function CalculateRateDifference(loan1 As LoanOffer, loan2 As LoanOffer) As String
//!         Dim diff As Double
//!         Dim savingsPerMonth As Double
//!         Dim totalSavings As Double
//!         
//!         diff = Abs(loan1.EffectiveRate - loan2.EffectiveRate)
//!         savingsPerMonth = Abs(loan1.Payment - loan2.Payment)
//!         totalSavings = savingsPerMonth * loan1.Months
//!         
//!         CalculateRateDifference = "Rate Difference: " & Format(diff * 100, "0.00") & "%" & vbCrLf & _
//!                                  "Monthly Savings: $" & Format(savingsPerMonth, "#,##0.00") & vbCrLf & _
//!                                  "Total Savings: $" & Format(totalSavings, "#,##0.00")
//!     End Function
//! End Module
//! ```
//!
//! ### Example 3: Investment Rate Calculator
//! ```vb
//! ' Calculate rate of return on investments
//! Class InvestmentRateCalculator
//!     Private m_initialInvestment As Double
//!     Private m_monthlyContribution As Double
//!     Private m_finalValue As Double
//!     Private m_months As Integer
//!     
//!     Public Sub Initialize(initialInvestment As Double, monthlyContribution As Double, _
//!                          finalValue As Double, years As Integer)
//!         m_initialInvestment = initialInvestment
//!         m_monthlyContribution = monthlyContribution
//!         m_finalValue = finalValue
//!         m_months = years * 12
//!     End Sub
//!     
//!     Public Function GetMonthlyRate() As Double
//!         ' Calculate monthly rate of return
//!         On Error Resume Next
//!         GetMonthlyRate = Rate(m_months, -m_monthlyContribution, -m_initialInvestment, m_finalValue)
//!         
//!         If Err.Number <> 0 Then
//!             GetMonthlyRate = -999
//!             Err.Clear
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function GetAnnualRate() As Double
//!         Dim monthlyRate As Double
//!         monthlyRate = GetMonthlyRate()
//!         
//!         If monthlyRate = -999 Then
//!             GetAnnualRate = -999
//!         Else
//!             GetAnnualRate = monthlyRate * 12
//!         End If
//!     End Function
//!     
//!     Public Function GetEffectiveAnnualRate() As Double
//!         ' Calculate with compounding
//!         Dim monthlyRate As Double
//!         monthlyRate = GetMonthlyRate()
//!         
//!         If monthlyRate = -999 Then
//!             GetEffectiveAnnualRate = -999
//!         Else
//!             GetEffectiveAnnualRate = ((1 + monthlyRate) ^ 12) - 1
//!         End If
//!     End Function
//!     
//!     Public Function GetTotalContributed() As Double
//!         GetTotalContributed = m_initialInvestment + (m_monthlyContribution * m_months)
//!     End Function
//!     
//!     Public Function GetTotalReturn() As Double
//!         GetTotalReturn = m_finalValue - GetTotalContributed()
//!     End Function
//!     
//!     Public Function GenerateReport() As String
//!         Dim report As String
//!         Dim annualRate As Double
//!         Dim effectiveRate As Double
//!         
//!         annualRate = GetAnnualRate()
//!         effectiveRate = GetEffectiveAnnualRate()
//!         
//!         If annualRate = -999 Then
//!             GenerateReport = "Error: Could not calculate rate of return"
//!             Exit Function
//!         End If
//!         
//!         report = "Investment Rate of Return Analysis" & vbCrLf
//!         report = report & String(50, "=") & vbCrLf
//!         report = report & "Initial Investment: $" & Format(m_initialInvestment, "#,##0.00") & vbCrLf
//!         report = report & "Monthly Contribution: $" & Format(m_monthlyContribution, "#,##0.00") & vbCrLf
//!         report = report & "Investment Period: " & (m_months / 12) & " years" & vbCrLf
//!         report = report & "Final Value: $" & Format(m_finalValue, "#,##0.00") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Total Contributed: $" & Format(GetTotalContributed(), "#,##0.00") & vbCrLf
//!         report = report & "Total Return: $" & Format(GetTotalReturn(), "#,##0.00") & vbCrLf
//!         report = report & String(50, "-") & vbCrLf
//!         report = report & "Annual Rate of Return: " & Format(annualRate * 100, "0.00") & "%" & vbCrLf
//!         report = report & "Effective Annual Rate: " & Format(effectiveRate * 100, "0.00") & "%"
//!         
//!         GenerateReport = report
//!     End Function
//! End Class
//! ```
//!
//! ### Example 4: Credit Card Rate Analyzer
//! ```vb
//! ' Analyze credit card interest rates from payment information
//! Class CreditCardRateAnalyzer
//!     Private m_balance As Double
//!     Private m_minimumPayment As Double
//!     Private m_monthsToPayoff As Integer
//!     
//!     Public Sub SetCardDetails(balance As Double, minimumPayment As Double, _
//!                              monthsToPayoff As Integer)
//!         m_balance = balance
//!         m_minimumPayment = minimumPayment
//!         m_monthsToPayoff = monthsToPayoff
//!     End Sub
//!     
//!     Public Function GetImplicitRate() As Double
//!         ' Calculate the implicit interest rate
//!         Dim monthlyRate As Double
//!         
//!         On Error Resume Next
//!         monthlyRate = Rate(m_monthsToPayoff, -m_minimumPayment, m_balance)
//!         
//!         If Err.Number = 0 Then
//!             GetImplicitRate = monthlyRate * 12  ' Annual rate
//!         Else
//!             GetImplicitRate = -1
//!             Err.Clear
//!         End If
//!         On Error GoTo 0
//!     End Function
//!     
//!     Public Function GetTotalInterest() As Double
//!         GetTotalInterest = (m_minimumPayment * m_monthsToPayoff) - m_balance
//!     End Function
//!     
//!     Public Function GetInterestAsPercent() As Double
//!         GetInterestAsPercent = GetTotalInterest() / m_balance
//!     End Function
//!     
//!     Public Function CompareToFixedPayment(fixedPayment As Double) As String
//!         Dim result As String
//!         Dim currentRate As Double
//!         Dim fixedMonths As Integer
//!         Dim savings As Double
//!         
//!         currentRate = GetImplicitRate()
//!         
//!         If currentRate < 0 Then
//!             CompareToFixedPayment = "Error calculating current rate"
//!             Exit Function
//!         End If
//!         
//!         ' Calculate months to pay off with fixed payment
//!         fixedMonths = NPer(currentRate / 12, -fixedPayment, m_balance)
//!         savings = (m_minimumPayment * m_monthsToPayoff) - (fixedPayment * fixedMonths)
//!         
//!         result = "Current Plan:" & vbCrLf
//!         result = result & "  Payment: $" & Format(m_minimumPayment, "#,##0.00") & vbCrLf
//!         result = result & "  Months: " & m_monthsToPayoff & vbCrLf
//!         result = result & "  Total: $" & Format(m_minimumPayment * m_monthsToPayoff, "#,##0.00") & vbCrLf
//!         result = result & vbCrLf & "Fixed Payment Plan:" & vbCrLf
//!         result = result & "  Payment: $" & Format(fixedPayment, "#,##0.00") & vbCrLf
//!         result = result & "  Months: " & fixedMonths & vbCrLf
//!         result = result & "  Total: $" & Format(fixedPayment * fixedMonths, "#,##0.00") & vbCrLf
//!         result = result & vbCrLf & "Savings: $" & Format(savings, "#,##0.00")
//!         
//!         CompareToFixedPayment = result
//!     End Function
//! End Class
//! ```
//!
//! ## Error Handling
//!
//! The `Rate` function can raise errors in the following situations:
//!
//! - **Invalid Procedure Call (Error 5)**: When:
//!   - The function cannot find a solution after 20 iterations
//!   - `nper` is 0 or negative
//!   - `pmt` and `pv` have the same sign (both positive or both negative)
//! - **Type Mismatch (Error 13)**: When arguments cannot be converted to numeric values
//! - **Overflow (Error 6)**: When calculated values exceed Double range
//!
//! Always use error handling when calling `Rate`:
//!
//! ```vb
//! On Error Resume Next
//! interestRate = Rate(nper, pmt, pv, fv, type, guess)
//! If Err.Number <> 0 Then
//!     MsgBox "Error calculating rate: " & Err.Description
//!     interestRate = -1  ' Error indicator
//!     Err.Clear
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - The `Rate` function uses an iterative algorithm (Newton-Raphson method)
//! - Each call may require multiple iterations (up to 20) to converge
//! - Providing a good `guess` value can significantly speed up convergence
//! - Poor initial guesses can cause failure to converge or slower performance
//! - Consider caching results if the same parameters are used repeatedly
//!
//! ## Best Practices
//!
//! 1. **Validate Inputs**: Check that payment and PV have opposite signs
//! 2. **Use Error Handling**: Always wrap Rate calls in error handlers
//! 3. **Provide Good Guesses**: Supply reasonable guess values for faster convergence
//! 4. **Retry with Different Guesses**: If Rate fails, try different guess values
//! 5. **Convert to Annual Rate**: Multiply monthly rate by 12 for APR
//! 6. **Check for Reasonableness**: Validate that calculated rate is realistic
//! 7. **Include Fees in APR**: Calculate effective APR by including all fees
//! 8. **Use APY for Compounding**: Calculate APY when showing compound returns
//! 9. **Document Assumptions**: Clearly state what the rate represents
//! 10. **Validate Results**: Verify Rate result by using it in Pmt or PV calculation
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Returns | Use Case |
//! |----------|---------|---------|----------|
//! | **Rate** | Interest rate per period | Double (rate) | Find rate from payment info |
//! | **Pmt** | Payment amount | Double (payment) | Calculate payment from rate |
//! | **PV** | Present value | Double (current value) | Find loan amount from payment |
//! | **FV** | Future value | Double (future value) | Find final value from payments |
//! | **`NPer`** | Number of periods | Double (period count) | Find term from payment/rate |
//! | **IRR** | Internal rate of return | Double (rate) | Find rate from irregular cash flows |
//!
//! ## Platform and Version Notes
//!
//! - Available in all versions of VBA and VB6
//! - Uses iterative algorithm that may not always converge
//! - Maximum 20 iterations before failure
//! - Results may vary slightly between platforms due to floating-point precision
//! - Default guess of 0.1 (10%) works well for most common scenarios
//!
//! ## Limitations
//!
//! - May fail to converge for some parameter combinations
//! - Limited to 20 iterations maximum
//! - Cannot handle multiple solutions (returns first solution found)
//! - Sensitive to initial guess value
//! - May return unrealistic rates if inputs are invalid
//! - Cannot handle variable rate scenarios
//! - Assumes constant periodic payments
//!
//! ## Related Functions
//!
//! - `Pmt`: Returns the periodic payment for an annuity
//! - `PV`: Returns the present value of an annuity
//! - `FV`: Returns the future value of an annuity
//! - `NPer`: Returns the number of periods for an annuity
//! - `IRR`: Returns internal rate of return for irregular cash flows
//! - `MIRR`: Returns modified internal rate of return

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn rate_basic() {
        let source = r#"
Dim interestRate As Double
interestRate = Rate(60, -200, 10000)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_with_all_parameters() {
        let source = r#"
Dim rate As Double
rate = Rate(48, -250, 10000, 0, 0, 0.08)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_if_statement() {
        let source = r#"
If Rate(nper, pmt, pv) > 0.06 Then
    MsgBox "Rate is above 6%"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_function_return() {
        let source = r#"
Function CalculateAPR(loan As Double, payment As Double) As Double
    Dim monthlyRate As Double
    monthlyRate = Rate(60, -payment, loan)
    CalculateAPR = monthlyRate * 12
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_variable_assignment() {
        let source = r#"
Dim monthlyRate As Double
monthlyRate = Rate(months, payment, principal)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_msgbox() {
        let source = r#"
MsgBox "APR: " & Format(Rate(60, -500, 20000) * 12 * 100, "0.00") & "%"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_debug_print() {
        let source = r#"
Debug.Print "Monthly Rate: " & Rate(n, p, pv)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_select_case() {
        let source = r#"
Dim apr As Double
apr = Rate(months, -payment, loan) * 12
Select Case apr
    Case Is > 0.1
        category = "High"
    Case Is > 0.05
        category = "Medium"
    Case Else
        category = "Low"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_class_usage() {
        let source = r#"
Private m_rate As Double

Public Sub Calculate()
    m_rate = Rate(m_nper, m_pmt, m_pv)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_with_statement() {
        let source = r#"
With loanCalc
    .InterestRate = Rate(.NumPayments, .Payment, .Principal)
End With
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_elseif() {
        let source = r#"
If amount > 100000 Then
    r = 0.05
ElseIf Rate(60, -pmt, loan) > 0.08 Then
    r = 0.06
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_for_loop() {
        let source = r#"
For payment = 100 To 500 Step 50
    r = Rate(60, -payment, 10000)
    Debug.Print payment, r * 12
Next payment
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_do_while() {
        let source = r#"
Do While Rate(nper, -pmt, pv) < targetRate
    pmt = pmt + 10
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_do_until() {
        let source = r#"
Do Until Rate(n, -p, amount) >= minRate
    amount = amount - 100
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_while_wend() {
        let source = r#"
While Rate(periods, payment, principal) > 0
    principal = principal + 1000
Wend
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_parentheses() {
        let source = r#"
Dim result As Double
result = (Rate(nper, pmt, pv))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_iif() {
        let source = r#"
Dim annualRate As Double
annualRate = IIf(useGuess, Rate(n, p, pv, fv, t, g), Rate(n, p, pv))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_comparison() {
        let source = r#"
If Rate(60, -200, 10000) > Rate(48, -250, 10000) Then
    MsgBox "First loan has higher rate"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_array_assignment() {
        let source = r#"
Dim rates(10) As Double
rates(i) = Rate(periods(i), payments(i), principals(i))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_property_assignment() {
        let source = r#"
Set obj = New LoanAnalyzer
obj.InterestRate = Rate(obj.Term, obj.Payment, obj.Principal)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_function_argument() {
        let source = r#"
Call AnalyzeLoan(Rate(60, -payment, principal) * 12, loanAmount)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_arithmetic() {
        let source = r#"
Dim apr As Double
apr = Rate(months, -pmt, pv) * 12 * 100
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_concatenation() {
        let source = r#"
Dim msg As String
msg = "APR is: " & Format(Rate(n, p, pv) * 12, "0.00%")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_format_function() {
        let source = r#"
Dim display As String
display = Format(Rate(60, -500, 20000) * 12 * 100, "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_error_handling() {
        let source = r#"
On Error Resume Next
interestRate = Rate(nper, pmt, pv, fv, type, guess)
If Err.Number <> 0 Then
    MsgBox "Error calculating rate"
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn rate_on_error_goto() {
        let source = r#"
Sub CalculateRate()
    On Error GoTo ErrorHandler
    Dim r As Double
    r = Rate(numMonths, monthlyPayment, loanAmount)
    Exit Sub
ErrorHandler:
    MsgBox "Error in rate calculation"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("Rate"));
        assert!(text.contains("Identifier"));
    }
}

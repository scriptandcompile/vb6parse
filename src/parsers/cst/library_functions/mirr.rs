//! # MIRR Function
//!
//! Returns a Double specifying the modified internal rate of return for a series of periodic
//! cash flows (payments and receipts).
//!
//! ## Syntax
//!
//! ```vb
//! MIRR(values(), finance_rate, reinvest_rate)
//! ```
//!
//! ## Parameters
//!
//! - **values()** (Required) - Array of Double specifying cash flow values. The array must contain
//!   at least one negative value (payment) and one positive value (receipt).
//! - **finance_rate** (Required) - Double specifying the interest rate paid as the cost of financing.
//! - **reinvest_rate** (Required) - Double specifying the interest rate received on gains from cash
//!   reinvestment.
//!
//! ## Return Value
//!
//! Returns a **Variant (Double)** representing the modified internal rate of return, expressed as
//! a decimal per period. For example, 0.1 represents 10%.
//!
//! ## Remarks
//!
//! The Modified Internal Rate of Return (MIRR) is a variation of the internal rate of return (IRR)
//! that addresses some of IRR's limitations. Unlike IRR, MIRR assumes:
//! - Negative cash flows (investments) are financed at the finance_rate
//! - Positive cash flows (returns) are reinvested at the reinvest_rate
//!
//! This makes MIRR more realistic than IRR for most real-world investment scenarios where the
//! cost of capital and reinvestment rate differ.
//!
//! ### Key Characteristics:
//! - Returns a rate per period (e.g., if values are monthly cash flows, result is monthly rate)
//! - To convert to annual percentage: multiply by periods per year and by 100
//! - Values must include at least one positive and one negative value
//! - Values are assumed to occur at regular intervals (end of each period)
//! - First value occurs at end of first period (not at time zero like NPV)
//! - Error 5 (Invalid procedure call) if values array contains no positive or no negative values
//! - Finance_rate and reinvest_rate should be expressed as decimals (e.g., 0.1 for 10%)
//!
//! ### MIRR vs IRR:
//! - **IRR** assumes all cash flows are reinvested at the IRR itself (often unrealistic)
//! - **MIRR** allows separate rates for financing costs and reinvestment gains (more realistic)
//! - **MIRR** is generally more stable and easier to interpret than IRR
//! - **MIRR** always returns a single value, while IRR can have multiple solutions
//!
//! ### When to Use:
//! - Evaluating capital investments with different borrowing and reinvestment rates
//! - Comparing mutually exclusive projects with different cash flow patterns
//! - Analysis of real estate investments, business projects, or equipment purchases
//! - Any scenario where the cost of capital differs from the expected return on reinvestment
//! - When you need a more conservative and realistic return measure than IRR
//!
//! ## Typical Uses
//!
//! 1. **Capital Budgeting** - Evaluate investment projects with realistic rate assumptions
//! 2. **Real Estate Analysis** - Calculate returns on property investments
//! 3. **Equipment Purchase Decisions** - Assess whether to buy or lease equipment
//! 4. **Business Valuation** - Determine value of business investments
//! 5. **Portfolio Management** - Evaluate investment performance with actual reinvestment rates
//! 6. **Loan vs Investment Comparison** - Compare cost of financing to investment returns
//! 7. **Project Ranking** - Rank multiple projects by profitability
//! 8. **Sensitivity Analysis** - Test how changes in rates affect investment viability
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple investment analysis
//! Dim cashFlows(4) As Double
//! cashFlows(0) = -100000  ' Initial investment
//! cashFlows(1) = 30000    ' Year 1 return
//! cashFlows(2) = 35000    ' Year 2 return
//! cashFlows(3) = 40000    ' Year 3 return
//! cashFlows(4) = 25000    ' Year 4 return
//!
//! Dim financeRate As Double
//! Dim reinvestRate As Double
//! Dim result As Double
//!
//! financeRate = 0.08   ' 8% cost of capital
//! reinvestRate = 0.05  ' 5% reinvestment rate
//! result = MIRR(cashFlows(), financeRate, reinvestRate)
//! ' Result is approximately 0.1065 (10.65% annual return)
//! ```
//!
//! ```vb
//! ' Example 2: Real estate investment
//! Dim propertyFlows(9) As Double
//! Dim i As Integer
//!
//! propertyFlows(0) = -500000  ' Purchase price
//! For i = 1 To 8
//!     propertyFlows(i) = 50000  ' Annual rental income
//! Next i
//! propertyFlows(9) = 600000   ' Sale price in year 9
//!
//! Dim mortgageRate As Double
//! Dim marketRate As Double
//!
//! mortgageRate = 0.06   ' 6% mortgage cost
//! marketRate = 0.04     ' 4% market return
//!
//! If MIRR(propertyFlows(), mortgageRate, marketRate) > mortgageRate Then
//!     MsgBox "Property investment beats financing cost"
//! End If
//! ```
//!
//! ```vb
//! ' Example 3: Monthly cash flows (convert to annual)
//! Dim monthlyCashFlows(11) As Double
//! monthlyCashFlows(0) = -10000
//! ' ... populate remaining months
//!
//! Dim monthlyFinanceRate As Double
//! Dim monthlyReinvestRate As Double
//!
//! monthlyFinanceRate = 0.08 / 12    ' Annual rate / 12
//! monthlyReinvestRate = 0.05 / 12   ' Annual rate / 12
//!
//! Dim monthlyReturn As Double
//! Dim annualReturn As Double
//!
//! monthlyReturn = MIRR(monthlyCashFlows(), monthlyFinanceRate, monthlyReinvestRate)
//! annualReturn = (1 + monthlyReturn) ^ 12 - 1  ' Convert to annual
//! ```
//!
//! ```vb
//! ' Example 4: Comparing two projects
//! Dim projectA(3) As Double
//! Dim projectB(3) As Double
//!
//! projectA(0) = -50000: projectA(1) = 20000
//! projectA(2) = 25000: projectA(3) = 30000
//!
//! projectB(0) = -50000: projectB(1) = 15000
//! projectB(2) = 20000: projectB(3) = 40000
//!
//! Dim costOfCapital As Double
//! Dim reinvestmentRate As Double
//!
//! costOfCapital = 0.1
//! reinvestmentRate = 0.06
//!
//! Dim mirrA As Double, mirrB As Double
//! mirrA = MIRR(projectA(), costOfCapital, reinvestmentRate)
//! mirrB = MIRR(projectB(), costOfCapital, reinvestmentRate)
//!
//! If mirrA > mirrB Then
//!     Debug.Print "Project A has better MIRR"
//! Else
//!     Debug.Print "Project B has better MIRR"
//! End If
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Safe MIRR calculation with validation
//! Function SafeMIRR(cashFlows() As Double, finRate As Double, _
//!                   reinvRate As Double) As Variant
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
//!     If Not hasPositive Or Not hasNegative Then
//!         SafeMIRR = Null
//!         Exit Function
//!     End If
//!     
//!     On Error Resume Next
//!     SafeMIRR = MIRR(cashFlows(), finRate, reinvRate)
//!     If Err.Number <> 0 Then
//!         SafeMIRR = Null
//!     End If
//!     On Error GoTo 0
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Convert MIRR to annualized percentage
//! Function AnnualizedMIRR(cashFlows() As Double, finRate As Double, _
//!                         reinvRate As Double, periodsPerYear As Integer) As Double
//!     Dim periodMIRR As Double
//!     periodMIRR = MIRR(cashFlows(), finRate, reinvRate)
//!     
//!     ' Convert to effective annual rate
//!     AnnualizedMIRR = ((1 + periodMIRR) ^ periodsPerYear - 1) * 100
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 3: MIRR-based investment decision
//! Function ShouldInvest(cashFlows() As Double, finRate As Double, _
//!                       reinvRate As Double, hurdle As Double) As Boolean
//!     Dim projectMIRR As Double
//!     projectMIRR = MIRR(cashFlows(), finRate, reinvRate)
//!     ShouldInvest = (projectMIRR >= hurdle)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 4: Sensitivity analysis on rates
//! Sub AnalyzeMIRRSensitivity(cashFlows() As Double)
//!     Dim finRate As Double
//!     Dim reinvRate As Double
//!     Dim result As Double
//!     
//!     Debug.Print "Finance Rate", "Reinvest Rate", "MIRR"
//!     
//!     For finRate = 0.05 To 0.15 Step 0.01
//!         For reinvRate = 0.03 To 0.1 Step 0.01
//!             result = MIRR(cashFlows(), finRate, reinvRate)
//!             Debug.Print Format(finRate, "0.00%"), _
//!                        Format(reinvRate, "0.00%"), _
//!                        Format(result, "0.00%")
//!         Next reinvRate
//!     Next finRate
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 5: Break-even analysis
//! Function BreakEvenFinanceRate(cashFlows() As Double, _
//!                               targetMIRR As Double, _
//!                               reinvRate As Double) As Double
//!     Dim lowRate As Double, highRate As Double
//!     Dim midRate As Double, testMIRR As Double
//!     Dim tolerance As Double
//!     
//!     lowRate = 0
//!     highRate = 1
//!     tolerance = 0.0001
//!     
//!     ' Binary search for break-even rate
//!     Do While (highRate - lowRate) > tolerance
//!         midRate = (lowRate + highRate) / 2
//!         testMIRR = MIRR(cashFlows(), midRate, reinvRate)
//!         
//!         If testMIRR > targetMIRR Then
//!             lowRate = midRate
//!         Else
//!             highRate = midRate
//!         End If
//!     Loop
//!     
//!     BreakEvenFinanceRate = midRate
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Compare MIRR to IRR
//! Sub CompareMIRRtoIRR(cashFlows() As Double, finRate As Double, reinvRate As Double)
//!     Dim irrValue As Double
//!     Dim mirrValue As Double
//!     
//!     irrValue = IRR(cashFlows())
//!     mirrValue = MIRR(cashFlows(), finRate, reinvRate)
//!     
//!     Debug.Print "IRR: " & Format(irrValue * 100, "0.00") & "%"
//!     Debug.Print "MIRR: " & Format(mirrValue * 100, "0.00") & "%"
//!     Debug.Print "Difference: " & Format((irrValue - mirrValue) * 100, "0.00") & "%"
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 7: Multi-year project evaluation
//! Function EvaluateMultiYearProject(initialInvestment As Double, _
//!                                   annualReturns() As Double, _
//!                                   salvageValue As Double, _
//!                                   finRate As Double, _
//!                                   reinvRate As Double) As String
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     Dim years As Integer
//!     
//!     years = UBound(annualReturns) - LBound(annualReturns) + 1
//!     ReDim cashFlows(0 To years)
//!     
//!     cashFlows(0) = -initialInvestment
//!     For i = LBound(annualReturns) To UBound(annualReturns)
//!         cashFlows(i - LBound(annualReturns) + 1) = annualReturns(i)
//!     Next i
//!     cashFlows(years) = cashFlows(years) + salvageValue
//!     
//!     Dim projectMIRR As Double
//!     projectMIRR = MIRR(cashFlows(), finRate, reinvRate)
//!     
//!     EvaluateMultiYearProject = "Project MIRR: " & _
//!         Format(projectMIRR * 100, "0.00") & "%"
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: Ranking multiple investments
//! Type Investment
//!     Name As String
//!     CashFlows() As Double
//!     MIRR As Double
//! End Type
//!
//! Function RankInvestments(investments() As Investment, _
//!                         finRate As Double, _
//!                         reinvRate As Double) As Investment()
//!     Dim i As Integer, j As Integer
//!     Dim temp As Investment
//!     
//!     ' Calculate MIRR for each investment
//!     For i = LBound(investments) To UBound(investments)
//!         investments(i).MIRR = MIRR(investments(i).CashFlows(), finRate, reinvRate)
//!     Next i
//!     
//!     ' Bubble sort by MIRR (descending)
//!     For i = LBound(investments) To UBound(investments) - 1
//!         For j = i + 1 To UBound(investments)
//!             If investments(j).MIRR > investments(i).MIRR Then
//!                 temp = investments(i)
//!                 investments(i) = investments(j)
//!                 investments(j) = temp
//!             End If
//!         Next j
//!     Next i
//!     
//!     RankInvestments = investments
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: Net Present Value equivalent using MIRR
//! Function MIRRtoNPV(initialInvestment As Double, mirrRate As Double, _
//!                    periods As Integer) As Double
//!     ' Convert MIRR back to equivalent NPV
//!     ' Useful for comparing MIRR-based and NPV-based analyses
//!     MIRRtoNPV = initialInvestment * ((1 + mirrRate) ^ periods)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Quarterly to annual MIRR conversion
//! Function QuarterlyToAnnualMIRR(quarterlyCashFlows() As Double, _
//!                                quarterlyFinRate As Double, _
//!                                quarterlyReinvRate As Double) As Double
//!     Dim quarterlyMIRR As Double
//!     
//!     quarterlyMIRR = MIRR(quarterlyCashFlows(), quarterlyFinRate, quarterlyReinvRate)
//!     
//!     ' Convert quarterly rate to effective annual rate
//!     QuarterlyToAnnualMIRR = (1 + quarterlyMIRR) ^ 4 - 1
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Investment Analysis Class
//!
//! ```vb
//! ' Class: InvestmentAnalyzer
//! ' Provides comprehensive investment analysis using MIRR
//!
//! Option Explicit
//!
//! Private m_cashFlows() As Double
//! Private m_financeRate As Double
//! Private m_reinvestRate As Double
//! Private m_periods As Integer
//!
//! Public Sub Initialize(cashFlows() As Double, finRate As Double, reinvRate As Double)
//!     Dim i As Integer
//!     m_periods = UBound(cashFlows) - LBound(cashFlows) + 1
//!     ReDim m_cashFlows(0 To m_periods - 1)
//!     
//!     For i = 0 To m_periods - 1
//!         m_cashFlows(i) = cashFlows(LBound(cashFlows) + i)
//!     Next i
//!     
//!     m_financeRate = finRate
//!     m_reinvestRate = reinvRate
//! End Sub
//!
//! Public Function GetMIRR() As Double
//!     GetMIRR = MIRR(m_cashFlows(), m_financeRate, m_reinvestRate)
//! End Function
//!
//! Public Function GetIRR() As Double
//!     GetIRR = IRR(m_cashFlows())
//! End Function
//!
//! Public Function GetNPV(discountRate As Double) As Double
//!     GetNPV = NPV(discountRate, m_cashFlows())
//! End Function
//!
//! Public Function GetPaybackPeriod() As Integer
//!     Dim cumulative As Double
//!     Dim i As Integer
//!     
//!     cumulative = 0
//!     For i = 0 To m_periods - 1
//!         cumulative = cumulative + m_cashFlows(i)
//!         If cumulative >= 0 Then
//!             GetPaybackPeriod = i + 1
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     GetPaybackPeriod = -1 ' Never pays back
//! End Function
//!
//! Public Function GenerateReport() As String
//!     Dim report As String
//!     report = "Investment Analysis Report" & vbCrLf
//!     report = report & String(40, "-") & vbCrLf
//!     report = report & "Periods: " & m_periods & vbCrLf
//!     report = report & "Finance Rate: " & Format(m_financeRate * 100, "0.00") & "%" & vbCrLf
//!     report = report & "Reinvest Rate: " & Format(m_reinvestRate * 100, "0.00") & "%" & vbCrLf
//!     report = report & vbCrLf
//!     report = report & "MIRR: " & Format(GetMIRR() * 100, "0.00") & "%" & vbCrLf
//!     report = report & "IRR: " & Format(GetIRR() * 100, "0.00") & "%" & vbCrLf
//!     report = report & "NPV @ Finance Rate: " & _
//!         Format(GetNPV(m_financeRate), "$#,##0.00") & vbCrLf
//!     report = report & "Payback Period: " & GetPaybackPeriod() & " periods" & vbCrLf
//!     
//!     GenerateReport = report
//! End Function
//!
//! Public Function IsViable(hurdleRate As Double) As Boolean
//!     IsViable = (GetMIRR() >= hurdleRate)
//! End Function
//! ```
//!
//! ### Example 2: Real Estate Investment Calculator
//!
//! ```vb
//! ' Class: RealEstateInvestment
//! ' Specialized calculator for real estate MIRR analysis
//!
//! Option Explicit
//!
//! Private m_purchasePrice As Double
//! Private m_downPayment As Double
//! Private m_annualRent As Double
//! Private m_holdingPeriod As Integer
//! Private m_appreciationRate As Double
//! Private m_mortgageRate As Double
//! Private m_reinvestmentRate As Double
//!
//! Public Sub ConfigureInvestment(purchasePrice As Double, downPayment As Double, _
//!                                annualRent As Double, holdingYears As Integer, _
//!                                appreciation As Double)
//!     m_purchasePrice = purchasePrice
//!     m_downPayment = downPayment
//!     m_annualRent = annualRent
//!     m_holdingPeriod = holdingYears
//!     m_appreciationRate = appreciation
//! End Sub
//!
//! Public Sub ConfigureRates(mortgageRate As Double, reinvestmentRate As Double)
//!     m_mortgageRate = mortgageRate
//!     m_reinvestmentRate = reinvestmentRate
//! End Sub
//!
//! Public Function CalculateMIRR() As Double
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     Dim salePrice As Double
//!     
//!     ReDim cashFlows(0 To m_holdingPeriod)
//!     
//!     ' Initial investment (down payment)
//!     cashFlows(0) = -m_downPayment
//!     
//!     ' Annual rental income
//!     For i = 1 To m_holdingPeriod - 1
//!         cashFlows(i) = m_annualRent
//!     Next i
//!     
//!     ' Final year: rent + sale proceeds
//!     salePrice = m_purchasePrice * ((1 + m_appreciationRate) ^ m_holdingPeriod)
//!     cashFlows(m_holdingPeriod) = m_annualRent + (salePrice - (m_purchasePrice - m_downPayment))
//!     
//!     CalculateMIRR = MIRR(cashFlows(), m_mortgageRate, m_reinvestmentRate)
//! End Function
//!
//! Public Function GetAnnualizedReturn() As String
//!     Dim mirrValue As Double
//!     mirrValue = CalculateMIRR()
//!     GetAnnualizedReturn = Format(mirrValue * 100, "0.00") & "% per year"
//! End Function
//!
//! Public Function CompareToStocks(stockMarketReturn As Double) As String
//!     Dim propertyReturn As Double
//!     propertyReturn = CalculateMIRR()
//!     
//!     If propertyReturn > stockMarketReturn Then
//!         CompareToStocks = "Property investment outperforms stocks by " & _
//!             Format((propertyReturn - stockMarketReturn) * 100, "0.00") & "%"
//!     Else
//!         CompareToStocks = "Stocks outperform property by " & _
//!             Format((stockMarketReturn - propertyReturn) * 100, "0.00") & "%"
//!     End If
//! End Function
//! ```
//!
//! ### Example 3: Project Portfolio Optimizer
//!
//! ```vb
//! ' Module: PortfolioOptimizer
//! ' Selects optimal mix of projects given budget constraint
//!
//! Option Explicit
//!
//! Type Project
//!     ID As String
//!     Name As String
//!     InitialCost As Double
//!     CashFlows() As Double
//!     MIRR As Double
//!     Selected As Boolean
//! End Type
//!
//! Function OptimizePortfolio(projects() As Project, budget As Double, _
//!                           finRate As Double, reinvRate As Double) As Project()
//!     Dim i As Integer, j As Integer
//!     Dim totalCost As Double
//!     Dim temp As Project
//!     
//!     ' Calculate MIRR for each project
//!     For i = LBound(projects) To UBound(projects)
//!         projects(i).MIRR = MIRR(projects(i).CashFlows(), finRate, reinvRate)
//!         projects(i).Selected = False
//!     Next i
//!     
//!     ' Sort by MIRR descending (greedy approach)
//!     For i = LBound(projects) To UBound(projects) - 1
//!         For j = i + 1 To UBound(projects)
//!             If projects(j).MIRR > projects(i).MIRR Then
//!                 temp = projects(i)
//!                 projects(i) = projects(j)
//!                 projects(j) = temp
//!             End If
//!         Next j
//!     Next i
//!     
//!     ' Select projects until budget exhausted
//!     totalCost = 0
//!     For i = LBound(projects) To UBound(projects)
//!         If totalCost + projects(i).InitialCost <= budget Then
//!             projects(i).Selected = True
//!             totalCost = totalCost + projects(i).InitialCost
//!         End If
//!     Next i
//!     
//!     OptimizePortfolio = projects
//! End Function
//!
//! Function GetPortfolioMIRR(projects() As Project) As Double
//!     Dim combinedFlows() As Double
//!     Dim maxPeriods As Integer
//!     Dim i As Integer, p As Integer
//!     
//!     ' Find maximum periods across all selected projects
//!     For i = LBound(projects) To UBound(projects)
//!         If projects(i).Selected Then
//!             If UBound(projects(i).CashFlows) > maxPeriods Then
//!                 maxPeriods = UBound(projects(i).CashFlows)
//!             End If
//!         End If
//!     Next i
//!     
//!     ReDim combinedFlows(0 To maxPeriods)
//!     
//!     ' Combine cash flows from all selected projects
//!     For i = LBound(projects) To UBound(projects)
//!         If projects(i).Selected Then
//!             For p = 0 To UBound(projects(i).CashFlows)
//!                 combinedFlows(p) = combinedFlows(p) + projects(i).CashFlows(p)
//!             Next p
//!         End If
//!     Next i
//!     
//!     GetPortfolioMIRR = MIRR(combinedFlows(), 0.08, 0.05) ' Example rates
//! End Function
//! ```
//!
//! ### Example 4: Monte Carlo Simulation with MIRR
//!
//! ```vb
//! ' Module: MIRRSimulation
//! ' Performs Monte Carlo simulation on MIRR with uncertain cash flows
//!
//! Option Explicit
//!
//! Function SimulateMIRR(baseCashFlows() As Double, volatility As Double, _
//!                       finRate As Double, reinvRate As Double, _
//!                       simulations As Long) As Double()
//!     Dim results() As Double
//!     Dim sim As Long, i As Integer
//!     Dim simulatedFlows() As Double
//!     Dim periods As Integer
//!     
//!     periods = UBound(baseCashFlows) - LBound(baseCashFlows) + 1
//!     ReDim results(1 To simulations)
//!     ReDim simulatedFlows(0 To periods - 1)
//!     
//!     Randomize Timer
//!     
//!     For sim = 1 To simulations
//!         ' Generate random cash flows based on volatility
//!         For i = 0 To periods - 1
//!             Dim randomFactor As Double
//!             randomFactor = 1 + (Rnd() - 0.5) * 2 * volatility
//!             simulatedFlows(i) = baseCashFlows(i) * randomFactor
//!         Next i
//!         
//!         On Error Resume Next
//!         results(sim) = MIRR(simulatedFlows(), finRate, reinvRate)
//!         If Err.Number <> 0 Then
//!             results(sim) = 0 ' Invalid scenario
//!         End If
//!         On Error GoTo 0
//!     Next sim
//!     
//!     SimulateMIRR = results
//! End Function
//!
//! Function AnalyzeSimulationResults(results() As Double) As String
//!     Dim i As Long
//!     Dim sum As Double, sumSq As Double
//!     Dim mean As Double, stdDev As Double
//!     Dim count As Long
//!     
//!     count = UBound(results) - LBound(results) + 1
//!     
//!     For i = LBound(results) To UBound(results)
//!         sum = sum + results(i)
//!         sumSq = sumSq + results(i) ^ 2
//!     Next i
//!     
//!     mean = sum / count
//!     stdDev = Sqr((sumSq / count) - (mean ^ 2))
//!     
//!     Dim report As String
//!     report = "Monte Carlo MIRR Analysis" & vbCrLf
//!     report = report & "Simulations: " & count & vbCrLf
//!     report = report & "Mean MIRR: " & Format(mean * 100, "0.00") & "%" & vbCrLf
//!     report = report & "Std Dev: " & Format(stdDev * 100, "0.00") & "%" & vbCrLf
//!     report = report & "95% Confidence Interval: " & _
//!         Format((mean - 1.96 * stdDev) * 100, "0.00") & "% to " & _
//!         Format((mean + 1.96 * stdDev) * 100, "0.00") & "%" & vbCrLf
//!     
//!     AnalyzeSimulationResults = report
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! result = MIRR(cashFlows(), finRate, reinvRate)
//! If Err.Number = 5 Then
//!     MsgBox "Invalid procedure call - check that cash flows " & _
//!            "contain both positive and negative values"
//! ElseIf Err.Number <> 0 Then
//!     MsgBox "Error calculating MIRR: " & Err.Description
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - MIRR calculation is iterative and can be computationally intensive for large arrays
//! - Cache MIRR results if using the same cash flows repeatedly
//! - For sensitivity analysis with many rate variations, consider pre-calculating once
//! - Array size affects performance - MIRR on 1000 periods is slower than 10 periods
//! - Consider using simplified models for real-time calculations
//!
//! ## Best Practices
//!
//! 1. **Validate inputs** - Ensure array contains both positive and negative values before calling MIRR
//! 2. **Use realistic rates** - Finance and reinvestment rates should reflect actual market conditions
//! 3. **Match rate periods** - If cash flows are monthly, use monthly rates (annual rate / 12)
//! 4. **Compare to hurdle rate** - Set minimum acceptable MIRR based on cost of capital
//! 5. **Consider risk** - Higher risk projects should have higher required MIRR
//! 6. **Document assumptions** - Clearly state finance and reinvestment rate assumptions
//! 7. **Use with other metrics** - Combine MIRR with NPV, IRR, and payback period for complete analysis
//! 8. **Test sensitivity** - Vary rates to see how robust the investment is to assumption changes
//! 9. **Account for inflation** - Use real rates (adjusted for inflation) for long-term projects
//! 10. **Round appropriately** - Display MIRR as percentage with 2 decimal places for clarity
//!
//! ## Comparison with Other Financial Functions
//!
//! | Function | Purpose | Key Difference from MIRR |
//! |----------|---------|--------------------------|
//! | **IRR** | Internal Rate of Return | Assumes reinvestment at IRR (often unrealistic); MIRR uses separate reinvestment rate |
//! | **NPV** | Net Present Value | Returns dollar amount, not percentage; uses single discount rate |
//! | **PV** | Present Value | Works with annuities/single payments, not irregular cash flows |
//! | **FV** | Future Value | Forward-looking value calculation; MIRR calculates rate of return |
//! | **Rate** | Interest rate for annuity | For regular payments only; MIRR handles irregular cash flows |
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, etc.)
//! - Not available in VBScript
//! - Part of VBA Financial functions library
//! - Requires at least one positive and one negative value in the array
//! - Arrays can be 0-based or 1-based (function handles both)
//!
//! ## Limitations
//!
//! - Requires at least one positive and one negative cash flow value
//! - All cash flows must occur at regular intervals
//! - Does not account for irregular timing between cash flows (see XIRR for that)
//! - Result is sensitive to both finance_rate and reinvest_rate assumptions
//! - Cannot handle multiple sign changes as robustly as some other methods
//! - First cash flow is assumed to occur at end of first period (not time zero)
//!
//! ## Related Functions
//!
//! - **IRR** - Calculate internal rate of return (assumes reinvestment at IRR)
//! - **NPV** - Calculate net present value using discount rate
//! - **PV** - Calculate present value of an investment
//! - **FV** - Calculate future value of an investment
//! - **Rate** - Calculate interest rate for an annuity
//! - **Pmt** - Calculate payment for a loan/annuity
//! - **XIRR** - IRR for irregular cash flow timing (Excel only)
//! - **XNPV** - NPV for irregular cash flow timing (Excel only)
//!
//! ## VB6 Parser Notes
//!
//! MIRR is parsed as a regular function call (CallExpression). This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers
//! working with financial calculations involving modified internal rate of return analysis.

#[cfg(test)]
mod tests {
    use crate::parsers::cst::ConcreteSyntaxTree;

    #[test]
    fn test_mirr_basic() {
        let source = r#"
Dim result As Double
result = MIRR(cashFlows(), 0.08, 0.05)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_variable_assignment() {
        let source = r#"
Dim investmentReturn As Double
investmentReturn = MIRR(projectFlows(), financeRate, reinvestRate)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_literal_rates() {
        let source = r#"
Dim rate As Double
rate = MIRR(cashFlows(), 0.1, 0.06)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_if_statement() {
        let source = r#"
If MIRR(cashFlows(), 0.08, 0.05) > 0.1 Then
    MsgBox "Good investment"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
    }

    #[test]
    fn test_mirr_function_return() {
        let source = r#"
Function CalculateReturn() As Double
    CalculateReturn = MIRR(flows(), 0.08, 0.05)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_comparison() {
        let source = r#"
Dim acceptable As Boolean
acceptable = MIRR(cashFlows(), finRate, reinvRate) >= hurdle
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_debug_print() {
        let source = r#"
Debug.Print MIRR(cashFlows(), 0.08, 0.05)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_with_statement() {
        let source = r#"
With investmentRecord
    .MIRR = MIRR(cashFlows(), .FinanceRate, .ReinvestRate)
End With
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_select_case() {
        let source = r#"
Select Case MIRR(cashFlows(), 0.08, 0.05)
    Case Is > 0.15
        MsgBox "Excellent"
    Case Is > 0.1
        MsgBox "Good"
    Case Else
        MsgBox "Poor"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_elseif() {
        let source = r#"
If x > 0 Then
    y = 1
ElseIf MIRR(cashFlows(), 0.08, 0.05) > 0.1 Then
    y = 2
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_parentheses() {
        let source = r#"
Dim result As Double
result = (MIRR(cashFlows(), 0.08, 0.05))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_iif() {
        let source = r#"
Dim msg As String
msg = IIf(MIRR(cashFlows(), 0.08, 0.05) > 0.1, "Good", "Bad")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_class_usage() {
        let source = r#"
Private m_mirr As Double

Public Sub Calculate()
    m_mirr = MIRR(m_cashFlows(), m_finRate, m_reinvRate)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_function_argument() {
        let source = r#"
Call ProcessReturn(MIRR(cashFlows(), 0.08, 0.05))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_property_assignment() {
        let source = r#"
Set obj = New Investment
obj.ReturnRate = MIRR(cashFlows(), 0.08, 0.05)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_array_assignment() {
        let source = r#"
Dim returns(10) As Double
Dim i As Integer
returns(i) = MIRR(cashFlows(), 0.08, 0.05)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_for_loop() {
        let source = r#"
Dim i As Integer
For i = 0 To 10
    results(i) = MIRR(scenarios(i), 0.08, 0.05)
Next i
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_while_wend() {
        let source = r#"
While MIRR(cashFlows(), rate, reinvRate) < targetReturn
    rate = rate + 0.01
Wend
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_do_while() {
        let source = r#"
Do While MIRR(cashFlows(), rate, reinvRate) < 0.1
    rate = rate + 0.005
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_do_until() {
        let source = r#"
Do Until MIRR(cashFlows(), rate, reinvRate) >= 0.1
    rate = rate + 0.005
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_msgbox() {
        let source = r#"
MsgBox "MIRR: " & MIRR(cashFlows(), 0.08, 0.05)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_concatenation() {
        let source = r#"
Dim report As String
report = "Return: " & Format(MIRR(cashFlows(), 0.08, 0.05) * 100, "0.00") & "%"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_comparison_expression() {
        let source = r#"
If MIRR(projectA(), 0.08, 0.05) > MIRR(projectB(), 0.08, 0.05) Then
    MsgBox "Project A is better"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_format() {
        let source = r#"
Dim formatted As String
formatted = Format(MIRR(cashFlows(), 0.08, 0.05), "0.00%")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_arithmetic() {
        let source = r#"
Dim annualizedReturn As Double
annualizedReturn = ((1 + MIRR(monthlyFlows(), monthlyFin, monthlyReinv)) ^ 12) - 1
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_label_caption() {
        let source = r#"
lblReturn.Caption = "MIRR: " & CStr(MIRR(cashFlows(), 0.08, 0.05))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn test_mirr_calculation() {
        let source = r#"
Dim percentReturn As Double
percentReturn = MIRR(cashFlows(), finRate, reinvRate) * 100
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("MIRR"));
        assert!(text.contains("Identifier"));
    }
}

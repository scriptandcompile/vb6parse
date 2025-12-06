//! # NPV Function
//!
//! Returns a Double specifying the net present value of an investment based on a series of periodic cash flows (payments and receipts) and a discount rate.
//!
//! ## Syntax
//!
//! ```vb
//! NPV(rate, values())
//! ```
//!
//! ## Parameters
//!
//! - **rate** (Required) - Double specifying discount rate over the length of the period, expressed as a decimal. For example, use 0.10 for 10 percent.
//! - **`values()`** (Required) - Array of Double specifying cash flow values. The array must contain at least one negative value (a payment) and one positive value (a receipt). The values must be equally spaced in time and occur at the end of each period.
//!
//! ## Return Value
//!
//! Returns a **Double** specifying the net present value of the investment.
//!
//! ## Remarks
//!
//! The NPV (Net Present Value) function calculates the present value of a series of future cash flows, discounted at a specified rate. It's a fundamental tool in capital budgeting and investment analysis.
//!
//! ### Key Characteristics:
//! - Returns present value of future cash flows
//! - All cash flows must be equally spaced in time
//! - Cash flows occur at the end of each period
//! - First cash flow (period 0) is NOT included - use separate calculation for initial investment
//! - Negative values represent cash outflows (payments)
//! - Positive values represent cash inflows (receipts)
//! - Rate should match the period of cash flows (e.g., monthly rate for monthly cash flows)
//! - Higher discount rates result in lower NPV
//! - NPV > 0 generally indicates a good investment
//! - NPV = 0 means investment breaks even
//! - NPV < 0 indicates investment loses money
//!
//! ### Important Note on Initial Investment:
//! Unlike some implementations, VB6's NPV does NOT include an initial investment (period 0) in the values array.
//! If you have an initial investment, subtract it from the NPV result:
//! ```vb
//! netPV = NPV(rate, cashFlows()) - initialInvestment
//! ```
//!
//! ### Common Use Cases:
//! - Capital budgeting decisions
//! - Investment project evaluation
//! - Equipment purchase analysis
//! - Real estate investment analysis
//! - Business valuation
//! - Compare alternative investments
//! - Determine project profitability
//! - Calculate present value of future revenues
//!
//! ## Typical Uses
//!
//! 1. **Investment Evaluation** - Determine if an investment is worthwhile
//! 2. **Project Comparison** - Compare multiple investment opportunities
//! 3. **Capital Budgeting** - Decide which projects to fund
//! 4. **Equipment Purchases** - Evaluate equipment vs. leasing decisions
//! 5. **Real Estate** - Analyze property investment returns
//! 6. **Business Acquisition** - Value businesses based on projected cash flows
//! 7. **Cost-Benefit Analysis** - Compare costs and benefits over time
//! 8. **Break-even Analysis** - Find discount rate where NPV = 0 (IRR)
//!
//! ## Basic Examples
//!
//! ```vb
//! ' Example 1: Simple investment with 5-year cash flows
//! Dim cashFlows(1 To 5) As Double
//! Dim npvResult As Double
//! Dim initialInvestment As Double
//!
//! initialInvestment = 10000
//! cashFlows(1) = 3000
//! cashFlows(2) = 3000
//! cashFlows(3) = 3000
//! cashFlows(4) = 3000
//! cashFlows(5) = 3000
//!
//! ' Calculate NPV at 10% discount rate
//! npvResult = NPV(0.1, cashFlows) - initialInvestment
//! ' Result: approximately $1,372 (good investment)
//! ```
//!
//! ```vb
//! ' Example 2: Evaluate equipment purchase
//! Dim savings(1 To 3) As Double
//! Dim equipmentCost As Double
//! Dim netValue As Double
//!
//! equipmentCost = 5000
//! savings(1) = 2000 ' Year 1 savings
//! savings(2) = 2500 ' Year 2 savings
//! savings(3) = 3000 ' Year 3 savings
//!
//! netValue = NPV(0.08, savings) - equipmentCost
//! If netValue > 0 Then
//!     MsgBox "Good investment: NPV = $" & Format(netValue, "#,##0.00")
//! End If
//! ```
//!
//! ```vb
//! ' Example 3: Compare two projects
//! Dim project1(1 To 4) As Double
//! Dim project2(1 To 4) As Double
//! Dim npv1 As Double, npv2 As Double
//!
//! ' Project 1: Higher initial investment, steady returns
//! project1(1) = 4000: project1(2) = 4000: project1(3) = 4000: project1(4) = 4000
//! npv1 = NPV(0.1, project1) - 12000
//!
//! ' Project 2: Lower investment, increasing returns
//! project2(1) = 2000: project2(2) = 3000: project2(3) = 4000: project2(4) = 5000
//! npv2 = NPV(0.1, project2) - 10000
//!
//! If npv2 > npv1 Then
//!     MsgBox "Project 2 has better NPV"
//! End If
//! ```
//!
//! ```vb
//! ' Example 4: Real estate investment analysis
//! Dim rentalIncome(1 To 10) As Double
//! Dim purchasePrice As Double
//! Dim i As Integer
//!
//! purchasePrice = 200000
//!
//! ' Annual rental income for 10 years
//! For i = 1 To 10
//!     rentalIncome(i) = 24000 ' $2,000/month
//! Next i
//!
//! Dim propertyNPV As Double
//! propertyNPV = NPV(0.08, rentalIncome) - purchasePrice
//! MsgBox "Property NPV: $" & Format(propertyNPV, "#,##0.00")
//! ```
//!
//! ## Common Patterns
//!
//! ```vb
//! ' Pattern 1: Basic NPV calculation with initial investment
//! Function CalculateInvestmentNPV(initialCost As Double, _
//!                                cashFlows() As Double, _
//!                                discountRate As Double) As Double
//!     CalculateInvestmentNPV = NPV(discountRate, cashFlows) - initialCost
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 2: Investment decision helper
//! Function ShouldInvest(initialInvestment As Double, _
//!                      cashFlows() As Double, _
//!                      requiredReturn As Double) As Boolean
//!     Dim npvResult As Double
//!     npvResult = NPV(requiredReturn, cashFlows) - initialInvestment
//!     ShouldInvest = (npvResult > 0)
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 3: Sensitivity analysis
//! Sub AnalyzeNPVSensitivity(initialCost As Double, cashFlows() As Double)
//!     Dim rate As Double
//!     Dim npvResult As Double
//!     
//!     Debug.Print "NPV Sensitivity Analysis"
//!     Debug.Print String(40, "-")
//!     
//!     For rate = 0.05 To 0.20 Step 0.01
//!         npvResult = NPV(rate, cashFlows) - initialCost
//!         Debug.Print Format(rate * 100, "0.0") & "%: $" & Format(npvResult, "#,##0.00")
//!     Next rate
//! End Sub
//! ```
//!
//! ```vb
//! ' Pattern 4: Profitability Index
//! Function CalculateProfitabilityIndex(initialInvestment As Double, _
//!                                      cashFlows() As Double, _
//!                                      discountRate As Double) As Double
//!     Dim presentValue As Double
//!     presentValue = NPV(discountRate, cashFlows)
//!     
//!     If initialInvestment = 0 Then
//!         CalculateProfitabilityIndex = 0
//!     Else
//!         CalculateProfitabilityIndex = presentValue / initialInvestment
//!     End If
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 5: Compare multiple projects
//! Function SelectBestProject(projects As Collection, discountRate As Double) As String
//!     Dim project As Variant
//!     Dim bestNPV As Double
//!     Dim bestProject As String
//!     Dim currentNPV As Double
//!     
//!     bestNPV = -999999
//!     
//!     For Each project In projects
//!         currentNPV = NPV(discountRate, project.CashFlows) - project.InitialCost
//!         
//!         If currentNPV > bestNPV Then
//!             bestNPV = currentNPV
//!             bestProject = project.Name
//!         End If
//!     Next project
//!     
//!     SelectBestProject = bestProject
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 6: Break-even discount rate (approximation of IRR)
//! Function FindBreakEvenRate(initialInvestment As Double, _
//!                           cashFlows() As Double) As Double
//!     Dim rate As Double
//!     Dim npvResult As Double
//!     Dim increment As Double
//!     
//!     increment = 0.001
//!     
//!     For rate = 0 To 1 Step increment
//!         npvResult = NPV(rate, cashFlows) - initialInvestment
//!         If npvResult <= 0 Then
//!             FindBreakEvenRate = rate
//!             Exit Function
//!         End If
//!     Next rate
//!     
//!     FindBreakEvenRate = -1 ' Not found
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 7: NPV with growing cash flows
//! Function NPVWithGrowth(initialInvestment As Double, _
//!                        firstYearCashFlow As Double, _
//!                        growthRate As Double, _
//!                        years As Integer, _
//!                        discountRate As Double) As Double
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     
//!     ReDim cashFlows(1 To years)
//!     
//!     For i = 1 To years
//!         cashFlows(i) = firstYearCashFlow * ((1 + growthRate) ^ (i - 1))
//!     Next i
//!     
//!     NPVWithGrowth = NPV(discountRate, cashFlows) - initialInvestment
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 8: Payback period with NPV
//! Function CalculateDiscountedPayback(initialInvestment As Double, _
//!                                     cashFlows() As Double, _
//!                                     discountRate As Double) As Double
//!     Dim i As Integer
//!     Dim cumulativeNPV As Double
//!     Dim periodCashFlows() As Double
//!     
//!     For i = LBound(cashFlows) To UBound(cashFlows)
//!         ReDim periodCashFlows(LBound(cashFlows) To i)
//!         
//!         Dim j As Integer
//!         For j = LBound(cashFlows) To i
//!             periodCashFlows(j) = cashFlows(j)
//!         Next j
//!         
//!         cumulativeNPV = NPV(discountRate, periodCashFlows) - initialInvestment
//!         
//!         If cumulativeNPV >= 0 Then
//!             CalculateDiscountedPayback = i
//!             Exit Function
//!         End If
//!     Next i
//!     
//!     CalculateDiscountedPayback = -1 ' Never pays back
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 9: NPV with terminal value
//! Function NPVWithTerminalValue(initialInvestment As Double, _
//!                               cashFlows() As Double, _
//!                               terminalValue As Double, _
//!                               discountRate As Double) As Double
//!     Dim i As Integer
//!     Dim modifiedCashFlows() As Double
//!     
//!     ReDim modifiedCashFlows(LBound(cashFlows) To UBound(cashFlows))
//!     
//!     For i = LBound(cashFlows) To UBound(cashFlows) - 1
//!         modifiedCashFlows(i) = cashFlows(i)
//!     Next i
//!     
//!     ' Add terminal value to final period
//!     modifiedCashFlows(UBound(cashFlows)) = cashFlows(UBound(cashFlows)) + terminalValue
//!     
//!     NPVWithTerminalValue = NPV(discountRate, modifiedCashFlows) - initialInvestment
//! End Function
//! ```
//!
//! ```vb
//! ' Pattern 10: Risk-adjusted NPV
//! Function RiskAdjustedNPV(initialInvestment As Double, _
//!                         expectedCashFlows() As Double, _
//!                         riskFreeRate As Double, _
//!                         riskPremium As Double) As Double
//!     Dim adjustedRate As Double
//!     adjustedRate = riskFreeRate + riskPremium
//!     RiskAdjustedNPV = NPV(adjustedRate, expectedCashFlows) - initialInvestment
//! End Function
//! ```
//!
//! ## Advanced Usage
//!
//! ### Example 1: Investment Analyzer Class
//!
//! ```vb
//! ' Class: InvestmentAnalyzer
//! ' Comprehensive investment analysis with NPV and related metrics
//!
//! Option Explicit
//!
//! Private m_initialInvestment As Double
//! Private m_cashFlows() As Double
//! Private m_discountRate As Double
//!
//! Public Sub Initialize(initialInvestment As Double, _
//!                      cashFlows() As Double, _
//!                      discountRate As Double)
//!     m_initialInvestment = initialInvestment
//!     m_cashFlows = cashFlows
//!     m_discountRate = discountRate
//! End Sub
//!
//! Public Function GetNPV() As Double
//!     GetNPV = NPV(m_discountRate, m_cashFlows) - m_initialInvestment
//! End Function
//!
//! Public Function GetProfitabilityIndex() As Double
//!     Dim pvCashFlows As Double
//!     pvCashFlows = NPV(m_discountRate, m_cashFlows)
//!     GetProfitabilityIndex = pvCashFlows / m_initialInvestment
//! End Function
//!
//! Public Function IsAcceptable() As Boolean
//!     IsAcceptable = (GetNPV() > 0)
//! End Function
//!
//! Public Function GetApproximateIRR() As Double
//!     Dim rate As Double
//!     Dim npvResult As Double
//!     Dim lastRate As Double
//!     Dim lastNPV As Double
//!     
//!     ' Binary search for IRR
//!     Dim lowRate As Double, highRate As Double
//!     lowRate = 0
//!     highRate = 1
//!     
//!     Do While (highRate - lowRate) > 0.0001
//!         rate = (lowRate + highRate) / 2
//!         npvResult = NPV(rate, m_cashFlows) - m_initialInvestment
//!         
//!         If npvResult > 0 Then
//!             lowRate = rate
//!         Else
//!             highRate = rate
//!         End If
//!     Loop
//!     
//!     GetApproximateIRR = rate
//! End Function
//!
//! Public Function GenerateReport() As String
//!     Dim report As String
//!     Dim npvValue As Double
//!     Dim pi As Double
//!     Dim irr As Double
//!     
//!     npvValue = GetNPV()
//!     pi = GetProfitabilityIndex()
//!     irr = GetApproximateIRR()
//!     
//!     report = "Investment Analysis Report" & vbCrLf
//!     report = report & String(50, "=") & vbCrLf & vbCrLf
//!     report = report & "Initial Investment: $" & Format(m_initialInvestment, "#,##0.00") & vbCrLf
//!     report = report & "Discount Rate: " & Format(m_discountRate * 100, "0.00") & "%" & vbCrLf
//!     report = report & "Number of Periods: " & UBound(m_cashFlows) & vbCrLf & vbCrLf
//!     report = report & "Net Present Value: $" & Format(npvValue, "#,##0.00") & vbCrLf
//!     report = report & "Profitability Index: " & Format(pi, "0.00") & vbCrLf
//!     report = report & "Approx. IRR: " & Format(irr * 100, "0.00") & "%" & vbCrLf & vbCrLf
//!     
//!     If IsAcceptable() Then
//!         report = report & "Recommendation: ACCEPT - Positive NPV"
//!     Else
//!         report = report & "Recommendation: REJECT - Negative NPV"
//!     End If
//!     
//!     GenerateReport = report
//! End Function
//!
//! Public Function RunSensitivityAnalysis() As String
//!     Dim result As String
//!     Dim rate As Double
//!     Dim npvValue As Double
//!     
//!     result = "NPV Sensitivity to Discount Rate" & vbCrLf
//!     result = result & String(40, "-") & vbCrLf
//!     
//!     For rate = 0.05 To 0.25 Step 0.05
//!         npvValue = NPV(rate, m_cashFlows) - m_initialInvestment
//!         result = result & Format(rate * 100, "0.0") & "%: $" & _
//!                  Format(npvValue, "#,##0.00") & vbCrLf
//!     Next rate
//!     
//!     RunSensitivityAnalysis = result
//! End Function
//! ```
//!
//! ### Example 2: Project Portfolio Manager
//!
//! ```vb
//! ' Class: ProjectPortfolioManager
//! ' Manages and ranks multiple investment projects
//!
//! Option Explicit
//!
//! Private Type Project
//!     name As String
//!     initialInvestment As Double
//!     cashFlows() As Double
//!     npv As Double
//!     pi As Double
//!     rank As Integer
//! End Type
//!
//! Private m_projects As Collection
//! Private m_discountRate As Double
//!
//! Public Sub Initialize(discountRate As Double)
//!     Set m_projects = New Collection
//!     m_discountRate = discountRate
//! End Sub
//!
//! Public Sub AddProject(name As String, _
//!                      initialInvestment As Double, _
//!                      cashFlows() As Double)
//!     Dim proj As Project
//!     
//!     proj.name = name
//!     proj.initialInvestment = initialInvestment
//!     proj.cashFlows = cashFlows
//!     proj.npv = NPV(m_discountRate, cashFlows) - initialInvestment
//!     
//!     Dim pvCashFlows As Double
//!     pvCashFlows = NPV(m_discountRate, cashFlows)
//!     proj.pi = pvCashFlows / initialInvestment
//!     
//!     m_projects.Add proj, name
//! End Sub
//!
//! Public Sub RankByNPV()
//!     ' Simple ranking based on NPV
//!     Dim projects() As Project
//!     Dim i As Integer, j As Integer
//!     
//!     ReDim projects(1 To m_projects.Count)
//!     
//!     For i = 1 To m_projects.Count
//!         projects(i) = m_projects(i)
//!     Next i
//!     
//!     ' Sort by NPV (descending)
//!     For i = 1 To UBound(projects) - 1
//!         For j = i + 1 To UBound(projects)
//!             If projects(j).npv > projects(i).npv Then
//!                 Dim temp As Project
//!                 temp = projects(i)
//!                 projects(i) = projects(j)
//!                 projects(j) = temp
//!             End If
//!         Next j
//!     Next i
//!     
//!     ' Assign ranks
//!     For i = 1 To UBound(projects)
//!         projects(i).rank = i
//!     Next i
//! End Sub
//!
//! Public Function GetPortfolioNPV() As Double
//!     Dim proj As Project
//!     Dim totalNPV As Double
//!     Dim i As Integer
//!     
//!     totalNPV = 0
//!     
//!     For i = 1 To m_projects.Count
//!         proj = m_projects(i)
//!         totalNPV = totalNPV + proj.npv
//!     Next i
//!     
//!     GetPortfolioNPV = totalNPV
//! End Function
//!
//! Public Function SelectProjectsWithBudget(budget As Double) As Collection
//!     Dim selectedProjects As New Collection
//!     Dim remainingBudget As Double
//!     Dim proj As Project
//!     Dim i As Integer
//!     
//!     RankByNPV
//!     remainingBudget = budget
//!     
//!     For i = 1 To m_projects.Count
//!         proj = m_projects(i)
//!         
//!         If proj.npv > 0 And proj.initialInvestment <= remainingBudget Then
//!             selectedProjects.Add proj
//!             remainingBudget = remainingBudget - proj.initialInvestment
//!         End If
//!     Next i
//!     
//!     Set SelectProjectsWithBudget = selectedProjects
//! End Function
//!
//! Public Function GenerateRankingReport() As String
//!     Dim report As String
//!     Dim proj As Project
//!     Dim i As Integer
//!     
//!     RankByNPV
//!     
//!     report = "Project Portfolio Ranking" & vbCrLf
//!     report = report & String(80, "=") & vbCrLf
//!     report = report & "Rank  Project Name          Investment       NPV          PI" & vbCrLf
//!     report = report & String(80, "-") & vbCrLf
//!     
//!     For i = 1 To m_projects.Count
//!         proj = m_projects(i)
//!         report = report & Format(proj.rank, "0") & "     "
//!         report = report & Left(proj.name & String(20, " "), 20) & "  "
//!         report = report & Format(proj.initialInvestment, "$#,##0") & "  "
//!         report = report & Format(proj.npv, "$#,##0") & "  "
//!         report = report & Format(proj.pi, "0.00") & vbCrLf
//!     Next i
//!     
//!     report = report & vbCrLf & "Total Portfolio NPV: $" & _
//!              Format(GetPortfolioNPV(), "#,##0.00")
//!     
//!     GenerateRankingReport = report
//! End Function
//! ```
//!
//! ### Example 3: Real Estate Investment Analyzer
//!
//! ```vb
//! ' Module: RealEstateAnalyzer
//! ' Analyzes real estate investments using NPV
//!
//! Option Explicit
//!
//! Public Function AnalyzeRentalProperty(purchasePrice As Double, _
//!                                      annualRent As Double, _
//!                                      annualExpenses As Double, _
//!                                      holdingYears As Integer, _
//!                                      appreciationRate As Double, _
//!                                      discountRate As Double) As String
//!     Dim cashFlows() As Double
//!     Dim i As Integer
//!     Dim salePrice As Double
//!     Dim npvValue As Double
//!     Dim result As String
//!     
//!     ReDim cashFlows(1 To holdingYears)
//!     
//!     ' Calculate annual net cash flows
//!     For i = 1 To holdingYears - 1
//!         cashFlows(i) = annualRent - annualExpenses
//!     Next i
//!     
//!     ' Final year includes property sale
//!     salePrice = purchasePrice * ((1 + appreciationRate) ^ holdingYears)
//!     cashFlows(holdingYears) = (annualRent - annualExpenses) + salePrice
//!     
//!     npvValue = NPV(discountRate, cashFlows) - purchasePrice
//!     
//!     result = "Real Estate Investment Analysis" & vbCrLf
//!     result = result & String(50, "=") & vbCrLf
//!     result = result & "Purchase Price: $" & Format(purchasePrice, "#,##0") & vbCrLf
//!     result = result & "Annual Rent: $" & Format(annualRent, "#,##0") & vbCrLf
//!     result = result & "Annual Expenses: $" & Format(annualExpenses, "#,##0") & vbCrLf
//!     result = result & "Holding Period: " & holdingYears & " years" & vbCrLf
//!     result = result & "Appreciation Rate: " & Format(appreciationRate * 100, "0.0") & "%" & vbCrLf
//!     result = result & "Discount Rate: " & Format(discountRate * 100, "0.0") & "%" & vbCrLf
//!     result = result & "Estimated Sale Price: $" & Format(salePrice, "#,##0") & vbCrLf & vbCrLf
//!     result = result & "Net Present Value: $" & Format(npvValue, "#,##0.00") & vbCrLf
//!     
//!     If npvValue > 0 Then
//!         result = result & "Recommendation: Good investment opportunity"
//!     Else
//!         result = result & "Recommendation: Consider other opportunities"
//!     End If
//!     
//!     AnalyzeRentalProperty = result
//! End Function
//!
//! Public Function CompareBuyVsLease(equipmentCost As Double, _
//!                                  leaseCosts() As Double, _
//!                                  buyingCashFlows() As Double, _
//!                                  discountRate As Double) As String
//!     Dim buyNPV As Double
//!     Dim leaseNPV As Double
//!     Dim result As String
//!     
//!     buyNPV = NPV(discountRate, buyingCashFlows) - equipmentCost
//!     leaseNPV = NPV(discountRate, leaseCosts)
//!     
//!     result = "Buy vs. Lease Analysis" & vbCrLf
//!     result = result & String(40, "-") & vbCrLf
//!     result = result & "Buying NPV: $" & Format(buyNPV, "#,##0.00") & vbCrLf
//!     result = result & "Leasing NPV: $" & Format(leaseNPV, "#,##0.00") & vbCrLf
//!     result = result & vbCrLf
//!     
//!     If buyNPV > leaseNPV Then
//!         result = result & "Recommendation: BUY (NPV advantage: $" & _
//!                  Format(buyNPV - leaseNPV, "#,##0.00") & ")"
//!     Else
//!         result = result & "Recommendation: LEASE (NPV advantage: $" & _
//!                  Format(leaseNPV - buyNPV, "#,##0.00") & ")"
//!     End If
//!     
//!     CompareBuyVsLease = result
//! End Function
//! ```
//!
//! ### Example 4: Business Valuation Tool
//!
//! ```vb
//! ' Class: BusinessValuationTool
//! ' Values businesses using discounted cash flow (DCF) method
//!
//! Option Explicit
//!
//! Private m_projectedCashFlows() As Double
//! Private m_terminalValue As Double
//! Private m_discountRate As Double
//!
//! Public Sub ProjectCashFlows(baseCashFlow As Double, _
//!                            growthRate As Double, _
//!                            years As Integer)
//!     Dim i As Integer
//!     
//!     ReDim m_projectedCashFlows(1 To years)
//!     
//!     For i = 1 To years
//!         m_projectedCashFlows(i) = baseCashFlow * ((1 + growthRate) ^ i)
//!     Next i
//! End Sub
//!
//! Public Sub CalculateTerminalValue(finalYearCashFlow As Double, _
//!                                  perpetualGrowthRate As Double, _
//!                                  discountRate As Double)
//!     ' Terminal value using Gordon Growth Model
//!     m_terminalValue = (finalYearCashFlow * (1 + perpetualGrowthRate)) / _
//!                       (discountRate - perpetualGrowthRate)
//! End Sub
//!
//! Public Function GetEnterpriseValue(discountRate As Double) As Double
//!     Dim i As Integer
//!     Dim cashFlowsWithTerminal() As Double
//!     
//!     m_discountRate = discountRate
//!     
//!     ' Create array including terminal value
//!     ReDim cashFlowsWithTerminal(1 To UBound(m_projectedCashFlows))
//!     
//!     For i = 1 To UBound(m_projectedCashFlows) - 1
//!         cashFlowsWithTerminal(i) = m_projectedCashFlows(i)
//!     Next i
//!     
//!     ' Add terminal value to final year
//!     cashFlowsWithTerminal(UBound(m_projectedCashFlows)) = _
//!         m_projectedCashFlows(UBound(m_projectedCashFlows)) + m_terminalValue
//!     
//!     GetEnterpriseValue = NPV(discountRate, cashFlowsWithTerminal)
//! End Function
//!
//! Public Function GetEquityValue(enterpriseValue As Double, _
//!                               debt As Double, _
//!                               cash As Double) As Double
//!     GetEquityValue = enterpriseValue - debt + cash
//! End Function
//!
//! Public Function GenerateValuationReport(companyName As String, _
//!                                        debt As Double, _
//!                                        cash As Double) As String
//!     Dim report As String
//!     Dim ev As Double
//!     Dim equity As Double
//!     
//!     ev = GetEnterpriseValue(m_discountRate)
//!     equity = GetEquityValue(ev, debt, cash)
//!     
//!     report = "Business Valuation: " & companyName & vbCrLf
//!     report = report & String(60, "=") & vbCrLf & vbCrLf
//!     report = report & "Discounted Cash Flow Analysis" & vbCrLf
//!     report = report & "Discount Rate (WACC): " & Format(m_discountRate * 100, "0.00") & "%" & vbCrLf
//!     report = report & "Projection Period: " & UBound(m_projectedCashFlows) & " years" & vbCrLf
//!     report = report & vbCrLf
//!     report = report & "Enterprise Value: $" & Format(ev, "#,##0,000") & vbCrLf
//!     report = report & "Less: Debt: $" & Format(debt, "#,##0,000") & vbCrLf
//!     report = report & "Plus: Cash: $" & Format(cash, "#,##0,000") & vbCrLf
//!     report = report & String(60, "-") & vbCrLf
//!     report = report & "Equity Value: $" & Format(equity, "#,##0,000")
//!     
//!     GenerateValuationReport = report
//! End Function
//! ```
//!
//! ## Error Handling
//!
//! ```vb
//! On Error Resume Next
//! Dim npvResult As Double
//! npvResult = NPV(rate, cashFlows())
//! If Err.Number <> 0 Then
//!     MsgBox "Error calculating NPV: " & Err.Description & vbCrLf & _
//!            "Ensure cash flows array is properly dimensioned and rate is valid."
//! End If
//! On Error GoTo 0
//! ```
//!
//! ## Performance Considerations
//!
//! - NPV calculation is relatively fast
//! - Performance scales linearly with number of cash flows
//! - For large arrays (1000+ periods), consider optimization
//! - Caching results when using same inputs repeatedly
//! - More efficient than manually discounting each period
//!
//! ## Best Practices
//!
//! 1. **Remember initial investment** - Subtract it from NPV result (not in array)
//! 2. **Use appropriate discount rate** - Match to risk and period
//! 3. **Validate cash flow array** - Ensure proper dimensioning
//! 4. **Check for mixed signs** - Need both positive and negative flows
//! 5. **Consider risk** - Use higher discount rate for riskier investments
//! 6. **Document assumptions** - State discount rate rationale
//! 7. **Sensitivity analysis** - Test multiple discount rates
//! 8. **Compare to alternatives** - Use NPV to rank projects
//! 9. **Account for inflation** - Use real rates or inflate cash flows
//! 10. **Validate results** - Ensure NPV makes sense given inputs
//!
//! ## Comparison with Related Functions
//!
//! | Function | Purpose | Key Difference |
//! |----------|---------|----------------|
//! | **NPV** | Net present value | Assumes end-of-period cash flows |
//! | **PV** | Present value | Single/annuity payments only |
//! | **IRR** | Internal rate of return | Finds rate where NPV = 0 |
//! | **MIRR** | Modified IRR | Uses separate financing/reinvestment rates |
//! | **FV** | Future value | Calculates future amount, not present |
//!
//! ## Platform Notes
//!
//! - Available in VBA (Excel, Access, Word, etc.)
//! - Available in VB6
//! - **Not available in `VBScript`**
//! - Equivalent to Excel's NPV function
//! - Part of VBA financial functions library
//! - Requires array parameter (not `ParamArray`)
//!
//! ## Limitations
//!
//! - Assumes equally spaced time periods
//! - Cash flows occur at end of each period
//! - Does NOT include initial investment in values array
//! - Cannot handle irregular time periods (use XNPV in Excel for that)
//! - Assumes constant discount rate
//! - Does not account for taxes directly
//! - Requires at least one positive and one negative value
//!
//! ## Related Functions
//!
//! - **IRR** - Finds internal rate of return (discount rate where NPV = 0)
//! - **MIRR** - Modified internal rate of return
//! - **PV** - Present value of annuity or lump sum
//! - **FV** - Future value of investment
//! - **PMT** - Payment amount for loan/investment
//! - **Rate** - Interest rate per period
//!
//! ## VB6 Parser Notes
//!
//! NPV is parsed as a regular function call (`CallExpression`). This module exists primarily
//! for documentation purposes to provide comprehensive reference material for VB6 developers
//! working with financial analysis, investment evaluation, capital budgeting, and discounted
//! cash flow calculations.

#[cfg(test)]
mod tests {
    use crate::*;

    #[test]
    fn npv_basic() {
        let source = r#"
Dim result As Double
result = NPV(0.1, cashFlows)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_with_initial_investment() {
        let source = r#"
Dim netValue As Double
netValue = NPV(0.08, cashFlows) - initialInvestment
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_if_statement() {
        let source = r#"
If NPV(rate, values) - cost > 0 Then
    MsgBox "Good investment"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_function_return() {
        let source = r#"
Function CalculateNPV(rate As Double, flows() As Double) As Double
    CalculateNPV = NPV(rate, flows)
End Function
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_variable_assignment() {
        let source = r#"
Dim presentValue As Double
presentValue = NPV(0.12, projectedCashFlows)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_msgbox() {
        let source = r#"
MsgBox "NPV: $" & Format(NPV(discountRate, flows), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_debug_print() {
        let source = r#"
Debug.Print "Net Present Value: " & NPV(r, cf)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_select_case() {
        let source = r#"
Select Case NPV(rate, cashFlows) - investment
    Case Is > 10000
        decision = "Excellent"
    Case Is > 0
        decision = "Good"
    Case Else
        decision = "Poor"
End Select
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_class_usage() {
        let source = r#"
Private m_npv As Double

Public Sub Calculate()
    m_npv = NPV(m_rate, m_cashFlows)
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_with_statement() {
        let source = r#"
With investment
    .NetPresentValue = NPV(.DiscountRate, .CashFlows) - .InitialCost
End With
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_elseif() {
        let source = r#"
If x > 0 Then
    y = 1
ElseIf NPV(r, flows) > threshold Then
    y = 2
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_for_loop() {
        let source = r#"
For rate = 0.05 To 0.15 Step 0.01
    npvValue = NPV(rate, cashFlows) - initialCost
    Debug.Print rate, npvValue
Next rate
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_do_while() {
        let source = r#"
Do While NPV(currentRate, flows) - cost > 0
    currentRate = currentRate + 0.01
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_do_until() {
        let source = r#"
Do Until NPV(r, values) - investment <= 0
    r = r + 0.001
Loop
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_while_wend() {
        let source = r#"
While NPV(discountRate, projections) > minValue
    discountRate = discountRate + 0.005
Wend
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_parentheses() {
        let source = r#"
Dim result As Double
result = (NPV(rate, cashFlows))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_iif() {
        let source = r#"
Dim decision As String
decision = IIf(NPV(r, cf) - cost > 0, "Accept", "Reject")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_comparison() {
        let source = r#"
If NPV(rate1, flows1) > NPV(rate2, flows2) Then
    MsgBox "Project 1 is better"
End If
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_array_assignment() {
        let source = r#"
Dim npvValues(10) As Double
npvValues(i) = NPV(rates(i), cashFlows)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_property_assignment() {
        let source = r#"
Set obj = New Investment
obj.PresentValue = NPV(obj.Rate, obj.CashFlows)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_function_argument() {
        let source = r#"
Call DisplayInvestmentAnalysis(NPV(discountRate, projectedFlows))
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_format() {
        let source = r#"
Dim formatted As String
formatted = "NPV: " & Format(NPV(0.1, flows), "0.00")
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_arithmetic() {
        let source = r#"
Dim profitabilityIndex As Double
profitabilityIndex = NPV(rate, flows) / initialInvestment
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_concatenation() {
        let source = r#"
Dim msg As String
msg = "Present Value: $" & NPV(r, cf) & " at " & r * 100 & "%"
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_abs_function() {
        let source = r#"
Dim absValue As Double
absValue = Abs(NPV(discountRate, cashFlows) - targetValue)
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_error_handling() {
        let source = r#"
On Error Resume Next
result = NPV(rate, cashFlows)
If Err.Number <> 0 Then
    MsgBox "NPV calculation error"
End If
On Error GoTo 0
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }

    #[test]
    fn npv_on_error_goto() {
        let source = r#"
Sub CalculateProjectNPV()
    On Error GoTo ErrorHandler
    Dim netValue As Double
    netValue = NPV(discountRate, cashFlows) - initialCost
    Exit Sub
ErrorHandler:
    MsgBox "Error calculating NPV"
End Sub
"#;
        let tree = ConcreteSyntaxTree::from_source("test.bas", source).unwrap();
        let text = tree.debug_tree();
        assert!(text.contains("NPV"));
        assert!(text.contains("Identifier"));
    }
}

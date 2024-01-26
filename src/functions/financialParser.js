import { numberToExcelColumn } from "./util"

const template = []

const headerRows = [
    "Non-Interest Revenue",
    "Cost of Revenue",
    "Gross Profit",
    "Research and Development",
    "Selling, General and Administrative",
    "Other Operating Expenses",
    "Operating Expenses",
    "Operating Income",
    "Interest Income",
    "Interest Expense",
    "Other Income",
    "Total Other Income",
    "Pretax Income",
    "Taxes",
    "Discontinued operations",
    "Non-controlling interests",
    "Equity income in investee",
    "Net Income",
    undefined,
    "Depreciation & Amortization",
    "Stock Based Compensation",
    "Other Operating Activities",
    "Cash Flow from Operations",
    "Capital Expenditures",
    "Free Cash Flow to Firm",
    undefined,
    "Revenue Y/Y",
    "Gross Margin",
    "Operating Margin",
    "Tax Rate",
    "Net Margin",
    "Free Cash Flow Y/Y",
    undefined,
    "Cash and Cash Equivalents",
    "Short-Term Investments",
    "Short-Term Debt",
    "Long-Term Debt",
    "Property, Plant and Equipment Net",
    "Current Assets",
    "Non-current Assets",
    "Assets",
    "Current Liabilities",
    "Non-current Liabilities",
    "Liabilities",
    "Shareholder's equity"
]
function setCell(entries, row, column, entry) {
  if (!entries[row]) {
    entries[row] = []
  }
  entries[row][column] = entry
}
for(let i=1; i<=46; i++) {
    setCell(template, i+1, 1, headerRows[i-1])
}

export function generateEntries(incomeStatements, cashFlowStatements, balanceSheets, maximumDivider) {
    const entries = [...template]

    let maximumDividerString
    if (maximumDivider === 1000) {
        maximumDividerString = "(in thousands)"
    } else if (maximumDivider === 1000000) {
        maximumDividerString = "(in millions)"
    } else if (maximumDivider === 1000000000) {
        maximumDividerString = "(in billions)"
    } else if (maximumDividerString === 1000000000000) {
        maximumDividerString = "(in trillions)"
    }

    const headerColumns = []
    const revenues = []
    const cogs = []
    const grossProfit = []
    const rAndD = []
    const sgAndA = []
    const otherOpex = []
    const operatingExpenses = []
    const operatingIncome = []
    const interestIncome = []
    const interestExpense = []
    const otherIncome = []
    const totalOtherIncome = []
    const pretaxIncome = []
    const taxes = []
    const discontinuedOperations = []
    const nonControllingInterests = []
    const equityIncomeInInvestee = []
    const netIncome = []

    const revenueYoY = []
    const grossMargin = []
    const operatingMargin = []
    const taxRate = []
    const netMargin = []
    const freeCashFlowYoY = []

    const depreciationAndAmortization = []
    const stockBasedCompensation = []
    const otherOperatingActivities = []
    const cashFlowFromOperations = []
    const capitalExpenditures = []
    const freeCashFlowToFirm = []

    const cashAndCashEquivalents = []
    const shortTermInvestments = []
    const shortTermDebt = []
    const longTermDebt = []
    const propertyPlantAndEquipment = []
    const currentAssets = []
    const nonCurrentAssets = []
    const assets = []
    const currentLiabilities = []
    const nonCurrentLiabilities = []
    const liabilities = []
    const shareholderEquity = []

    const currentYear = new Date().getFullYear()
    const startingYear = Math.min(...Object.keys(incomeStatements))

    for(let i=startingYear; i <= currentYear + 10; i++) {
        const previousColumn = i-startingYear+1
        const previousExcelColumn = numberToExcelColumn(previousColumn)

        const column = i-startingYear+2
        const excelColumn = numberToExcelColumn(column)

        const incomeStatement = incomeStatements[i]
        const cashFlowStatement = cashFlowStatements[i]
        const balanceSheet = balanceSheets[i]

        if (incomeStatement !== undefined) {
            revenues.push(incomeStatement.Revenue/maximumDivider)
            cogs.push(incomeStatement.CostOfRevenue/maximumDivider)
            rAndD.push(incomeStatement.ResearchAndDevelopment/maximumDivider)
            sgAndA.push(incomeStatement.SellingGeneralAndAdministrative/maximumDivider)
            otherOpex.push(incomeStatement.OtherOperatingIncomeExpenses/maximumDivider)
            interestIncome.push(incomeStatement.InterestIncome/maximumDivider)
            interestExpense.push(incomeStatement.InterestExpense/maximumDivider)
            otherIncome.push(incomeStatement.OtherIncome/maximumDivider)
            taxes.push(incomeStatement.IncomeTaxExpenseBenefit/maximumDivider)
            discontinuedOperations.push(incomeStatement.IncomeLossFromDiscontinuedOperationsNetOfTax/maximumDivider)
            nonControllingInterests.push(incomeStatement.NetIncomeLossAttributableToNoncontrollingInterest/maximumDivider)
            equityIncomeInInvestee.push(incomeStatement.EquityIncomeInInvestee/maximumDivider)
        }

        grossProfit.push("=IF(COUNTA("+excelColumn+"2;"+excelColumn+"3)>0;"+excelColumn+"2-"+excelColumn+'3; "")')      
        operatingExpenses.push("=IF(COUNTA("+excelColumn+"5;"+excelColumn+"7)>0; "+excelColumn+"5+"+excelColumn+"6+"+excelColumn+'7; "")')
        operatingIncome.push("=IF(COUNTA("+excelColumn+"4;"+excelColumn+"8)>0;"+excelColumn+"4-"+excelColumn+'8; "")')
        totalOtherIncome.push("=IF(COUNTA("+excelColumn+"10;"+excelColumn+"11;"+excelColumn+"12)>0;"+excelColumn+"10-"+excelColumn+"11+"+excelColumn+'12; "")')
        pretaxIncome.push("=IF(COUNTA("+excelColumn+"9;"+excelColumn+"13)>0;"+excelColumn+"9+"+excelColumn+'13; "")')
        netIncome.push("=IF(COUNTA("+excelColumn+"14;"+excelColumn+"15;"+excelColumn+"16;"+excelColumn+"17;"+excelColumn+"18)>0;"+excelColumn+"14-"+excelColumn+"15+"+excelColumn+"16-"+excelColumn+"17+"+excelColumn+'18; "")')
        if (i !== startingYear) {
            revenueYoY.push(`=IF(COUNTA(${previousExcelColumn}2;${excelColumn}2)==2 AND B2 != 0; CONCAT((${excelColumn}2/${previousExcelColumn}2-1)*100; "%"); "")`)
        }

        if (cashFlowStatement !== undefined) {
            depreciationAndAmortization.push(cashFlowStatement.depreciationAndAmortization/maximumDivider)
            stockBasedCompensation.push(cashFlowStatement.stockBasedCompensation/maximumDivider)
            otherOperatingActivities.push(cashFlowStatement.otherOperatingActivities/maximumDivider)
            capitalExpenditures.push(cashFlowStatement.capex/maximumDivider)
        }

        cashFlowFromOperations.push("=IF(COUNTA("+excelColumn+"19;"+excelColumn+"21;"+excelColumn+"22;"+excelColumn+"23)>0;"+excelColumn+"19+"+excelColumn+"21+"+excelColumn+"22+"+excelColumn+'23; "")')
        freeCashFlowToFirm.push("=IF(COUNTA("+excelColumn+"24;"+excelColumn+"25)>0;"+excelColumn+"24-"+excelColumn+'25; "")')

        if (balanceSheet !== undefined) {
            cashAndCashEquivalents.push(balanceSheet.cashAndCashEquivalents/maximumDivider)
            shortTermInvestments.push(balanceSheet.shortTermInvestments/maximumDivider)
            shortTermDebt.push(balanceSheet.shortTermDebt/maximumDivider)
            longTermDebt.push(balanceSheet.longTermDebt/maximumDivider)
            propertyPlantAndEquipment.push(balanceSheet.propertyPlantAndEquipmentNet/maximumDivider)
            currentAssets.push(balanceSheet.currentAssets/maximumDivider)
            nonCurrentAssets.push(balanceSheet.nonCurrentAssets/maximumDivider)
            currentLiabilities.push(balanceSheet.currentLiabilities/maximumDivider)
            nonCurrentLiabilities.push(balanceSheet.nonCurrentLiabilities/maximumDivider)
        }

        assets.push("=IF(COUNTA("+excelColumn+"40;"+excelColumn+"41)>0;"+excelColumn+"40+"+excelColumn+'41; "")')
        liabilities.push("=IF(COUNTA("+excelColumn+"43;"+excelColumn+"44)>0;"+excelColumn+"43+"+excelColumn+'44; "")')
        shareholderEquity.push("=IF(COUNTA("+excelColumn+"42;"+excelColumn+"45)>0;"+excelColumn+"42-"+excelColumn+'45; "")')

        headerColumns.push('="'+i.toString()+'"')
    }

    for(let j=2; j<=currentYear - startingYear + 10 + 2; j++) {
        setCell(entries, 1, j, headerColumns[j-2])
        setCell(entries, 2, j, revenues[j-2]?.toString())
        setCell(entries, 3, j, cogs[j-2]?.toString())
        setCell(entries, 4, j, grossProfit[j-2]?.toString())
        setCell(entries, 5, j, rAndD[j-2]?.toString())
        setCell(entries, 6, j, sgAndA[j-2]?.toString())
        setCell(entries, 7, j, otherOpex[j-2]?.toString())
        setCell(entries, 8, j, operatingExpenses[j-2]?.toString())
        setCell(entries, 9, j, operatingIncome[j-2]?.toString())
        setCell(entries, 10, j, interestIncome[j-2]?.toString())
        setCell(entries, 11, j, interestExpense[j-2]?.toString())
        setCell(entries, 12, j, otherIncome[j-2]?.toString())
        setCell(entries, 13, j, totalOtherIncome[j-2]?.toString())
        setCell(entries, 14, j, pretaxIncome[j-2]?.toString())
        setCell(entries, 15, j, taxes[j-2]?.toString())
        setCell(entries, 16, j, discontinuedOperations[j-2]?.toString())
        setCell(entries, 17, j, nonControllingInterests[j-2]?.toString())
        setCell(entries, 18, j, equityIncomeInInvestee[j-2]?.toString())
        setCell(entries, 19, j, netIncome[j-2]?.toString())

        setCell(entries, 21, j, depreciationAndAmortization[j-2]?.toString())
        setCell(entries, 22, j, stockBasedCompensation[j-2]?.toString())
        setCell(entries, 23, j, otherOperatingActivities[j-2]?.toString())
        setCell(entries, 24, j, cashFlowFromOperations[j-2]?.toString())
        setCell(entries, 25, j, capitalExpenditures[j-2]?.toString())
        setCell(entries, 26, j, freeCashFlowToFirm[j-2]?.toString())

        setCell(entries, 28, j, revenueYoY[j-3]?.toString())

        setCell(entries, 35, j, cashAndCashEquivalents[j-2]?.toString())
        setCell(entries, 36, j, shortTermInvestments[j-2]?.toString())
        setCell(entries, 37, j, shortTermDebt[j-2]?.toString())
        setCell(entries, 38, j, longTermDebt[j-2]?.toString())
        setCell(entries, 39, j, propertyPlantAndEquipment[j-2]?.toString())
        setCell(entries, 40, j, currentAssets[j-2]?.toString())
        setCell(entries, 41, j, nonCurrentAssets[j-2]?.toString())
        setCell(entries, 42, j, assets[j-2]?.toString())
        setCell(entries, 43, j, currentLiabilities[j-2]?.toString())
        setCell(entries, 44, j, nonCurrentLiabilities[j-2]?.toString())
        setCell(entries, 45, j, liabilities[j-2]?.toString())
        setCell(entries, 46, j, shareholderEquity[j-2]?.toString())
    }

    return [entries, 46, currentYear - startingYear + 10 + 2]
}
$symbols = 'aapl','msft','goog','amzn','tsla','brk-b','unh','jnj','v','meta','tsm','xom','tcehy','wmt','pg','cvx'
$date = "2022-11-*"
$MissingType = [System.Type]::Missing
$WorksheetCount = 3
$excel = New-Object -ComObject excel.application
$excel.Visible = $True

### Add a workbook
$Workbook = $Excel.Workbooks.Add()

#Add worksheets
$null = $Excel.Worksheets.Add($MissingType, $Excel.Worksheets.Item($Excel.Worksheets.Count), 
$WorksheetCount - $Excel.Worksheets.Count, $Excel.Worksheets.Item(1).Type)

$Sheet = $Excel.Worksheets.Item(1)
$Sheet1 = $Excel.Worksheets.Item(2)
$Sheet2 = $Excel.Worksheets.Item(3)

$Sheet.Name  = "Insider"
$Sheet1.Name  = "Chiefs"
$Sheet2.Name = "Financials"

###########################################################################################
########### Insider #################
###########################################################################################
$intRow = 2
$Sheet.Cells.Item(1,1) = "Buyers"
$Sheet.Cells.Item(1,2) = "Sellers"
$Sheet.Cells.Item(1,3) = "Symbol"
$Sheet.Cells.Item(1,4) = "Date"
$Sheet.Cells.Item(1,5) = "Company"
$Sheet.Cells.Item(1,6) = "Title"
$Sheet.Cells.Item(1,7) = "Name"
$Sheet.Cells.Item(1,8) = "Type"
$Sheet.Cells.Item(1,9) = "Shares"
$WorkBook = $Sheet.UsedRange
$Sheet.Activate()

$stocks = foreach($s in $symbols){
    try{
        Invoke-RestMethod -Uri "https://query2.finance.yahoo.com/v11/finance/quoteSummary/${s}?modules=assetProfile,balanceSheetHistory,balanceSheetHistoryQuarterly,calendarEvents,cashflowStatementHistory,cashflowStatementHistoryQuarterly,defaultKeyStatistics,earnings,earningsHistory,earningsTrend,financialData,fundOwnership,incomeStatementHistory,incomeStatementHistoryQuarterly,indexTrend,industryTrend,insiderHolders,insiderTransactions,institutionOwnership,majorDirectHolders,majorHoldersBreakdown,netSharePurchaseActivity,price,quoteType,recommendationTrend,secFilings,sectorTrend,summaryDetail,summaryProfile,symbol,upgradeDowngradeHistory,fundProfile,topHoldings,fundPerformance" -TimeoutSec 3
    }
    catch{
        $s | Out-File 'c:\temp\badsymbol-insider.txt'
    }
}
$insider = foreach($stock in $stocks){
    $symbol = $stock.quoteSummary.result.quoteType.symbol
    $transactions = $stock.quoteSummary.result.insiderTransactions.transactions
    $company = $stock.quoteSummary.result.quoteType.shortName
    $sector = $stock.quoteSummary.result.assetProfile.sector
    $industry = $stock.quoteSummary.result.assetProfile.industry
    $stock.quoteSummary.result.insiderHolders.holders
    $stock.quoteSummary.result.earningsTrend.trend
    foreach($i in $transactions){
        if($i.startDate.fmt -like $date){
            if (!([string]::IsNullOrWhiteSpace($i.transactionText))){
                    $shares = $i.shares | Select-Object -ExpandProperty fmt
                    [PSCustomObject]@{
                        Symbol = $symbol
                        Company = $company
                        Sector = $sector
                        Industry = $industry
                        Name = $i.filerName
                        Date = $i.startDate.fmt
                        Type = $i.transactionText
                        Title = $i.filerRelation
                        Shares = $shares
                    }
                    $intRow ++ 
                    $Sheet.Cells.Item($intRow, 1) = $null
                    $Sheet.Cells.Item($intRow, 2) = $null
                    $Sheet.Cells.Item($intRow, 3) = $symbol
                    $Sheet.Cells.Item($intRow, 4) = $i.startDate.fmt
                    $Sheet.Cells.Item($intRow, 5) = $company
                    $Sheet.Cells.Item($intRow, 6) = $i.filerRelation
                    $Sheet.Cells.Item($intRow, 7) = $i.filerName
                    $Sheet.Cells.Item($intRow, 8) = $i.transactionText
                    $Sheet.Cells.Item($intRow, 9) = $shares

                    if($i.transactionText -like '*Sale*'){
                        $Sheet.Cells.Item($intRow, 3).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 4).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 5).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 6).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 7).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 8).Font.ColorIndex = 3
                        $Sheet.Cells.Item($intRow, 9).Font.ColorIndex = 3
                    }elseif($i.transactionText -like '*Purchase*'){
                        $Sheet.Cells.Item($intRow, 3).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 4).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 5).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 6).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 7).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 8).Font.ColorIndex = 10
                        $Sheet.Cells.Item($intRow, 9).Font.ColorIndex = 10
                    }elseif(($i.transactionText -like '*Conversion*') -or ($i.transactionText -like '*Award*') -or ($i.transactionText -like '*Gift*') -or ($i.transactionText -like '*Exercise*') -or ($i.transactionText -like '*Acquisition*')){
                        $Sheet.Cells.Item($intRow, 3).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 4).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 5).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 6).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 7).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 8).Font.ColorIndex = 5
                        $Sheet.Cells.Item($intRow, 9).Font.ColorIndex = 5
                    }else{
                        $Sheet.Cells.Item($intRow, 3).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 4).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 5).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 6).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 7).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 8).Font.ColorIndex = 1
                        $Sheet.Cells.Item($intRow, 9).Font.ColorIndex = 1
                    }
            }
        }
    }
}

$buy = $insider | Where-Object {(($_.Type -notlike '*Sale at*') -and ($_.Date -like $date))} | Select-Object -ExpandProperty $_.Symbol -Unique
$sales = $insider | Where-Object {(($_.Type -like 'Sale*') -and ($_.Date -like $date))}|Measure-Object|Select-Object -ExpandProperty Count
$purchase = $insider | Where-Object {(($_.Type -notlike '*Sale at*') -and ($_.Date -like $date))}|Measure-Object|Select-Object -ExpandProperty Count -Unique

$Sheet.Cells.Item(2, 1) = $purchase
$Sheet.Cells.Item(2, 2) = $sales

$range = $Sheet.Range("a1","i1")
$range.Style = 'Title'
$range.Font.Bold = $True
$range.Font.ColorIndex = 1
$range.Interior.ColorIndex = 20
$WorkBook.EntireColumn.AutoFit() | Out-Null
$Sheet.Cells().HorizontalAlignment = -4108

###########################################################################################
########### Chiefs #################
###########################################################################################
$intRow = 2
$Sheet1.Cells.Item(1,1) = "Buyers"
$Sheet1.Cells.Item(1,2) = "Sellers"
$Sheet1.Cells.Item(1,3) = "Symbol"
$Sheet1.Cells.Item(1,4) = "Date"
$Sheet1.Cells.Item(1,5) = "Company"
$Sheet1.Cells.Item(1,6) = "Title"
$Sheet1.Cells.Item(1,7) = "Name"
$Sheet1.Cells.Item(1,8) = "Type"
$Sheet1.Cells.Item(1,9) = "Shares"
$WorkBook = $Sheet1.UsedRange
$Sheet1.Activate()

$chiefs = foreach($stock in $stocks){
    $symbol = $stock.quoteSummary.result.quoteType.symbol
    $transactions = $stock.quoteSummary.result.insiderTransactions.transactions
    $company = $stock.quoteSummary.result.quoteType.shortName
    foreach($i in $transactions){
        if($i.startDate.fmt -like $date) {
            if (!([string]::IsNullOrWhiteSpace($i.transactionText))){
                if(($i.filerRelation -like '*Chief*') -or ($i.filerRelation -like '*Beneficial*')){
                    $shares = $i.shares | Select-Object -ExpandProperty fmt
                    [PSCustomObject]@{
                        Symbol = $symbol
                        Company = $company
                        Sector = $sector
                        Industry = $industry
                        Name = $i.filerName
                        Date = $i.startDate.fmt
                        Type = $i.transactionText
                        Title = $i.filerRelation
                        Shares = $shares
                    }
                    $intRow ++ 
                    $Sheet1.Cells.Item($intRow, 1) = $null
                    $Sheet1.Cells.Item($intRow, 2) = $null
                    $Sheet1.Cells.Item($intRow, 3) = $symbol
                    $Sheet1.Cells.Item($intRow, 4) = $i.startDate.fmt
                    $Sheet1.Cells.Item($intRow, 5) = $company
                    $Sheet1.Cells.Item($intRow, 6) = $i.filerRelation
                    $Sheet1.Cells.Item($intRow, 7) = $i.filerName
                    $Sheet1.Cells.Item($intRow, 8) = $i.transactionText
                    $Sheet1.Cells.Item($intRow, 9) = $shares

                    if($i.transactionText -like '*Sale*'){
                        $Sheet1.Cells.Item($intRow, 3).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 4).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 5).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 6).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 7).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 8).Font.ColorIndex = 3
                        $Sheet1.Cells.Item($intRow, 9).Font.ColorIndex = 3
                    }elseif($i.transactionText -like '*Purchase*'){
                        $Sheet1.Cells.Item($intRow, 3).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 4).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 5).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 6).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 7).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 8).Font.ColorIndex = 10
                        $Sheet1.Cells.Item($intRow, 9).Font.ColorIndex = 10
                    }elseif(($i.transactionText -like '*Conversion*') -or ($i.transactionText -like '*Award*') -or ($i.transactionText -like '*Gift*') -or ($i.transactionText -like '*Exercise*') -or ($i.transactionText -like '*Acquisition*')){
                        $Sheet1.Cells.Item($intRow, 3).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 4).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 5).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 6).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 7).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 8).Font.ColorIndex = 5
                        $Sheet1.Cells.Item($intRow, 9).Font.ColorIndex = 5
                    }else{
                        $Sheet1.Cells.Item($intRow, 3).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 4).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 5).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 6).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 7).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 8).Font.ColorIndex = 1
                        $Sheet1.Cells.Item($intRow, 9).Font.ColorIndex = 1
                    }
                }
            }
        }
    }
}

$sales = $chiefs | Where-Object {(($_.Type -like 'Sale*') -and ($_.Date -like $date))}|Measure-Object|Select-Object -ExpandProperty Count
$purchase = $chiefs | Where-Object {(($_.Type -notlike '*Sale at*') -and ($_.Date -like $date))}|Measure-Object|Select-Object -ExpandProperty Count -Unique

$Sheet1.Cells.Item(2, 1) = $purchase
$Sheet1.Cells.Item(2, 2) = $sales

$range = $Sheet1.Range("a1","i1")
$range.Style = 'Title'
$range.Font.Bold = $True
$range.Font.ColorIndex = 1
$range.Interior.ColorIndex = 20
$WorkBook.EntireColumn.AutoFit() | Out-Null
$Sheet1.Cells().HorizontalAlignment = -4108

###########################################################################################
########### Financials #################
###########################################################################################
$intRow = 1
$Sheet2.Cells.Item(1,1) = "Symbol"
$Sheet2.Cells.Item(1,2) = "Company"
$Sheet2.Cells.Item(1,3) = "Sector"
$Sheet2.Cells.Item(1,4) = "Industry"
$Sheet2.Cells.Item(1,5) = "Action"
$Sheet2.Cells.Item(1,6) = "Analyst"
$Sheet2.Cells.Item(1,7) = "Price"
$Sheet2.Cells.Item(1,8) = "HighPrice"
$Sheet2.Cells.Item(1,9) = "LowPrice"
$Sheet2.Cells.Item(1,10) = "Earnings"
$Sheet2.Cells.Item(1,11) = "Cap"
$Sheet2.Cells.Item(1,12) = "Revenue"
$Sheet2.Cells.Item(1,13) = "DebtEquity"
$Sheet2.Cells.Item(1,14) = "ebitda"
$Sheet2.Activate()
$WorkBook = $Sheet2.UsedRange

$fin = foreach($stock in $stocks){
    $financialData = $stock.quoteSummary.result.financialData
    $symbol = $stock.quoteSummary.result.quoteType.symbol
    $sector = $stock.quoteSummary.result.assetProfile.sector
    $industry = $stock.quoteSummary.result.assetProfile.industry
    $ebitda = $stock.quoteSummary.result.financialData.ebitdaMargins.fmt
    $revenueGrowth = $stock.quoteSummary.result.financialData.revenueGrowth.fmt
    $debttoequity = $stock.quoteSummary.result.financialData.debtToEquity.fmt
    $company = $stock.quoteSummary.result.quoteType.shortName
    $Cap = $stock.quoteSummary.result.summaryDetail.marketCap.longFmt
    foreach($f in $financialData){
        $action = $f.recommendationKey
            [PSCustomObject]@{
                Symbol = $symbol
                Company = $company
                Sector = $sector
                Industry = $industry
                Action = $action
                CurrentPrice = $f.currentPrice.fmt
                HighPrice = $f.targetHighPrice.fmt
                LowPrice = $f.targetLowPrice.fmt
                EarningsGrowth = $f.earningsGrowth.fmt
                Cap = $cap
                AnalystMean = $f.recommendationMean.fmt
                ebitda = $ebitda
                revenueGrowth = $revenueGrowth
                DebtToEquity = $debttoequity              
            }
            $intRow ++ 
            $Sheet2.Cells.Item($intRow, 1) = $symbol
            $Sheet2.Cells.Item($intRow, 2) = $company
            $Sheet2.Cells.Item($intRow, 3) = $sector
            $Sheet2.Cells.Item($intRow, 4) = $industry
            $Sheet2.Cells.Item($intRow, 5) = $f.recommendationKey
            $Sheet2.Cells.Item($intRow, 6) = $f.recommendationMean.fmt
            $Sheet2.Cells.Item($intRow, 7) = $f.currentPrice.fmt
            $Sheet2.Cells.Item($intRow, 8) = $f.targetHighPrice.fmt
            $Sheet2.Cells.Item($intRow, 9) = $f.targetLowPrice.fmt
            $Sheet2.Cells.Item($intRow, 10) = $f.earningsGrowth.fmt
            $Sheet2.Cells.Item($intRow, 11) = $cap 
            $Sheet2.Cells.Item($intRow, 12) = $revenueGrowth
            $Sheet2.Cells.Item($intRow, 13) = $debtToEquity
            $Sheet2.Cells.Item($intRow, 14) = $ebitda
    }
    if($f.earningsGrowth.raw -gt 1.00){
        $Sheet2.Cells.Item($intRow, 1).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 2).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 3).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 4).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 5).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 6).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 7).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 8).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 9).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 10).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 11).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 12).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 13).Font.ColorIndex = 10
        $Sheet2.Cells.Item($intRow, 14).Font.ColorIndex = 10
    }
    elseif(($f.earningsGrowth.raw -lt 1.00 -eq $true) -and ($f.earningsGrowth.raw -gt 0.01 -eq $true)){
        $Sheet2.Cells.Item($intRow, 1).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 2).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 3).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 4).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 5).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 6).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 7).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 8).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 9).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 10).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 11).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 12).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 13).Font.ColorIndex = 5
        $Sheet2.Cells.Item($intRow, 14).Font.ColorIndex = 5
    }
    elseif($f.EarningsGrowth.raw -lt 0.01){
        $Sheet2.Cells.Item($intRow, 1).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 2).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 3).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 4).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 5).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 6).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 7).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 8).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 9).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 10).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 11).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 12).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 13).Font.ColorIndex = 3
        $Sheet2.Cells.Item($intRow, 14).Font.ColorIndex = 3
    }
}
 
$range2 = $Sheet2.Range("a1","n1")
$range2.Style = 'Title'
$range2.Font.Bold = $True
$range2.Font.ColorIndex = 1
$range2.Interior.ColorIndex = 20
$WorkBook.EntireColumn.AutoFit() | Out-Null
$Sheet2.Cells().HorizontalAlignment = -4108

$ts = get-date -f "MMddyyyyhhmmss"
$Excel.ActiveWorkbook.SaveAs("c:\temp\stock-report-$ts.xlsx")
$excel.Quit()
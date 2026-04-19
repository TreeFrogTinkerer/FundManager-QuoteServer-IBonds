#######################################################################
##
## I-Bond to FundManager Quote "Server" Tool
## v1.1
## by ballardian
## 05/09/2025
##
## Script is provided as-is with no warranties or guarantees
##
## To Use:
## Self-Hosted Option
## 1) Run every 6 months
## 2) Point the <local file> Internet Retrieve Option to the generated IbondFM-QuoteServer.csv in this folder
## 3) Set your I-Bonds in FM to the following ticker format IB + MON + YYYY -- for example April 2023 I-bond would be IBAPR2023
## 4) Set your I-Bonds in FM to use the <local file> option for current and historical prices 
## 
## NOTES:  
##
##    A) Interest will only appear at the 12 month mark and up to 5 years it takes the 3 months penalty out -- the interest has accrued but isn't shown due to the data style from the Treasury
##
#######################################################################
# This it the function that re-formats the data and returns a formatted CSV line called below
function Create-CSVLine {
  param(
    [string]$Mo,
    [string]$IssueYear,
    [string]$MoAmount,
    [string]$RedemptionMo,
    [string]$RedemptionYr
  )
    # Auto-Generates a Ticker
    $Ticker = "IB"+$Mo + $IssueYear
    # If the amount isn't a null then continue
    if ($MoAmount -ne "null")
    {
        # The data is in units referencing $25 so this multiplies it to make it the $100 base FM tells you to use
        $HundrededAmount = [decimal]$MoAmount*4
        # If the amount isn't zero (I'm not sure why some data comes back as 0 as the price) continue on
        if ([decimal]$MoAmount -eq 0)
        {
            # Combibne the above strings into a csv "line" with the correct format
            $CSVLine = $Ticker +"," + $RedemptionMo + "/1/"+$RedemptionYr+",100"
        }
        else
        {
            # Combibne the above strings into a csv "line" with the correct format
            $CSVLine = $Ticker +"," + $RedemptionMo + "/1/"+$RedemptionYr+","+$HundrededAmount
        }
        # Output the line and return it from the function that goes into the $l variables above
        $CSVLine
    }
}

# Retrieve the data from the US Treasury API -- this API returns redemption values
$TreasuryRAW = Invoke-RestMethod -Method 'GET' -Uri "https://api.fiscaldata.treasury.gov/services/api/fiscal_service/v2/accounting/od/sb_value?filter=series_cd:eq:I&page[number]=1&page[size]=10000 "

# Count how many sets of data there are - it is returned not line but in JSON matrices
$Rounds = $TreasuryRAW.data.Length

# Configures the array for our FM formatted data before it is exported to a UTF-8 csv file
[System.Collections.ArrayList]$exportcsv = @()
# Adds the header to the output array
$null = $exportcsv.Add('SYMB,MM/DD/YY,NAV')

#Sets the index to zero to loop through the data sets
$Counter=0

# This do/while loop goes through the smaller data sets in the JSON file
do
{
    # Mostly just reduces the typing we have to do by removing the single data set we are using into the $series variable
    $series = $TreasuryRAW.data[$Counter]
    
    ## This block processes each months data in the set by sending it and the correct parameters to a function that re-formats the data int a line for the csv file
    ## if the data coming back is null then it doesn't write to the csv file and it has a line of text it adds it to the array buffer
    
    ## Jan
    $l = Create-CSVLine -Mo "JAN" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_jan_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Feb
    $l = Create-CSVLine -Mo "FEB" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_feb_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Mar
    $l = Create-CSVLine -Mo "MAR" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_mar_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Apr
    $l = Create-CSVLine -Mo "APR" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_apr_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## May
    $l = Create-CSVLine -Mo "MAY" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_may_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Jun
    $l = Create-CSVLine -Mo "JUN" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_jun_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Jul
    $l = Create-CSVLine -Mo "JUL" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_jul_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Aug
    $l = Create-CSVLine -Mo "AUG" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_aug_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Sep
    $l = Create-CSVLine -Mo "SEP" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_sep_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Oct
    $l = Create-CSVLine -Mo "OCT" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_oct_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Nov
    $l = Create-CSVLine -Mo "NOV" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_nov_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    ## Dec
    $l = Create-CSVLine -Mo "DEC" -IssueYear $series.issue_year -RedemptionMo $series.redemption_month -RedemptionYr $series.redemption_year -MoAmount $series.issue_dec_amt
    if ($l -ne $null){$null = $exportcsv.Add($l)}
    
    # Moves the counter to process the next data set
    $Counter++
}
until($Counter -eq $Rounds)



# Export the data to a csv file in the root of this script
$exportcsv -join "`r`n" | Out-File -FilePath "IbondFM-QuoteServer.csv" -Encoding UTF8

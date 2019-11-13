#================================================================
#   Get-MonthDates - Return the start and end dates of the  month
#================================================================

function Get-MonthDates ( [datetime]$Date=(Get-Date) )
{
    $MonthStart = [datetime]$Date.ToString("yyyy-MM-01")

    write-output $MonthStart
    write-output $MonthStart.AddMonths(1).AddDays(-1)
}
	

#===============================================================
#    Main Program - Script execution starts here
#===============================================================

# Test the function:
# -----------------
    $Dates = Get-MonthDates
    write-host "The first day of the month is:" $Dates[0]
    write-host " The last day of the month is:" $Dates[1]

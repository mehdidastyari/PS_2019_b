#===================================================================
#
#  Sum - Display the sum of all numbers given on the command line
#
#===================================================================


#-------------------------------------------------------------------
#  Show-Error - display error & usage info, then  exit with status
#-------------------------------------------------------------------
function Show-Error ( [string]$Message, [int]$ExitStatus=1 )
{
     write-host -foregroundcolor red "`n$Message"
	 write-host -foregroundcolor white "`nHow to use this script:"
	 write-host "`n    Sum  <number> [<number>] [<number>...]"
	 write-host "`nwhere:"
	 write-host "`n   <number>s are integer or floating-point values`n"
	 
	 exit $ExitStatus
}

#-------------------------------------------------------------------
#  Main program - execution starts here
#-------------------------------------------------------------------

# Verify the arguments:
# --------------------
	if ( $args.count -lt 2 )
	{
	    Show-Error "Must have two or more script arguments"
	}
	
	
# Sum the arguments:
# -----------------
	$Sum = 0.0
	foreach( $arg in $args )
	{
		$Number = $arg -as [double]
		if ($Number -eq $Null)
		{
		    Show-Error "'$arg' is not a valid number"
		}
		
		$Sum += $Number
	}
	
# Show the sum:
# ------------
	write-host "`nThe sum is $Sum`n"
	
	exit 0
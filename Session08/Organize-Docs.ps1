#===================================================================
#
#  Organize-Docs - Put MS-Office documents into appropriate subfolders
#
#	  This script places any office documents into subfolders.  The
#	  folder to be organized must be given as a command argument.
#
#===================================================================


#-------------------------------------------------------------------
#  Show-Error - display error & usage info, then  exit with status
#-------------------------------------------------------------------
function Show-Error ( [string]$Message, [int]$ExitStatus=1 )
{
     write-host -foregroundcolor red "`n$Message"
	 write-host -foregroundcolor white "`nHow to use this script:"
	 write-host "`n    Organize-OfficeDocs  <folder>"
	 write-host "`nwhere:"
	 write-host "`n  <folder>"
	 write-host "     Is the the folder to be organized.  Office docs"
	 write-host "     in this folder are put into Word, Excel or PowerPoint"
	 write-host "     subfolders"
	 write-host ""
	 
	 exit $ExitStatus
}

#-------------------------------------------------------------------
#  Main program - execution starts here
#-------------------------------------------------------------------

# Verify the arguments:
# --------------------
	if ( $args.count -lt 1 )
	{
	    Show-Error "Missing folder name"
	}

	$FolderName = $args[0]
	if ( -not (test-path $FolderName -PathType Container) )
	{
		Show-Error "Folder '$FolderName' not found"
	}
	
# Create the collections of file types:
# ------------------------------------
	$WordDocTypes = ".doc", ".docx", ".rtf"
	$ExcelDocTypes = ".xls", ".xlsx", ".csv"
	$PPDocTypes = ".ppt", ".pptx"
	$AccessDocTypes = ".mdb", ".accdb"
	
# Initialize counts:
# -----------------
	write-host ""
	$FileCount = 0
	$MoveCount = 0
	
# Go through all the files in the given folder:
# --------------------------------------------
	get-childitem $FolderName  |
	   where-object { $_ -is [IO.FileInfo] }  |
	   foreach-object `
	{		
		$FileCount++

	# Get the document type of the file based on it's file type:
    # Note that a drawback of this script is that we need statements
    # for each document type, which means that if we have to add
    # a new document type ("Scripts", for example) we have to edit
    # these statements.  We can eliminate this dependency using
    # multidimensional arrays ("collections of collections").
    # See the Organize-Docs-MD.ps1 script for an example.
	# ---------------------------------------------------------------
	    $FileType = $_.Extension.ToLower()
		$DocType = ""
		if ( $WordDocTypes -contains $FileType ) { $DocType = "Word" }
		if ( $ExcelDocTypes -contains $FileType ) { $DocType = "Excel" }
		if ( $PPDocTypes -contains $FileType ) { $DocType = "PowerPoint" }
		if ( $AccessDocTypes -contains $FileType ) { $DocType = "Access" }
	
	# Ignore the file if it's not a recognized type:
	# ---------------------------------------------
		if ( $DocType -ne "" )
		{
		# Create the target folder if it doesn't exist:
		# --------------------------------------------
			$TargetFolderName = join-path $FolderName $DocType
			if ( -not (test-path $TargetFolderName -PathType Container) )
			{
				new-item $TargetFolderName -Type Directory | out-null
			}
			
		# Move the file to the target folder:
		# ----------------------------------
			move-item $_.FullName $TargetFolderName
			write-host ( $_.Name.PadLeft(22) + " - moved to " + $TargetFolderName )
			$MoveCount++
		}
		else
		{
			write-host ( $_.Name.PadLeft(22) + " - not an Office document" )
		}
	}

# Show totals and exit:
# --------------------
	write-host "`n$MoveCount of $FileCount file(s) moved.`n"
	exit 0
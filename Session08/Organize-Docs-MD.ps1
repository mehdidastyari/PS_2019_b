#=======================================================================
#
#  Organize-Docs - Put various document types into different subfolders
#
#	  This script places any office documents into subfolders.  The
#	  folder to be organized must be given as a command argument.
#
#	  This version of the script uses a multidimensional array.
#     This makes the script more flexible - it doesn't need specific
#     statements in the script to deal with Word, Excel, etc. documents.
#     We can add new kinds of documents ("Scripts", for example) simply
#     by changing how the collections are initialized without having to
#     change any of the script logic.
#
#=======================================================================


#-------------------------------------------------------------------
#  Show-Error - display error & usage info, then  exit with status
#-------------------------------------------------------------------
function Show-Error ( [string]$Message, [int]$ExitStatus=1 )
{
     write-host -foregroundcolor red "`n$Message"
	 write-host -foregroundcolor white "`nHow to use this script:"
	 write-host "`n    Organize-Docs  <folder>"
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
	$FileTypes = (".doc", ".docx", ".rtf"),  # Collection of file types for word
	             (".xls", ".xlsx", ".csv"),  # Collection of file types for Excel
				 (".ppt", ".pptx" ),         # Collection of file types for PowerPoint
				 (".mdb", ".accdb" )         # Collection of file types for Access

    $Doctypes = "Word", "Excel", "PowerPoint", "Access"
                                             # Plain-text names for each type of document.
                                             # These must appear in the same order as the
                                             # collections given above (i.e., "Word" is
                                             # first in this collection because the first
                                             # collection of file types above are those
                                             # for Word documents).
	
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
	# ---------------------------------------------------------
	    $FileType = $_.Extension.ToLower()
		$DocType = ""
		for( $Index=0; $Index -lt $FileTypes.Count; $Index++ )
		{
		    if ( $FileTypes[$Index] -contains $FileType ) { $DocType = $DocTypes[$Index] }
		}
	
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
			write-host ("{0,-20} - moved to {1}" -f $_.Name, $TargetFolderName )
			$MoveCount++
		}
		else
		{
			write-host ("{0,-20} - not an Office document" -f $_.Name )
		}
	}

# Show totals and exit:
# --------------------
	write-host "`n$MoveCount of $FileCount file(s) moved.`n"
	exit 0
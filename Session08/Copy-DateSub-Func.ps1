#===================================================================
#
#  Copy-DateSub-Func - Copy files from a source folder to dated subfolders
#
#    Files are copied from the given source folder to dated
#    subfolders in the given target folder based on their
#    last modified dates.
#
#    For example, if the script is invoked using:
#
#       Copy-DateSub  d:\Source  e:\Target
#
#    Then files from d:\Source will be copied into folders such as:
#
#        e:\Target\2010-01-15
#        e:\Target\2010-01-18
#
#    ...based on their modification dates.  The target folders
#    will be created if they do not exist.   If a target folder
#    already exists with a folder name that BEGINs with the correct
#    date, that folder will be used.  For example, if this folder:
#
#         e:\Target\2010-01-18_Birthday-Party
#
#    ...exists, then files last modified on 2010-01-18 will be
#    copied into it.
#
#===================================================================


#-------------------------------------------------------------------
#  Show-Error - display error & usage info, then  exit with status
#-------------------------------------------------------------------
function Show-Error ( [string]$Message, [int]$ExitStatus=1 )
{
    write-host -foregroundcolor red "`n$Message"
    write-host -foregroundcolor white "`nHow to use this script:"
    write-host "`n    Copy-DateSub  <source-folder> <target-folder>"
    write-host "`nwhere:"
    write-host "`n  <source-folder>"
    write-host "     Is the folder that files will be copied from"
    write-host "`n  <target-folder>"
    write-host "     Is the top-level folder that files will be copied to."
    write-host "     It will be created if it does not already exist."
    write-host ""
    write-host "Files are copied to dated subfolders beneath the given target"
    write-host "based on their modification dates.  If a subfolder whose name"
    write-host "begins with the correct date  (in yyyy-mm-dd format)  already"
    write-host "exists, that folder will be used."
	 
    exit $ExitStatus
}

#-------------------------------------------------------------------
#  Get-ToFolderName  -  Get the target folder name for the given date
#-------------------------------------------------------------------
function Get-ToFolderName ( [string]$TargetFolderName, [datetime]$Date )
{
# Get the file date in yyyy-mm-dd format - this will be the subfolder name:
# ------------------------------------------------------------------------
    $FileDate = $Date.ToString( "yyyy-MM-dd" )

# Build a wildcard search string to see if there's an
# existing folder name that starts with the required date:
# -------------------------------------------------------
    $SearchFolderName = [IO.Path]::Combine( $TargetFolderName, $FileDate + "*" )

# Look for an existing folder name that starts with the required date.
# If there are multiple folders, get the one with the shortest name:
# -----------------------------------------------------------------
    $ToFolderName = get-childitem $SearchFolderName   |
       where-object { $_ -is [IO.DirectoryInfo] }   |
       foreach-object { $_.FullName }   |
       sort-object Length   |
       select-object -first 1

# If we didn't find an existing subfolder with the right name,
# then the subfolder name will simply be the modification date:
# ------------------------------------------------------------
    if ( $ToFolderName -eq $Null )
    {
        $ToFolderName = [IO.Path]::Combine( $TargetFolderName, $FileDate )
    }
	
# Return the selected to folder:
# -----------------------------
    return $ToFolderName
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
	
    $SourceFolderName = $args[0]
    $TargetFolderName = $args[1]
	
    if ( -not (test-path $SourceFolderName -pathtype container) )
    {
        Show-Error "Folder '$SourceFolderName' does not exist"
    }
	
    if ( test-path $TargetFolderName -pathtype leaf )
    {
        Show-Error "Target folder '$TargetFolderName' can't be created`n   because a file of the same name already exists"
    }
	
# Create the target folder if necessary:
# -------------------------------------

    if ( -not (test-path $TargetFolderName -pathtype container) )
    {
        new-item $TargetFolderName -type Directory | out-null
    }

# Copy all the files:
# ------------------
    write-host "`nCopying files from '$SourceFolderName'..."
    $CopyCount = 0

    get-childitem $SourceFolderName   |
        where-object { $_ -is [IO.FileInfo] }   |
        foreach-object `
        {
	# Get the destination folder based on the file's date:
	# ---------------------------------------------------
            $ToFolderName = Get-ToFolderName $TargetFolderName $_.LastWriteTime
	
	# If the target folder doesn't exist, then create it:
	# --------------------------------------------------
            if ( -not (test-path $ToFolderName -type Container) )
            {
                new-item $ToFolderName -type directory | out-null
            }

	# Copy this file from the source folder to the appropriate subfolder:
	# ------------------------------------------------------------------
            copy-item $_.FullName $ToFolderName
            $CopyCount++
            write-host "   $($_.Name) copied to`n      $(split-path $ToFolderName -leaf)`n"
        }

# Show totals and exit:
# --------------------
    write-host "$CopyCount files copied`n"
    exit 0
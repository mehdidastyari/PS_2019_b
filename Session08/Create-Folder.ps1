#===============================================================
#    Create-Folder - Create a folder in a given parent directory
#
#        This function creates the named folder if necessary and
#        returns one of the following status codes:
#               0 - Success
#           not 0 - Failure
#===============================================================

function Create-Folder ( [string]$ParentFolderName, [string]$FolderName )
{
# Check to ensure that the parent folder exists:
# ---------------------------------------------
    if ( -not (test-path $ParentFolderName -pathtype container ) )
    {
        write-host "Parent folder '$ParentFolderName' does not exist"
        return 1
    }
	
# Make the new folder name:
# ------------------------
    $NewFolderName = [IO.Path]::Combine( $ParentFolderName, $FolderName )

# Don't bother creating the folder if it already exists:
# -----------------------------------------------------
    if ( test-path $NewFolderName ) { return 0 }

# Otherwise, create the folder and return success status:
# ------------------------------------------------------
    new-item $NewFolderName -type Directory
    return 0	
}

#===============================================================
#	Main Program - Script execution starts here
#===============================================================

# Test the Create-Folder Function:
# -------------------------------
    $Status = Create-Folder ".\" "Test-Folder"
    if ($Status -eq 0)
    {
        write-host "Folder created successfully"
    }
    else
    {
        write-host "Folder could not be created"
    }

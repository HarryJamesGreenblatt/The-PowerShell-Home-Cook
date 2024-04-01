function Find-HexColorCodes {
<#
    .SYNOPSIS
        Searches for hex color codes in all files within a specified directory.

    .DESCRIPTION
        The Find-HexColorCodes function crawls through a directory, checking the contents of each file for patterns matching hex codes for colors, and returns a list of all unique color matches found.

    .PARAMETER directoryPath
        The path to the directory that will be searched.

    .EXAMPLE
        Find-HexColorCodes -directoryPath "C:\path\to\your\directory"
        This command will search for hex color codes in the specified directory and return a list of unique matches.

    .NOTES
        This function uses a regular expression to match hex color codes in the format #FFF or #FFFFFF.
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string] 
        $directoryPath
    )

    begin{

        # Define the regex pattern for hex color codes
        $hexPattern = "#([A-Fa-f0-9]{6}|[A-Fa-f0-9]{3})"
    
        # Collect all matching hex codes
        $colorMatches = @()
    
        # Get all files in the directory recursively
        $files = Get-ChildItem -Path $directoryPath -Recurse -File

    }
    
    process{

        # Iterate over each file
        foreach ($file in $files) {
    
            # Check the contents of the file for hex color codes
            $matches = Select-String -Path $file.FullName -Pattern $hexPattern -AllMatches
    
            # If matches are found, add them to the list
            foreach ($match in $matches.Matches) {
                $colorMatches += $match.Value
            }

        }
    }
    
    end{
        # Remove duplicate color matches
        $uniqueColorMatches = $colorMatches | Sort-Object -Unique

        # Output the list of unique color matches
        $uniqueColorMatches
    }

}

Export-ModuleMember -Function Find-HexColorCodes
function Find-HexColorCodes {
<#
    .SYNOPSIS
    Searches for hex color codes in all files within a specified directory and optionally displays a webpage depicting it as a color table.

    .DESCRIPTION
    The Find-HexColorCodes function crawls through a directory, checking the contents of each file for patterns matching hex codes for colors. It returns a list of all unique color matches found and can optionally display a color table for a visual representation of the colors.

    .PARAMETER directoryPath
    The path to the directory that will be searched.

    .PARAMETER displayColorTable
    A switch parameter that, when used, will display a color table with a visual representation of each color found.

    .EXAMPLE
    Find-HexColorCodes -directoryPath "C:\path\to\your\directory"
    This command will search for hex color codes in the specified directory and return a list of unique matches.

    .EXAMPLE
    Find-HexColorCodes -directoryPath "C:\path\to\your\directory" -displayColorTable
    This command will search for hex color codes in the specified directory, return a list of unique matches, and display a color table with each color.

    .NOTES
    This function uses a regular expression to match hex color codes in the format #FFF or #FFFFFF. The color table is displayed in the console using ASCII characters. The color table can also be output to an HTML file and viewed in a web browser for a more visual representation.
#>

    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        [string] 
        $directoryPath,

        [switch]
        $displayColorTable
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

        # Remove duplicate color matches
        $uniqueColorMatches = $colorMatches | Sort-Object -Unique
    
        # Create a custom object for each unique color match
        $colorObjects = $uniqueColorMatches | ForEach-Object {
            [PSCustomObject]@{
                colorCode = $_
            }
        }


    }
    
    end{

        # Output the list of unique color matches, should any exist.
        # Otherwise, print a message informing the user nothing was found.
        $colorObjects.Length -gt 0 `
            ? $colorObjects `
            : "No hex color codes could be found."

        # Check if the displayColorTable switch is used
        if($displayColorTable){
            '<style>'+
               'body{background: #111111; font-family: "Segoe UI"}'+
            '</style>'+
            '<table style="background: #333333; color: white;" border="1">' +
            '<tr>' +
                '<th style="padding: 0 1em;">Code</th>'+
                '<th style="padding: 0 1em;">Color</th>'+
            '</tr>'+
            $( Find-HexColorCodes -directoryPath $directoryPath |% {
                '<tr>'+
                "<td>$($_.colorCode)</td>"+
                "<td style=""background-color: $($_.colorCode);""></td>"+
                '</tr>'
            }) +
            "</table>" | Out-File ~\Documents\colors.html `
                       && Invoke-Item ~\Documents\colors.html `
                       && sleep 1 `
                       && Remove-Item ~\Documents\colors.html
        }
    }
    
}


Export-ModuleMember -Function Find-HexColorCodes


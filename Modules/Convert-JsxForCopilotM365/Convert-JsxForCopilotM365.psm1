function Convert-JsxForCopilotM365 {
<#
    .SYNOPSIS
    Converts JSX tags to HTML entities using a Python script.

    .DESCRIPTION
    The Convert-JsxForCopilotM365 function takes JSX input, either as an array of objects or from the clipboard, and converts the JSX tags to HTML entities using the "sub_tags_with_entities.py" Python script. The converted JSX is then outputted.

    .PARAMETER inputJsx
    An array of objects representing the JSX to be converted. If not provided, the function uses the contents of the clipboard as the input JSX.

    .EXAMPLE
    Convert-JsxForCopilotM365 -inputJsx "<h1>Hello, World!</h1>"
    Converts the input JSX string "<h1>Hello, World!</h1>" to HTML entities using the "sub_tags_with_entities.py" Python script.

    .INPUTS
    System.Object[]

    .OUTPUTS
    System.String

    .NOTES
    The "sub_tags_with_entities.py" Python script must take a JSX string as a command-line argument and output the converted JSX to the standard output.
#>
    [CmdletBinding()]
    param (
        [Object[]] $inputJsx = (Get-Clipboard)
    )

    begin {
        $jsxAsString = $inputJsx -join "`n"
    }

    process {
        $convertedJsx = python "$PyPath\sub_tags_with_entities.py" $jsxAsString
    }

    end {
        $convertedJsx
    }
}
    
Export-ModuleMember -Function Convert-JsxForCopilotM365
function Invoke-CodeChunker {
<#
    .SYNOPSIS
    Invokes a Python script to chunk code into 2000 character segments for sharing in the Chat window.

    .DESCRIPTION
    This function is designed to facilitate the sharing of large code blocks in the Chat window by automatically chunking them into segments of 2000 characters. It eliminates the need to manually track breakpoints in code samples that exceed this character limit, ensuring seamless communication of code snippets with Microsoft 365 Copilot.

    .PARAMETER codeToChunk
    An array of code lines to be chunked. If not provided, the function defaults to using the content from the clipboard.

    .EXAMPLE
    Invoke-CodeChunker -codeToChunk $myCodeArray
    Chunks the code contained in $myCodeArray and prepares it for sharing with Microsoft 365 Copilot in the Chat window.

    .EXAMPLE
    Invoke-CodeChunker
    Automatically chunks the clipboard content for easy sharing with Microsoft 365 Copilot.

    .INPUTS
    Object[]
    You can pipe objects to Invoke-CodeChunker.

    .OUTPUTS
    None
    This function does not produce any output to the pipeline. It performs an action by invoking a Python script and preparing code segments for sharing.

    .NOTES
    Ensure that the Python script path is correctly specified in the $PyPath variable. The Python environment must be activated before running the script and deactivated after completion. The primary use case is to interact with Microsoft 365 Copilot by sharing code segments without manual intervention.
#>
    [CmdletBinding()]
    param (
        [Object[]] $codeToChunk = ( Get-Clipboard )
    )

    begin {
        # Activate the python environment containing the required dependencies
        & $PyEnv

        # Convert the code array into a single string with newline characters
        $codeToChunkAsString = $codeToChunk -join "`n"
    }

    process {
        # Call the Python script with the code string as an argument
        python "$PyPath\code_chunker.py" $codeToChunkAsString
    }

    end {
        # Deactivate the python environmwnt after processing
        deactivate
    }
}

# Export Invoke-CodeChunker as a custom module
Export-ModuleMember -Function Invoke-CodeChunker
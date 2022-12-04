function Write-Figlet {
<#
    .SYNOPSIS
    Converts a provided string input to generative ASCII Art produced by the python pyfiglet module.
    

    .DESCRIPTION
    Calls an external python script, Write_Figlet.py, passing the Text and Font parameters as arguments.
    Following the execution of script, the Text provided is then printed to console in the ASCII typesetting 
    specified by the provided Font.   


    .PARAMETER Text
    The string input which will be converted to ASCII typesetting.


    .PARAMETER Font
    A string representing which of the available Font Options will be used.


    .PARAMETER FontOptions
    A switch indicating if the full list of Font Options should be retreived.
    If provided, the list of Font Options will be returned as a System.Object[] 


    .PARAMETER PythonScript
    The path the the Write_Fig;et.py script


    .PARAMETER PythonEnvironment
    The path to the Python Virtual Environment which stores the module dependencies


    .EXAMPLE
    Write-Figlet "Example Text" 
     ____                        _      _        
    |  _ \ __ _ ___ ___  ___  __| |    / \   ___
    | |_) / _` / __/ __|/ _ \/ _` |   / _ \ / __|
    |  __/ (_| \__ \__ \  __/ (_| |  / ___ \\__ \
    |_|   \__,_|___/___/\___|\__,_| /_/   \_\___/

         _                                         _
        / \   _ __ __ _ _   _ _ __ ___   ___ _ __ | |_
       / _ \ | '__/ _` | | | | '_ ` _ \ / _ \ '_ \| __|
      / ___ \| | | (_| | |_| | | | | | |  __/ | | | |_
     /_/   \_\_|  \__, |\__,_|_| |_| |_|\___|_| |_|\__|
                  |___/


    .EXAMPLE
    "Pased via Pipeline" | Write-Figlet
    ____                        _         _       
    |  _ \ __ _ ___ ___  ___  __| | __   _(_) __ _
    | |_) / _` / __/ __|/ _ \/ _` | \ \ / / |/ _` |
    |  __/ (_| \__ \__ \  __/ (_| |  \ V /| | (_| |
    |_|   \__,_|___/___/\___|\__,_|   \_/ |_|\__,_|

    ____  _            _ _
    |  _ \(_)_ __   ___| (_)_ __   ___
    | |_) | | '_ \ / _ \ | | '_ \ / _ \
    |  __/| | |_) |  __/ | | | | |  __/
    |_|   |_| .__/ \___|_|_|_| |_|\___|
            |_|
    

    .EXAMPLE
    ( Write-Figlet -FontOptions ) |? {$_ -match 'star'}

    kik_star
    starwars
    star_war

    
    .NOTES
    Write-Figlet is functionally dependent on the following: 
        
        python script: 
            Write_Figlet.py

        python modules:
            sys
            json
#>
    
    [CmdletBinding()]

    param (

        [Parameter(
            ValueFromPipeline=$true
        )]
        [string]
        $Text,

        [string]
        $Font,

        [switch]
        $FontOptions,

        [string]
        $PythonScript = "$($HOME)\Dev\projects\Snake_Bytes\scripts\Write_Figlet.py",

        [string]
        $PythonEnvironment = "$($HOME)\Dev\python\thunderDome\Scripts\Activate.ps1"

    )
    
    begin {
        
        Write-Verbose "Loading the Python Environment at $PythonEnvironment"
        & $PythonEnvironment

    }
    
    process {
    
        if( -not ( $FontOptions ) ){
        
            Write-Verbose "Passing '$Text' and '$Font' as arguments to the Python Script at $PythonScript"
            python $PythonScript $Text $Font

        }

    }
    
    end {

        if( $FontOptions ){

            Write-Verbose (
                "A list of Font Options were requested. " +
                "Passing 'show_font_options' argument to Python Script at $PythonScript"
            )


            Write-Verbose (
                "The full list of Figlet Font Options will now be converted to a JSON object" +
                " and then a PowerShell Object"
            )

            python $PythonScript 'show_font_options' | ConvertFrom-Json

        }


        else{

            if( $Font.Length -lt 1 ){

                $Font = "Default"    
    
            }
    
    
            Write-Verbose "'$Text' was printed using the '$Font' figlet font"
    
        }
            
    }

}

Export-ModuleMember -Function Write-Figlet
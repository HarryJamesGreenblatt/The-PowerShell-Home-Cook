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
    
    
        .PARAMETER Width
        An integer representing the number of input characters to account for in a line of output.
        The default value is 80. 
    
    
        .PARAMETER FontOptions
        A switch indicating if the full list of Font Options should be retreived.
        If provided, the list of Font Options will be returned as a System.Object[] 
    
    
        .PARAMETER PythonScript
        The path the the Write_Figlet.py script
    
    
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
        (Write-Figlet -FontOptions) |? {$_ -match 'banner'} 
    
        banner
        banner3-D
        banner3
        banner4
    
        
        .EXAMPLE
        "Passed via Pipeline" | Write-Figlet -Font "banner3-D"
    
        ########:::::'###:::::'######:::'######::'########:'########::
        ##.... ##:::'## ##:::'##... ##:'##... ##: ##.....:: ##.... ##:
        ##:::: ##::'##:. ##:: ##:::..:: ##:::..:: ##::::::: ##:::: ##:
        ########::'##:::. ##:. ######::. ######:: ######::: ##:::: ##:
        ##.....::: #########::..... ##::..... ##: ##...:::: ##:::: ##:
        ##:::::::: ##.... ##:'##::: ##:'##::: ##: ##::::::: ##:::: ##:
        ##:::::::: ##:::: ##:. ######::. ######:: ########: ########::
        ..:::::::::..:::::..:::......::::......:::........::........:::
        '##::::'##:'####::::'###::::
         ##:::: ##:. ##::::'## ##:::
         ##:::: ##:: ##:::'##:. ##::
         ##:::: ##:: ##::'##:::. ##:
        . ##:: ##::: ##:: #########:
        :. ## ##:::: ##:: ##.... ##:
        ::. ###::::'####: ##:::: ##:
        :::...:::::....::..:::::..::
        ########::'####:'########::'########:'##:::::::'####:'##::: ##:'########:
        ##.... ##:. ##:: ##.... ##: ##.....:: ##:::::::. ##:: ###:: ##: ##.....::
        ##:::: ##:: ##:: ##:::: ##: ##::::::: ##:::::::: ##:: ####: ##: ##:::::::
        ########::: ##:: ########:: ######::: ##:::::::: ##:: ## ## ##: ######:::
        ##.....:::: ##:: ##.....::: ##...:::: ##:::::::: ##:: ##. ####: ##...::::
        ##::::::::: ##:: ##:::::::: ##::::::: ##:::::::: ##:: ##:. ###: ##:::::::
        ##::::::::'####: ##:::::::: ########: ########:'####: ##::. ##: ########:
        ..:::::::::....::..:::::::::........::........::....::..::::..::........::
        
    
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
    
            [int]
            $Width = 85,
    
            [switch]
            $FontOptions,
    
            [string]
            $PythonScript = "$PyPath\Write_Figlet.py",
    
            [string]
            $PythonEnvironment = $PyEnv
    
        )
        
        begin {
            
            Write-Verbose "Loading the Python Environment at $PythonEnvironment"
            & $PythonEnvironment
    
        }
        
        process {
        
            if( -not ( $FontOptions ) ){
            
                Write-Verbose "Passing '$Text' and '$Font' as arguments to the Python Script at $PythonScript"
                python $PythonScript $Text $Font $Width
    
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

            Write-Verbose "deactivating the Python Environment"
            deactivate
                
        }
    
    }
    
    Export-ModuleMember -Function Write-Figlet
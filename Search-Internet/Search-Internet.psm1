function Search-Internet {
    <#
        .SYNOPSIS
        Searches the web for the given Query String on an optionally specified Search Engine and Browser.
    
        .DESCRIPTION
        Given an input 'Question' string, converts that input to a URI Encoded String by passing it into JavaScript's built-in 
        encodeURIComponent() function, which returns a string replacing all non-alphabetic characters with 
        specifiers designated with a Percent Sign  (i.e '<blank space>' =  %20). 
        
        The Encoded String is then added to the provided Search Engine's associated base URL, thus completing a 'Search Query',
        which itself can be passed directly into the specified Browser's address bar.
        
        By invoking the Start-Proceess cmndlet both with the speified Browser's Process Name and the constructed Search Queary 
        provided as input parameters, the specified Browser thereby opens itself and executes a Search for the provided Question
        in the specified Browser using the specified Search Engine.   
    
        .PARAMETER Question
        The Search Term or Search Query to be looked up on the Internet.
    
        .PARAMETER SearchEngine
        The Search Engine that will be used to look the Question up. 
    
            Options: 
                - Bing
                - Google
                - DuckDuckGo (Default)
    
        .PARAMETER Browser
        The Browser that will be used to look the Question up. 
    
            Options: 
                - Edge (Default)
                - Chrome
    
        .EXAMPLE
        Search-Internet "What's the weather today?"
    
        (By default opens Edge Browser and used the DuckDuckGO Search Engine)
    
        .EXAMPLE
        Search-Internet "wWhat's the weather today?" -SearchEngine Google -Browser Chrome
    
        (Launches the Chrome Browser and executes a Search on the Google Search Engine)
    
    #>
        [CmdletBinding()]
        param (
    
            [Parameter(Mandatory=$true)]
            [string]
            $Question,
    
            [string] 
            $SearchEngine = "DuckDuckGo",
    
            [string] 
            $Browser = 'Edge'
    
        )
        
        begin {
    
    
            Write-Verbose "The given Question is: $Question`n`n"
            Write-Verbose "Now Encoding the given Question to be URI compatible`n`n"
    
            $URIEncodedQuestion =  "console.log(encodeURIComponent(`"$Question`"))" | node
    
            Write-Verbose "The URI Encoded Question is:  $URIEncodedQuestion`n`n"
            
            
            $SearchQuery = ''
            $BrowserProcessName = ''
            
            
            Write-Verbose "Now Constructing the Search Query string based on the given SearchEngine parameter`n`n"
            
            switch ($SearchEngine) {
                
                "Bing"     { $SearchQuery = 'www.bing.com/search\?q='+$URIEncodedQuestion }
                "Google"   { $SearchQuery = 'www.google.com/search?hl=en&q='+$URIEncodedQuestion }
                "DuckDuckGo" { $SearchQuery = 'www.duckduckgo.com/?q='+$URIEncodedQuestion }
                default  {  
                    Write-Error -ErrorAction Stop -Message "
                    Please specify one of the following Search Engines:`n
                    `t- Bing`n
                    `t- Google 
                    "
                }
                
            }
            
            
            Write-Verbose "Now aligning the given Browser parameter with it's associated Process Name`n`n"
            
            switch ($Browser) {
                
                "Edge"       { $BrowserProcessName = 'msedge' }
                "Chrome"     { $BrowserProcessName = 'chrome' }
                default  { 
                    Write-Error -ErrorAction Stop -Message "
                    Please specify one of the following Search Engines:`n
                    `t- Edge`n
                    `t- Chrome 
                    "
                }
                
            }
    
        }
        
        process {
    
            Write-Verbose "Now opening up a Search Query on '$Question' in '$Browser' using '$SearchEngine's Search Engine`n`n"
           
            Start-Process $BrowserProcessName -ArgumentList $SearchQuery
    
        }
        
        end {
    
            Write-Verbose "The the Search Query is now deployed in the Browser window" 
        
        }
        
    }
    
    Export-ModuleMember -Function Search-Internet
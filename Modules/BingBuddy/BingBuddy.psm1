 # BingBuddy.psm1
<#
    .SYNOPSIS
    BingBuddy is a PowerShell module that provides functions to interact with the Bing Search API.

    .DESCRIPTION
    BingBuddy is designed to simplify the process of making various types of searches using the Bing Search API. It includes functions to invoke searches, process results, and open search result URLs in a web browser.

    .FUNCTIONS
    - Get-BingSearchResults: Performs a search using the Bing Search API and returns the results.
    - Open-BingSearchResult: Opens the URL from a Bing search result in the default web browser.

    .EXAMPLE
    # Example of using Get-BingSearchResults
    $results = Get-BingSearchResults -Query "PowerShell" -Service "web"
    $results | Format-Table

    # Example of using Open-BingSearchResult
    $results | Open-BingSearchResult

    .NOTES
    To use the BingBuddy module, you must have a valid Bing Search API key. Ensure that you handle the API key securely and do not expose it in scripts or logs.

    .LINK
    https://docs.microsoft.com/en-us/azure/cognitive-services/bing-web-search/

#>



function Get-BingSearchResults {
<#
    .SYNOPSIS
    Invokes a Bing Search and returns results based on the specified query and service type.

    .DESCRIPTION
    This function makes a call to the Bing Search API and returns results for a specified query. 
    It supports various service types including web, image, video, news, and more.

    .PARAMETER Query
    The search query string to be submitted to the Bing Search API.

    .PARAMETER Service
    The type of search service to use. Valid options are web, image, video, news, custom, entity, auto, spell, visual, and local.

    .PARAMETER ApiKey
    The API key for authenticating with the Bing Search API. If not specified, the function will use the value of the $BingSearchApiKey variable.

    .PARAMETER ResultsCount
    The number of search results to return. If not specified, the default number of results defined by the API will be returned.

    .PARAMETER NSFW
    A switch to include Not Safe For Work (NSFW) content in the search results. If not specified, NSFW content will be excluded.

    .PARAMETER Market
    The geographic region to which the result data is localized. 

    .EXAMPLE
    Get-BingSearchResults -Query "PowerShell" -Service "web"

    This example invokes a web search for the query "PowerShell" and returns the results.

    .EXAMPLE
    Get-BingSearchResults -Query "Cats" -Service "image" -ResultsCount 10 -NSFW

    This example invokes an image search for the query "Cats", returns 10 results, and includes NSFW content.

    .NOTES
    This function requires an active internet connection and a valid Bing Search API key to function.

    .LINK
    https://docs.microsoft.com/en-us/azure/cognitive-services/bing-web-search/
#>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Query,
        
        [Parameter(Mandatory)]
        [ValidateSet(
            "web",
            "images",
            "videos",
            "news",
            "entities",
            "suggestions",
            "spelling",
            "visual",
            "local")]
        [string]
        $Service,
            
        [Parameter()]
        [string]
        $ApiKey = $BingSearchApiKey,

        [Parameter()]
        [int]
        $ResultsCount,

        [Parameter()]
        [switch]
        $NSFW,

        [Parameter()]
        [string]
        $Market = "en-US"
    )

    begin {

        # Helper function to add the Service property to each result object
        function Add-ServiceProperty {
            param (
                [PSCustomObject]$ResultObject,
                [string]$Service
            )
            $ResultObject | Add-Member `
                -NotePropertyName 'service' `
                -NotePropertyValue $Service `
                -PassThru
        }


        # Validate API Key. Exit program if found not to be valid.
        if (-not $ApiKey ){
            Write-Error "You need to provide a valid Bing Search API key." -ErrorAction Stop
        }


        # Create the headers hash using the API key
        $headers = @{
            "Ocp-Apim-Subscription-Key" = $ApiKey
        }


        # URL encode the Query using Python 
        $QueryEncoded = python -c "import urllib.parse as up; print(up.quote(up.quote_plus('$Query')))"

        
        Write-Verbose("`nQuery: $Query`nurl encoded: $QueryEncoded")


        # Set the appropriate endpoint string based on the service path
        $baseUrl = "https://api.bing.microsoft.com/v7.0"

        switch ($Service) {
            "web"      { $endpoint = "/search" }
            "images"   { $endpoint = "/images/search" }
            "videos"   { $endpoint = "/videos/search" }
            "news"     { $endpoint = "/news/search" }
            "entities" { $endpoint = "/entities" }
            "suggestions"     { $endpoint = "/suggestions" }
            "spelling" { $endpoint = "/spellcheck" }
            "local"    { $endpoint = "/localbusiness/search" }
            default    { throw "Invalid service path provided." }
        }

        Write-Verbose("`nservice path: $Service`nendpoint: $endpoint")

    }

    process {
        # Construct the request URL
        $url = $Service -eq "spelling" `
            ? "$($baseUrl)$($endpoint)?text=$($QueryEncoded)" 
            : "$($baseUrl)$($endpoint)?q=$($QueryEncoded)"

        # Add optional parameters to the URL
        if ($ResultsCount) {
            $url += "&count=$ResultsCount"
        }
        if ($NSFW) {
            $url += "&safeSearch=Off"
        }
        $url += "&mkt=$Market"    

        Write-Verbose "`nurl: $url"
        
        # Make the API call
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method 'GET'

        # Process the response based on the service type
        switch -Regex ($Service) {
            "web"         { $response.webPages.value      | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "images"      { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "videos"      { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "news"        { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "entities"    { $response.entities.value      | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "spelling"    { $response.flaggedTokens       | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "suggestions" { 
                            $response.suggestionGroups.searchSuggestions 
                                | ForEach-Object { Add-ServiceProperty $_ $Service } 
                          }
            default       { throw "Invalid service path provided." }
        }

        Write-Verbose "`nurl: $url"

    }

    end {
        # Define a hashtable mapping the service paths to their expected response properties
        $responsePropertyMap = @{
            "web"                 = 'webPages';
            "images"              = 'images';
            "videos"              = 'videos';
            "news"                = 'news';
            "entities"            = 'entities';
            "suggestions"         = 'suggestions';
            "spelling"            = 'flaggedTokens';
        }

        # Determine the expected response property based on the service path
        $expectedProperty = $responsePropertyMap[$Service]


        Write-Verbose "the response is $response "


        # Check if the response service was 'suggestion' and return its value if so
        if ($response.PSObject.Properties.Name -match 'suggestion') {
            $response.suggestionGoups.searchSuggestions
        } 

        # Otherwise, check if the response contains the expected property and return its value
        elseif ($response.PSObject.Properties.Name -contains $expectedProperty) {
            $response.$expectedProperty.value
        } 
        
        # Fallback to 'value' if the expected property is not found
        elseif ($response.PSObject.Properties.Name -contains 'value') {
            $response.value
        } 
        
        else {
            # Handle cases where the response structure is unexpected
            Write-Error "Unexpected response structure for service path: $Service"
        }
    }
}




function Open-BingSearchResult {
<#
    .SYNOPSIS
    Opens the content URL from a Bing search result in the default web browser.

    .DESCRIPTION
    This function takes a single record from the output of Get-BingSearchResults, 
    which should contain a URL property, and opens it in the default web browser using Start-Process.

    .PARAMETER SearchResult
    The search result object that contains the URL to be opened. 
    This parameter can accept input directly or via the pipeline.

    .EXAMPLE
    $bingSearchResult | Open-BingSearchResult

    This example takes a Bing search result object and opens the content URL in the default web browser.

    .EXAMPLE
    Open-BingSearchResult -SearchResult $bingSearchResult

    This example opens the content URL from the specified Bing search result in the default web browser.

    .NOTES
    Ensure that the search result object contains a valid URL property. 
    The function will output an error if the URL is not found or is invalid.

    .LINK
    https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-arrays?view=powershell-7.1
#>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline)]
        [PSCustomObject]$SearchResult
    )
        
    begin {  
        # Define a hashtable mapping the service types to their URL properties
        $urlPropertyMap = @{
            "web"          = 'url';
            "images"       = 'contentUrl';
            "videos"       = 'contentUrl';
            "news"         = 'url';
            "entities"     = 'url';
            "suggestions"  = 'url';
            "spelling"     = 'url';
        }
    }
    
    process {
        #Retreive the Service property of the SearchResult parameter 
        $Service = $SearchResult.Service

        # Determine the URL property based on the service type
        $urlProperty = $urlPropertyMap[$Service]

        # Check if the search result contains the URL property and open it
        if ($SearchResult.PSObject.Properties.Name -contains $urlProperty) {
            Start-Process $SearchResult.$urlProperty
        } else {
            Write-Error "No URL found in the search result for service type: $Service"
        }
    }
}



Export-ModuleMember -Function `
    Get-BingSearchResults, 
    Open-BingSearchResult
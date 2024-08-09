# BingBuddy.psm1
<#
    .SYNOPSIS
    BingBuddy is a PowerShell module that provides functions to interact with the Bing Search API.

    .DESCRIPTION
    BingBuddy is designed to simplify the process of making various types of searches using the Bing Search API. 
    It includes functions to invoke searches, process results, and open search result URLs in a web browser.

    .FUNCTIONALITY
    - Get-BingSearchResults: Performs a search using the Bing Search API and returns unique results.
    - Invoke-BingSearch: Invokes a Bing Search and returns results based on the specified query and service type.
    - Open-BingSearchResult: Opens the URL from a Bing search result in the default web browser.
    - Receive-BingNews: Retrieves news articles using the Bing News Search API, optionally filtered by a specific category.
    - Receive-BingNewsTrendingTopics: Retrieves trending news topics using the Bing Search API.

    .EXAMPLE
    # Example of using Get-BingSearchResults
    $results = Get-BingSearchResults -Query "PowerShell" -Service "web"
    $results | Format-Table

    # Example of using Open-BingSearchResult
    $results | Open-BingSearchResult

    # Example of using Receive-BingNews
    $news = Receive-BingNews -Category "Technology" -ApiKey "YourApiKey"
    $news | Format-Table

    # Example of using Receive-BingNewsTrendingTopics
    $trendingTopics = Receive-BingNewsTrendingTopics -ApiKey "YourApiKey"
    $trendingTopics | Format-Table

    .NOTES
    To use the BingBuddy module, you must have a valid Bing Search API key. 
    Ensure that you handle the API key securely and do not expose it in scripts or logs.

    .LINK
    https://docs.microsoft.com/en-us/azure/cognitive-services/bing-web-search/
#>



function Invoke-BingSearch {
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
        [boolean]
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
            "web"         { $endpoint = "/search" }
            "images"      { $endpoint = "/images/search" }
            "videos"      { $endpoint = "/videos/search" }
            "news"        { $endpoint = "/news/search" }
            "entities"    { $endpoint = "/entities" }
            "suggestions" { $endpoint = "/suggestions" }
            "spelling"    { $endpoint = "/spellcheck" }
            "local"       { $endpoint = "/localbusiness/search" }
            default       { throw "Invalid service path provided." }
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


function Get-BingSearchResults {
    <#
    .SYNOPSIS
    Performs a search using the Bing Search API and returns unique results.

    .DESCRIPTION
    This function is a wrapper around the Invoke-BingSearch function. 
    It performs a search using the Bing Search API and filters out duplicate results.

    .PARAMETER Query
    The search query string to be submitted to the Bing Search API.

    .PARAMETER Service
    The type of search service to use. 
    Valid options are web, images, videos, news, entities, suggestions, spelling, visual, and local.

    .PARAMETER ApiKey
    The API key for authenticating with the Bing Search API. 
    If not specified, the function will use the value of the $BingSearchApiKey variable.

    .PARAMETER ResultsCount
    The number of search results to return. If not specified, 
    the default number of results defined by the API will be returned.

    .PARAMETER NSFW
    A switch to include Not Safe For Work (NSFW) content in the search results. 
    If not specified, NSFW content will be excluded.

    .PARAMETER Market
    The geographic region to which the result data is localized. 

    .EXAMPLE
    Get-BingSearchResults -Query "PowerShell" -Service "web"

    This example performs a web search for the query "PowerShell" and returns unique results.

    .EXAMPLE
    Get-BingSearchResults -Query "Cats" -Service "images" -ResultsCount 10 -NSFW

    This example performs an image search for the query "Cats", returns 10 unique results, and includes NSFW content.

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
  
    process {
        # Call the Invoke-BingSearch function to get raw search results
        Invoke-BingSearch `
           -Query $Query `
           -Service $Service `
           -ApiKey $BingSearchApiKey `
           -ResultsCount $ResultsCount `
           -NSFW $NSFW `
           -Market $Market 
                | 
                Select-Object * -Unique
    }
}


function Receive-BingNewsTrendingTopics {
    <#
    .SYNOPSIS
    Retrieves trending news topics using the Bing Search API.

    .DESCRIPTION
    This function makes a call to the Bing Search API to retrieve trending news topics.

    .PARAMETER ApiKey
    The API key for authenticating with the Bing Search API. If not specified, the function will use the value of the $BingSearchApiKey variable.

    .PARAMETER Market
    The geographic region to which the result data is localized. 

    .EXAMPLE
    Receive-BingTrendingTopics

    This example retrieves trending news topics.

    .NOTES
    This function requires an active internet connection and a valid Bing Search API key to function.
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [string]
        $ApiKey = $BingSearchApiKey,

        [Parameter()]
        [string]
        $Market = "en-US"
    )

    begin {
        # Validate API Key
        if (-not $ApiKey) {
            Write-Error "You need to provide a valid Bing Search API key." -ErrorAction Stop
        }

        # Create the headers hash using the API key
        $headers = @{
            "Ocp-Apim-Subscription-Key" = $ApiKey
        }

        # Set the endpoint URL
        $url = "https://api.bing.microsoft.com/v7.0/news/trendingtopics?mkt=$Market"
    }

    process {
        # Make the API call
        $response = Invoke-RestMethod -Uri $url -Headers $headers -Method 'GET'

        # Return the trending topics
        return $response.value
    }
}


function Receive-BingNews {
    <#
    .SYNOPSIS
    Retrieves news articles using the Bing News Search API, optionally filtered by a specific category.

    .DESCRIPTION
    This function makes a call to the Bing News Search API to retrieve news articles. 
    If a category is specified, it retrieves news articles for that category. Otherwise, it retrieves general news articles.

    .PARAMETER Category
    The news category to retrieve articles for. Valid options include:
    - Business
    - Entertainment
    - Health
    - Politics
    - Products
    - Technology
    - Science
    - Sports
    - US
    - World

    .PARAMETER Trending
    A switch to retrieve trending news topics instead of regular news articles.
    
    .PARAMETER ApiKey
    The API key for authenticating with the Bing News Search API. 
    If not specified, the function will use the value of the $BingSearchApiKey variable.

    .EXAMPLE
    Receive-BingNews -Category "Technology" -ApiKey "YourApiKey"

    This example retrieves technology news articles using the specified API key.

    .EXAMPLE
    Receive-BingNews -Trending -ApiKey "YourApiKey"

    This example retrieves trending news topics using the specified API key.

    .NOTES
    This function requires an active internet connection and a valid Bing Search API key to function.

    .LINK
    https://docs.microsoft.com/en-us/azure/cognitive-services/bing-news-search/
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet(
            "Business",
            "Entertainment",
            "Health",
            "Politics",
            "Products",
            "Technology",
            "Science",
            "Sports",
            "US",
            "World")]
        [string]
        $Category,
        
        [Parameter()]
        [switch]
        $Trending,

        [Parameter()]
        [string]
        $ApiKey = $BingSearchApiKey,

        [Parameter()]
        [string]
        $Market = "en-US"
    )

    process {

        if($Trending){
            Receive-BingNewsTrendingTopics -ApiKey $ApiKey -Market $Market
        }

        else{
        
            # Validate API Key. Exit program if found not to be valid.
            if (-not $ApiKey) {
                Write-Error "You need to provide a valid Bing Search API key." -ErrorAction Stop
            }

            # Create the headers hash using the API key
            $headers = @{
                "Ocp-Apim-Subscription-Key" = $ApiKey
            }

            # Set the base URL for the Bing News Search API
            $baseUrl = "https://api.bing.microsoft.com/v7.0/news"

            # Construct the request URL
            $url = $baseUrl

            # Add category to the URL if specified
            if ($Category) {
                $url += "?category=$Category"
            }

            # Construct market parameter to the URL
            $marketParam = $url.Length -gt $baseUrl.Length ? "&mkt=$Market" : "?mkt=$Market"

            # Add market parameter to the URL
            $url += $marketParam

            Write-Verbose "`nurl: $url"

            # Make the API call
            $response = Invoke-RestMethod -Uri $url -Headers $headers -Method 'GET'

            # Process the response
            $results = $response.value

            # Return the results
            $results
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
    It also handles the edge case where the -Trending switch is provided to the Receive-BingNews function,
    which may return objects with both webSearchUrl and newsSearchUrl properties.

    .PARAMETER SearchResult
    The search result object that contains the URL to be opened. 
    This parameter can accept input directly or via the pipeline.

    .PARAMETER Source
    Specifies which URL property to use when the search result contains either webSearchUrl or newsSearchUrl properties.

    .EXAMPLE
    $bingSearchResult | Open-BingSearchResult

    This example takes a Bing search result object and opens the content URL in the default web browser.

    .EXAMPLE
    Open-BingSearchResult -SearchResult $bingSearchResult

    This example opens the content URL from the specified Bing search result in the default web browser.

    .EXAMPLE
    Receive-BingNews -Trending -ApiKey "YourApiKey" | Open-BingSearchResult -TrendSource "webSearchUrl"

    This example retrieves trending news topics and opens the web search URL in the default web browser.

    .NOTES
    Ensure that the search result object contains a valid URL property. 
    The function will output an error if the URL is not found or is invalid.

    .LINK
    https://docs.microsoft.com/en-us/powershell/scripting/learn/deep-dives/everything-about-arrays?view=powershell-7.1
    #>
    [CmdletBinding()]
    param (
        [Parameter(
            ValueFromPipeline, 
            Mandatory)]
        [PSCustomObject]
        $SearchResult,

        [Parameter()]
        [ValidateSet(
            "web",
            "news"
        )]
        [string]
        $Source
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

        # Retreive the Service property of the SearchResult parameter,
        # should any such Service property exist.
        $Service = $SearchResult.Service

        if($Service){

            # Determine the URL property based on the service type
            $urlProperty = $urlPropertyMap[$Service]

            # Check if the search result contains the URL property and open it
            if ($SearchResult.PSObject.Properties.Name -contains $urlProperty) {
                Start-Process $SearchResult.$urlProperty
            } else {
                Write-Error "No URL found in the search result for service type: $Service"
            }     
            
        }

        else{

            # Otherwise,
            # Assess wether or not the SearchResult is a Trending Topic,
            # then open the result accordingly
            switch ($Source) {

                web   { Start-Process $SearchResult.webSearchUrl  }
                news  { Start-Process $SearchResult.newsSearchUrl }

                Default        {

                    if ($SearchResult.PSObject.Properties.Name -ccontains "url") {
                        Start-Process $SearchResult.url   
                    } 
                    else {

                        $choices = @("webSearchUrl", "newsSearchUrl")
                        Write-Host "A Source Selection Menu has been launched, please check your taskbar for a PowerShell icon."

                        $selection = $choices | Out-GridView -Title "Which Source do you want to check?" -PassThru
                        Start-Process $SearchResult.$selection           
                    }  
                    
                }
            }           
        }

    }
}


Export-ModuleMember -Function `
    Get-BingSearchResults, 
    Receive-BingNews,
    Open-BingSearchResult

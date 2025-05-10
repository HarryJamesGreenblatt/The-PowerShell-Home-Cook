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
    The API key for authenticating with the Bing Search API. If not specified, the function will use the value of the $env:BingSearchApiKey variable.

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
            "Web",
            "Images",
            "Videos",
            "News",
            "Entities",
            "Suggestions",
            "Spelling")]
        [string]
        $Service,
            
        [Parameter()]
        [string]
        $ApiKey = $env:BingSearchApiKey,

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
            "Web"         { $endpoint = "/search" }
            "Images"      { $endpoint = "/images/search" }
            "Videos"      { $endpoint = "/videos/search" }
            "News"        { $endpoint = "/news/search" }
            "Entities"    { $endpoint = "/entities" }
            "Suggestions" { $endpoint = "/suggestions" }
            "Spelling"    { $endpoint = "/spellcheck" }
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
            "Web"         { $response.webPages.value      | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "Images"      { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "Videos"      { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "News"        { $response.value               | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "Entities"    { $response.entities.value      | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "Spelling"    { $response.flaggedTokens       | ForEach-Object { Add-ServiceProperty $_ $Service } }
            "Suggestions" { 
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
            "Web"                 = 'webPages';
            "Images"              = 'images';
            "Videos"              = 'videos';
            "News"                = 'news';
            "Entities"            = 'entities';
            "Suggestions"         = 'suggestions';
            "Spelling"            = 'flaggedTokens';
        }

        # Determine the expected response property based on the service path
        $expectedProperty = $responsePropertyMap[$Service]


        Write-Verbose "the response is $response"


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


function Receive-BingNewsTrendingTopics {
    <#
    .SYNOPSIS
    Retrieves trending news topics using the Bing Search API.

    .DESCRIPTION
    This function makes a call to the Bing Search API to retrieve trending news topics.

    .PARAMETER ApiKey
    The API key for authenticating with the Bing Search API. If not specified, the function will use the value of the $env:BingSearchApiKey variable.

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
        $ApiKey = $env:BingSearchApiKey,

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


function Get-MarketCode {
    param (
        [Parameter(Mandatory)]
        [ValidateSet(
            "United States",
            "United Kingdom",
            "Canada",
            "Australia",
            "France",
            "Germany",
            "Spain",
            "Italy",
            "Brazil",
            "Mexico",
            "India",
            "China",
            "Japan",
            "Russia",
            "Finland",
            "Denmark",
            "Worldwide"
        )]
        [string]
        $Market
    )

    # Map of country/region names to market codes
    $marketMap = @{
        "United States"  = "en-US"
        "United Kingdom" = "en-GB"
        "Canada"         = "en-CA"
        "Australia"      = "en-AU"
        "France"         = "fr-FR"
        "Germany"        = "de-DE"
        "Spain"          = "es-ES"
        "Italy"          = "it-IT"
        "Brazil"         = "pt-BR"
        "Mexico"         = "es-MX"
        "India"          = "en-IN"
        "China"          = "zh-CN"
        "Japan"          = "ja-JP"
        "Russia"         = "ru-RU"
        "Finland"        = "fi-FI"
        "Denmark"        = "da-DK"
        "Worldwide"      = "en-WW"
    }

    # Convert the selected country/region name to its corresponding market code
    if ($marketMap.ContainsKey($Market)) {
        return $marketMap[$Market]
    } else {
        Write-Verbose "Market '$Market' not found in explicit map, returning input value."
        return $Market
    }
}

function Get-LanguageCode {
    <#
    .SYNOPSIS
    Maps a language name to its corresponding Bing API language code.

    .DESCRIPTION
    This function takes a language name and returns the appropriate language code
    for use with the Bing API's setLang parameter.

    .PARAMETER Language
    The language to get the code for.

    .EXAMPLE
    Get-LanguageCode -Language "English"
    # Returns "en"
    
    .EXAMPLE
    Get-LanguageCode -Language "Chinese (Simplified)"
    # Returns "zh-hans"

    .NOTES
    Reference: https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/query-parameters
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [ValidateSet(
            "Arabic", "Basque", "Bengali", "Bulgarian", "Catalan", 
            "Chinese (Simplified)", "Chinese (Traditional)", "Croatian", "Czech", "Danish",
            "Dutch", "English", "English-United Kingdom", "Estonian", "Finnish",
            "French", "Galician", "German", "Gujarati", "Hebrew",
            "Hindi", "Hungarian", "Icelandic", "Italian", "Japanese",
            "Kannada", "Korean", "Latvian", "Lithuanian", "Malay",
            "Malayalam", "Marathi", "Norwegian (Bokmål)", "Polish", 
            "Portuguese (Brazil)", "Portuguese (Portugal)", "Punjabi", "Romanian", "Russian",
            "Serbian (Cyrylic)", "Slovak", "Slovenian", "Spanish", "Swedish",
            "Tamil", "Telugu", "Thai", "Turkish", "Ukrainian", "Vietnamese"
        )]
        [string]
        $Language
    )

    # Map of language names to language codes
    $languageMap = @{
        "Arabic"                = "ar"
        "Basque"                = "eu"
        "Bengali"               = "bn"
        "Bulgarian"             = "bg"
        "Catalan"               = "ca"
        "Chinese (Simplified)"  = "zh-hans"
        "Chinese (Traditional)" = "zh-hant"
        "Croatian"              = "hr"
        "Czech"                 = "cs"
        "Danish"                = "da"
        "Dutch"                 = "nl"
        "English"               = "en"
        "English-United Kingdom" = "en-gb"
        "Estonian"              = "et"
        "Finnish"               = "fi"
        "French"                = "fr"
        "Galician"              = "gl"
        "German"                = "de"
        "Gujarati"              = "gu"
        "Hebrew"                = "he"
        "Hindi"                 = "hi"
        "Hungarian"             = "hu"
        "Icelandic"             = "is"
        "Italian"               = "it"
        "Japanese"              = "jp"
        "Kannada"               = "kn"
        "Korean"                = "ko"
        "Latvian"               = "lv"
        "Lithuanian"            = "lt"
        "Malay"                 = "ms"
        "Malayalam"             = "ml"
        "Marathi"               = "mr"
        "Norwegian (Bokmål)"    = "nb"
        "Polish"                = "pl"
        "Portuguese (Brazil)"   = "pt-br"
        "Portuguese (Portugal)" = "pt-pt"
        "Punjabi"               = "pa"
        "Romanian"              = "ro"
        "Russian"               = "ru"
        "Serbian (Cyrylic)"     = "sr"
        "Slovak"                = "sk"
        "Slovenian"             = "sl"
        "Spanish"               = "es"
        "Swedish"               = "sv"
        "Tamil"                 = "ta"
        "Telugu"                = "te"
        "Thai"                  = "th"
        "Turkish"               = "tr"
        "Ukrainian"             = "uk"
        "Vietnamese"            = "vi"
    }

    # Convert the selected language name to its corresponding language code
    if ($languageMap.ContainsKey($Language)) {
        return $languageMap[$Language]
    } else {
        Write-Verbose "Language '$Language' not found in explicit map, returning input value."
        return $Language
    }
}

function Get-MarketCategoryInfo {
    <#
    .SYNOPSIS
    Returns information about available categories for a specific market.

    .DESCRIPTION
    This function provides information about the news categories available for a given market code.
    Different markets support different sets of news categories based on Bing's implementation.

    .PARAMETER MarketCode
    The market code to get category information for (e.g., "en-US", "zh-CN").

    .EXAMPLE
    Get-MarketCategoryInfo -MarketCode "en-US"

    .NOTES
    Reference: https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/query-parameters#news-categories-by-market
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $MarketCode
    )

    # Comprehensive mapping of market codes to their supported categories
    # Based on https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/query-parameters#news-categories-by-market
    $marketCategoryMap = @{
        # United States (English)
        "en-US" = @{
            "DisplayName" = "United States (English)"
            "Categories" = @(
                "Business", "Entertainment", "Health", "Politics", "ScienceAndTechnology",
                "Sports", "US", "World"
            )
            "Subcategories" = @{
                "Entertainment" = @("Entertainment_MovieAndTV", "Entertainment_Music") # Added Entertainment subcategories
                "ScienceAndTechnology" = @("Technology", "Science") # Added ScienceAndTechnology subcategories
                "Sports" = @("Sports_Golf", "Sports_MLB", "Sports_NBA", "Sports_NFL", "Sports_NHL", "Sports_Soccer", "Sports_Tennis", "Sports_CFB", "Sports_CBB")
                "US" = @("US_Northeast", "US_South", "US_Midwest", "US_West")
                "World" = @("World_Africa", "World_Americas", "World_Asia", "World_Europe", "World_MiddleEast")
            }
        }
        # United Kingdom (English)
        "en-GB" = @{
            "DisplayName" = "United Kingdom (English)"
            "Categories" = @(
                "Business", "Entertainment", "Health", "Politics", "ScienceAndTechnology", 
                "Sports", "UK", "World"
            )
        }
        # Canada (English)
        "en-CA" = @{
            "DisplayName" = "Canada (English)"
            "Categories" = @(
                "Business", "Canada", "Entertainment", "LifeStyle", "Politics", 
                "ScienceAndTechnology", "Sports", "World"
            )
        }
        # China (Chinese)
        "zh-CN" = @{
            "DisplayName" = "China (Chinese)"
            "Categories" = @(
                "Auto", "Business", "China", "Education", "Entertainment", "Military",
                "RealEstate", "ScienceAndTechnology", "Society", "Sports", "World"
            )
        }
        # Japan (Japanese)
        "ja-JP" = @{
            "DisplayName" = "Japan (Japanese)"
            "Categories" = @(
                "Business", "Entertainment", "Japan", "LifeStyle", "Politics",
                "ScienceAndTechnology", "Sports", "World"
            )
        }
        # India (English)
        "en-IN" = @{
            "DisplayName" = "India (English)"
            "Categories" = @(
                "Business", "Entertainment", "India", "LifeStyle", "Politics",
                "ScienceAndTechnology", "Sports", "World"
            )
        }
        # Default fallback for any other market
        "default" = @{
            "DisplayName" = "Global"
            "Categories" = @(
                "Business", "Entertainment", "Politics", "ScienceAndTechnology", 
                "Sports", "World"
            )
        }
    }

    # Return the category info for the specified market, or the default if not found
    if ($marketCategoryMap.ContainsKey($MarketCode)) {
        return $marketCategoryMap[$MarketCode]
    }
    else {
        Write-Verbose "Market code '$MarketCode' not found in category mapping. Using default categories."
        return $marketCategoryMap["default"]
    }
}

function Test-CategoryForMarket {
    <#
    .SYNOPSIS
    Tests if a category is valid for a specific market.

    .DESCRIPTION
    This function checks if a specified category is supported for a given market.
    It helps ensure that users only specify valid categories for their chosen market.

    .PARAMETER Category
    The category to test.

    .PARAMETER MarketCode
    The market code to check against (e.g., "en-US", "zh-CN").

    .EXAMPLE
    Test-CategoryForMarket -Category "RealEstate" -MarketCode "zh-CN"

    .NOTES
    Reference: https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/query-parameters#news-categories-by-market
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory)]
        [string]
        $Category,
        
        [Parameter(Mandatory)]
        [string]
        $MarketCode
    )

    $marketInfo = Get-MarketCategoryInfo -MarketCode $MarketCode
    
    # Check if the category is in the main category list
    if ($marketInfo.Categories -contains $Category) {
        return $true
    }
    
    # Check subcategories if available
    if ($marketInfo.ContainsKey("Subcategories")) {
        foreach ($subcategoryList in $marketInfo.Subcategories.Values) {
            if ($subcategoryList -contains $Category) {
                return $true
            }
        }
    }
    
    return $false
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
    If not specified, the function will use the value of the $env:BingSearchApiKey variable.

    .PARAMETER ResultsCount
    The number of search results to return. If not specified, 
    the default number of results defined by the API will be returned.

    .PARAMETER NSFW
    A switch to include Not Safe For Work (NSFW) content in the search results. 
    If not specified, NSFW content will be excluded.

    .PARAMETER Market
    The geographic region to which the result data is localized. 

    .PARAMETER Language
    The language in which to return search results. By default, the function will use the language
    associated with the selected market. This parameter can be used to override that behavior and
    request results in a specific language.

    .EXAMPLE
    Get-BingSearchResults -Query "PowerShell" -Service "web"

    This example performs a web search for the query "PowerShell" and returns unique results.

    .EXAMPLE
    Get-BingSearchResults -Query "Cats" -Service "images" -ResultsCount 10 -NSFW

    This example performs an image search for the query "Cats", returns 10 unique results, and includes NSFW content.

    .EXAMPLE
    Get-BingSearchResults -Query "Paris tourism" -Service "Web" -Market "France" -Language "English"

    This example performs a web search for "Paris tourism" in the French market but returns results in English.

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
            "Web",
            "Images",
            "Videos",
            "News",
            "Entities",
            "Suggestions",
            "Spelling")]
        [string]
        $Service,
            
        [Parameter()]
        [string]
        $ApiKey = $env:BingSearchApiKey,

        [Parameter()]
        [int]
        $ResultsCount,

        [Parameter()]
        [switch]
        $NSFW,

        [Parameter()]
        [ValidateSet(
            "United States",
            "United Kingdom",
            "Canada",
            "Australia",
            "France",
            "Germany",
            "Spain",
            "Italy",
            "Brazil",
            "Mexico",
            "India",
            "China",
            "Japan",
            "Russia",
            "Finland",
            "Denmark",
            "Worldwide"
        )]
        [string]
        $Market = "United States",
        
        [Parameter()]
        [ValidateSet(
            "Arabic", "Basque", "Bengali", "Bulgarian", "Catalan", 
            "Chinese (Simplified)", "Chinese (Traditional)", "Croatian", "Czech", "Danish",
            "Dutch", "English", "English-United Kingdom", "Estonian", "Finnish",
            "French", "Galician", "German", "Gujarati", "Hebrew",
            "Hindi", "Hungarian", "Icelandic", "Italian", "Japanese",
            "Kannada", "Korean", "Latvian", "Lithuanian", "Malay",
            "Malayalam", "Marathi", "Norwegian (Bokmål)", "Polish", 
            "Portuguese (Brazil)", "Portuguese (Portugal)", "Punjabi", "Romanian", "Russian",
            "Serbian (Cyrylic)", "Slovak", "Slovenian", "Spanish", "Swedish",
            "Tamil", "Telugu", "Thai", "Turkish", "Ukrainian", "Vietnamese"
        )]
        [string]
        $Language
    )

    begin {
        # Convert the selected country/region name to its corresponding market code
        $marketCode = Get-MarketCode -Market $Market
        
        # Convert language name to language code if specified
        $languageCode = if ($Language) {
            Get-LanguageCode -Language $Language
        } else {
            # Default to no specific language code
            $null
        }
    }
  
    process {
        # Call the Invoke-BingSearch function to get raw search results
        # Pass the language code if specified
        $params = @{
            Query = $Query
            Service = $Service
            ApiKey = $ApiKey
            ResultsCount = $ResultsCount
            NSFW = $NSFW
            Market = $marketCode
        }
        
        # Only add the language parameter if it was specified
        if ($languageCode) {
            $params.Language = $languageCode
        }
        
        Invoke-BingSearch @params | Select-Object * -Unique
    }
}


function Receive-BingNews {
    <#
    .SYNOPSIS
    Retrieves news articles using the Bing News Search API, optionally filtered by a specific category.

    .DESCRIPTION
    This function makes a call to the Bing News Search API to retrieve news articles. 
    If a category is specified, it retrieves news articles for that category. Otherwise, it retrieves general news articles.
    The available categories depend on the selected market.

    .PARAMETER Category
    The news category to retrieve articles for. Tab completion shows US market categories, but the 
    actual available categories depend on your chosen market. The function validates if your selected
    category is available for your chosen market and will suggest alternatives if necessary.
    
    Subcategories like US_Midwest, Sports_NFL are handled with client-side filtering for better results.

    .PARAMETER Trending
    A switch to retrieve trending news topics instead of regular news articles.
    Note: Trending news is only available for specific markets (Canada, France, Germany, United Kingdom,
    People's Republic of China, and United States). If you specify an unsupported market with this switch,
    the function will warn you and suggest using a supported market.
    
    .PARAMETER ApiKey
    The API key for authenticating with the Bing News Search API. 
    If not specified, the function will use the value of the $env:BingSearchApiKey variable.

    .PARAMETER Market
    The geographic region to which the result data is localized. 
    Different markets support different features:
    - For category-based news: Australia, Canada (fr-CA), Chile, Denmark, Finland, France, Germany, 
      Italy, Mexico, People's Republic of China, Brazil, United Kingdom, United States, Worldwide
    - For trending news: Canada (en-CA, fr-CA), France, Germany, People's Republic of China, United Kingdom,
      United States (en-US)

    .PARAMETER Language
    The language in which to return search results. By default, the function will use the language
    associated with the selected market. This parameter can be used to override that behavior and
    request results in a specific language.

    .EXAMPLE
    Receive-BingNews -Category "ScienceAndTechnology" -ApiKey "YourApiKey" -Market "United States"

    This example retrieves technology news articles using the specified API key from the US market.

    .EXAMPLE
    Receive-BingNews -Category "US_Midwest" -Market "United States"

    This example retrieves news specific to the Midwest region of the US using client-side filtering.

    .EXAMPLE
    Receive-BingNews -Trending -ApiKey "YourApiKey" -Market "United Kingdom"

    This example retrieves trending news topics for the UK market using the specified API key.

    .EXAMPLE
    Receive-BingNews -Market "France" -Language "English"

    This example retrieves news articles from the French market but in English language.

    .NOTES
    This function requires an active internet connection and a valid Bing Search API key to function.
    Different markets support different categories. See the Microsoft documentation for details.

    .LINK
    https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/market-codes#news-category-api-markets
    .LINK
    https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/market-codes#trending-news-api-markets
    #>
    [CmdletBinding()]
    param (
        [Parameter()]
        [ValidateSet(
            # Markets that support News Category API
            "Australia", "Canada", "Chile", "Denmark", "Finland", "France", "Germany", 
            "Italy", "Mexico", "People's Republic of China", "Brazil", 
            "United Kingdom", "United States", "Worldwide"
        )]
        [string]
        $Market = "United States",
        
        [Parameter()]
        [ValidateSet(
            # US categories (offering the most options for tab completion)
            "Business", "Entertainment", "Health", "Politics", "Products", "ScienceAndTechnology", 
            "Sports", "US", "World",
            # Entertainment subcategories
            "Entertainment_MovieAndTV", "Entertainment_Music",
            # ScienceAndTechnology subcategories  
            "Technology", "Science",
            # Sports subcategories for the US
            "Sports_Golf", "Sports_MLB", "Sports_NBA", "Sports_NFL", "Sports_NHL", 
            "Sports_Soccer", "Sports_Tennis", "Sports_CFB", "Sports_CBB",
            # US regional subcategories
            "US_Northeast", "US_South", "US_Midwest", "US_West",
            # World subcategories
            "World_Africa", "World_Americas", "World_Asia", "World_Europe", "World_MiddleEast"
        )]
        [string]
        $Category,
        
        [Parameter()]
        [switch]
        $Trending,

        [Parameter()]
        [string]
        $ApiKey = $env:BingSearchApiKey,

        [Parameter()]
        [ValidateSet(
            "Arabic", "Basque", "Bengali", "Bulgarian", "Catalan", 
            "Chinese (Simplified)", "Chinese (Traditional)", "Croatian", "Czech", "Danish",
            "Dutch", "English", "English-United Kingdom", "Estonian", "Finnish",
            "French", "Galician", "German", "Gujarati", "Hebrew",
            "Hindi", "Hungarian", "Icelandic", "Italian", "Japanese",
            "Kannada", "Korean", "Latvian", "Lithuanian", "Malay",
            "Malayalam", "Marathi", "Norwegian (Bokmål)", "Polish", 
            "Portuguese (Brazil)", "Portuguese (Portugal)", "Punjabi", "Romanian", "Russian",
            "Serbian (Cyrylic)", "Slovak", "Slovenian", "Spanish", "Swedish",
            "Tamil", "Telugu", "Thai", "Turkish", "Ukrainian", "Vietnamese"
        )]
        [string]
        $Language
    )

    begin {
        # Convert the selected country/region name to its corresponding market code
        $marketCode = Get-MarketCode -Market $Market
        
        # Convert language name to language code if specified
        $languageCode = if ($Language) {
            Get-LanguageCode -Language $Language
        } else {
            # Default to no specific language code
            $null
        }
        
        # Define markets that support trending news
        # According to https://learn.microsoft.com/en-us/bing/search-apis/bing-news-search/reference/market-codes#trending-news-api-markets
        $trendingSupportedMarkets = @(
            "Canada",
            "France",
            "Germany", 
            "People's Republic of China",
            "United Kingdom", 
            "United States"
        )
        
        # Check if Trending is requested but not supported in the selected market
        if ($Trending -and $trendingSupportedMarkets -notcontains $Market) {
            Write-Warning "Trending news is not supported for the market '$Market'."
            Write-Warning "Supported markets for trending news are: $($trendingSupportedMarkets -join ', ')"
            
            $prompt = "Would you like to use 'United States' market for trending news instead? (Y/N)"
            $response = Read-Host -Prompt $prompt
            
            if ($response -eq "Y" -or $response -eq "y") {
                Write-Verbose "Using 'United States' market for trending news."
                $Market = "United States"
                $marketCode = Get-MarketCode -Market $Market
            }
            else {
                Write-Error "Cannot proceed with trending news for market '$Market'. Please select a supported market." -ErrorAction Stop
            }
        }
        
        # If a category is specified, validate it against the market's supported categories
        if ($Category) {
            $marketInfo = Get-MarketCategoryInfo -MarketCode $marketCode
            $isValidCategory = Test-CategoryForMarket -Category $Category -MarketCode $marketCode
            
            if (-not $isValidCategory) {
                # Get list of valid categories for this market for the error message
                $validCategories = $marketInfo.Categories -join ", "
                
                Write-Warning "Category '$Category' is not supported for the market '$Market' ($marketCode)."
                Write-Warning "Valid categories for this market are: $validCategories"
                
                # If subcategories exist, provide that information too
                if ($marketInfo.ContainsKey("Subcategories")) {
                    foreach ($parentCategory in $marketInfo.Subcategories.Keys) {
                        $subcats = $marketInfo.Subcategories[$parentCategory] -join ", "
                        Write-Warning "Subcategories for '$parentCategory': $subcats"
                    }
                }
                
                # Offer to use a default category instead
                $defaultCategory = $marketInfo.Categories[0]
                $prompt = "Would you like to use '$defaultCategory' instead? (Y/N)"
                $response = Read-Host -Prompt $prompt
                
                if ($response -eq "Y" -or $response -eq "y") {
                    Write-Verbose "Using default category '$defaultCategory' instead."
                    $Category = $defaultCategory
                } else {
                    Write-Error "Please specify a valid category for the selected market." -ErrorAction Stop
                }
            }
        }
    }

    process {
        if ($Trending) {
            # Validate API Key. Exit program if found not to be valid.
            if (-not $ApiKey) {
                Write-Error "You need to provide a valid Bing Search API key." -ErrorAction Stop
            }

            # Create the headers hash using the API key
            $headers = @{
                "Ocp-Apim-Subscription-Key" = $ApiKey
            }

            # Set the endpoint URL
            $url = "https://api.bing.microsoft.com/v7.0/news/trendingtopics?mkt=$marketCode"
            
            # Add language parameter if specified
            if ($languageCode) {
                $url += "&setLang=$languageCode"
                Write-Verbose "Using language: $Language ($languageCode) for trending news"
            }
            
            Write-Verbose "API URL: $url"

            # Make the API call
            try {
                $response = Invoke-RestMethod -Uri $url -Headers $headers -Method 'GET'
                
                # Return the trending topics
                return $response.value
            } catch {
                Write-Error "Failed to retrieve trending news: $_"
                Write-Warning "API URL used: $url"
                
                if ($_.ErrorDetails.Message) {
                    try {
                        $errorInfo = $_.ErrorDetails.Message | ConvertFrom-Json
                        Write-Warning "API Error: $($errorInfo | ConvertTo-Json -Depth 3)"
                    } catch {
                        Write-Warning "Error details: $($_.ErrorDetails.Message)"
                    }
                }
            }
        }
        else {
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

            # Whether we need client-side filtering
            $needsFiltering = $false
            $filterType = ""
            $filterValue = ""

            # Add category to the URL if specified
            if ($Category) {
                # For subcategories, there's a special handling in the API
                if ($Category -match "_") {
                    $mainCategory = ($Category -split "_")[0]
                    $subCategory = ($Category -split "_")[1]
                    Write-Verbose "Using main category: $mainCategory for subcategory: $Category"
                    
                    # Use the main category for the API request
                    $url += "?category=$mainCategory"
                    
                    # Flag that we need to filter results later
                    $needsFiltering = $true
                    $filterType = $mainCategory
                    $filterValue = $subCategory
                } else {
                    # Standard category
                    $url += "?category=$Category"
                }
            }

            # Construct market parameter to the URL
            $urlQueryChar = if ($url.Length -gt $baseUrl.Length) { "&" } else { "?" }
            $url += "$urlQueryChar" + "mkt=$marketCode"
            
            # Add language parameter if specified
            if ($languageCode) {
                $url += "&setLang=$languageCode"
                Write-Verbose "Using language: $Language ($languageCode) for news results"
            }

            Write-Verbose "API URL: $url"

            # Make the API call
            try {
                $apiResponse = Invoke-RestMethod -Uri $url -Headers $headers -Method 'GET'
                
                # Check if we have any results
                if ($null -eq $apiResponse.value -or $apiResponse.value.Count -eq 0) {
                    Write-Warning "No results found for category '$Category' in market '$Market' ($marketCode)."
                    if ($languageCode) {
                        Write-Warning "You specified language '$Language'. Try a different language or remove the language constraint."
                    }
                } else {
                    # Store unfiltered results first
                    $unfilteredResults = $apiResponse.value
                    $results = $unfilteredResults
                    
                    # Apply client-side filtering if needed (for subcategories)
                    if ($needsFiltering) {
                        Write-Verbose "Applying client-side filtering for subcategory: $filterValue"
                        
                        # Store the count before filtering
                        $originalCount = $results.Count
                        
                        # Apply client-side filtering based on the subcategory
                        switch ($filterType) {
                            "US" {
                                switch ($filterValue) {
                                    "Northeast" { 
                                        $filter = "New York|Boston|Philadelphia|Maine|Vermont|Connecticut|Massachusetts|Rhode Island|New Hampshire|NY|NYC|New England" 
                                    }
                                    "South" { 
                                        $filter = "Atlanta|Miami|Texas|Florida|Louisiana|Georgia|Alabama|Mississippi|Tennessee|Kentucky|Carolina|SC|NC|GA|FL|TX|VA|TN|AR|OK" 
                                    }
                                    "Midwest" { 
                                        $filter = "Chicago|Detroit|Ohio|Indiana|Illinois|Michigan|Wisconsin|Minnesota|Iowa|Missouri|Minneapolis|Cleveland|Cincinnati|St Louis|Kansas City|OH|MI|IL|IN|WI|MN|IA|MO|KS|NE|ND|SD" 
                                    }
                                    "West" { 
                                        $filter = "California|Los Angeles|San Francisco|Seattle|Oregon|Washington|Colorado|Nevada|Arizona|Utah|Portland|Denver|Phoenix|Las Vegas|San Diego|LA|CA|OR|WA|CO|NV|AZ|UT|ID|MT|WY|NM|HI|AK" 
                                    }
                                    default { $filter = $filterValue }
                                }
                                
                                Write-Verbose "Using regional filter: $filter"
                                $originalCount = $results.Count
                                
                                $results = $results | Where-Object { 
                                    $_.name -match $filter -or 
                                    $_.description -match $filter -or
                                    $_.provider.name -match $filter
                                }
                                
                                Write-Verbose "Filtered from $originalCount results to $($results.Count) results"
                            }
                            "Sports" {
                                Write-Verbose "Filtering for sports subcategory: $filterValue"
                                
                                # Sports subcategory filtering
                                $sportMap = @{
                                    "Golf" = "golf|PGA|LPGA|Masters|US Open|British Open|The Open Championship"
                                    "MLB" = "MLB|baseball|pitcher|batter|inning|home run|Major League Baseball"
                                    "NBA" = "NBA|basketball|court|dunk|three-pointer|National Basketball Association"
                                    "NFL" = "NFL|football|touchdown|quarterback|field goal|National Football League"
                                    "NHL" = "NHL|hockey|ice hockey|puck|goal|National Hockey League"
                                    "Soccer" = "soccer|football|FIFA|Premier League|La Liga|Bundesliga|Serie A|goal|match|championship"
                                    "Tennis" = "tennis|racket|court|serve|Grand Slam|ATP|WTA|US Open|Wimbledon|French Open|Australian Open"
                                    "CFB" = "college football|NCAA football|CFB|university football"
                                    "CBB" = "college basketball|NCAA basketball|university basketball|March Madness"
                                }
                                
                                if ($sportMap.ContainsKey($filterValue)) {
                                    $filter = $sportMap[$filterValue]
                                } else {
                                    $filter = $filterValue
                                }
                                
                                $results = $results | Where-Object { 
                                    $_.name -match $filter -or 
                                    $_.description -match $filter
                                }
                            }
                            "World" {
                                Write-Verbose "Filtering for world region: $filterValue"
                                
                                # World region filtering
                                $regionMap = @{
                                    "Africa" = "Africa|African|Nigeria|Kenya|South Africa|Egypt|Ethiopia|Ghana|Morocco|Algeria|Tunisia|Libya|Sudan"
                                    "Americas" = "Latin America|South America|Central America|Caribbean|Brazil|Mexico|Argentina|Colombia|Peru|Venezuela|Chile|Cuba|Haiti"
                                    "Asia" = "Asia|Asian|China|Japan|India|South Korea|Indonesia|Pakistan|Bangladesh|Vietnam|Thailand|Malaysia|Singapore|Philippines"
                                    "Europe" = "Europe|European|UK|Germany|France|Italy|Spain|Russia|Poland|Ukraine|Netherlands|Belgium|Sweden|Norway|Finland|Denmark"
                                    "MiddleEast" = "Middle East|Iran|Iraq|Saudi Arabia|Turkey|Israel|Syria|UAE|Qatar|Kuwait|Lebanon|Jordan|Yemen|Oman|Bahrain"
                                }
                                
                                if ($regionMap.ContainsKey($filterValue)) {
                                    $filter = $regionMap[$filterValue]
                                } else {
                                    $filter = $filterValue
                                }
                                
                                $results = $results | Where-Object { 
                                    $_.name -match $filter -or 
                                    $_.description -match $filter
                                }
                            }
                            "Entertainment" {
                                Write-Verbose "Filtering for entertainment subcategory: $filterValue"
                                
                                $entertainmentMap = @{
                                    "MovieAndTV" = "movie|film|cinema|TV|television|series|show|actor|actress|director|Hollywood|Netflix|HBO|Disney|streaming"
                                    "Music" = "music|song|album|artist|band|concert|tour|singer|rapper|musician|Grammy|Billboard|Spotify|iTunes"
                                }
                                
                                if ($entertainmentMap.ContainsKey($filterValue)) {
                                    $filter = $entertainmentMap[$filterValue]
                                } else {
                                    $filter = $filterValue
                                }
                                
                                $results = $results | Where-Object { 
                                    $_.name -match $filter -or 
                                    $_.description -match $filter
                                }
                            }
                            "ScienceAndTechnology" {
                                Write-Verbose "Filtering for science/tech subcategory: $filterValue"
                                
                                $sciTechMap = @{
                                    "Technology" = "technology|tech|software|hardware|app|smartphone|computer|AI|artificial intelligence|robot|automation|internet|digital|cyber|mobile"
                                    "Science" = "science|scientific|research|study|discovery|physics|chemistry|biology|astronomy|space|NASA|experiment|theory|quantum|molecule|climate"
                                }
                                
                                if ($sciTechMap.ContainsKey($filterValue)) {
                                    $filter = $sciTechMap[$filterValue]
                                } else {
                                    $filter = $filterValue
                                }
                                
                                $results = $results | Where-Object { 
                                    $_.name -match $filter -or 
                                    $_.description -match $filter
                                }
                            }
                            default {
                                Write-Warning "Unknown filter type: $filterType, using simple text matching"
                                $results = $results | Where-Object { 
                                    $_.name -match $filterValue -or 
                                    $_.description -match $filterValue
                                }
                            }
                        }
                        
                        if ($results.Count -eq 0) {
                            Write-Warning "No results found after filtering for subcategory '$filterValue'."
                            Write-Warning "Try using the main category '$filterType' instead or run with -Verbose for more information."
                            
                            $prompt = "Would you like to see the unfiltered results for '$filterType' instead? (Y/N)"
                            $userInput = Read-Host -Prompt $prompt
                            
                            if ($userInput -eq "Y" -or $userInput -eq "y") {
                                $results = $unfilteredResults
                                Write-Verbose "Returning all results for main category '$filterType'"
                            } else {
                                $results = $null
                            }
                        }
                    }
                    
                    $results
                }
            } catch {
                Write-Error "Failed to retrieve news: $_"
                Write-Warning "API URL used: $url"
                
                if ($_.ErrorDetails.Message) {
                    try {
                        $errorInfo = $_.ErrorDetails.Message | ConvertFrom-Json
                        Write-Warning "API Error: $($errorInfo | ConvertTo-Json -Depth 3)"
                    } catch {
                        Write-Warning "Error details: $($_.ErrorDetails.Message)"
                    }
                }
            }
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
            "Web",
            "News"
        )]
        [string]
        $Source
    )
        
    begin {
        
        # Define a hashtable mapping the service types to their URL properties
        $urlPropertyMap = @{
            "Web"          = 'url';
            "Images"       = 'contentUrl';
            "Videos"       = 'contentUrl';
            "News"         = 'url';
            "Entities"     = 'webSearchUrl';
            "Suggestions"  = 'url';
            "Spelling"     = 'url';
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

                Web   { Start-Process $SearchResult.webSearchUrl  }
                News  { Start-Process $SearchResult.newsSearchUrl }

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

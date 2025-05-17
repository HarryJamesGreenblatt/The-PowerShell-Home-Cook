# BingBuddy.ArgumentCompleters.ps1
# Contains argument completers for BingBuddy module parameters

# Define a shared script block for Market parameter tab completion
$MarketCompleter = {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Map of market names to codes for tooltip display
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

    # Use the keys from the market map for completion - only show friendly names
    # This matches the expected behavior where the user enters a friendly name
    # which will be converted to a market code by Get-MarketCode when needed
    $marketMap.Keys | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object {
        # Return completion results with tooltips that show the market code
        $marketCode = $marketMap[$_]
        # Use the friendly name as the completion text, with the market code in the tooltip
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', "$_ ($marketCode)")
    }
}

# Register argument completer for all commands that use the Market parameter
@(
    'Get-BingSearchResults',
    'Invoke-BingSearch',
    'Receive-BingNews',
    'Receive-BingNewsTrendingTopics'
) | ForEach-Object {
    Register-ArgumentCompleter -CommandName $_ -ParameterName Market -ScriptBlock $MarketCompleter
}

# Register argument completer for the Category parameter in Receive-BingNews
Register-ArgumentCompleter -CommandName Receive-BingNews -ParameterName Category -ScriptBlock {
    param($commandName, $parameterName, $wordToComplete, $commandAst, $fakeBoundParameters)

    # Determine the market (default to United States if not specified)
    $market = $fakeBoundParameters['Market']
    if (-not $market) { $market = 'United States' }

    # Get the market code and valid categories
    $marketCode = Get-MarketCode -Market $market
    $marketInfo = Get-MarketCategoryInfo -MarketCode $marketCode
    
    # Get both main categories and subcategories
    $allCategories = $marketInfo.Categories.Clone()
    
    # Add subcategories if available
    if ($marketInfo.ContainsKey("Subcategories")) {
        foreach ($parentCategory in $marketInfo.Subcategories.Keys) {
            $marketInfo.Subcategories[$parentCategory] | ForEach-Object {
                $allCategories += $_
            }
        }
    }
    
    # Return matching categories for tab completion
    $allCategories | Where-Object { $_ -like "$wordToComplete*" } | ForEach-Object {
        [System.Management.Automation.CompletionResult]::new($_, $_, 'ParameterValue', $_)
    }
}

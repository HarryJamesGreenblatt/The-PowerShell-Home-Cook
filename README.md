# The-PowerShell-Home-Cook
You know, if PowerShell was a musical instrument, I think it would be a Conductor's Baton :heavy_minus_sign::ok_hand:

## Modules
### [`Search-Internet`](./Modules/Search-Internet/Search-Internet.psm1)
Searches the web for the given Query String on an optionally specified Search Engine and Browser.

<br>

### [`Write-Figlet`](./Modules/Write-Figlet/Write-Figlet.psm1)
Converts a provided string input to generative ASCII Typesetting produced by the python pyfiglet module.

<br>

### [`BingBuddy`](./Modules/BingBuddy/BingBuddy.psm1)
Provides functions to interact with the Bing Search API, including web, image, video, and news searches.

- #### [`Get-BingSearchResults`](./Modules/BingBuddy/BingBuddy.psm1)
    Performs a search using the Bing Search API and returns unique results.

- #### [`Invoke-BingSearch`](./Modules/BingBuddy/BingBuddy.psm1)
    Invokes a Bing Search and returns results based on the specified query and service type.

- #### [`Open-BingSearchResult`](./Modules/BingBuddy/BingBuddy.psm1)
    Opens the URL from a Bing search result in the default web browser.

- #### [`Receive-BingNews`](./Modules/BingBuddy/BingBuddy.psm1)
    Retrieves news articles using the Bing News Search API, optionally filtered by a specific category.

- #### [`Receive-BingNewsTrendingTopics`](./Modules/BingBuddy/BingBuddy.psm1)
    Retrieves trending news topics using the Bing Search API.

<br>

### [`Convert-JsxForCopilotM365`](./Modules/Convert-JsxForCopilotM365/Convert-JsxForCopilotM365.psm1)
Converts JSX tags to HTML entities using a Python script.

<br>

### [`Invoke-CodeChunker`](./Modules/Invoke-CodeChunker/Invoke-CodeChunker.psm1)
Invokes a Python script to chunk code into 2000 character segments for sharing in the Copilot chat window.

<br>

### [`Find-HexColorCodes`](./Modules/Find-HexColorCodes/Find-HexColorCodes.psm1)
Searches for hex color codes in all files within a specified directory and optionally displays a webpage depicting it as a color table.

<br>

### [`Import-Outlook`](./Modules/Import-Outlook/Import-Outlook.psm1)
A collection of modules which facilitate the management of Outlook messages using Outlook's Message Application Programming Interface (MAPI).

- #### [`Receive-OutlookMailbox`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Connects to the user's Outlook account and retrieves the specified mail folder as a powershell Com Object.

- #### [`Limit-OutlookMailbox`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Limits the quantity of ComObjects collected from the user's Outlook Mail Folder using a DASL Filter Query string.

- #### [`Send-OutlookMail`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Sends an Outlook message, including any attachments, by way of either a plaintext or HTML Body format, to the given recipient address.

- #### [`Import-WPFDataFromXAML`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Loads a WPF Form using the XAML file at specified file path and creates script-scoped variables which reference all the form's controls by name.

- #### [`Exit-OutlookMailbox`](./Modules/Import-Outlook/Import-Outlook.psm1)
    Stops all Outlook running Outlook processes as to not leave persistent connections to the Exchange Server.

<br>

## Scripts
### [`Test-HelloWorldForm.ps1`](./Scripts/Test-HelloWorldForm.ps1)
Uses Windows Forms to generate a simple button GUI using 100% vanilla PowerShell.

<br>

### [`Test-HelloWorldXMLForm.ps1`](./Scripts/Test-HelloWorldXMLForm.ps1)
Loads an externally developed WPF Form saved in XAML format into a simple button GUI whose control logic is handled by PowerShell.

<br>

### [`Test-ServiceHealthForm.ps1`](./Scripts/Test-ServiceHealthForm.ps1)
Uses Windows Forms to generate a PowerShell GUI that implements a Combo Box to list the Health of Services currently running on the user's machine.
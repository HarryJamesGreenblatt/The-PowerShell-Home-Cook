function Get-HealthAndStatus {
<#
.SYNOPSIS
        Initiates a remote session between endpoints and packages a collection of health and status metrics 
        and log data belonging to both into a Custom Object.

.DESCRIPTION
        The Client will have it's System Information, derived from the  Get-ComputerInfo Cmdlet, packaged into a
        Hash Table called ClientHealthAndStatus, along with a List summarizing all of its Active Event Logs. 
        The ClientHealthAndStatus will then be passed as an input to a PowerShell Remoting session concuted between
        the Client and the Server over SSH. 
        
        During the remoting session, the Server will go through a similar process,
        producing it's own Hash table called ClientHealthAndStatus, which will be exported back to the Client following 
        the termination of the remoting session.
        
        Finally, a Custom PSObject is created to store both Hash Tables, effectively enabling to 
        capability of accessing each endpoint's data by it's role and the type of report,  
                
                i.e.  $HealthAndStatus.Server.SystemInfo   or    $HealthAndStatus.Client.LogList 

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.NOTES
        May be used concurrently with Write-ToTransferDirectory.
        (See Examples)

.EXAMPLE
        Get-HealthAndStatus `
                UserName = My_Name `
                HostName = Server_Name `
                PathToTransferDirectory = Path\to\directory

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params -Verbose

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params | Write-ToTransferDirectory @params
#>
        [CmdletBinding()]
        
        param (

                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory

        )
        
        begin {

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."
                
                $Session = New-PSSession -UserName $UserName -HostName $HostName


                Write-Verbose "Packaging Client's System Information and Log Data into a Health and Status Hash Table."
                
                $ClientHealthAndStatus = @{

                        SystemInfo = Get-ComputerInfo;

                        LogList = Get-WinEvent -ListLog * | 
                                                Where-Object RecordCount -gt 0 | 
                                                                Sort-Object RecordCount -Descending  
                }
                
                Write-Verbose "The Client's Health and Status is now stored in a:`n  $ClientHealthAndStatus."
                        
        }
        

        process {
                
                Write-Verbose "Invoking a Command to the Server to package it's System Information and Log Data into a seperate Hash Table."
                
                Invoke-Command `
                        -Session $Session `
                        -OutVariable ServerHealthAndStatus `
                        -ScriptBlock {
                        
                                @{
                                        SystemInfo = Get-ComputerInfo;
                                        
                                        LogList = Get-WinEvent -ListLog * | 
                                                        Where-Object RecordCount -gt 0 | 
                                                                        Sort-Object recordcount -Descending
                                }

                        }

                Exit-PSSession

        }
        
        end {

                Write-Verbose "Returning from the remoting session, the Server's Health and Status is now held in a:`n  $ServerHealthAndStatus."
                Write-Verbose "Packaging both the Client and Server Health and Status Hashes into a Custom PSObject."
                
                [PSCustomObject]@{
                        Client = $ClientHealthAndStatus;
                        Server = $ServerHealthAndStatus
                }
                
                Write-Verbose "The Health and Status Custom PSObject is now available."
        }
}
Export-ModuleMember -Function Get-HealthAndStatus                     





function Write-ToTransferDirectory {
<#
.SYNOPSIS
        Initiates a remote session between endpoints, unpackages a Custom PSObject into a specified directory 
        designated as a staging area to facilitate data transfer, and modifies the unpacked data to remove undesirable 
        non-printing characters.

.DESCRIPTION
        The Client will initiate a remoting session with the Server and check for the existence of a specified Transfer Directory.
        If it doesn't exist, one will be created.

        Following that, the contents of the $HealthAndStatus input, whether passed in as Pipeline Input, or as 
        a Parameter Value, are unpacakged and written as files respective to the System Information and Log Lists included  
        with the Client and Server Health and Status Hash Tables that are returned from Get-HealthAndStatus.
        
        However, given that PowerShell 7 unfortantely includes ANSI Color Codes in its File Representation of it's Objects,
        Additional measures are taken to Remove these ANSI Color Codes using a REGEX Subsitution via  -replace  '',''.

.PARAMETER HealthAndStatus,
        A Custom PSObject containing Hash Tables corresponding to the Client's and Server's Health and Status Data, respectively.

.PARAMETER UserName
        The Client's User Account Name belonging to the User who may authenticate to the Server.

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.PARAMETER ANSIColorCodes
        The Regular Expression matching the ANSI Color Code jusnk characters included by default in PowerShell 7 output. 

.NOTES
        Get-HealthAndStatus is used concurrently with this function.
        (See Examples)

.EXAMPLE
        Write-ToTransferDirectory `
                HealthAndStatus = Get-HealthAndStatus @params `
                UserName = My_Name `
                HostName = Server_Name `
                PathToTransferDirectory = Path\to\directory

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Write-ToTransferDirectory @params -Verbose

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params | Write-ToTransferDirectory @params
#>
        [CmdletBinding()]
        
        param (
                [Parameter(ValueFromPipeline=$true)]
                [System.Object[]]
                $HealthAndStatus,
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [string]
                $ANSIColorCodes = '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]'           
        
        )
        
        begin {

                Write-Verbose "The path to the Server's Transfer Directory is:  $PathToTransferDirectory."
                Write-Verbose "Initiating Session over SSH to  $UserName@$HostName."
                
                $Session = New-PSSession -UserName $UserName -HostName $HostName


                Write-Verbose "Setting up an array of File Names to facilitate File Write operations."
                
                $Deliverables = @(
                        "Client Info",
                        "Client Logs",
                        "Server Info",
                        "Server Logs"
                )   
        }
        
        process {

                Write-Verbose ("Invoking a Command to the Server to unpackage it's Health and Status Object " +
                "into Files stored in the specified Transfer Directory.")
                
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $HealthAndStatus,$Deliverables,$PathToTransferDirectory,$ANSIColorCodes `
                        -ScriptBlock {

                                Param(
                                        [PSCustomObject]$HealthAndStatus, 
                                        [System.Object[]]$Deliverables,
                                        $PathToTransferDirectory,
                                        $ANSIColorCodes
                                )

                                
                                Write-Verbose "Checking If a Transfer Directory exists."
                                
                                If( -not (Test-Path $PathToTransferDirectory) ){
                                        New-Item -ItemType Directory $PathToTransferDirectory
                                }


                                Write-Verbose "Unpackaging the Health and Status Object into individual Transfer Files."
                                
                                $HealthAndStatus.Client.SystemInfo | Out-File "$PathToTransferDirectory\Client Info" -Encoding ASCII
                                $HealthAndStatus.Client.LogList    | Out-File "$PathToTransferDirectory\Client Logs" -Encoding ASCII
                                $HealthAndStatus.Server.SystemInfo | Out-File "$PathToTransferDirectory\Server Info" -Encoding ASCII
                                $HealthAndStatus.Server.LogList    | Out-File "$PathToTransferDirectory\Server Logs" -Encoding ASCII


                                Write-Verbose "Modifying all Transfer Files to remove the uunwanted ANSI Color Codes."
                                
                                $Deliverables | ForEach-Object -Process {

                                        (Get-Content "$PathToTransferDirectory\$_")  `
                                                -replace $ANSIColorCodes, '' | 
                                                        Out-File "$PathToTransferDirectory\$_" -Encoding ASCII 

                                }

                        }

                Exit-PSSession

        }
        
        end {
                Write-Verbose "The Health and Status Custom PSObject Has now beem fully unpacked into the Transfer Directory."    
        }

}
Export-ModuleMember -Function Write-ToTransferDirectory




function Compare-LastWriteTimes {
<#
.SYNOPSIS
        A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
        A longer description of the function, its purpose, common use cases, etc.
.NOTES
        Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
        Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
        Test-MyTestFunction -Verbose
        Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>

        [CmdletBinding()]
        
        param (
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [Parameter(Mandatory=$true)]
                [string]
                $PathToFileDemonstratingSystemActivity
        
        )
        
        begin {
                $Session = New-PSSession -UserName $UserName -HostName $HostName                
        }
        
        process {
                
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $PathToTransferDirectory,$PathToFileDemonstratingSystemActivity `
                        -OutVariable MostRecentlyModified `
                        -ScriptBlock {

                                Param($PathToTransferDirectory,$PathToFileDemonstratingSystemActivity)
                                
                                $LastTransferDirectoryWriteTime = 
                                        (Get-ChildItem "$PathToTransferDirectory")[0].LastWriteTime
                                

                                $LastActivityWriteTime = 
                                        (Get-ChildItem "$PathToFileDemonstratingSystemActivity").LastWriteTime
                        

                                (
                                        ($LastTransferDirectoryWriteTime -lt $LastActivityWriteTime) `
                                        ? `
                                        "Activity Probe" `
                                        : `
                                        "Transfer Directory" `
                                )

                        }

                Exit-PSSession  
        }
                
        end {           
        }

}
Export-ModuleMember -Function Compare-LastWriteTimes




function Update-HealthAndStatus {
<#
.SYNOPSIS
        A short one-line action-based description, e.g. 'Tests if a function is valid'
.DESCRIPTION
        A longer description of the function, its purpose, common use cases, etc.
.NOTES
        Information or caveats about the function e.g. 'This function is not supported in Linux'
.LINK
        Specify a URI to a help page, this will show when Get-Help -Online is used.
.EXAMPLE
        Test-MyTestFunction -Verbose
        Explanation of the function or its result. You can include multiple examples with additional .EXAMPLE lines
#>

        [CmdletBinding()]

        param (
                
                [Parameter(Mandatory=$true)]
                [string]
                $MostRecentlyModified,
        
                [Parameter(Mandatory=$true)]
                [string]
                $UserName,

                [Parameter(Mandatory=$true)]
                [string]
                $HostName,

                [Parameter(Mandatory=$true)]
                [string]
                $PathToTransferDirectory,
                
                [Parameter(Mandatory=$true)]
                [string]
                $PathToFileDemonstratingSystemActivity
        
        )
        
        begin {
                $Session = New-PSSession -UserName $UserName -HostName $HostName 

                $ActivityHealthAndStatus = 
                
                        ($MostRecentlyModified -match 'Transfer Directory') `
                        ? `
                        "As of $(Get-Date), there has been NO additional activity observed." `
                        : `
                        "Recent System Activity observed on $(Get-Date)"

        }
        
        process {
                
                Invoke-Command `
                        -Session $Session `
                        -ArgumentList $PathToTransferDirectory, $ActivityHealthAndStatus `
                        -ScriptBlock {
                                Param($PathToTransferDirectory, $ActivityHealthAndStatus)
                                $PathToActivityFile = "$PathToTransferDirectory\System Activity"
                                
                                $ActivityHealthAndStatus | Out-File $PathToActivityFile -Append
                        }

                Exit-PSSession
        }
        
        end {
        }
}
Export-ModuleMember -Function Update-HealthAndStatus
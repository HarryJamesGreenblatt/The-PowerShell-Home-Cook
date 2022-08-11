function Get-HealthAndStatus {
<#
.SYNOPSIS
        Initiates a remote session between endpoints and packages a collection of health and status metrics 
        and log data belonging to both into a Custom Object.

.DESCRIPTION
        A Client will have it's System Information, derived from the  Get-ComputerInfo Cmdlet, packaged into a
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
                i.e.  $HealthAndStatus.Server.SystemInfo   or    $HealthAndStatus.Client.LogList 

.PARAMETER HostName
        The "Computer Name" assigned to the Server.

.PARAMETER PathToTransferDirectory
        The Path to a File Transfer Directory located on the Server. 

.NOTES
        May be used concurrently with Write-ToTransferDirectory.
        (See Examples)

.EXAMPLE
        $params = @{
                UserName = My_Name;
                HostName = Server_Name;
                PathToTransferDirectory= Path\to\directory
        }

        Get-HealthAndStatus @params

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

                $Session = New-PSSession -UserName $UserName -HostName $HostName

                $ClientHealthAndStatus = @{

                        SystemInfo = Get-ComputerInfo;

                        LogList = Get-WinEvent -ListLog * | 
                                                Where-Object RecordCount -gt 0 | 
                                                         Sort-Object RecordCount -Descending  
                }
                    
        }
        

        process {
                
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

                [PSCustomObject]@{
                        Client = $ClientHealthAndStatus;
                        Server = $ServerHealthAndStatus
                }
                
        }
}
Export-ModuleMember -Function Get-HealthAndStatus                     





function Write-ToTransferDirectory {
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

                $Session = New-PSSession -UserName $UserName -HostName $HostName

                $Deliverables = @(
                        "Client Info",
                        "Client Logs",
                        "Server Info",
                        "Server Logs"
                )   
        }
        
        process {

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


                                If( -not (Test-Path $PathToTransferDirectory) ){
                                        New-Item -ItemType Directory $PathToTransferDirectory
                                }


                                $HealthAndStatus.Client.SystemInfo | Out-File "$PathToTransferDirectory\Client Info" -Encoding ASCII
                                $HealthAndStatus.Client.LogList    | Out-File "$PathToTransferDirectory\Client Logs" -Encoding ASCII
                                $HealthAndStatus.Server.SystemInfo | Out-File "$PathToTransferDirectory\Server Info" -Encoding ASCII
                                $HealthAndStatus.Server.LogList    | Out-File "$PathToTransferDirectory\Server Logs" -Encoding ASCII


                                $Deliverables | ForEach-Object -Process {

                                        (Get-Content "$PathToTransferDirectory\$_")  `
                                                -replace $ANSIColorCodes, '' | 
                                                        Out-File "$PathToTransferDirectory\$_" -Encoding ASCII 

                                }

                        }

                        
                Exit-PSSession

        }
        
        end {
                
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
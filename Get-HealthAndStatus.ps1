$session = New-PSSession -UserName Harry -HostName Joanne1

$client_health_and_status = @{
    
    client_info = Get-ComputerInfo;
    client_logs = Get-WinEvent -ListLog * | 
                            Where-Object RecordCount -gt 0 | 
                                     Sort-Object recordcount -Descending

}


Invoke-Command -Session $session -ArgumentList $client_health_and_status -ScriptBlock {

    Param($client_health_and_status)


    $client_health_and_status.client_info | 
            Out-File 'C:\Users\harry\Desktop\Transfer Directory\Client Info' -Encoding ASCII

    (Get-Content 'C:\Users\harry\Desktop\Transfer Directory\Client Info') -replace '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]','' | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Client Info' -Encoding ASCII
    

    $client_health_and_status.client_logs | 
            Out-File 'C:\Users\harry\Desktop\Transfer Directory\Client Logs' -Encoding ASCII

    (Get-Content 'C:\Users\harry\Desktop\Transfer Directory\Client Logs') -replace '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]','' | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Client Logs' -Encoding ASCII
    
    

    $server_info = Get-ComputerInfo
    

    $server_logs = Get-WinEvent -ListLog * | 
    Where-Object RecordCount -gt 0 | 
    Sort-Object recordcount -Descending
    
    $server_info | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Server Info' -Encoding ASCII
    (Get-Content 'C:\Users\harry\Desktop\Transfer Directory\Server Info')  -replace '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]','' | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Server Info' -Encoding ASCII
    
    $server_logs | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Server Logs' -Encoding ASCII
    (Get-Content 'C:\Users\harry\Desktop\Transfer Directory\Server Logs') -replace '\x1B\[([0-9]{1,3}((;[0-9]{1,3})*)?)?[m|K]','' | Out-File 'C:\Users\harry\Desktop\Transfer Directory\Server Logs' -Encoding ASCII

}
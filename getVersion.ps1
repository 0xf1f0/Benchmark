function getPSversion
{
    psVersion = Get-Host | Select-Object Version
    return psVersion
}


function getWinversion
{
    try
    {
        $osVersion = Get-CimInstance Win32_OperatingSystem | Select-Object Caption
    }
    catch [System.Exception]
    {
       
    }
}

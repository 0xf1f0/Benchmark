#Get the benchmark files for the appropriate windows product version
cls
$getPV = (Get-CimInstance Win32_OperatingSystem).Caption
switch -Regex ($getPV.ToLower())
{
    "w(indows 8)" {"Yes"}
    default {"No"}
}
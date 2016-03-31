$filePath = "K:\MIS\Ola\cis_win7.csv"
<# Create an Object Excel.Application using Com interface
$objExcel = New-Object -ComObject Excel.Application

# Disable the 'visible' property so the document won't open in excel
$objExcel.Visible = $false
#>
#Import the data from csv file and store in a variable
$rawData = Import-Csv -path $filePath

#select all hive objects
$cisData = ($rawData | Where-Object{$_.hive, $_.key, $_.name , $_.type, $_.value})

#create objects of each column entry in $cisData
$hive = ($cisData | ForEach-Object{$_.hive})
$key = ($cisData | ForEach-Object{$_.key})
$name = ($cisData | ForEach-Object{$_.name})
$value = ($cisData | ForEach-Object{$_.value})

#Join the cell elements to make a full path to the registry
$fullPath = ($cisData | ForEach-Object{"Registry::" +  $_.hive + "\" + $_.key}) 
$rootPath = ($cisData | ForEach-Object{$_.hive + "\" + $_.key})


#The total  number of rows in the csv file
$count = $fullPath.count

cls #Clear the screen

Write-Host "Attempting to backup the registry" 
for($index = 0 ; $index -lt $count; $index++)
{
    #check if the registry path exists
    if((Test-path $fullPath[$index]))
    {
        #get non-null values from the .csv column entries
        if($key[$index] -ne "" -and $name[$index] -ne "" -and $value[$index] -ne "")
        {                  
            #check if a property[name] exist in the registry for a corresponding valid registry path
            if(Get-ItemProperty -Path $fullPath[$index] | Select -ExpandProperty $name[$index] -EA SilentlyContinue)
            {
                #Export each of the found registry values
                new-item C:\Users\OlaO\Desktop\regBackup -type directory -Force
                reg export $rootPath[$index] C:\Users\OlaO\Desktop\regBackup\$index.reg /y
                $inputFile = "C:\Users\OlaO\Desktop\regBackup\*"
                
                #Create a folder on the desktop + Merge individual backup file into a single file
                new-item C:\Users\OlaO\Desktop\mergedBackup -type directory -force
                $outputFile = "C:\Users\OlaO\Desktop\mergedBackup\mergedReg $(get-date -f MM-dd-yyyy-HH-mm).reg" 
                 
                #Display the output
                write-host "PSParentPath: "$fullPath[$index]
                write-host "RootPath: "$rootPath[$index]
                write-host "Expanded Property: "$name[$index]
                write-host "File Name: $index.reg"   
                        
                Try
                {
                    write-host "Current Value: " (Get-ItemProperty -Path $fullPath[$index] | Select -ExpandProperty $name[$index] -EA stop)  
                    #Set-ItemProperty -Path $fullPath[$index] -Name $name[$index] -Value $value[$index]
                    write-host "New Value: " (Get-ItemProperty -Path $fullPath[$index] | Select -ExpandProperty $name[$index] -EA stop)
                    Write-host ""                                  
                }
                
                Catch [System.Exception]
                {
                    $_.Exception >> "C:\Users\olao\Desktop\exceptionGet.txt"                                     
                } 
            }
        } 
    }
}
                
# Merge each backup reg files into a single reg file
# Set an header and remove multiple instances of the header in each file
'Windows Registry Editor Version 5.00' | Set-Content $outputFile 
get-content $inputFile -include "*.reg" | Where-Object{$_ -notmatch 'Windows Registry Editor Version 5.00'} | Add-Content $outputFile


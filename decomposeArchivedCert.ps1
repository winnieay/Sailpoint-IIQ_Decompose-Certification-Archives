$csvFilePath = "C:\Users\winnieauyeung\Desktop\certificationData.csv"
$outputFilePath = "C:\Users\winnieauyeung\Desktop\output.csv"


# Read the CSV file line by line
$csvContent = Get-Content -Path $csvFilePath

$csvHeader="AttributeName","Entitlement Description","Application","Decision","Access Review","Decision Maker","Display Name","Certification Name"
$propertiyName1="exceptionAttributeName","exceptionAttributeValue","exceptionApplication"
$propertiyName2="status","actorDisplayName","actorName"
$count=0
$countEntity=@()
$te=@()
$EntityTargetname=@()

foreach ($line in $csvContent) {
    if ($line -match "<Reference ") {
        if ($line -match 'name=""([^"]+)""') {
            $value = $matches[1]
            if(!$certificationname){$certificationname = $value}
        }
    }
    if ($line -match "<CertificationStatistics ") {
        if ($line -match 'totalItems=""([^"]+)""') {
             $Items= $matches[1]
        }
    }
    if ($line -match "<CertificationEntity") {
        if ($line -match 'targetDisplayName=""([^"]+)""') {
            $EntityTargetname += $matches[1]
        }
    }
    if ($line -match "</CertificationItem>") {
        $count+=1;
    }
    
    if ($line -match "</CertificationEntity>") {
        $countEntity+=$count
        $count=0;
    }
}

Function getValue{
  param ($data,$keyName,$propertiyNames)
  $result=@()
    foreach ($line in $data) {
            $values=$null
            if ($line -match $keyName) {
                    foreach($propertiyName in $propertiyNames){
                        if ($line -cmatch $propertiyName+'=""([^"]+)""') {
                            $values+=$matches[1]+"!"
                        }
                    }
                    $result+=$values
            }
        }
    return $result
} 


<# $certificationname
$Items
$countEntity
$EntityTargetname
write-host "------------------------" #>

$te+=getValue $csvContent "CertificationItem" $propertiyName1
$te+=getValue $csvContent "CertificationAction" $propertiyName2
$te=$te | Where-Object { $_ -match "\S" }
# $te.count


# Initialize result array
$result = @()
for ($i = 0; $i -lt $Items; $i++) {
    $concat = ""
    for ($j = $i; $j -lt $te.Count; $j += $Items) {
        $concat += $te[$j]
    }
    $result += $concat.Trim()
}


$result2 = @()
$currentIndex = 0
for ($i = 0; $i -lt $countEntity.Count; $i++) {
    $chunkSize = $countEntity[$i]
    $currentTarget = $EntityTargetname[$i]
    # Get elements for current chunk
    $chunk = $result[$currentIndex..($currentIndex + $chunkSize - 1)]
    # Concatenate and add to result
    foreach ($element in $chunk) {
        $result2 += "$element$currentTarget!$certificationname"
    }
    # Move index to next chunk start
    $currentIndex += $chunkSize
}

# Convert $result2 into custom objects
$objects = $result2 | ForEach-Object {
    $values = $_.Split("!")
    $obj = [ordered]@{}
    $values[4] = "Access Review for " + $values[4].Trim() 
    $values[1] = "Value $($values[1].Trim()) on $($values[0].Trim())"
    for ($i = 0; $i -lt $csvHeader.Count; $i++) {
        $obj[$csvHeader[$i]] = $values[$i]
    }
    [PSCustomObject]$obj
}

# Export objects to a CSV file
$objects | Export-Csv -Path $outputFilePath -NoTypeInformation -Encoding UTF8
#(Get-Content -Path $outputFilePath) | ForEach-Object { $_ -replace '"', '' } | Out-File -FilePath $outputFilePath


$newOrder = @("Display Name","Access Review","Certification Name","Entitlement Description","Decision","Decision Maker","Application")
$data = Import-Csv -Path $outputFilePath
$data | Select-Object -Property $newOrder | Export-Csv -Path $outputFilePath -NoTypeInformation

Clear-Variable -Name "certificationname"
Clear-Variable -Name "te"
Clear-Variable -Name "result"
Clear-Variable -Name "result2"
Clear-Variable -Name "objects"
Clear-Variable -Name "data"

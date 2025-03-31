$csvFilePath = "C:/Users/winnieay/Desktop/cert_latest.csv"
$outputFilePath = "C:/Users/winnieay/Desktop/cert/"

# Read the CSV file line by line
$csvContent = Get-Content -Path $csvFilePath -Raw

$blocks = [regex]::Matches($csvContent, '<Certification activated=.*?</Certification>', [System.Text.RegularExpressions.RegexOptions]::Singleline)

$csvHeader="Entitlement Type","Entitlement Description","Application","Decision","Access Review","Decision Maker","Account Name","Certification Name","Sign Off Date"
$propertiyName1="exceptionAttributeName","exceptionAttributeValue","exceptionApplication"
$propertiyName2="status","actorDisplayName","actorName"
$count=0
$countEntity=@()
$te=@()
$EntityTargetname=@()

function getValue{
  param ($data,$keyName,$propertiyNames,$c)
  $result=@()
    foreach ($line in $data) {
            $values=$null
            if ($line -match $keyName) {
                    foreach($propertiyName in $propertiyNames){
                        if ($line -cmatch $propertiyName+'=""([^"]+)""') {
                            if($matches[1] -and $matches[1] -ne ""){
                                $values+=$($matches[1].Trim())+"!!"
                            }
                        }
                    }
                    if($keyName -eq "CertificationItem"){
                        #write-host "or: $values"
                        $count = ($values | Select-String -Pattern "!!" -AllMatches).Matches.Count
                        if($count -lt 3 -and $count -gt 0){
                            $values="Account!!Account!!$values"
                            #write-host "new: $values"
                        } 
                    }
                    $result+=$values
            }
        }
    return $result
} 
function InitArray{
    param ($Items,$te)
    $result = @()
    for ($i = 0; $i -lt $Items; $i++) {
        $concat = $null
        for ($j = $i; $j -lt $te.Count; $j += $Items) {
            $concat += $te[$j].Trim()
        }
        $result += $concat.Trim()
    }
    return $result
}
function ConcatArray{
    param ($countEntity,$EntityTargetname,$result,$signeddate,$c)
    $epoch = New-Object DateTime 1970, 1, 1, 0, 0, 0, ([DateTimeKind]::Utc)
    $dateTime = $epoch.AddSeconds($signeddate / 1000)
    $signeddate = $dateTime.ToString("MM/dd/yy, h:mm tt")
    $result2 = @()
    $currentIndex = 0
    for ($i = 0; $i -lt $countEntity.Count; $i++) {
        $chunkSize = $countEntity[$i]
        $currentTarget = $EntityTargetname[$i]
        $chunk = $result[$currentIndex..($currentIndex + $chunkSize - 1)]
        foreach ($element in $chunk) {
            $result2 += "$element$currentTarget!!$certificationname!!$signeddate"
        }
        $currentIndex += $chunkSize
    }
    return $result2
}
function DataProcessing {
    param ($result,$csvHeader,$c)
    $objects = $result | ForEach-Object {
        $values = $_.Split("!!").Trim()
        $values = $values | Where-Object { $_.Trim() -ne "" }
        $obj = [ordered]@{}
        if($values[4]){$values[4] = "Access Review for " + $values[4]}

        if($values[1] -and $values[0]){
            if($values[1].Trim() -eq "Account"){
                $values[1] = "$($values[0]) $($values[6].Trim()) on $($values[2])"
            }else{
                $values[1] = "$($values[1].Trim()) on $($values[0])"
            }
        }

        if($values[0] -ne "Account"){$values[0] = "Additional Entitlement"}
        for ($i = 0; $i -lt $csvHeader.Count; $i++) {
            $obj[$csvHeader[$i]] = $values[$i]
        }
        [PSCustomObject]$obj
    }
    return $objects
    
}
function ExportCSV {
    param ($objects,$outputFilePath,$certificationname,$c)
    $certificationname = $certificationname -replace '[#-.\[\],:/]', '_'
    $objects | Export-Csv -Path "$outputFilePath$certificationname$c.csv" -NoTypeInformation -Encoding UTF8
    $newOrder = @("Account Name","Sign Off Date","Access Review","Certification Name","Entitlement Type","Entitlement Description","Decision","Decision Maker","Application")
    $data = Import-Csv -Path "$outputFilePath$certificationname$c.csv"
    $data | Select-Object -Property $newOrder | Export-Csv -Path "$outputFilePath$certificationname$c.csv" -NoTypeInformation
}
function loopCert{
    param ($csvContent,$csvHeader,$propertiyName1,$propertiyName2,$c)
    $certificationname=$null
    $signeddate=$null
    $count=0
    $countEntity=@()
    $te=@()
    $EntityTargetname=@()
    foreach ($line in $csvContent) {
        if ($line -match '<Certification activated=""') {
            if ($line -match 'signed=""([^"]+)""') {
                $value = $matches[1]
                if(!$signeddate){$signeddate = $value}
            }
        }
        if ($line -match '<Reference class=""sailpoint.object.CertificationGroup""') {
            if ($line -match 'name=""([^"]+)""') {
                $value = $matches[1]
                if(!$certificationname){$certificationname = $value}
            }
        }
        if ($line -match "<CertificationStatistics ") {
            if ($line -match 'totalItems=""([^"]+)""') {
                 $Items+= $matches[1]/1
            }
        }
        if ($line -match "<CertificationEntity ") {
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
    
    write-host "The certification name: $certificationname"
    write-host "The total Items: $Items"
    #$countEntity
    #$EntityTargetname
    #write-host "------------------------"

    
    $te+=getValue $csvContent "CertificationItem" $propertiyName1 $c
    $te+=getValue $csvContent "CertificationAction" $propertiyName2 $c
    $te=$te | Where-Object { $_ -match "\S" }
    $te.count

    # Initialize result array
    $result =InitArray $Items $te
    $result2 =ConcatArray $countEntity $EntityTargetname $result $signeddate $c
    $objects =DataProcessing $result2 $csvHeader $c

    ExportCSV $objects $outputFilePath $certificationname $c
    
    
    Clear-Variable -Name "certificationname"
    Clear-Variable -Name "result"
    Clear-Variable -Name "result2"
    Clear-Variable -Name "te"
    Clear-Variable -Name "countEntity"
    Clear-Variable -Name "EntityTargetname"
}

$c=0
foreach ($block in $blocks) {
    $block = $block -split "`n"
    $c++
    write-host "The block number: $c"
    loopCert $block $csvHeader $propertiyName1 $propertiyName2 $c
} 

Clear-Variable -Name "te"
Clear-Variable -Name "csvContent"
Clear-Variable -Name "blocks"
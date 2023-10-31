$Medium = Test-Path -Path $PSScriptRoot\Medium
$High = Test-Path -Path $PSScriptRoot\High
$Critical = Test-Path -Path $PSScriptRoot\Critical

if($Medium -eq $false){
    New-Item -Path $PSScriptRoot -Name "Medium" -ItemType "directory"
}

if($High -eq $false){
    New-Item -Path $PSScriptRoot -Name "High" -ItemType "directory"
}

if($Critical -eq $false){
    New-Item -Path $PSScriptRoot -Name "Critical" -ItemType "directory"
}

$InputDataInfo = Get-ChildItem $PSScriptRoot\InputData | Measure-Object
$ReportsDataInfo = Get-ChildItem $PSScriptRoot\Reports | Measure-Object

if($InputDataInfo.count -eq 0){
    Write-Host "Please provide csv report in InpudData folder!" -BackgroundColor Red
    exit
}

if($ReportsDataInfo.count -ne 0){
    Get-ChildItem -Path $PSScriptRoot\Reports -Include *.* -Recurse | foreach { $_.Delete()}
}

$Path = $PSScriptRoot + "\InputData"
$FullPath = $Path + "\" + (Get-ChildItem -Path $Path -Filter *.csv | Sort-Object -Descending -Property LastWriteTime | select -First 1).Name
$DestPathHighFinalReport = $PSScriptRoot+'\'+'Reports\'+'ReportHigh.html'
$DestPathCriticalFinalReport = $PSScriptRoot+'\'+'Reports\'+'ReportCritical.html'
$DestPathMediumFinalReport = $PSScriptRoot+'\'+'Reports\'+'ReportMedium.html'
$DestinationPathCriticalReport = $PSScriptRoot+'\' + "Critical"+'\'
$DestinationPathHighReport = $PSScriptRoot+'\' + "High"+'\'
$DestinationPathMediumReport = $PSScriptRoot+'\' + "Medium"+'\'
$fullPathTemplateCritical = $PSScriptRoot+'\'+'Templates\'+'Critical.htm'
$fullPathTemplateHigh = $PSScriptRoot+'\'+'Templates\'+'High.htm'
$fullPathTemplateMedium = $PSScriptRoot+'\'+'Templates\'+'Medium.htm'

$inputData=import-csv -Path $FullPath

$data = $inputData  | Where-Object {$_.Risk -eq 'Medium' -or $_.Risk -eq 'High' -or $_.Risk -eq 'Critical'} 
$grouppedData = $data | Group-Object -Property Name 

for($i=0;$i -lt $grouppedData.Count;$i++){
  
  $risk_rating = $grouppedData[$i].Group | Select-Object  -Property Risk -Unique
  
  if($risk_rating.Risk -contains 'High'){
    $content = Get-Content -Path $fullPathTemplateHigh
    $name = $grouppedData[$i].Name 
    $hosts = $grouppedData[$i].Group | Select-Object  -Property Host -Unique
    $Description = $grouppedData[$i].Group | Select-Object  -Property Description -Unique
    $Recommendations = $grouppedData[$i].Group | Select-Object  -Property Solution -Unique 
    $References =  $grouppedData[$i].Group | Select-Object  -Property "See Also" -Unique 
    $content | ForEach-Object {$_.replace("NameV",$name).replace("IPV",$hosts.Host).replace("DescriptionV",$Description.Description).replace("RecommendationsV",$Recommendations.Solution).replace("LinksV",$References.'See Also')} | Set-Content $DestinationPathHighReport\$i.htm

  }elseif ($risk_rating.Risk -contains 'Medium'){
    $content = Get-Content -Path $fullPathTemplateMedium
    $name = $grouppedData[$i].Name 
    $hosts = $grouppedData[$i].Group | Select-Object  -Property Host -Unique
    $Description = $grouppedData[$i].Group | Select-Object  -Property Description -Unique
    $Recommendations = $grouppedData[$i].Group | Select-Object  -Property Solution -Unique 
    $References =  $grouppedData[$i].Group | Select-Object  -Property "See Also" -Unique 
    $content | ForEach-Object {$_.replace("NameV",$name).replace("IPV",$hosts.Host).replace("DescriptionV",$Description.Description).replace("Re_comendationsV",$Recommendations.Solution).replace("LinksV",$References.'See Also')} | Set-Content $DestinationPathMediumReport\$i.htm

  }else {
    
    $content = Get-Content -Path $fullPathTemplateCritical
    $name = $grouppedData[$i].Name 
    $hosts = $grouppedData[$i].Group | Select-Object  -Property Host -Unique
    $Description = $grouppedData[$i].Group | Select-Object  -Property Description -Unique
    $Recommendations = $grouppedData[$i].Group | Select-Object  -Property Solution -Unique 
    $References =  $grouppedData[$i].Group | Select-Object  -Property "See Also" -Unique 
    $content | ForEach-Object {$_.replace("NameV",$name).replace("IPV",$hosts.Host).replace("DescriptionV",$Description.Description).replace("RecV",$Recommendations.Solution).replace("LinksV",$References.'See Also')} | Set-Content $DestinationPathCriticalReport\$i.htm

  }
}

$DestinationPathCriticalReportDirectory = Get-ChildItem $DestinationPathCriticalReport | Measure-Object

if($DestinationPathCriticalReportDirectory.Count -ne 0){
    $sw = New-Object System.IO.StreamWriter $DestPathCriticalFinalReport, $true  
    Get-ChildItem -Path $DestinationPathCriticalReport -Filter '*.htm' -File | ForEach-Object {
    Get-Content -Path $_.FullName -Encoding UTF8 | ForEach-Object {
            $sw.WriteLine($_)
        }
    }
    $sw.Dispose()
    Remove-Item $DestinationPathCriticalReport\*.*
}

$DestinationPathHighReportDirectory = Get-ChildItem $DestinationPathHighReport | Measure-Object

if($DestinationPathHighReportDirectory.Count -ne 0){
    $sw = New-Object System.IO.StreamWriter $DestPathHighFinalReport, $true  
    Get-ChildItem -Path $DestinationPathHighReport -Filter '*.htm' -File | ForEach-Object {
    Get-Content -Path $_.FullName -Encoding UTF8 | ForEach-Object {
            $sw.WriteLine($_)
        }
    }
    $sw.Dispose()
    Remove-Item $DestinationPathHighReport\*.*
}

$DestinationPathMediumReportDirectory = Get-ChildItem $DestinationPathMediumReport | Measure-Object

if($DestinationPathMediumReportDirectory.Count -ne 0){
    $sw = New-Object System.IO.StreamWriter $DestPathMediumFinalReport, $true  
    Get-ChildItem -Path $DestinationPathMediumReport -Filter '*.htm' -File | ForEach-Object {
    Get-Content -Path $_.FullName -Encoding UTF8 | ForEach-Object {
            $sw.WriteLine($_)
        }
    }
    $sw.Dispose()
    Remove-Item $DestinationPathMediumReport\*.*
}

Write-Host "Reports are done! Find in Reports folder!" -BackgroundColor Green
Read-Host -Prompt "Press any key to continue"



$RootPath = "C:\Report"
$OutFunc = "SystemReport" 
$tpSec10 = Test-Path "$RootPath \$OutFunc\"
if ($tpSec10 -eq $false)
{
New-Item -Path "$RootPath \$OutFunc\" -ItemType Directory -Force
}
$working = "$RootPath \$OutFunc\"
$Report = "$RootPath \$OutFunc\"+ "$OutFunc.html"

$Intro = "The results in this report are a guide and not a guarantee that the tested system is not without further defect or vulnerabilities."

$hn = Get-CimInstance -ClassName win32_computersystem 
$os = Get-CimInstance -ClassName win32_operatingsystem
$bios = Get-CimInstance -ClassName win32_bios
$cpu = Get-CimInstance -ClassName win32_processor

Foreach ($hfitem in $getHF)
{
$hfid = $hfitem.hotfixid
$hfdate = ($hfitem.installedon).ToShortDateString()
$hfurl = $hfitem.caption
$newObjHF = $hfid, $hfdate,$hfurl 
$HotFix += $newObjHF
} 

$HotFix=@()
$getHF = Get-HotFix | Select-Object HotFixID,InstalledOn,Caption 

Foreach ($hfitem in $getHF)
{
$hfid = $hfitem.hotfixid
$hfdate = $hfitem.installedon
$hfurl = $hfitem.caption

$newObjHF = New-Object psObject
Add-Member -InputObject $newObjHF -Type NoteProperty -Name HotFixID -Value $hfid
Add-Member -InputObject $newObjHF -Type NoteProperty -Name InstalledOn -Value ($hfdate).Date.ToString("dd-MM-yyyy")
Add-Member -InputObject $newObjHF -Type NoteProperty -Name Caption -Value $hfurl 
$HotFix += $newObjHF
}

$getUnin = Get-ChildItem "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"
$UninChild = $getUnin.Name.Replace("HKEY_LOCAL_MACHINE","HKLM:")
$InstallApps =@()
Foreach ( $uninItem in $UninChild)
{
$getUninItem = Get-ItemProperty $uninItem

$UninDisN = $getUninItem.DisplayName -replace "$null",""
$UninDisVer = $getUninItem.DisplayVersion -replace "$null",""
$UninPub = $getUninItem.Publisher -replace "$null",""
$UninDate = $getUninItem.InstallDate -replace "$null",""
$newObjInstApps = New-Object -TypeName PSObject

Add-Member -InputObject $newObjInstApps -Type NoteProperty -Name Publisher -Value $UninPub 
Add-Member -InputObject $newObjInstApps -Type NoteProperty -Name DisplayName -Value $UninDisN
Add-Member -InputObject $newObjInstApps -Type NoteProperty -Name DisplayVersion -Value $UninDisVer
Add-Member -InputObject $newObjInstApps -Type NoteProperty -Name InstallDate -Value $UninDate
$InstallApps += $newObjInstApps
}


$style = @"
<Style>
body
{
background-color:#250F00; 
color:#B87333;
font-size:100%;
font-family:helvetica;
margin:0,0,10px,0;
word-break:normal; 
word-wrap:break-word
}
table
{
border-width: 1px;
padding: 7px;
border-style: solid;
border-color:#B87333;
border-collapse:collapse;
width:auto
}
h1
{
background-color:#250F00; 
color:#B87333;
font-size:150%;
font-family:helvetica;
margin:0,0,10px,0;
word-break:normal; 
word-wrap:break-word
}
h2
{
background-color:#250F00; 
color:#4682B4;
font-size:120%;
font-family:helvetica;
margin:0,0,10px,0; 
word-break:normal; 
word-wrap:break-word
}
h3
{
background-color:#250F00; 
color:#B87333;
font-size:100%;
font-family:helvetica;
margin:0,0,10px,0; 
word-break:normal; 
word-wrap:break-word;
font-weight: normal;
width:auto
}
th
{
border-width: 1px;
padding: 7px;
border-style: solid;
border-color:#B87333;
background-color:#250F00
}
td
{
border-width: 1px;
padding:7px;
border-style: solid; 
border-style: #B87333
}
tr:nth-child(odd) 
{
background-color:#250F00;
}
tr:nth-child(even) 
{
background-color:#181818;
}

</Style>
"@

$FragDescrip1 = $Descrip1 | ConvertTo-Html -as table -Fragment -PreContent "<h3><span>$Intro</span></h3>" | Out-String
$fragHost = $hn | ConvertTo-Html -As table -Property Name,Domain,Model -fragment -PreContent "<h2><span>Host Details</span></h2>"  | Out-String
$fragOS = $OS | ConvertTo-Html -As table -property Caption,Version,OSArchitecture,InstallDate -fragment -PreContent "<h2><span>Windows Details</span></h2>" | Out-String
$fragBios = $bios | ConvertTo-Html -As table -property Name,Manufacturer,SerialNumber,SMBIOSBIOSVersion,ReleaseDate -fragment -PreContent "<h2><span>Bios Details</span></h2>" | Out-String
$fragCpu = $cpu | ConvertTo-Html -As table -property Name,MaxClockSpeed,NumberOfCores,ThreadCount -fragment -PreContent "<h2><span>Processor Details</span></h2>" | Out-String
$fragHotFix = $HotFix | ConvertTo-Html -As table -property HotFixID,InstalledOn,Caption -fragment -PreContent "<h2><span>Installed Updates</span></h2>" | Out-String
$fragInstaApps = $InstallApps | Sort-Object publisher,displayname -Unique | ConvertTo-Html -As Table -fragment -PreContent "<h2><span>Installed Applications</span></h2>" | Out-String


ConvertTo-Html -Head $style -Body "<h1 align=center style='text-align:center'><span style='color:#4682B4;'>TENAKA.NET</span><h1>", 
$fragDescrip1, 
$fraghost, 
$fragOS, 
$fragInstaApps,
$fragHotFix,
$fragbios, 
$fragcpu | out-file $Report
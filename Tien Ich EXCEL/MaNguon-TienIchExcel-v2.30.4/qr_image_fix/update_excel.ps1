$ErrorActionPreference='Stop'
$p='HKCU:\Software\Microsoft\Office\16.0\Excel\Options'
$k=Get-Item $p
foreach($n in $k.GetValueNames()){
 if($n -match '^OPEN\d*$' -and [string]$k.GetValue($n) -match 'TienIchExcel-v2\.(28|29)(\.1)?\.xlam'){
  Set-ItemProperty -Path $p -Name $n -Value ('"'+(Join-Path (Get-Location) 'output/TienIchExcel-v2.28.1.xlam')+'"')
  'Updated own Excel add-in startup entry '+$n
 }
}

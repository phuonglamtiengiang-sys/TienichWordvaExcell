$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$excel=$null;$addin=$null
try{
 Copy-Item "$root/output/TienIchExcel-v2.28.xlam" "$root/output/TienIchExcel-v2.28.1.xlam" -Force
 $excel=New-Object -ComObject Excel.Application
 $addin=$excel.Workbooks.Open("$root/output/TienIchExcel-v2.28.1.xlam")
 foreach($name in @('mod_Core','mod_QRImageFormat')){
  $component=$null;foreach($existing in $addin.VBProject.VBComponents){if($existing.Name -eq $name){$component=$existing;break}}
  if(-not $component){$component=$addin.VBProject.VBComponents.Add(1);$component.Name=$name}
  if($component.CodeModule.CountOfLines -gt 0){$component.CodeModule.DeleteLines(1,$component.CodeModule.CountOfLines)}
  $component.CodeModule.AddFromString([IO.File]::ReadAllText("$root/src/$name.txt"))
 }
 $addin.Save()
 'Built Excel classic VietQR image fix'
}finally{if($addin){$addin.Close($false)};if($excel){$excel.Quit()}}

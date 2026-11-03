$ErrorActionPreference='Stop'
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) 'output/Word-1.3.2/TienIchWord.dll'))|Out-Null
$w=New-Object -ComObject Word.Application;$d=$w.Documents.Add();$z=0
try{
 $engine=New-Object TienIchWord.WordEngine -ArgumentList $w
 foreach($mode in @('accent','clean','normalize')){
  $d.Content.InsertXML([IO.File]::ReadAllText((Join-Path (Get-Location) 'qr_image_fix/user-document.xml')))
  $r=$d.Range(0,$d.Content.End-1);$before=$r.Text;$shapes=$d.Shapes.Count
  if($mode -eq 'normalize'){$p=[TienIchWord.TextNormalization]::Create($before,$true,$true,$false,$false);$expected=$p.Result;$engine.ApplyPlan($r,$p,$false)}
  else{
   if($mode -eq 'accent'){$expected=[TienIchWord.TextTools]::StripAccents($before)}else{$expected=[string]::Concat([TienIchWord.TextTools]::CleanTokens($before))}
   $engine.Transform($r,$mode)
  }
  if($d.Range(0,$d.Content.End-1).Text -cne $expected){throw "Mismatch $mode"}
  if($d.Shapes.Count -ne $shapes){throw 'Lost floating object'}
  if($before -cne $expected){$d.Undo(1)|Out-Null;if($d.Range(0,$d.Content.End-1).Text -cne $before){throw "Undo $mode"}}
  "PASS actual copy $mode and Undo"
 }
 foreach($source in @('vni','tcvn')){
  if($source -eq 'vni'){$s="  Coäng hoøa xaõ hoäi chuû nghóa Vieät Nam`rÑaëng Vaên An  "}else{$s="  Céng hßa x· héi chñ nghÜa ViÖt Nam`r  "}
  $d.Content.Text=$s
  $anchor=$d.Range(0,0);$shape=$d.Shapes.AddShape(1,10,10,20,20,$anchor)
  $r=$d.Range(0,$d.Content.End-1);$before=$r.Text;$count=$d.Shapes.Count
  $r.Font.Bold=-1;$r.Font.Name='Arial'
  $p=[TienIchWord.VietnameseEncoding]::Convert($before,$source,$false)
  $engine.ApplyPlan($r,$p,$true)
  if($d.Range(0,$d.Content.End-1).Text -cne $p.Result){throw "Conversion mismatch $source"}
  if($d.Shapes.Count -ne $count){throw 'Lost shape during encoding'}
  if($d.Range(0,1).Font.Bold -ne -1){throw 'Lost bold'}
  $d.Undo(1)|Out-Null
  if($d.Range(0,$d.Content.End-1).Text -cne $before){throw 'Encoding Undo mismatch'}
  "PASS $source with floating object, formatting and Undo"
 }
}finally{$d.Close([ref]$z);if($w.Documents.Count -eq 0){$w.Quit([ref]$z)}}

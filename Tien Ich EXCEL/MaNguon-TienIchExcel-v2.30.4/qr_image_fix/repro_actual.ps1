$ErrorActionPreference='Stop'
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) 'output/Word-1.3/TienIchWord.dll'))|Out-Null
$w=[Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
$d=$w.Documents.Add();$z=0
try{
 $d.Content.InsertXML([IO.File]::ReadAllText((Join-Path (Get-Location) 'qr_image_fix/user-document.xml')))
 $r=$d.Range(0,$d.Content.End-1);$before=$r.Text;$span=$r.End-$r.Start
 $paragraphs=$d.Paragraphs.Count;$shapes=$d.Shapes.Count
 $engine=New-Object TienIchWord.WordEngine -ArgumentList $w
 $engine.Transform($r,'upper')
 if($d.Range(0,$d.Content.End-1).Text -cne $before.ToUpperInvariant()){throw 'Uppercase mismatch'}
 if($d.Paragraphs.Count -ne $paragraphs -or $d.Shapes.Count -ne $shapes){throw 'Structure changed'}
 $d.Undo(1)|Out-Null
 if($d.Range(0,$d.Content.End-1).Text -cne $before){throw 'Undo mismatch'}
 'PASS actual document copy: uppercase, paragraphs, floating objects, single Undo; text length='+$before.Length+' Word span='+$span
 foreach($tracked in @($false,$true)){
  $d.Content.Text="Cộng hòa xã hội chủ nghĩa Việt Nam`rĐộc lập - Tự do - Hạnh phúc`rchữ u và ù, ư, đ; tiếng Việt"
  $d.TrackRevisions=$tracked;$r=$d.Range(0,$d.Content.End-1);$original=$r.Text
  $engine.Transform($r,'upper');$d.TrackRevisions=$false;$d.AcceptAllRevisions()
  if($d.Range(0,$d.Content.End-1).Text -cne $original.ToUpperInvariant()){throw 'Vietnamese mismatch'}
  $engine.Transform($d.Range(0,$d.Content.End-1),'lower')
  if($d.Range(0,$d.Content.End-1).Text -cne $original.ToLowerInvariant()){throw 'Lowercase mismatch'}
  'PASS Vietnamese upper/lower tracked='+$tracked
 }
}finally{$d.Close([ref]$z)}

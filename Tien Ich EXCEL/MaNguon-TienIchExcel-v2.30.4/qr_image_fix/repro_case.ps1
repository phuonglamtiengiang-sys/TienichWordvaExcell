$ErrorActionPreference='Stop'
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) 'output/Word-1.3.1/TienIchWord.dll'))|Out-Null
$w=New-Object -ComObject Word.Application
$d=$w.Documents.Add();$z=0
try{
 $s="Cộng hòa xã hội chủ nghĩa Việt Nam`rĐộc lập - Tự do - Hạnh phúc`rHọ tên: Hồ Phương Lâm"
 $d.Content.Text=$s;$d.TrackRevisions=$true
 $e=New-Object TienIchWord.WordEngine -ArgumentList $w
 $e.Transform($d.Range(0,$d.Content.End-1),'upper')
 $d.TrackRevisions=$false;$d.AcceptAllRevisions()
 if($d.Content.Text.TrimEnd([char]13) -cne $s.ToUpperInvariant()){'REPRODUCED: tracked changes corrupts character offsets'}else{'NO REPRO'}
}finally{$d.Close([ref]$z);$w.Quit([ref]$z)}

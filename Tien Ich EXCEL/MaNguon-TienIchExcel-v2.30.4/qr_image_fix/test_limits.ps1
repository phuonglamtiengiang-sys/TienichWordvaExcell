$ErrorActionPreference='Stop'
[Reflection.Assembly]::LoadFrom((Join-Path (Get-Location) 'output/Word-1.3.6/TienIchWord.dll'))|Out-Null
$p=New-Object TienIchWord.Preferences;$p.SelectionWarningLimit=100000
if($p.ShouldWarn(100000) -or !$p.ShouldWarn(100001)){throw 'Warning limit'}
$p.Target='vi';$s='a'*20000
if([TienIchWord.Translator]::Translate($s,$p,[Func[bool]]{return $true},[Action[int,int]]{param($a,$b)}) -cne $s){throw 'Translation boundary'}
$rejected=$false;try{[TienIchWord.Translator]::Translate(('a'*20001),$p,[Func[bool]]{return $true},[Action[int,int]]{param($a,$b)})|Out-Null}catch{$rejected=$true};if(!$rejected){throw 'Translation over limit'}
$w=New-Object -ComObject Word.Application;$d=$w.Documents.Add();$zero=0
try{
 $d.Content.Text='a'*100000;$r=$d.Range(0,100000);[TienIchWord.WordEngine]::Validate($r,$true)
 $engine=New-Object TienIchWord.WordEngine -ArgumentList $w;$engine.Transform($r,'upper')
 if($d.Range(0,100000).Text -cne ('A'*100000)){throw 'Uppercase large selection'}
 $d.Undo(1)|Out-Null;if($d.Range(0,100000).Text -cne ('a'*100000)){throw 'Undo'}
 $d.Content.Text='a'*100001;$rejected=$false;try{[TienIchWord.WordEngine]::Validate($d.Range(0,100001),$true)}catch{$rejected=$true};if(!$rejected){throw 'Text over limit'}
 'PASS 100000 text/uppercase/Undo; 100001 rejected; warning boundary; translation 20000 accepted and 20001 rejected without external requests'
}finally{$d.Close([ref]$zero);if($w.Documents.Count -eq 0){$w.Quit([ref]$zero)}}

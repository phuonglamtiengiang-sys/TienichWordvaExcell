$ErrorActionPreference='Stop'
$root=Get-Location
$excel=$null;$word=$null;$book=$null;$doc=$null;$zero=0
try{
 $path=(Resolve-Path 'qr_image_fix/blank.svg').Path
 $excel=New-Object -ComObject Excel.Application;$book=$excel.Workbooks.Add()
 $pic=$book.Worksheets.Item(1).Shapes.AddPicture($path,$false,$true,10,10,-1,-1)
 'Excel SVG: '+$pic.Width+' x '+$pic.Height
 $word=New-Object -ComObject Word.Application;$doc=$word.Documents.Add()
 $r=$doc.Range(0,0);$pic2=$doc.InlineShapes.AddPicture($path,$false,$true,$r)
 'Word SVG: '+$pic2.Width+' x '+$pic2.Height
}finally{if($book){$book.Close($false)};if($excel){$excel.Quit()};if($doc){$doc.Close([ref]$zero)};if($word){$word.Quit([ref]$zero)}}

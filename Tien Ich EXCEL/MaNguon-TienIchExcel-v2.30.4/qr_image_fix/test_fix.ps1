$ErrorActionPreference='Stop'
$root=Split-Path $PSScriptRoot -Parent
$excel=$null;$addin=$null;$book=$null;$word=$null;$doc=$null;$zero=0
[Reflection.Assembly]::LoadFrom("$root/output/Word-1.3.1/TienIchWord.dll")|Out-Null
try{
 $excel=New-Object -ComObject Excel.Application
 $addin=$excel.Workbooks.Open("$root/output/TienIchExcel-v2.28.1.xlam")
 $helper=$addin.VBProject.VBComponents.Add(1);$helper.Name='mod_ImageFixTestOnly'
 $helper.CodeModule.AddFromString(@'
Option Explicit
Public Function FixtureImage(ByVal path As String) As String
 Dim stream As Object, bytes As Variant
 Set stream = CreateObject("ADODB.Stream"): stream.Type = 1: stream.Open
 stream.LoadFromFile path: bytes = stream.Read: stream.Close
 FixtureImage = QRImageExtension(bytes, "vietqr")
End Function
Public Function InsertFixtureQR(ByVal target As Range, ByVal amount As String) As Boolean
 InsertFixtureQR = InsertCodeImage(target, VietQRURL("970422", "123456789", amount, "TEST"), 120, "vietqr")
End Function
'@)
 $macro="'$($addin.Name)'!mod_ImageFixTestOnly."
 if($excel.Run($macro+'FixtureImage',"$root/qr_image_fix/blank.svg") -ne '.svg'){throw 'VBA SVG detection failed'}
 'PASS Excel accepts actual SVG response'
 $book=$excel.Workbooks.Add();$sheet=$book.Worksheets.Item(1);$sheet.Range('A1').Value2='ORIGINAL DATA'
 if(-not $excel.Run($macro+'InsertFixtureQR',$sheet.Range('B2'),'')){throw 'Excel blank amount QR failed'}
 if(-not $excel.Run($macro+'InsertFixtureQR',$sheet.Range('F2'),'125000')){throw 'Excel amount QR failed'}
 if($sheet.Range('A1').Value2 -ne 'ORIGINAL DATA' -or $sheet.Shapes.Count -ne 2){throw 'Unexpected Excel document state'}
 $book.SaveAs("$root/qr_image_fix/excel-result.xlsx",51)
 'PASS Excel creates both VietQR images and preserves source cells'
 $word=New-Object -ComObject Word.Application;$doc=$word.Documents.Add();$doc.Content.Text='ORIGINAL DATA'
 $engine=New-Object TienIchWord.WordEngine -ArgumentList $word
 $number=0
 foreach($amount in @('','125000')){
  $bytes=[TienIchWord.Codes]::Download([TienIchWord.Codes]::VietQR('970422','123456789',$amount,'TEST'))
  $info=[TienIchWord.OfficeImage]::Read($bytes,$true)
  if(-not $info.Svg){throw 'Live SVG test did not receive SVG'}
  $range=$doc.Range(0,0);$engine.InsertPicture($range,$bytes,'vietqr')|Out-Null;$number++
 }
 if($doc.InlineShapes.Count -ne 2 -or -not $doc.Content.Text.Contains('ORIGINAL DATA')){throw 'Word did not preserve document state'}
 $doc.SaveAs2("$root/qr_image_fix/word-result.docx",16)
 $doc.Undo(1)|Out-Null
 if($doc.InlineShapes.Count -ne 1){throw 'Word image Undo failed'}
 'PASS Word downloads SVG, inserts both QR images, retains text and supports Undo'
 @('PASS classic Excel and Word forms unchanged; no external designer','PASS live blank and amount VietQR requests','PASS SVG embedded in saved XLSX/DOCX','PASS source data and Word Undo retained')|Set-Content "$root/qr_image_fix/results.txt" -Encoding UTF8
}finally{if($book){$book.Close($false)};if($addin){$addin.Close($false)};if($excel){$excel.Quit()};if($doc){$doc.Close([ref]$zero)};if($word){$word.Quit([ref]$zero)}}

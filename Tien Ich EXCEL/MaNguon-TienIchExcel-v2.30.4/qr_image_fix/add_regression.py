from pathlib import Path
p=Path('qr_image_fix/test_text.ps1');s=p.read_text(encoding='utf-8-sig')
s=s.replace(" $d.Content.InsertXML([IO.File]::ReadAllText((Join-Path (Get-Location) 'qr_image_fix/user-document.xml')))",''' $d.Content.Text="  Cộng  hòa xã hội chủ nghĩa Việt Nam  `rĐộc lập - Tự do - Hạnh phúc`rTiếng Việt ,đúng rồi !  "
  $anchor=$d.Range(0,0);$null=$d.Shapes.AddShape(1,10,10,20,20,$anchor)''')
s=s.replace("(Join-Path (Get-Location) 'output/Word-1.3.2/TienIchWord.dll')", "(Join-Path (Split-Path $PSScriptRoot -Parent) 'output/Word-1.3.2/TienIchWord.dll')").replace('actual copy','synthetic floating anchor')
Path('tests/test_word_text_positions.ps1').write_text(s,encoding='utf-8-sig')
p=Path('word/package_word.py');s=p.read_text(encoding='utf-8-sig').replace("['test_word_ribbon_schema.ps1'","['test_word_text_positions.ps1','test_word_ribbon_schema.ps1'");p.write_text(s,encoding='utf-8-sig')

from pathlib import Path
p=Path('word/WordUtilities.cs');s=p.read_text(encoding='utf-8-sig').replace('            if(mode=="accent"){\n                var replacements=', '            if(mode=="accent"&&!original.Any(c=>CharUnicodeInfo.GetUnicodeCategory(c)==UnicodeCategory.NonSpacingMark)){\n                var replacements=');p.write_text(s,encoding='utf-8-sig')

// SPDX-License-Identifier: GPL-2.0-or-later
// Based on UniKey VnConv (UVConverter 1.1.3b), Copyright (C) 1998-2002 Pham Kim Long.
// C# adaptation (2026) for Tien Ich Word: SingleByteCharset constructor/nextInput,
// DoubleByteCharset constructor/nextInput and UnicodeCharset::putChar from charset.cpp.
// Changes: Word UTF-16/Windows-1252 bridge; positions for formatting-preserving edits;
// preserve characters outside Windows-1252; reject orphan VNI marks rather than drop them.
// This program is distributed under GPL version 2 or (at your option) any later version.
// No warranty. See vendor/unikey/upstream/uvconv/gpl.txt.
using System;
using System.Collections.Generic;
using System.Text;

namespace TienIchWord {
    public static class UniKeyDecoder {
        static readonly Encoding Western=Encoding.GetEncoding(1252,EncoderFallback.ExceptionFallback,DecoderFallback.ExceptionFallback);
        static readonly int[] TcvnMap=BuildSingle(),VniMap=BuildDouble();
        static readonly Dictionary<int,int> VniPairs=BuildPairs();
        static int[] BuildSingle(){
            var map=new int[256];var chars=UniKeyTables.Tcvn;
            for(int i=0;i<chars.Length;i++)if(chars[i]!=0&&(i==chars.Length-1||chars[i]!=chars[i+1]))map[chars[i]]=i+1;
            return map;
        }
        static int[] BuildDouble(){
            var map=new int[256];var chars=UniKeyTables.Vni;
            for(int i=0;i<chars.Length;i++){
                if((chars[i]>>8)!=0)map[chars[i]>>8]=-1;
                else if(map[chars[i]]==0)map[chars[i]]=i+1;
            }return map;
        }
        static Dictionary<int,int> BuildPairs(){var pairs=new Dictionary<int,int>();for(int i=0;i<UniKeyTables.Vni.Length;i++)if(UniKeyTables.Vni[i]>255)pairs.Add(UniKeyTables.Vni[i],i);return pairs;}
        static bool ByteAt(string text,int index,out int value){
            value=0;if(index>=text.Length)return false;
            char c=text[index];if(c<=255){value=c;return true;}
            try{byte[] bytes=Western.GetBytes(new[]{c});if(bytes.Length==1){value=bytes[0];return true;}}catch(EncoderFallbackException){}return false;
        }
        public static TextPlan Convert(string text,string source,bool uppercase){
            if(source!="tcvn"&&source!="vni")throw new ArgumentException("Chọn TCVN3 hoặc VNI Windows.");
            var plan=new TextPlan(text);var edits=new List<TextEdit>();
            for(int i=0;i<text.Length;){
                int ch;if(!ByteAt(text,i,out ch)){i++;continue;}
                int count=1,index=(source=="tcvn"?TcvnMap:VniMap)[ch]-1;
                if(source=="vni"){
                    if(index==-2)throw new ArgumentException("Dấu VNI đứng riêng tại ký tự "+(i+1)+". Kiểm tra bảng mã hoặc chọn lại vùng; chưa thay đổi tài liệu.");
                    if(index>=0){int hi,pair;if(ByteAt(text,i+1,out hi)&&hi>0&&VniPairs.TryGetValue(ch|(hi<<8),out pair)){index=pair;count=2;}}
                }
                string value=index>=0?((char)UniKeyTables.Unicode[index]).ToString():text.Substring(i,count);
                if(uppercase)value=value.ToUpperInvariant();
                if(value!=text.Substring(i,count))edits.Add(new TextEdit{Start=i,Length=count,Value=value});
                i+=count;
            }
            plan.Add(edits);return plan;
        }
    }
}

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Xml;
namespace TienIchWord {
 public sealed class OfficeImage {
  public bool Svg;public int Width,Height;
  public static OfficeImage Read(byte[] bytes,bool allowSvg){
   if(bytes==null||bytes.Length<24||bytes.Length>10000000)throw new IOException("Dữ liệu ảnh không hợp lệ.");
   if(bytes[0]==137&&bytes[1]==80&&bytes[2]==78&&bytes[3]==71){using(var stream=new MemoryStream(bytes))using(var image=Image.FromStream(stream))return new OfficeImage{Width=image.Width,Height=image.Height};}
   if(!allowSvg)throw new IOException("Dịch vụ không trả về ảnh PNG hợp lệ.");
   var doc=new XmlDocument{XmlResolver=null};
   using(var stream=new MemoryStream(bytes))using(var reader=XmlReader.Create(stream,new XmlReaderSettings{DtdProcessing=DtdProcessing.Prohibit,XmlResolver=null,MaxCharactersInDocument=10000000}))doc.Load(reader);
   if(doc.DocumentElement==null||doc.DocumentElement.LocalName!="svg"||doc.DocumentElement.NamespaceURI!="http://www.w3.org/2000/svg")throw new IOException("Dịch vụ không trả về ảnh VietQR hợp lệ.");
   string[] allowed={"svg","title","desc","rect","g","path","image","text","tspan","defs","linearGradient","stop","clipPath","circle","ellipse","line","polyline","polygon"};
   foreach(XmlElement e in doc.SelectNodes("//*")){
    if(e.NamespaceURI!="http://www.w3.org/2000/svg"||!allowed.Contains(e.LocalName))throw new IOException("Ảnh SVG có thành phần không hỗ trợ.");
    foreach(XmlAttribute a in e.Attributes){
     if(a.LocalName.StartsWith("on",StringComparison.OrdinalIgnoreCase)||a.Value.IndexOf("url(",StringComparison.OrdinalIgnoreCase)>=0)throw new IOException("Ảnh SVG có nội dung không hỗ trợ.");
     if(a.LocalName=="href"&&!a.Value.StartsWith("data:image/png;base64,",StringComparison.OrdinalIgnoreCase)&&!a.Value.StartsWith("data:image/jpeg;base64,",StringComparison.OrdinalIgnoreCase))throw new IOException("Không tải tài nguyên bên ngoài từ SVG.");
    }
   }
   int width=Dimension(doc.DocumentElement.GetAttribute("width")),height=Dimension(doc.DocumentElement.GetAttribute("height"));
   if(width<1||height<1||width>8192||height>8192)throw new IOException("Kích thước SVG không hợp lệ.");
   return new OfficeImage{Svg=true,Width=width,Height=height};
  }
  static int Dimension(string value){if(!Regex.IsMatch(value??"",@"^[0-9]+(?:px)?$"))throw new IOException("Kích thước SVG không hỗ trợ.");return int.Parse(value.Replace("px",""));}
 }
}

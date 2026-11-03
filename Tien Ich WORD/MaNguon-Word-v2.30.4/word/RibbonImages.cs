using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
namespace TienIchWord {
    // AxHost supplies the Office-compatible OLE IPictureDisp wrapper.
    internal sealed class RibbonPicture:AxHost {
        RibbonPicture():base(""){}
        internal static object Wrap(Image image){return GetIPictureDispFromPicture(image);}
    }
    internal static class RibbonImages {
        static readonly Dictionary<string,object> Pictures=new Dictionary<string,object>();
        static readonly List<Image> Bitmaps=new List<Image>();
        internal static object Load(string name){
            lock(Pictures){object picture;if(Pictures.TryGetValue(name,out picture))return picture;
                using(var source=Resources.Picture("icons."+name+".png")){
                    var bitmap=new Bitmap(source);picture=RibbonPicture.Wrap(bitmap);Bitmaps.Add(bitmap);Pictures.Add(name,picture);return picture;
                }
            }
        }
    }
}

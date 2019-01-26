using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PdfBackend
{
    class FontManager
    {
        static Dictionary<string, XFont> fonts = new Dictionary<string, XFont>();
        public static XFont GetFont(string name, double size, FontStyle fontStyle)
        {
            var k = name + size + fontStyle;
            XFont font;
            if (!fonts.TryGetValue(k, out font))
            {
                XFontStyle fs = XFontStyle.Regular;
                fs |= (fontStyle & FontStyle.Bold) == FontStyle.Bold ? XFontStyle.Bold : XFontStyle.Regular;
                fs |= (fontStyle & FontStyle.Underline) == FontStyle.Underline ? XFontStyle.Underline : XFontStyle.Regular;
                font = new XFont(name, size, fs);
                fonts.Add(k, font);
            }
            return font;
        }
    }
}

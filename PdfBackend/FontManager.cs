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
        public static XFont GetFont(string name, double size, bool bold)
        {
            var k = name + size + bold;
            XFont font;
            if (!fonts.TryGetValue(k, out font))
            {
                font = new XFont(name, size, bold ? XFontStyle.Bold : XFontStyle.Regular);
                fonts.Add(k, font);
            }
            return font;
        }
    }
}

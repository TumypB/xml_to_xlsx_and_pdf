using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PdfBackend
{
    class Helpers
    {
        public static double ParseSize(string size)
        {
            if (size.EndsWith("cm"))
                return 28.3465d * double.Parse(size.Substring(0, size.Length - 2), CultureInfo.InvariantCulture);
            else if (size.EndsWith("pt"))
                return double.Parse(size.Substring(0, size.Length - 2), CultureInfo.InvariantCulture);
            else
                return double.Parse(size, CultureInfo.InvariantCulture) * 6;
        }

        public static Tuple<string, string> WrapLine(XGraphics gfx, XFont font, double width, string input)
        {
            var words = Regex.Matches(input, @"(\W*)(\w+)(\W*)").Cast<Match>().Select(m => m.Value).ToArray();

            var line = "";
            var remainder = (string)null;
            foreach (var w in words)
            {
                var m = gfx.MeasureString(line + w, font);
                if (m.Width > width || remainder != null)
                {
                    remainder += w;
                }
                else
                {
                    line += w;
                }
            }
            return Tuple.Create(line, remainder);
        }

        public static Tuple<string, string> BreakLine(XGraphics gfx, XFont font, double width, string input)
        {
            int i = 2;
            for (; i < input.Length; i++)
            {
                var m = gfx.MeasureString(input.Substring(0, i), font);
                if (m.Width >= width)
                {
                    break;
                }
            }
            if (input.Length < 2)
                return Tuple.Create(input, (string)null);
            else
                return Tuple.Create(input.Substring(0, i - 1), input.Substring(i - 1));
        }

        public static string GetAttributeValue(XElement xStyle, string attr)
        {
            return ((string)xStyle.Attribute(attr))?.Trim().ToLower();
        }

    }
}
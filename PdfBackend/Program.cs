using PdfSharp.Drawing;
using PdfSharp.Drawing.Layout;
using PdfSharp.Pdf;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PdfBackend
{
    class Program
    {
        internal const int DefaultFontSize = 10;
        internal const string DefaultFontName = "Arial";

        //static XFont defaultFont = new XFont("Verdana", 20, XFontStyle.Bold);
        static void Main(string[] args)
        {
            var d = new PdfDocument();

            var page = d.AddPage();

             XGraphics gfx = XGraphics.FromPdfPage(page);

            var xBook = XElement.Load(@"C:\000\MyReport.xml");
            var xSheet = xBook.Element("sheet");
            var ss = StyleSheet.FromXElement(xBook.Element("styles"));
            var sheet = Sheet.FromXElement(xSheet, ss);

            Render.MeasureSheet(gfx, sheet);
            Render.RenderSheet(gfx, sheet);

            d.Save(@"C:\000\test.pdf");
        }
    }

    abstract class Block
    {
        public double X;
        public double Y;
        public abstract double Height { get; set; }
        public double Width { get; set; }
        public abstract double Ascent { get; }
        public abstract double Descent { get; }
    }

    class TextBlock : Block
    {
        public string Text;
        public XFont Font;
        public override double Height { get; set; }
        public override double Ascent => Font.GetHeight() * Font.CellAscent / Font.CellSpace;
        public override double Descent => Font.GetHeight() * Font.CellDescent / Font.CellSpace;
    }

    class ImageBlock : Block
    {
        public XImage Image;
        public override double Height { get; set; }
        public override double Ascent => Height;
        public override double Descent => 0;
    }
}

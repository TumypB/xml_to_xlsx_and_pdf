using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace PdfBackend
{
    class SpanInfo
    {
        public int StartRow;
        public int StartColumn;
        public int ColumnCount;
        public int RowCount;

        public int LastRow
        {
            get
            {
                return StartRow + RowCount - 1;
            }
        }

        public int LastColumn
        {
            get
            {
                return StartColumn + ColumnCount - 1;
            }
        }
    }

    class Sheet
    {
        public List<Column> Columns = new List<Column>();
        public List<Row> Rows = new List<Row>();
        public StyleSheet styleSheet = new StyleSheet();

        public double MarginLeft;
        public double MarginTop;

        public static Sheet FromXElement(XElement xSheet, StyleSheet styles)
        {
            var res = new Sheet();

            var marginLeft = (string)xSheet.Attribute("marginLeft");
            var marginTop = (string)xSheet.Attribute("marginTop");
            res.MarginLeft = string.IsNullOrWhiteSpace(marginLeft) ? 0 : Helpers.ParseSize(marginLeft);
            res.MarginTop = string.IsNullOrWhiteSpace(marginTop) ? 0 : Helpers.ParseSize(marginTop);

            var styleName = (string)xSheet.Attribute("style");

            var defaultStyle = styleName == null ? new Style()
            {
                FontName = "Arial",
                FontSize = 14,
                Borders = new double[] { 0d, 0d, 0d, 0d },
                Fill = "",
                Decorations = new string[] { "solid", "solid", "solid", "solid" },
                Name = "default"
            } : styles[styleName];

            res.Columns = xSheet.Element("colgroup").Elements().Select(Column.FromXElement).ToList();
            res.Rows = xSheet.Elements("tr").Select(xTr => Row.FromXElement(xTr, styles, defaultStyle)).ToList();
            return res;
        }

        public double GetColumnPosition(int i)
        {
            var lCol = Columns[Columns.Count - 1];
            return i < Columns.Count ? Columns[i].Left : lCol.Left + lCol.Width;
        }

        internal double GetDefaultRowHeight()
        {
            return 20;
        }
    }

    enum Align
    {
        Near, Center, Far
    }

    class Style
    {
        public string Name;
        public string Fill;
        public string FontName;
        public double FontSize;
        public double[] Borders;
        public string[] Decorations;
        public double MarginLeft;
        public double MarginTop;
        public double MarginRight;
        public double MarginBottom;

        public Align VerticalAlign;
        public Align HorizontalAlign;

        public static Style FromXElement(XElement xStyle)
        {
            var s = new Style();
            s.Name = (string)xStyle.Attribute("name");
            s.Fill = (string)xStyle.Attribute("fill");
            s.FontName = (string)xStyle.Attribute("font-name");
            s.FontSize = (double)xStyle.Attribute("font-size");
            s.Borders = ((string)xStyle.Attribute("border")).Split(' ', ',', ';').Select(x => double.Parse(x, CultureInfo.InvariantCulture)).ToArray();

            var margin = (string)xStyle.Attribute("margin");
            var margins = string.IsNullOrWhiteSpace(margin) ?
                new double[4] : margin.Split(';').Select(x => Helpers.ParseSize(x)).ToArray();
            s.MarginLeft = margins[0]; s.MarginTop = margins[1]; s.MarginRight = margins[2]; s.MarginBottom = margins[3];

            s.Decorations = ((string)xStyle.Attribute("decoration")).Split(new char[] { ' ', ',', ';' });

            switch (Helpers.GetAttributeValue(xStyle, "valign") ?? "")
            {
                case "near": s.VerticalAlign = Align.Near; break;
                case "far": s.VerticalAlign = Align.Far; break;
                case "center": s.VerticalAlign = Align.Center; break;
                default: s.VerticalAlign = Align.Near; break;
            }

            switch (Helpers.GetAttributeValue(xStyle, "align") ?? "")
            {
                case "near": s.HorizontalAlign = Align.Near; break;
                case "far": s.HorizontalAlign = Align.Far; break;
                case "center": s.HorizontalAlign = Align.Center; break;
                default: s.HorizontalAlign = Align.Near; break;
            }

            return s;
        }
    }

    class StyleSheet : Dictionary<string, Style>
    {
        public static StyleSheet FromXElement(XElement xStyles)
        {
            var ss = new StyleSheet();

            foreach (var xs in xStyles.Elements())
            {
                var s = Style.FromXElement(xs);
                ss.Add(s.Name, s);
            }

            return ss;
        }
    }

    class Column
    {
        public double Left { get; set; }
        public double Width { get; set; }

        internal static Column FromXElement(XElement xCol)
        {
            var col = new Column();

            var width = (string)xCol.Attribute("width");

            col.Width = Helpers.ParseSize(width);

            return col;
        }
    }

    class Row
    {
        //Sheet parent;

        public double Height { get; set; }
        public double RenderHeight { get; set; }
        public List<Cell> Cells = new List<Cell>();
        public int RowIndex;

        public Style Style;

        internal static Row FromXElement(XElement xTr, StyleSheet styles, Style defaultStyle)
        {
            var row = new Row();
            //row.parent = parent;
            var height = (string)xTr.Attribute("height");
            row.Height = string.IsNullOrWhiteSpace(height) ? 0 : Helpers.ParseSize(height);

            var styleName = (string)xTr.Attribute("style");
            row.Style = styleName == null ? defaultStyle : styles[styleName];
            row.Cells = xTr.Elements("td").Select(xTd => Cell.FromXElement(xTd, styles, row.Style)).ToList();

            return row;
        }
    }

    class Cell
    {
        public Span[] Text;

        public SpanInfo SpanInfo;
        public int ColumnIndex;

        public Style Style;

        public int ColSpan
        {
            get { return SpanInfo?.ColumnCount ?? 1; }
        }

        public int RowSpan
        {
            get { return SpanInfo?.RowCount ?? 1; }
        }

        internal static Cell FromXElement(XElement xTd, StyleSheet styles, Style defaultStyle)
        {
            var cell = new Cell();

            var rowSpan = (int?)xTd.Attribute("rowspan") ?? 1;
            var colSpan = (int?)xTd.Attribute("colspan") ?? 1;
            if (rowSpan > 1 || colSpan > 1)
                cell.SpanInfo = new SpanInfo() { ColumnCount = colSpan, RowCount = rowSpan };
            var styleName = (string)xTd.Attribute("style");

            cell.Style = styleName == null ? defaultStyle : styles[styleName];

            var lines = new List<Span>();
            Span lastSpan = null;
            foreach (var n in xTd.Nodes())
            {
                if (n.NodeType == System.Xml.XmlNodeType.Comment)
                {

                }
                else if (n.NodeType == System.Xml.XmlNodeType.Text)
                {
                    lastSpan = new TextSpan() { Text = ((XText)n).Value.Trim('\n', ' ') };
                    lines.Add(lastSpan);
                }
                else if (n.NodeType == System.Xml.XmlNodeType.Element)
                {
                    var x = ((XElement)n);
                    if (x.Name.LocalName == "br")
                    {
                        if (lastSpan != null && !lastSpan.NewLineAfter)
                            lastSpan.NewLineAfter = true;
                        else
                        {
                            lastSpan = new TextSpan() {  Text = "", NewLineAfter = true };
                            lines.Add(lastSpan);
                        }
                    }
                    else if (x.Name.LocalName == "span")
                    {
                        lastSpan = TextSpan.FromXElement(x, cell.Style.FontName, cell.Style.FontSize);
                        lines.Add(lastSpan);
                    }
                    else  if (x.Name.LocalName == "qrcode")
                    {
                        lastSpan = QrSpan.FromXElement(x);
                        lines.Add(lastSpan);
                    }
                    else
                        throw new NotImplementedException();
                }
                else
                    throw new NotImplementedException();
            }

            cell.Text = lines.ToArray();

            return cell;
        }

        public override string ToString()
        {
            return "Cell: " + ColumnIndex + (SpanInfo == null ? "" : " " + SpanInfo.ColumnCount + ", " + SpanInfo.RowCount);
        }
    }

    abstract class Span
    {
        public bool NewLineAfter;
    }

    class TextSpan : Span
    {
        public string Text;
        public string FontName;
        public bool IsBold;
        public double FontSize;

        public static Span FromXElement(XElement x, string defaultFontName, double defaultFontSize)
        {
            var res = new TextSpan();
            res.Text = x.Value;
            res.FontName = (string)x.Attribute("font-name") ?? defaultFontName;
            res.FontSize = (int?)x.Attribute("font-size") ?? defaultFontSize;
            res.IsBold = ((string)x.Attribute("font-weight") ?? "").ToLower() == "bold";
            return res;
        }
    }

    class QrSpan : Span
    {
        public double Width;
        public double Height;
        public System.IO.MemoryStream Data;

        internal static QrSpan FromXElement(XElement x)
        {
            var res = new QrSpan();
            res.Width = Helpers.ParseSize((string)x.Attribute("width"));
            res.Height = Helpers.ParseSize((string)x.Attribute("height"));

            var r = (string)x.Attribute("data");
            using (var q = new QRCoder.QRCodeGenerator())
            using (var qrcode = q.CreateQrCode(r, QRCoder.QRCodeGenerator.ECCLevel.L, forceUtf8: true, utf8BOM: true))
            using (var qc = new QRCoder.QRCode(qrcode))
            {
                var g = qc.GetGraphic(5);
                res.Data = new System.IO.MemoryStream();
                g.Save(res.Data, System.Drawing.Imaging.ImageFormat.Png);
                res.Data.Flush();
                res.Data.Seek(0, System.IO.SeekOrigin.Begin);
            }
            return res;
        }
    }
}

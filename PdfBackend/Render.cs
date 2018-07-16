using PdfSharp.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace PdfBackend
{
    class Render
    {
        private static List<Block> FormatParagraph(IEnumerable<Span> spans, XGraphics gfx, XFont defaultFont, Align align, double width)
        {
            if (!spans.Any())
                return new List<Block>(0);

            var res = new List<List<Block>>();
            var currentLine = new List<Block>();
            res.Add(currentLine);
            var x = 0.0d;

            foreach (var span in spans)
            {
                var imageSpan = span as QrSpan;
                var textSpan = span as TextSpan;

                if (imageSpan != null)
                {
                    if ((width - x) < imageSpan.Width)
                    {
                        x = 0;
                        currentLine = new List<Block>();
                        res.Add(currentLine);
                    }

                    currentLine.Add(new ImageBlock()
                    {
                        Height = imageSpan.Height,
                        Width = imageSpan.Width,
                        Image = XImage.FromStream(imageSpan.Data)
                    });
                    x += imageSpan.Width;

                    if (imageSpan.NewLineAfter)
                    {
                        x = 0;
                        currentLine = new List<Block>();
                        res.Add(currentLine);
                    }
                }
                else if (textSpan != null)
                {
                    var f = textSpan.FontName == null ? defaultFont : FontManager.GetFont(textSpan.FontName, textSpan.FontSize, textSpan.IsBold);
                    var text = textSpan.Text;

                    while (text != null)
                    {

                        var w = Helpers.WrapLine(gfx, f, width - x, text);
                        if (string.IsNullOrWhiteSpace(w.Item1) && x == 0)
                        {
                            w = Helpers.BreakLine(gfx, f, width, text);
                        }
                        if (!string.IsNullOrWhiteSpace(w.Item1) || textSpan.NewLineAfter)
                        {
                            var sz = gfx.MeasureString(w.Item1, f);
                            currentLine.Add(new TextBlock() { Height = f.GetHeight(), Width = sz.Width, X = x, Font = f, Text = w.Item1 });
                            x += sz.Width;
                        }

                        text = w.Item2;
                        if (text != null || textSpan.NewLineAfter)
                        {
                            x = 0;
                            currentLine = new List<Block>();
                            res.Add(currentLine);
                        }
                    }
                }
            }
                        
            var y = 0.0d;
            foreach (var line in res.Where(l => l.Any()))
            {
                var lineHeight = line.Max(block => block.Height);
                var maxAscent = line.Max(block => block.Ascent);
                x = 0;
                foreach (var block in line)
                {
                    block.Y = y + maxAscent;
                    block.X = x;
                    x += block.Width;
                }
                y += (lineHeight * 1.05);
                UpdateBlockAlignment(align, width, line);
            }

            return res.SelectMany(a=>a).ToList();
        }

        private static void UpdateBlockAlignment(Align align, double width, List<Block> currentLine)
        {
            if (align != Align.Near)
            {
                var lineWidth = currentLine.Sum(b => b.Width);
                var leftPadding = align == Align.Far ? width - lineWidth : (width - lineWidth) / 2;
                foreach (var b in currentLine)
                {
                    b.X += leftPadding;
                }
            }
        }

        public static void MeasureSheet(XGraphics gfx, Sheet t)
        {
            double left = t.MarginLeft;
            foreach (var c in t.Columns)
            {
                c.Left = left;
                left += c.Width;
            }

            var spans = new HashSet<SpanInfo>();
            int rowIndex = 0;
            foreach (var r in t.Rows)
            {
                foreach (var s in spans.ToArray())
                    if (s.LastRow < rowIndex)
                        spans.Remove(s);

                var maxH = 0d;
                int colIndex = 0;
                foreach (var c in r.Cells)
                {
                    foreach (var s in spans)
                    {
                        if (s.StartColumn <= colIndex && colIndex <= s.LastColumn)
                            colIndex += s.ColumnCount;
                    }

                    c.ColumnIndex = colIndex;

                    if (c.SpanInfo != null)
                    {
                        c.SpanInfo.StartRow = rowIndex;
                        c.SpanInfo.StartColumn = colIndex;
                        if (c.SpanInfo.RowCount > 1)
                            spans.Add(c.SpanInfo);
                        colIndex += c.SpanInfo.ColumnCount - 1;
                    }
                    else
                    {
                        maxH = Math.Max(maxH, MeasureCellHeight(gfx, c, t.Columns[colIndex].Width));
                    }
                    colIndex++;
                }
                r.RenderHeight = r.Height > 0 ? r.Height : maxH;
                r.RowIndex = rowIndex;
                rowIndex++;
            }
        }

        private static double MeasureCellHeight(XGraphics gfx, Cell c, double width)
        {
            var font = GetFont(c.Style);
            var b = FormatParagraph(c.Text, gfx, font, c.Style.HorizontalAlign, width - c.Style.MarginLeft - c.Style.MarginRight);
            return getFormattedTextHeight(b) + c.Style.MarginTop + c.Style.MarginBottom;
        }

        private static double getFormattedTextHeight(List<Block> paragraph)
        {
            return paragraph.Count > 0 
                ? paragraph.Max(z => z.Y + z.Descent)
                : 0;
        }

        static Dictionary<Style, XFont> fontCache = new Dictionary<Style, XFont>();

        private static XFont GetFont(Style style)
        {
            XFont font;
            if (!fontCache.TryGetValue(style, out font))
            {
                font = new XFont(style.FontName, style.FontSize);
                fontCache.Add(style, font);
            }
            return font;
        }

        private static List<string> WrapWords(XGraphics gfx, XFont font, double width, string input)
        {
            var words = Regex.Matches(input, @"(\w+)(\W*)").Cast<Match>().Select(m => m.Value).ToArray();
            var lines = new List<string>();
            var line = "";
            foreach (var w in words)
            {
                var m = gfx.MeasureString(line + w, font);
                if (m.Width > width)
                {
                    if (!string.IsNullOrWhiteSpace(line))
                        lines.Add(line);
                    line = w;
                }
                else
                {
                    line += w;
                }
            }
            lines.Add(line);
            return lines;
        }

        public static void RenderSheet(XGraphics gfx, Sheet t)
        {
            //gfx.DrawString("Hello, World!", font, XBrushes.Black, new XRect(0, 0, page.Width, page.Height), XStringFormats.BottomCenter);
            //gfx.DrawRectangle(XPens.Black, XBrushes.White, 0, 0, 100, 100);

            double y = 0;
            foreach (var r in t.Rows)
            {
                foreach (var c in r.Cells)
                {
                    RenderCell(gfx, t, r, c, y + t.MarginTop);
                }
                y += r.RenderHeight;
            }

        }

        static void RenderBorder(XGraphics gfx, double x, double y, double dx, double dy, Style s, int bIndex)
        {
            var width = s.Borders[bIndex];
            if (width > 0)
            {
                var pen = new XPen(XColors.Black, width);
                gfx.DrawLine(pen, x, y, dx, dy);
            }
        }

        static void RenderBorder(XGraphics gfx, double x, double y, double w, double h, Style style)
        {
            var firstBorder = style.Borders[0];
            var allEqual = style.Borders.All(b => b == firstBorder);

            if (allEqual)
            {
                if (firstBorder > 0)
                    gfx.DrawRectangle(new XPen(XColors.Black, firstBorder), x, y, w, h);
            }
            else
            {
                RenderBorder(gfx, x, y, x, y + h, style, 0);
                RenderBorder(gfx, x, y, x + w, y, style, 1);
                RenderBorder(gfx, x + w, y, x + w, y + h, style, 2);
                RenderBorder(gfx, x, y + h, x + w, y + h, style, 3);
            }
        }

        private static void RenderCell(XGraphics gfx, Sheet s, Row r, Cell c, double y)
        {
            var x = s.Columns[c.ColumnIndex].Left;
            var w = s.GetColumnPosition(c.ColumnIndex + (c.SpanInfo?.ColumnCount ?? 1)) - s.GetColumnPosition(c.ColumnIndex);
            var h = s.Rows.Skip(r.RowIndex).Take(c.SpanInfo?.RowCount ?? 1).Sum(rr => rr.RenderHeight);
            RenderBorder(gfx, x, y, w, h, c.Style);

            var font = GetFont(c.Style);
            var blocks = FormatParagraph(c.Text, gfx, font, c.Style.HorizontalAlign, w - c.Style.MarginLeft - c.Style.MarginRight);
            var th = getFormattedTextHeight(blocks);

            double yy;
            switch (c.Style.VerticalAlign)
            {
                case Align.Near: yy = y + c.Style.MarginTop; break;
                case Align.Center: yy = y + (h - c.Style.MarginTop - c.Style.MarginBottom) / 2 - th / 2 + c.Style.MarginTop; break;
                case Align.Far: yy = y + h - c.Style.MarginBottom - th; break;
                default: throw new NotImplementedException();
            }
            var format = new XStringFormat() { Alignment = XStringAlignment.Near, LineAlignment = XLineAlignment.BaseLine };
            var state = gfx.Save();
            gfx.IntersectClip(new XRect(x, y, w, h));

            foreach (var b in blocks)
            {
                var tb = b as TextBlock;
                if (tb != null)
                {
                    gfx.DrawString(tb.Text, tb.Font, XBrushes.Black, b.X + x + c.Style.MarginLeft, b.Y + yy, format);
                    continue;
                }
                var ib = b as ImageBlock;
                if (ib != null)
                {
                    gfx.DrawImage(ib.Image, 
                        ib.X + x + c.Style.MarginLeft, ib.Y + yy - ib.Height, ib.Width, ib.Height
                    );
                    continue;
                }
            }
            gfx.Restore(state);
        }
    }
}

namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using iTextSharp.text.html;
    using iTextSharp.text.pdf;
    using System;
    using System.Collections;
    using System.Globalization;
    using System.util;

    public class IncCell : ITextElementArray, IElement
    {
        private PdfPCell cell = new PdfPCell();
        private ArrayList chunks = new ArrayList();

        public IncCell(string tag, ChainedProperties props)
        {
            string s = props["colspan"];
            if (s != null)
            {
                this.cell.Colspan = int.Parse(s);
            }
            s = props["align"];
            if (tag.Equals("th"))
            {
                this.cell.HorizontalAlignment = 1;
            }
            if (s != null)
            {
                if (Util.EqualsIgnoreCase(s, "center"))
                {
                    this.cell.HorizontalAlignment = 1;
                }
                else if (Util.EqualsIgnoreCase(s, "right"))
                {
                    this.cell.HorizontalAlignment = 2;
                }
                else if (Util.EqualsIgnoreCase(s, "left"))
                {
                    this.cell.HorizontalAlignment = 0;
                }
                else if (Util.EqualsIgnoreCase(s, "justify"))
                {
                    this.cell.HorizontalAlignment = 3;
                }
            }
            s = props["valign"];
            this.cell.VerticalAlignment = 5;
            if (s != null)
            {
                if (Util.EqualsIgnoreCase(s, "top"))
                {
                    this.cell.VerticalAlignment = 4;
                }
                else if (Util.EqualsIgnoreCase(s, "bottom"))
                {
                    this.cell.VerticalAlignment = 6;
                }
            }
            s = props["border"];
            float num = 0f;
            if (s != null)
            {
                num = float.Parse(s, NumberFormatInfo.InvariantInfo);
            }
            this.cell.BorderWidth = num;
            s = props["cellpadding"];
            if (s != null)
            {
                this.cell.Padding = float.Parse(s, NumberFormatInfo.InvariantInfo);
            }
            this.cell.UseDescender = true;
            s = props["bgcolor"];
            this.cell.BackgroundColor = Markup.DecodeColor(s);
        }

        public bool Add(object o)
        {
            if (!(o is IElement))
            {
                return false;
            }
            this.cell.AddElement((IElement) o);
            return true;
        }

        public bool IsContent()
        {
            return true;
        }

        public bool IsNestable()
        {
            return true;
        }

        public bool Process(IElementListener listener)
        {
            return true;
        }

        public override string ToString()
        {
            return base.ToString();
        }

        public PdfPCell Cell
        {
            get
            {
                return this.cell;
            }
        }

        public ArrayList Chunks
        {
            get
            {
                return this.chunks;
            }
        }

        public int Type
        {
            get
            {
                return 30;
            }
        }
    }
}


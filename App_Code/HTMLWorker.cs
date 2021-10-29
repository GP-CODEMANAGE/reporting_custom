namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using iTextSharp.text.html;
    using iTextSharp.text.pdf;
    using iTextSharp.text.pdf.draw;
    using iTextSharp.text.xml.simpleparser;
    using System;
    using System.Collections;
    using System.Globalization;
    using System.IO;
    using System.Text;
    using System.util;
    using System.Web;

    public class HTMLWorker : ISimpleXMLDocHandler, IDocListener, IElementListener
    {
        private ChainedProperties cprops = new ChainedProperties();
        private Paragraph currentParagraph;
        protected IDocListener document;
        private FactoryProperties factoryProperties = new FactoryProperties();
        private Hashtable interfaceProps;
        private bool isPRE = false;
        protected ArrayList objectList;
        private bool pendingLI = false;
        private bool pendingTD = false;
        private bool pendingTR = false;
        private bool skipText = false;
        private Stack stack = new Stack();
        private StyleSheet style = new StyleSheet();
        private Stack tableState = new Stack();
        public static Hashtable tagsSupported = new Hashtable();
        public const string tagsSupportedString = "ol ul li a pre font span br p div body table td th tr i b u sub sup em strong s strike h1 h2 h3 h4 h5 h6 img hr";

        static HTMLWorker()
        {
            StringTokenizer tokenizer = new StringTokenizer("ol ul li a pre font span br p div body table td th tr i b u sub sup em strong s strike h1 h2 h3 h4 h5 h6 img hr");
            while (tokenizer.HasMoreTokens())
            {
                tagsSupported[tokenizer.NextToken()] = null;
            }
        }

        public HTMLWorker(IDocListener document)
        {
            this.document = document;
        }

        public bool Add(IElement element)
        {
            this.objectList.Add(element);
            return true;
        }

        public void ClearTextWrap()
        {
        }

        public void Close()
        {
        }

        public virtual void EndDocument()
        {
            foreach (IElement element in this.stack)
            {
                this.document.Add(element);
            }
            if (this.currentParagraph != null)
            {
                this.document.Add(this.currentParagraph);
            }
            this.currentParagraph = null;
        }

        public virtual void EndElement(string tag)
        {
            if (tagsSupported.ContainsKey(tag))
            {
                string key = (string) FactoryProperties.followTags[tag];
                if (key != null)
                {
                    this.cprops.RemoveChain(key);
                }
                else if (tag.Equals("font") || tag.Equals("span"))
                {
                    this.cprops.RemoveChain(tag);
                }
                else if (tag.Equals("a"))
                {
                    if (this.currentParagraph == null)
                    {
                        this.currentParagraph = new Paragraph();
                    }
                    IALink link = null;
                    bool flag = false;
                    if (this.interfaceProps != null)
                    {
                        link = (IALink) this.interfaceProps["alink_interface"];
                        if (link != null)
                        {
                            flag = link.Process(this.currentParagraph, this.cprops);
                        }
                    }
                    if (!flag)
                    {
                        string url = this.cprops["href"];
                        if (url != null)
                        {
                            ArrayList chunks = this.currentParagraph.Chunks;
                            for (int i = 0; i < chunks.Count; i++)
                            {
                                ((Chunk) chunks[i]).SetAnchor(url);
                            }
                        }
                    }
                    Paragraph paragraph = (Paragraph) this.stack.Pop();
                    Phrase o = new Phrase();
                    o.Add(this.currentParagraph);
                    paragraph.Add(o);
                    this.currentParagraph = paragraph;
                    this.cprops.RemoveChain("a");
                }
                else if (!tag.Equals("br"))
                {
                    if (this.currentParagraph != null)
                    {
                        if (this.stack.Count == 0)
                        {
                            this.document.Add(this.currentParagraph);
                        }
                        else
                        {
                            object obj2 = this.stack.Pop();
                            if (obj2 is ITextElementArray)
                            {
                                ((ITextElementArray) obj2).Add(this.currentParagraph);
                            }
                            this.stack.Push(obj2);
                        }
                    }
                    this.currentParagraph = null;
                    if (tag.Equals("ul") || tag.Equals("ol"))
                    {
                        if (this.pendingLI)
                        {
                            this.EndElement("li");
                        }
                        this.skipText = false;
                        this.cprops.RemoveChain(tag);
                        if (this.stack.Count != 0)
                        {
                            object obj3 = this.stack.Pop();
                            if (!(obj3 is List))
                            {
                                this.stack.Push(obj3);
                            }
                            else if (this.stack.Count == 0)
                            {
                                this.document.Add((IElement) obj3);
                            }
                            else
                            {
                                ((ITextElementArray) this.stack.Peek()).Add(obj3);
                            }
                        }
                    }
                    else if (tag.Equals("li"))
                    {
                        this.pendingLI = false;
                        this.skipText = true;
                        this.cprops.RemoveChain(tag);
                        if (this.stack.Count != 0)
                        {
                            object obj4 = this.stack.Pop();
                            if (!(obj4 is ListItem))
                            {
                                this.stack.Push(obj4);
                            }
                            else if (this.stack.Count == 0)
                            {
                                this.document.Add((IElement) obj4);
                            }
                            else
                            {
                                object obj5 = this.stack.Pop();
                                if (!(obj5 is List))
                                {
                                    this.stack.Push(obj5);
                                }
                                else
                                {
                                    ListItem item = (ListItem) obj4;
                                    ((List) obj5).Add(item);
                                    ArrayList list2 = item.Chunks;
                                    if (list2.Count > 0)
                                    {
                                        item.ListSymbol.Font = ((Chunk) list2[0]).Font;
                                    }
                                    this.stack.Push(obj5);
                                }
                            }
                        }
                    }
                    else if (tag.Equals("div") || tag.Equals("body"))
                    {
                        this.cprops.RemoveChain(tag);
                    }
                    else if (tag.Equals("pre"))
                    {
                        this.cprops.RemoveChain(tag);
                        this.isPRE = false;
                    }
                    else if (tag.Equals("p"))
                    {
                        this.cprops.RemoveChain(tag);
                    }
                    else if (((tag.Equals("h1") || tag.Equals("h2")) || (tag.Equals("h3") || tag.Equals("h4"))) || (tag.Equals("h5") || tag.Equals("h6")))
                    {
                        this.cprops.RemoveChain(tag);
                    }
                    else if (tag.Equals("table"))
                    {
                        if (this.pendingTR)
                        {
                            this.EndElement("tr");
                        }
                        this.cprops.RemoveChain("table");
                        PdfPTable element = ((IncTable) this.stack.Pop()).BuildTable();
                        element.SplitRows = true;
                        if (this.stack.Count == 0)
                        {
                            this.document.Add(element);
                        }
                        else
                        {
                            ((ITextElementArray) this.stack.Peek()).Add(element);
                        }
                        bool[] flagArray = (bool[]) this.tableState.Pop();
                        this.pendingTR = flagArray[0];
                        this.pendingTD = flagArray[1];
                        this.skipText = false;
                    }
                    else if (tag.Equals("tr"))
                    {
                        object obj6;
                        if (this.pendingTD)
                        {
                            this.EndElement("td");
                        }
                        this.pendingTR = false;
                        this.cprops.RemoveChain("tr");
                        ArrayList ncols = new ArrayList();
                        IncTable table3 = null;
                        do
                        {
                            obj6 = this.stack.Pop();
                            if (obj6 is IncCell)
                            {
                                ncols.Add(((IncCell) obj6).Cell);
                            }
                        }
                        while (!(obj6 is IncTable));
                        table3 = (IncTable) obj6;
                        table3.AddCols(ncols);
                        table3.EndRow();
                        this.stack.Push(table3);
                        this.skipText = true;
                    }
                    else if (tag.Equals("td") || tag.Equals("th"))
                    {
                        this.pendingTD = false;
                        this.cprops.RemoveChain("td");
                        this.skipText = true;
                    }
                }
            }
        }

        public bool NewPage()
        {
            return true;
        }

        public void Open()
        {
        }

        public void Parse(TextReader reader)
        {
            SimpleXMLParser.Parse(this, null, reader, true);
        }

        public static ArrayList ParseToList(TextReader reader, StyleSheet style)
        {
            return ParseToList(reader, style, null);
        }

        public static ArrayList ParseToList(TextReader reader, StyleSheet style, Hashtable interfaceProps)
        {
            HTMLWorker worker = new HTMLWorker(null);
            if (style != null)
            {
                worker.Style = style;
            }
            worker.document = worker;
            worker.InterfaceProps = interfaceProps;
            worker.objectList = new ArrayList();
            worker.Parse(reader);
            return worker.objectList;
        }

        public void ResetFooter()
        {
        }

        public void ResetHeader()
        {
        }

        public void ResetPageCount()
        {
        }

        public bool SetMarginMirroring(bool marginMirroring)
        {
            return false;
        }

        public bool SetMarginMirroringTopBottom(bool marginMirroring)
        {
            return false;
        }

        public bool SetMargins(float marginLeft, float marginRight, float marginTop, float marginBottom)
        {
            return true;
        }

        public bool SetPageSize(Rectangle pageSize)
        {
            return true;
        }

        public virtual void StartDocument()
        {
            Hashtable props = new Hashtable();
            this.style.ApplyStyle("body", props);
            this.cprops.AddToChain("body", props);
        }

        public virtual void StartElement(string tag, Hashtable h)
        {
            if (tagsSupported.ContainsKey(tag))
            {
                this.style.ApplyStyle(tag, h);
                string key = (string) FactoryProperties.followTags[tag];
                if (key != null)
                {
                    Hashtable prop = new Hashtable();
                    prop[key] = null;
                    this.cprops.AddToChain(key, prop);
                }
                else
                {
                    FactoryProperties.InsertStyle(h, this.cprops);
                    if (tag.Equals("a"))
                    {
                        this.cprops.AddToChain(tag, h);
                        if (this.currentParagraph == null)
                        {
                            this.currentParagraph = new Paragraph();
                        }
                        this.stack.Push(this.currentParagraph);
                        this.currentParagraph = new Paragraph();
                    }
                    else if (tag.Equals("br"))
                    {
                        if (this.currentParagraph == null)
                        {
                            this.currentParagraph = new Paragraph();
                        }
                        this.currentParagraph.Add(this.factoryProperties.CreateChunk("\n", this.cprops));
                    }
                    else if (tag.Equals("hr"))
                    {
                        bool flag = true;
                        if (this.currentParagraph == null)
                        {
                            this.currentParagraph = new Paragraph();
                            flag = false;
                        }
                        if (flag)
                        {
                            int count = this.currentParagraph.Chunks.Count;
                            if ((count == 0) || ((Chunk) this.currentParagraph.Chunks[count - 1]).Content.EndsWith("\n"))
                            {
                                flag = false;
                            }
                        }
                        string str2 = (string) h["align"];
                        int align = 1;
                        if (str2 != null)
                        {
                            if (Util.EqualsIgnoreCase(str2, "left"))
                            {
                                align = 0;
                            }
                            if (Util.EqualsIgnoreCase(str2, "right"))
                            {
                                align = 2;
                            }
                        }
                        string str = (string) h["width"];
                        float percentage = 1f;
                        if (str != null)
                        {
                            float num4 = Markup.ParseLength(str, 12f);
                            if (num4 > 0f)
                            {
                                percentage = num4;
                            }
                            if (!str.EndsWith("%"))
                            {
                                percentage = 100f;
                            }
                        }
                        string str4 = (string) h["size"];
                        float lineWidth = 1f;
                        if (str4 != null)
                        {
                            float num6 = Markup.ParseLength(str4, 12f);
                            if (num6 > 0f)
                            {
                                lineWidth = num6;
                            }
                        }
                        if (flag)
                        {
                            this.currentParagraph.Add(Chunk.NEWLINE);
                        }
                        this.currentParagraph.Add(new LineSeparator(lineWidth, percentage, null, align, this.currentParagraph.Leading / 2f));
                        this.currentParagraph.Add(Chunk.NEWLINE);
                    }
                    else if (tag.Equals("font") || tag.Equals("span"))
                    {
                        this.cprops.AddToChain(tag, h);
                    }
                    else if (tag.Equals("img"))
                    {
                        string src = (string) h["src"];
                        if (src != null)
                        {
                            this.cprops.AddToChain(tag, h);
                            Image instance = null;
                            if (this.interfaceProps != null)
                            {
                                IImageProvider provider = (IImageProvider) this.interfaceProps["img_provider"];
                                if (provider != null)
                                {
                                    instance = provider.GetImage(src, h, this.cprops, this.document);
                                }
                                if (instance == null)
                                {
                                    Hashtable hashtable2 = (Hashtable) this.interfaceProps["img_static"];
                                    if (hashtable2 != null)
                                    {
                                        Image image2 = (Image) hashtable2[src];
                                        if (image2 != null)
                                        {
                                            instance = Image.GetInstance(image2);
                                        }
                                    }
                                    else if (!src.StartsWith("http"))
                                    {
                                        string str6 = (string) this.interfaceProps["img_baseurl"];
                                        if (str6 != null)
                                        {
                                            src = str6 + src;
                                            instance = Image.GetInstance(src);
                                        }
                                    }
                                }
                            }
                            if (instance == null)
                            {
                                if (!src.StartsWith("http"))
                                {
                                    string str7 = this.cprops["image_path"];
                                    if (str7 == null)
                                    {
                                        str7 = "";
                                    }
                                    src = Path.Combine(str7, src);
                                }
                                instance = Image.GetInstance(src);
                            }
                            string str8 = (string) h["align"];
                            string str9 = (string) h["width"];
                            string str10 = (string) h["height"];
                            string s = this.cprops["before"];
                            string str12 = this.cprops["after"];
                            if (s != null)
                            {
                                instance.SpacingBefore = float.Parse(s, NumberFormatInfo.InvariantInfo);
                            }
                            if (str12 != null)
                            {
                                instance.SpacingAfter = float.Parse(str12, NumberFormatInfo.InvariantInfo);
                            }
                            float actualFontSize = Markup.ParseLength(this.cprops["size"], 12f);
                            if (actualFontSize <= 0f)
                            {
                                actualFontSize = 12f;
                            }
                            float newWidth = Markup.ParseLength(str9, actualFontSize);
                            float newHeight = Markup.ParseLength(str10, actualFontSize);
                            if ((newWidth > 0f) && (newHeight > 0f))
                            {
                                instance.ScaleAbsolute(newWidth, newHeight);
                            }
                            else if (newWidth > 0f)
                            {
                                newHeight = (instance.Height * newWidth) / instance.Width;
                                instance.ScaleAbsolute(newWidth, newHeight);
                            }
                            else if (newHeight > 0f)
                            {
                                newWidth = (instance.Width * newHeight) / instance.Height;
                                instance.ScaleAbsolute(newWidth, newHeight);
                            }
                            instance.WidthPercentage = 0f;
                            if (str8 != null)
                            {
                                this.EndElement("p");
                                int num10 = 1;
                                if (Util.EqualsIgnoreCase(str8, "left"))
                                {
                                    num10 = 0;
                                }
                                else if (Util.EqualsIgnoreCase(str8, "right"))
                                {
                                    num10 = 2;
                                }
                                instance.Alignment = num10;
                                IImg img = null;
                                bool flag2 = false;
                                if (this.interfaceProps != null)
                                {
                                    img = (IImg) this.interfaceProps["img_interface"];
                                    if (img != null)
                                    {
                                        flag2 = img.Process(instance, h, this.cprops, this.document);
                                    }
                                }
                                if (!flag2)
                                {
                                    this.document.Add(instance);
                                }
                                this.cprops.RemoveChain(tag);
                            }
                            else
                            {
                                this.cprops.RemoveChain(tag);
                                if (this.currentParagraph == null)
                                {
                                    this.currentParagraph = FactoryProperties.CreateParagraph(this.cprops);
                                }
                                this.currentParagraph.Add(new Chunk(instance, 0f, 0f));
                            }
                        }
                    }
                    else
                    {
                        this.EndElement("p");
                        if (((tag.Equals("h1") || tag.Equals("h2")) || (tag.Equals("h3") || tag.Equals("h4"))) || (tag.Equals("h5") || tag.Equals("h6")))
                        {
                            if (!h.ContainsKey("size"))
                            {
                                h["size"] = (7 - int.Parse(tag.Substring(1))).ToString();
                            }
                            this.cprops.AddToChain(tag, h);
                        }
                        else if (tag.Equals("ul"))
                        {
                            if (this.pendingLI)
                            {
                                this.EndElement("li");
                            }
                            this.skipText = true;
                            this.cprops.AddToChain(tag, h);
                            List list = new List(false);
                            try
                            {
                                list.IndentationLeft = float.Parse(this.cprops["indent"], NumberFormatInfo.InvariantInfo);
                            }
                            catch
                            {
                                list.Autoindent = true;
                            }
                            list.SetListSymbol("•");
                            this.stack.Push(list);
                        }
                        else if (tag.Equals("ol"))
                        {
                            if (this.pendingLI)
                            {
                                this.EndElement("li");
                            }
                            this.skipText = true;
                            this.cprops.AddToChain(tag, h);
                            List list2 = new List(true);
                            try
                            {
                                list2.IndentationLeft = float.Parse(this.cprops["indent"], NumberFormatInfo.InvariantInfo);
                            }
                            catch
                            {
                                list2.Autoindent = true;
                            }
                            this.stack.Push(list2);
                        }
                        else if (tag.Equals("li"))
                        {
                            if (this.pendingLI)
                            {
                                this.EndElement("li");
                            }
                            this.skipText = false;
                            this.pendingLI = true;
                            this.cprops.AddToChain(tag, h);
                            this.stack.Push(FactoryProperties.CreateListItem(this.cprops));
                        }
                        else if (tag.Equals("p"))
                        {
                            h["margin-left"] = "20px";
                            this.cprops.AddToChain(tag, h);
                        }
                        else if ((tag.Equals("div") || tag.Equals("body")))
                        {
                            this.cprops.AddToChain(tag, h);
                        }
                        else if (tag.Equals("pre"))
                        {
                            if (!h.ContainsKey("face"))
                            {
                                h["face"] = "Courier";
                            }
                            this.cprops.AddToChain(tag, h);
                            this.isPRE = true;
                        }
                        else if (tag.Equals("tr"))
                        {
                            if (this.pendingTR)
                            {
                                this.EndElement("tr");
                            }
                            this.skipText = true;
                            this.pendingTR = true;
                            this.cprops.AddToChain("tr", h);
                        }
                        else if (tag.Equals("td") || tag.Equals("th"))
                        {
                            if (this.pendingTD)
                            {
                                this.EndElement(tag);
                            }
                            this.skipText = false;
                            this.pendingTD = true;
                            this.cprops.AddToChain("td", h);
                            this.stack.Push(new IncCell(tag, this.cprops));
                        }
                        else if (tag.Equals("table"))
                        {
                            this.cprops.AddToChain("table", h);
                            IncTable table = new IncTable(h);
                            this.stack.Push(table);
                            this.tableState.Push(new bool[] { this.pendingTR, this.pendingTD });
                            this.pendingTR = this.pendingTD = false;
                            this.skipText = true;
                        }
                    }
                }
            }
        }

        public virtual void Text(string str)
        {
            if (!this.skipText)
            {
                string text = str;
                if (this.isPRE)
                {
                    if (this.currentParagraph == null)
                    {
                        this.currentParagraph = FactoryProperties.CreateParagraph(this.cprops);
                    }
                    this.currentParagraph.Add(this.factoryProperties.CreateChunk(text, this.cprops));
                }
                else if ((text.Length != 0) || (text.IndexOf(' ') >= 0))
                {
                    StringBuilder builder = new StringBuilder();
                    int length = text.Length;
                    bool flag = false;
                    for (int i = 0; i < length; i++)
                    {
                        char ch;
                        switch ((ch = text[i]))
                        {
                            case '\t':
                            case '\r':
                            {
                                continue;
                            }
                            case '\n':
                            {
                                if (i > 0)
                                {
                                    flag = true;
                                    builder.Append(' ');
                                }
                                continue;
                            }
                            case ' ':
                            {
                                if (!flag)
                                {
                                    builder.Append(ch);
                                }
                                continue;
                            }
                        }
                        flag = false;
                        builder.Append(ch);
                    }
                    if (this.currentParagraph == null)
                    {
                        this.currentParagraph = FactoryProperties.CreateParagraph(this.cprops);
                        //this.currentParagraph.Font = setFontsverdana();
                    }
                    this.currentParagraph.Add(this.factoryProperties.CreateChunk(builder.ToString(), this.cprops));
                }
            }
        }

        public HeaderFooter Footer
        {
            set
            {
            }
        }

        public HeaderFooter Header
        {
            set
            {
            }
        }

        public Hashtable InterfaceProps
        {
            get
            {
                return this.interfaceProps;
            }
            set
            {
                this.interfaceProps = value;
                FontFactoryImp imp = null;
                if (this.interfaceProps != null)
                {
                    imp = (FontFactoryImp) this.interfaceProps["font_factory"];
                }
                if (imp != null)
                {
                    this.factoryProperties.FontImp = imp;
                }
            }
        }

        public int PageCount
        {
            set
            {
            }
        }

        public StyleSheet Style
        {
            get
            {
                return this.style;
            }
            set
            {
                this.style = value;
            }
        }

        public iTextSharp.text.Font setFontsverdana()
        {
            #region WITH OLD FONTS FROM FRUTIGER
            //string fontpath = Server.MapPath(".");
            //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //if (bold == 1)
            //{
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
            //}
            //if (italic == 1)
            //{
            //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger_italic.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //}
            //if (bold == 1 && italic == 1)
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
            //return font; 
            #endregion

            #region WITH NEW FONTS FROM FRUTIGER
            string fontpath = HttpContext.Current.Server.MapPath(".");

            BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, 9, iTextSharp.text.Font.NORMAL);



            //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\d.ttf", BaseFont.CP1252, BaseFont.EMBEDDED);
            //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTR_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //if (bold == 1)
            //{
            //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBL____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLD);
            //}
            //if (italic == 1)
            //{
            //    //FTI_____.PFM
            //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTI_____.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //}
            //if (bold == 1 && italic == 1)
            //{
            //    customfont = BaseFont.CreateFont(fontpath + "\\Frutiger\\FTBLI___.PFM", BaseFont.CP1252, BaseFont.EMBEDDED);
            //    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            //    //font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.BOLDITALIC);
            //}

            return font;
            #endregion
        }
    }
}


namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using iTextSharp.text.html;
    using iTextSharp.text.pdf;
    using System;
    using System.Collections;
    using System.Globalization;
    using System.util;
    using System.Web;

    public class FactoryProperties
    {
        public static Hashtable followTags = new Hashtable();
        private FontFactoryImp fontImp = FontFactory.FontImp;

        static FactoryProperties()
        {
            followTags["i"] = "i";
            followTags["b"] = "b";
            followTags["u"] = "u";
            followTags["sub"] = "sub";
            followTags["sup"] = "sup";
            followTags["em"] = "i";
            followTags["strong"] = "b";
            followTags["s"] = "s";
            followTags["strike"] = "s";
        }

        public Chunk CreateChunk(string text, ChainedProperties props)
        {
            //string fontpath = HttpContext.Current.Server.MapPath(".");
            //BaseFont customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdana.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
            //iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, 9, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font font = this.GetFont(props);//setFontsverdana();// 
            float rise = font.Size / 2f;
            Chunk chunk = new Chunk(text,font);
            if (props.HasProperty("sub"))
            {
                chunk.SetTextRise(-rise);
            }
            else if (props.HasProperty("sup"))
            {
                chunk.SetTextRise(rise);
            }
            chunk.SetHyphenation(GetHyphenation(props));
            return chunk;
        }
        
      

        public static ListItem CreateListItem(ChainedProperties props)
        {
            ListItem p = new ListItem();
            CreateParagraph(p, props);
            return p;
        }

        public static Paragraph CreateParagraph(ChainedProperties props)
        {
            Paragraph p = new Paragraph();
            CreateParagraph(p, props);
            return p;
        }

        public static void CreateParagraph(Paragraph p, ChainedProperties props)
        {
            string str = props["align"];
            if (str != null)
            {
                if (Util.EqualsIgnoreCase(str, "center"))
                {
                    p.Alignment = 1;
                }
                else if (Util.EqualsIgnoreCase(str, "right"))
                {
                    p.Alignment = 2;
                }
                else if (Util.EqualsIgnoreCase(str, "justify"))
                {
                    p.Alignment = 3;
                }
            }
            p.Hyphenation = GetHyphenation(props);
            SetParagraphLeading(p, props["leading"]);
            str = props["before"];
            if (str != null)
            {
                try
                {
                    //p.FirstLineIndent = float.Parse(str, NumberFormatInfo.InvariantInfo);
                    p.SpacingBefore = float.Parse(str, NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                }
            }
            str = props["after"];
            if (str != null)
            {
                try
                {
                    p.SpacingAfter = float.Parse(str, NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                }
            }
            str = props["extraparaspace"];//text-indent
            if (str != null)
            {
                try
                {
                    p.ExtraParagraphSpace = float.Parse(str, NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                }
            }
            str = props["text-indent"];//text-indent
            if (str != null)
            {
                try
                {
                    p.Alignment = int.Parse(str, NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                }
            }

            str = props["margin-left"];//text-indent
            if (str != null)
            {
                try
                {
                    p.IndentationLeft = float.Parse(str, NumberFormatInfo.InvariantInfo);
                }
                catch
                {
                }
            }

        }

        public iTextSharp.text.Font GetFont(ChainedProperties props)
        {
            string str = props["face"];
            if (str == null || str != "Verdana")
            {
                str = "Verdana";
            }


            int istyle = 0, bstyle = 0, ustyle = 0, sstyle=0;
            if (props.HasProperty("i"))
            {
                istyle |= 1;
            }
            if (props.HasProperty("b"))
            {
                bstyle |= 1;
            }
            if (props.HasProperty("u"))
            {
                ustyle |= 1;
            }
            if (props.HasProperty("s"))
            {
                sstyle |= 1;
            }



            string s = "9";// props["size"];
            float size = 9f;
            if (s != null)
            {
                size = float.Parse(s, NumberFormatInfo.InvariantInfo);
            }
            Color color = Markup.DecodeColor(props["color"]);
            string encoding = props["encoding"];
            if (encoding == null)
            {
                encoding = "Cp1252";
            }

            if (str != null)
            {
                StringTokenizer tokenizer = new StringTokenizer(str, ",");
                while (tokenizer.HasMoreTokens())
                {
                    str = tokenizer.NextToken().Trim();
                    if (str.StartsWith("\""))
                    {
                        str = str.Substring(1);
                    }
                    if (str.EndsWith("\""))
                    {
                        str = str.Substring(0, str.Length - 1);
                    }
                    if (this.fontImp.IsRegistered(str))
                    {
                        break;
                    }
                }
            }

            return setFontsAll(9, bstyle, istyle, ustyle);// this.fontImp.GetFont(str, encoding, true, size, style, color);
        }

        public static IHyphenationEvent GetHyphenation(ChainedProperties props)
        {
            return GetHyphenation(props["hyphenation"]);
        }

        public static IHyphenationEvent GetHyphenation(Hashtable props)
        {
            return GetHyphenation((string)props["hyphenation"]);
        }

        public static IHyphenationEvent GetHyphenation(string s)
        {
            if ((s == null) || (s.Length == 0))
            {
                return null;
            }
            string lang = s;
            string country = null;
            int leftMin = 2;
            int rightMin = 2;
            int index = s.IndexOf('_');
            if (index != -1)
            {
                lang = s.Substring(0, index);
                country = s.Substring(index + 1);
                index = country.IndexOf(',');
                if (index == -1)
                {
                    return new HyphenationAuto(lang, country, leftMin, rightMin);
                }
                s = country.Substring(index + 1);
                country = country.Substring(0, index);
                index = s.IndexOf(',');
                if (index == -1)
                {
                    leftMin = int.Parse(s);
                }
                else
                {
                    leftMin = int.Parse(s.Substring(0, index));
                    rightMin = int.Parse(s.Substring(index + 1));
                }
            }
            return new HyphenationAuto(lang, country, leftMin, rightMin);
        }

        public static void InsertStyle(Hashtable h)
        {
            string str = (string)h["style"];
            if (str != null)
            {
                Properties properties = Markup.ParseAttributes(str);
                foreach (string str2 in properties.Keys)
                {
                    if (str2.Equals("font-family"))
                    {
                        h["face"] = "Verdana";
                        continue;
                    }
                    if (str2.Equals("font-size"))
                    {
                        h["size"] = Markup.ParseLength(properties[str2]).ToString(NumberFormatInfo.InvariantInfo) + "pt";
                        continue;
                    }
                    if (str2.Equals("font-style"))
                    {
                        string str3 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        if (str3.Equals("italic") || str3.Equals("oblique"))
                        {
                            h["i"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("font-weight"))
                    {
                        string str4 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        if ((str4.Equals("bold") || str4.Equals("700")) || (str4.Equals("800") || str4.Equals("900")))
                        {
                            h["b"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("text-decoration"))
                    {
                        if (properties[str2].Trim().ToLower(CultureInfo.InvariantCulture).Equals("underline"))
                        {
                            h["u"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("color"))
                    {
                        Color color = Markup.DecodeColor(properties[str2]);
                        if (color != null)
                        {
                            string str6 = "#" + ((color.ToArgb() & 0xffffff)).ToString("X06", NumberFormatInfo.InvariantInfo);
                            h["color"] = str6;
                        }
                        continue;
                    }
                    if (str2.Equals("line-height"))
                    {
                        string str7 = properties[str2].Trim();
                        float num2 = Markup.ParseLength(properties[str2]);
                        if (str7.EndsWith("%"))
                        {
                            num2 /= 100f;
                            h["leading"] = "0," + num2.ToString(NumberFormatInfo.InvariantInfo);
                        }
                        else if (Util.EqualsIgnoreCase("normal", str7))
                        {
                            h["leading"] = "0,1.5";
                        }
                        else
                        {
                            h["leading"] = num2.ToString(NumberFormatInfo.InvariantInfo) + ",0";
                        }
                        continue;
                    }
                    if (str2.Equals("text-align"))
                    {
                        string str8 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        h["align"] = str8;
                    }
                    else if (str2.Equals("padding-left"))
                    {
                        string str9 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        h["indent"] = str9;
                    }
                    else if (str2.Equals("text-indent"))
                    {
                        decimal dec;
                        string str10 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        str10 = str10.Replace("in", "").Replace("pt", "").Replace("-", "");
                        if (str10.Contains("."))
                        {
                            dec = Convert.ToDecimal(str10) + Convert.ToDecimal(50);
                        }
                        else
                        {
                            dec = Convert.ToDecimal(str10) + Convert.ToDecimal(50);
                        }
                        h["text-indent"] = Convert.ToString(dec);
                    }
                    else if (str2.Equals("margin"))
                    {
                        string str11 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        string[] strtest = str11.Split(' ');
                        h["margin-left"] = strtest[2];
                    }
                }
            }
        }

        public static void InsertStyle(Hashtable h, ChainedProperties cprops)
        {
            string str = (string)h["style"];
            if (str != null)
            {
                Properties properties = Markup.ParseAttributes(str);
                foreach (string str2 in properties.Keys)
                {
                    if (str2.Equals("font-family"))
                    {
                        h["face"] = "Verdana";// properties["Verdana"];
                        continue;
                    }
                    if (str2.Equals("font-size"))
                    {
                        float actualFontSize = Markup.ParseLength(cprops["size"], 12f);
                        if (actualFontSize <= 0f)
                        {
                            actualFontSize = 12f;
                        }
                        h["size"] = Markup.ParseLength(properties[str2], actualFontSize).ToString(NumberFormatInfo.InvariantInfo) + "pt";
                        continue;
                    }
                    if (str2.Equals("font-style"))
                    {
                        string str3 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        if (str3.Equals("italic") || str3.Equals("oblique"))
                        {
                            h["i"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("font-weight"))
                    {
                        string str4 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        if ((str4.Equals("bold") || str4.Equals("700")) || (str4.Equals("800") || str4.Equals("900")))
                        {
                            h["b"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("text-decoration"))
                    {
                        if (properties[str2].Trim().ToLower(CultureInfo.InvariantCulture).Equals("underline"))
                        {
                            h["u"] = null;
                        }
                        continue;
                    }
                    if (str2.Equals("color"))
                    {
                        Color color = Markup.DecodeColor(properties[str2]);
                        if (color != null)
                        {
                            string str6 = "#" + ((color.ToArgb() & 0xffffff)).ToString("X06", NumberFormatInfo.InvariantInfo);
                            h["color"] = str6;
                        }
                        continue;
                    }
                    if (str2.Equals("line-height"))
                    {
                        string str7 = properties[str2].Trim();
                        float num3 = Markup.ParseLength(cprops["size"], 12f);
                        if (num3 <= 0f)
                        {
                            num3 = 12f;
                        }
                        float num4 = Markup.ParseLength(properties[str2], num3);
                        if (str7.EndsWith("%"))
                        {
                            num4 /= 100f;
                            h["leading"] = "0," + num4.ToString(NumberFormatInfo.InvariantInfo);
                        }
                        else if (Util.EqualsIgnoreCase("normal", str7))
                        {
                            h["leading"] = "0,1.5";
                        }
                        else
                        {
                            h["leading"] = num4.ToString(NumberFormatInfo.InvariantInfo) + ",0";
                        }
                        continue;
                    }
                    if (str2.Equals("text-align"))
                    {
                        string str8 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        h["align"] = str8;
                    }
                    else if (str2.Equals("padding-left"))
                    {
                        string str9 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        h["indent"] = str9;
                    }
                    else if (str2.Equals("text-indent"))
                    {
                        decimal dec;
                        string str10 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        if (str10.Contains("in"))
                        {
                            str10 = str10.Replace("in", "").Replace("pt", "").Replace("auto", "0.0");//.Replace("-","");
                            str10 = Convert.ToString(Convert.ToDecimal(str10) * Convert.ToDecimal(72));
                        }
                        else
                        {
                            str10 = str10.Replace("in", "").Replace("pt", "").Replace("auto", "0.0");//.Replace("-","");
                        }
                        dec = Convert.ToDecimal(str10) + Convert.ToDecimal(15);
                        h["text-indent"] = Convert.ToString(dec);
                    }
                    else if (str2.Equals("margin"))
                    {
                        decimal dec0, dec1, dec2, dec3;
                        string str11 = properties[str2].Trim().ToLower(CultureInfo.InvariantCulture);
                        string[] strtest = str11.Split(' ');

                        for (int i = 0; i < strtest.Length; i++)
                        {
                            if (strtest[i].Contains("in"))
                            {
                                strtest[i] = strtest[i].Replace("in", "").Replace("pt", "").Replace("-", "").Replace("auto", "0.0");
                                dec0 = Convert.ToDecimal(strtest[i]) * Convert.ToDecimal(72);
                                dec0 = dec0 + Convert.ToDecimal(15);
                            }
                            else
                            {
                                strtest[i] = strtest[i].Replace("in", "").Replace("pt", "").Replace("-", "").Replace("auto","0.0");
                                dec0 = Convert.ToDecimal(strtest[i]) + Convert.ToDecimal(15);
                            }
                           
                            if (i == 0) { h["margin-top"] = Convert.ToString(dec0); }
                            if (i == 1) { h["margin-right"] = Convert.ToString(dec0); }
                            if (i == 2) { h["margin-bottom"] = Convert.ToString(dec0); }
                            if (i == 3) { h["margin-left"] = Convert.ToString(dec0); }

                            //dec0 = Convert.ToDecimal(strtest[i]) + Convert.ToDecimal(50);
                        }

                       
                    }
                  
                }
            }
        }

        private static void SetParagraphLeading(Paragraph p, string leading)
        {
            if (leading == null)
            {
                p.SetLeading(0f, 1.0f);
            }
            else
            {
                try
                {
                    StringTokenizer tokenizer = new StringTokenizer(leading, " ,");
                    float fixedLeading = float.Parse(tokenizer.NextToken(), NumberFormatInfo.InvariantInfo);
                    if (!tokenizer.HasMoreTokens())
                    {
                        p.SetLeading(fixedLeading, 0f);
                    }
                    else
                    {
                        float multipliedLeading = float.Parse(tokenizer.NextToken(), NumberFormatInfo.InvariantInfo);
                        p.SetLeading(fixedLeading, multipliedLeading);
                    }
                }
                catch
                {
                    p.SetLeading(0f, 1.5f);
                }
            }
        }

        public FontFactoryImp FontImp
        {
            get
            {
                return this.fontImp;
            }
            set
            {
                this.fontImp = value;
            }
        }


        public iTextSharp.text.Font setFontsAll(int size, int bold, int italic,int underline)
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
            iTextSharp.text.Font font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
            if (bold == 1)
            {
                customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanab.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
                font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
                if (underline == 1)
                {
                    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);
                }
            }
            if (italic == 1)
            {
                //FTI_____.PFM
                customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanai.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
                font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
                if (underline == 1)
                {
                    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);
                }
            }
            if (bold == 1 && italic == 1)
            {
                customfont = BaseFont.CreateFont(fontpath + "\\Verdana\\verdanaz.TTF", BaseFont.CP1252, BaseFont.EMBEDDED);
                font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.NORMAL);
                if (underline == 1)
                {
                    font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);
                }
            }

            if (underline == 1)
            {
                font = new iTextSharp.text.Font(customfont, size, iTextSharp.text.Font.UNDERLINE);
            }


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


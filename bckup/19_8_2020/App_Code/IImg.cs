namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using System;
    using System.Collections;

    public interface IImg
    {
        bool Process(Image img, Hashtable h, ChainedProperties cprops, IDocListener doc);
    }
}


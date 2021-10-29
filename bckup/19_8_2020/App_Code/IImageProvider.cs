namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using System;
    using System.Collections;

    public interface IImageProvider
    {
        Image GetImage(string src, Hashtable h, ChainedProperties cprops, IDocListener doc);
    }
}


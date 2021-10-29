namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text;
    using System;

    public interface IALink
    {
        bool Process(Paragraph current, ChainedProperties cprops);
    }
}


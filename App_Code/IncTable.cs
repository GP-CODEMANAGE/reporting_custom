namespace iTextSharp.text.html.simpleparser
{
    using iTextSharp.text.pdf;
    using System;
    using System.Collections;
    using System.Globalization;

    public class IncTable
    {
        private ArrayList cols;
        private Hashtable props = new Hashtable();
        private ArrayList rows = new ArrayList();

        public IncTable(Hashtable props)
        {
            foreach (DictionaryEntry entry in props)
            {
                this.props[entry.Key] = entry.Value;
            }
        }

        public void AddCol(PdfPCell cell)
        {
            if (this.cols == null)
            {
                this.cols = new ArrayList();
            }
            this.cols.Add(cell);
        }

        public void AddCols(ArrayList ncols)
        {
            if (this.cols == null)
            {
                this.cols = new ArrayList(ncols);
            }
            else
            {
                this.cols.AddRange(ncols);
            }
        }

        public PdfPTable BuildTable()
        {
            if (this.rows.Count == 0)
            {
                return new PdfPTable(1);
            }
            int numColumns = 0;
            ArrayList list = (ArrayList) this.rows[0];
            for (int i = 0; i < list.Count; i++)
            {
                numColumns += ((PdfPCell) list[i]).Colspan;
            }
            PdfPTable table = new PdfPTable(numColumns);
            string s = (string) this.props["width"];
            if (s == null)
            {
                table.WidthPercentage = 100f;
            }
            else if (s.EndsWith("%"))
            {
                table.WidthPercentage = float.Parse(s.Substring(0, s.Length - 1), NumberFormatInfo.InvariantInfo);
            }
            else
            {
                table.TotalWidth = float.Parse(s, NumberFormatInfo.InvariantInfo);
                table.LockedWidth = true;
            }
            for (int j = 0; j < this.rows.Count; j++)
            {
                ArrayList list2 = (ArrayList) this.rows[j];
                for (int k = 0; k < list2.Count; k++)
                {
                    table.AddCell((PdfPCell) list2[k]);
                }
            }
            return table;
        }

        public void EndRow()
        {
            if (this.cols != null)
            {
                this.cols.Reverse();
                this.rows.Add(this.cols);
                this.cols = null;
            }
        }

        public ArrayList Rows
        {
            get
            {
                return this.rows;
            }
        }
    }
}


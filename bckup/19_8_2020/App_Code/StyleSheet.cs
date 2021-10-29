namespace iTextSharp.text.html.simpleparser
{
    using System;
    using System.Collections;
    using System.Globalization;

    public class StyleSheet
    {
        public Hashtable classMap = new Hashtable();
        public Hashtable tagMap = new Hashtable();

        private void ApplyMap(Hashtable map, Hashtable props)
        {
        }

        public void ApplyStyle(string tag, Hashtable props)
        {
            Hashtable hashtable2;
            Hashtable d = (Hashtable) this.tagMap[tag.ToLower(CultureInfo.InvariantCulture)];
            if (d != null)
            {
                hashtable2 = new Hashtable(d);
                foreach (DictionaryEntry entry in props)
                {
                    hashtable2[entry.Key] = entry.Value;
                }
                foreach (DictionaryEntry entry2 in hashtable2)
                {
                    props[entry2.Key] = entry2.Value;
                }
            }
            string str = (string) props["class"];
            if (str != null)
            {
                d = (Hashtable) this.classMap[str.ToLower(CultureInfo.InvariantCulture)];
                if (d != null)
                {
                    props.Remove("class");
                    hashtable2 = new Hashtable(d);
                    foreach (DictionaryEntry entry3 in props)
                    {
                        hashtable2[entry3.Key] = entry3.Value;
                    }
                    foreach (DictionaryEntry entry4 in hashtable2)
                    {
                        props[entry4.Key] = entry4.Value;
                    }
                }
            }
        }

        public void LoadStyle(string style, Hashtable props)
        {
            this.classMap[style.ToLower(CultureInfo.InvariantCulture)] = props;
        }

        public void LoadStyle(string style, string key, string value)
        {
            style = style.ToLower(CultureInfo.InvariantCulture);
            Hashtable hashtable = (Hashtable) this.classMap[style];
            if (hashtable == null)
            {
                hashtable = new Hashtable();
                this.classMap[style] = hashtable;
            }
            hashtable[key] = value;
        }

        public void LoadTagStyle(string tag, Hashtable props)
        {
            this.tagMap[tag.ToLower(CultureInfo.InvariantCulture)] = props;
        }

        public void LoadTagStyle(string tag, string key, string value)
        {
            tag = tag.ToLower(CultureInfo.InvariantCulture);
            Hashtable hashtable = (Hashtable) this.tagMap[tag];
            if (hashtable == null)
            {
                hashtable = new Hashtable();
                this.tagMap[tag] = hashtable;
            }
            hashtable[key] = value;
        }
    }
}


namespace iTextSharp.text.html.simpleparser
{
    using System;
    using System.Collections;
    using System.Globalization;
    using System.Reflection;

    public class ChainedProperties
    {
        public ArrayList chain = new ArrayList();
        public static int[] fontSizes = new int[] {9};// { 8, 10, 12, 14, 0x12, 0x18, 0x24 };

        public void AddToChain(string key, Hashtable prop)
        {
            string s = (string) prop["size"];
            if (s != null)
            {
                if (s.EndsWith("pt"))
                {
                    prop["size"] = s.Substring(0, s.Length - 2);
                }
                else
                {
                    int index = 0;
                    if (!s.StartsWith("+") && !s.StartsWith("-"))
                    {
                        try
                        {
                            index = int.Parse(s) - 1;
                        }
                        catch
                        {
                            index = 0;
                        }
                    }
                    else
                    {
                        string str2 = this["basefontsize"];
                        if (str2 == null)
                        {
                            str2 = "12";
                        }
                        int num3 = (int) float.Parse(str2, NumberFormatInfo.InvariantInfo);
                        for (int i = fontSizes.Length - 1; i >= 0; i--)
                        {
                            if (num3 >= fontSizes[i])
                            {
                                index = i;
                                break;
                            }
                        }
                        int num5 = int.Parse(s.StartsWith("+") ? s.Substring(1) : s);
                        index += num5;
                    }
                    if (index < 0)
                    {
                        index = 0;
                    }
                    else if (index >= fontSizes.Length)
                    {
                        index = fontSizes.Length - 1;
                    }
                    prop["size"] = fontSizes[index].ToString();
                }
            }
            this.chain.Add(new object[] { key, prop });
        }

        public bool HasProperty(string key)
        {
            for (int i = this.chain.Count - 1; i >= 0; i--)
            {
                object[] objArray = (object[]) this.chain[i];
                Hashtable hashtable = (Hashtable) objArray[1];
                if (hashtable.ContainsKey(key))
                {
                    return true;
                }
            }
            return false;
        }

        public void RemoveChain(string key)
        {
            for (int i = this.chain.Count - 1; i >= 0; i--)
            {
                if (key.Equals(((object[]) this.chain[i])[0]))
                {
                    this.chain.RemoveAt(i);
                    return;
                }
            }
        }

        public string this[string key]
        {
            get
            {
                for (int i = this.chain.Count - 1; i >= 0; i--)
                {
                    object[] objArray = (object[]) this.chain[i];
                    Hashtable hashtable = (Hashtable) objArray[1];
                    string str = (string) hashtable[key];
                    if (str != null)
                    {
                        return str;
                    }
                }
                return null;
            }
        }
    }
}


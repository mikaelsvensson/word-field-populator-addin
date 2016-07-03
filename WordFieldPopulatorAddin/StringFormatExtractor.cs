using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace WordFieldPopulatorAddin
{
    internal class StringFormatExtractor : IFieldExtractor
    {
        public IList<string> Get(Document vstoDocument)
        {
            var values = new List<string>();
            var rng = vstoDocument.Content;
            var f = rng.Find;
            f.ClearFormatting();
            f.Forward = true;
            var sep = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            f.Text = "\\{[!}]{1" + sep + "}\\}";
            f.MatchWildcards = true;
            f.Execute();
            while (f.Found)
            {
                var key = GetKeyFromMatchRange(rng);
                values.Add(key);
                f.Execute(MatchWildcards: true);
            }
            return values;
        }

        public void Update(Document vstoDocument, IDictionary<string, string> values)
        {
            var rng = Find(vstoDocument);
            while (rng != null)
            {
                string key = GetKeyFromMatchRange(rng);
                if (values.ContainsKey(key))
                {
                    int numericValue;
                    rng.Text = String.Format(rng.Text.Replace(key, "0"), int.TryParse(values[key], out numericValue) ? (object)numericValue : values[key]);
                }

                rng = Find(vstoDocument);
            }
        }

        private static string GetKeyFromMatchRange(Range rng)
        {
            return rng.Text.Substring(1, rng.Text.IndexOfAny(new char[] { ',', ':', '}' }) - 1);
        }

        private Range Find(Document vstoDocument)
        {
            var rng = vstoDocument.Content;
            var f = rng.Find;
            f.ClearFormatting();
            f.Forward = true;
            var sep = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            f.Text = "\\{[!}]{1" + sep + "}\\}";
            f.MatchWildcards = true;
            f.Execute();
            return f.Found ? rng : null;
        }
    }
}
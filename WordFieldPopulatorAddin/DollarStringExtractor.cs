using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace WordFieldPopulatorAddin
{
    internal class DollarStringExtractor : IFieldExtractor
    {
        public IList<string> Get(Document vstoDocument)
        {
            var values = new List<string>();
            var rng = vstoDocument.Content;
            var f = rng.Find;
            f.ClearFormatting();
            f.Forward = true;
            var sep = System.Globalization.CultureInfo.CurrentCulture.TextInfo.ListSeparator;
            f.Text = "$[A-Z]{1" + sep + "}";

            f.MatchWildcards = true;
            f.Execute();
            while (f.Found)
            {
                System.Diagnostics.Debug.WriteLine(rng.Text);
                var key = rng.Text.Substring(1);
                values.Add(key);
                f.Execute(MatchWildcards: true);
            }
            return values;
        }

        public void Update(Document vstoDocument, IDictionary<string, string> values)
        {
            foreach (var item in values)
            {
                var finder = vstoDocument.Content.Find;
                finder.ClearFormatting();
                finder.Text = "$" + item.Key;
                finder.Replacement.ClearFormatting();
                finder.Replacement.Text = item.Value;
                var found = finder.Execute(Replace: WdReplace.wdReplaceAll);
            }

        }
    }
}
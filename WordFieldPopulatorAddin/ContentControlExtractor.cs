using System;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace WordFieldPopulatorAddin
{
    internal class ContentControlExtractor : IFieldExtractor
    {
        public IList<string> Get(Document vstoDocument)
        {
            var keys = new List<string>();
            foreach (ContentControl control in vstoDocument.ContentControls)
            {
                if (!String.IsNullOrEmpty(control.Tag))
                {
                    keys.Add(control.Tag);
                }
            }
            return keys;
        }

        public void Update(Document vstoDocument, IDictionary<string, string> values)
        {
            foreach (var item in values)
            {
                foreach (ContentControl control in vstoDocument.ContentControls)
                {
                    if (control.Tag == item.Key)
                    {
                        control.Range.Text = item.Value;
                    }
                }
            }
        }
    }
}
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;

namespace WordFieldPopulatorAddin
{
    internal interface IFieldExtractor
    {
        IList<string> Get(Document vstoDocument);
        void Update(Document vstoDocument, IDictionary<string, string> values);
    }
}
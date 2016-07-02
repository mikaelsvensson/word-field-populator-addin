using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.Serialization.Json;

namespace WordFieldPopulatorAddin
{
    public partial class FieldPopulatorRibbon
    {
        private void FieldPopulatorRibbon_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Document nativeDocument = Globals.ThisAddIn.Application.ActiveDocument;
            Microsoft.Office.Tools.Word.Document vstoDocument = Globals.Factory.GetVstoObject(nativeDocument);


            var extractors = new List<IFieldExtractor>(new IFieldExtractor[] { new DollarStringExtractor(), new ContentControlExtractor() });

            var clipDict = GetValuesFromClipboard();

            var keys = extractors.SelectMany(extractor => extractor.Get(nativeDocument));

            IDictionary<string, string> values = keys.ToDictionary(
                k => k,
                v => clipDict.ContainsKey(v) ? clipDict[v] : "");

            if (values.Count > 0)
            {
                var form = new FieldsForm(values);
                form.ShowDialog();

                extractors.ForEach(extractor => extractor.Update(nativeDocument, form.Values));
            }
            else
            {
                System.Windows.Forms.MessageBox.Show(
                    "Det finns inga $VARIABLER eller innehållskontroller för text i detta dokument.", 
                    "Inga variabler eller fält", 
                    System.Windows.Forms.MessageBoxButtons.OK, 
                    System.Windows.Forms.MessageBoxIcon.Information);
            }

        }

        private Dictionary<string, string> GetValuesFromClipboard()
        {
            String clipText = System.Windows.Forms.Clipboard.GetText(System.Windows.Forms.TextDataFormat.UnicodeText);

            try
            {
                // Try parsing clipboard as JSON
                var serSettings = new DataContractJsonSerializerSettings();
                serSettings.UseSimpleDictionaryFormat = true;
                var serializer = new DataContractJsonSerializer(typeof(Dictionary<string, string>), serSettings);
                Dictionary<string, string> result = (Dictionary<string, string>)serializer.ReadObject(new MemoryStream(Encoding.UTF8.GetBytes(clipText)));
                return result;
            }
            catch (System.Runtime.Serialization.SerializationException ex)
            {
                // Try parsing clipboard content as FIELD=VALUE;FIELD=VALUE;...
                return
                    !String.IsNullOrEmpty(clipText) && clipText.Contains('=')
                    ?
                    clipText.Split(';')
                        .Select(s => s.Split('='))
                        .ToDictionary(s => s[0], s => s[s.Length > 1 ? 1 : 0])
                    :
                    new Dictionary<string, string>();
            }
        }
    }
}

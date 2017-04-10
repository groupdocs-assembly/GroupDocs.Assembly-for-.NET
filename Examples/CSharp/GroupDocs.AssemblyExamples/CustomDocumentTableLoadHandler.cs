using GroupDocs.Assembly.Data;
using System;

namespace GroupDocs.AssemblyExamples
{
    public class CustomDocumentTableLoadHandler : IDocumentTableLoadHandler
    {
        public void Handle(DocumentTableLoadArgs args)
        {
            switch (args.TableIndex)
            {
                case 0:
                    // Do nothing. The table is to be loaded with default options.
                    break;
                case 1:
                    // Discard loading of the table completely.
                    args.IsLoaded = false;
                    break;
                case 2:
                    // Load the table with custom options.
                    args.Options = new DocumentTableOptions();
                    args.Options.FirstRowContainsColumnNames = true;
                    break;
                default:
                    throw new InvalidOperationException();
            }
        }
    }
}

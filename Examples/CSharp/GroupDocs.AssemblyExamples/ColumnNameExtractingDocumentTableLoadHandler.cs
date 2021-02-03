using GroupDocs.Assembly.Data;

namespace GroupDocs.AssemblyExamples
{
    public class ColumnNameExtractingDocumentTableLoadHandler : IDocumentTableLoadHandler
    {
        public void Handle(DocumentTableLoadArgs args)
        {
            args.Options = new DocumentTableOptions { FirstRowContainsColumnNames = true };
        }
    }
}

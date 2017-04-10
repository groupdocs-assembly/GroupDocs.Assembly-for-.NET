
Imports GroupDocs.Assembly.Data

Public Class ColumnNameExtractingDocumentTableLoadHandler
    Implements IDocumentTableLoadHandler

    Public Sub Handle(args As DocumentTableLoadArgs) Implements IDocumentTableLoadHandler.Handle
        args.Options = New DocumentTableOptions()
        args.Options.FirstRowContainsColumnNames = True
    End Sub
End Class

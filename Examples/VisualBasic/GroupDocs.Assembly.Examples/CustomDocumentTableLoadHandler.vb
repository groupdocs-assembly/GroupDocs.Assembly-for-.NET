
Imports GroupDocs.Assembly.Data

Public Class CustomDocumentTableLoadHandler
    Implements IDocumentTableLoadHandler
    Public Sub Handle(args As DocumentTableLoadArgs) Implements IDocumentTableLoadHandler.Handle
        Select Case args.TableIndex
            Case 0
                ' Do nothing. The table is to be loaded with default options.
                Exit Select
            Case 1
                ' Discard loading of the table completely.
                args.IsLoaded = False
                Exit Select
            Case 2
                ' Load the table with custom options.
                args.Options = New DocumentTableOptions()
                args.Options.FirstRowContainsColumnNames = True
                Exit Select
            Case Else
                Throw New InvalidOperationException()
        End Select
    End Sub

End Class


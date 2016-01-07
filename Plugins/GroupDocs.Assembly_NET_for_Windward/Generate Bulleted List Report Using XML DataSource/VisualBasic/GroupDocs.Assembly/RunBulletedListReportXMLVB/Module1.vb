Imports System.IO
Imports GroupDocs.Assembly
'ExStart:Generate bulleted list in groupdocs.assembly-VB
Module Module1

    Public Const productXMLfile As String = "../../../../../Data/Data Source/Products.xml"
    Sub Main()

        'Setting up source open presentation template
        Dim template As FileStream = File.OpenRead("../../../../../Data/Samples/Source/Bulleted List_XML.docx")
        'Setting up destination open presentation report 
        Dim output As FileStream = File.Create("../../../../../Data/Samples/Destination/Xml Bulleted List.docx")

        Try
            'Instantiate DocumentAssembler class
            Dim assembler As New DocumentAssembler()
            'Call AssembleDocument to generate Bulleted List Report in open presentation format
            assembler.AssembleDocument(template, output, GetProductsData(), "ds")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try


    End Sub

    Public Function GetProductsData() As DataSet
        Try
            Dim mainDs As New DataSet()
            Dim fsReadXml As New System.IO.FileStream(productXMLfile, System.IO.FileMode.Open)
            mainDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
            mainDs.Tables(0).TableName = "Product"
            Return mainDs
        Catch
            Return Nothing
        End Try
    End Function
End Module
'ExEnd:Generate bulleted list in groupdocs.assembly-VB

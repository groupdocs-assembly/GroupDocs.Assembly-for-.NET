Imports System.IO
Imports GroupDocs.Assembly

'ExStart:Generate CommonList in groupdocs.assembly-VB
Module Module1

    Public Const CustomerXMLfile As String = "../../../../../Data/Data Source/Customers.xml"
    Sub Main()

        'Setting up source open presentation template
        Dim template As FileStream = File.OpenRead("../../../../../Data/Samples/Source/Common List.docx")
        'Setting up destination open presentation report 
        Dim output As FileStream = File.Create("../../../../../Data/Samples/Destination/Xml Common List.docx")

        Try
            'Instantiate DocumentAssembler class
            Dim assembler As New DocumentAssembler()
            'Call AssembleDocument to generate single row Report in open presentation format
            assembler.AssembleDocument(template, output, GetCustomListData(), "customer")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try


    End Sub
    Public Function GetCustomListData() As DataSet
        Try
            Dim mainDs As New DataSet()
            Dim fsReadXml As New System.IO.FileStream(CustomerXMLfile, System.IO.FileMode.Open)
            mainDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
            mainDs.Tables(0).TableName = "Customers"
            Return mainDs
        Catch
            Return Nothing
        End Try
    End Function


End Module
'ExEnd:Generate CommonList in groupdocs.assembly-VB

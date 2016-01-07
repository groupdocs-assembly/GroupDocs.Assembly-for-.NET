Imports System.IO
Imports GroupDocs.Assembly

'ExStart:Generate single row in groupdocs.assembly-VB
Module Module1

    Public Const CustomerXMLfile As String = "../../../../../Data/Data Source/Customers.xml"
    Sub Main()

        'Setting up source open presentation template
        Dim template As FileStream = File.OpenRead("../../../../../Data/Samples/Source/Single row.docx")
        'Setting up destination open presentation report 
        Dim output As FileStream = File.Create("../../../../../Data/Samples/Destination/Xml Single row.docx")

        Try
            'Instantiate DocumentAssembler class
            Dim assembler As New DocumentAssembler()
            'Call AssembleDocument to generate single row Report in open presentation format
            assembler.AssembleDocument(template, output, GetSingleRow(), "customer")
        Catch ex As Exception
            Console.WriteLine(ex.Message)
        End Try


    End Sub


    Public Function GetSingleRow() As DataRow
        Try
            Dim singleCustomer As New DataSet()
            Dim readProductXML As New FileStream(CustomerXMLfile, FileMode.Open)
            singleCustomer.ReadXml(readProductXML, XmlReadMode.ReadSchema)
            singleCustomer.Tables(0).TableName = "Customers"
            Return singleCustomer.Tables("Customers").Rows(0)
        Catch
            Return Nothing
        End Try
    End Function

End Module
'ExEnd:Generate single row in groupdocs.assembly-VB

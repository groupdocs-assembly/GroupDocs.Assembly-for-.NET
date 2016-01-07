Imports System
Imports System.Data
Imports Net.windward.api.csharp
Imports System.IO
Imports WindwardReportsDrivers.net.windward.datasource.ado
Imports WindwardInterfaces.net.windward.api.csharp

Module ModuleXml

    Sub Main()
        'ExStart:Generate bulleted list in windward-VB
        'Initilize the engine
        Report.Init()
        'open a inputfilestream for our template file
        Dim rtf As FileStream = File.OpenRead("../../../../../Data/Samples/Source/Bulleted List.docx")
        'open an inputfilestream for our data file
        Dim Xml As FileStream = File.OpenRead("../../../../../Data/Data Source/WW-Products.xml")
        'open an outputfilestream for our output
        Dim output As FileStream = File.Create("../../../../../Data/Samples/Destination/Xml Bulleted List.docx")

        'Create a report process
        Dim myReport As Report = New ReportPdf(rtf, output)

        'open a data object to connect to our xml file
        Dim data As IReportDataSource = New XmlDataSourceImpl(Xml, False)

        'run the report process
        myReport.ProcessSetup()
        'the second parameter is "" to tell the process that our data is the default data source
        myReport.ProcessData(data, "SW")
        myReport.ProcessComplete()

        'close out of our template file and output
        output.Close()
        rtf.Close()
        Xml.Close()

        'Open finished report
        Dim fullPath As String = Path.GetFullPath("../../../../../Data/Samples/Destination/Xml Bulleted List.docx")
        System.Diagnostics.Process.Start(fullPath)
        'ExEnd:Generate bulleted list in windward-VB
    End Sub

End Module

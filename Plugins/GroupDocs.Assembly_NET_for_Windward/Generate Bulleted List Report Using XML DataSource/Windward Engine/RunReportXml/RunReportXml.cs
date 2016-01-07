using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using net.windward.api.csharp;
using WindwardReportsDrivers.net.windward.datasource;
using WindwardInterfaces.net.windward.api.csharp;
using System.IO;

namespace RunReportXml
{
    class RunReportXml
    {
        static void Main(string[] args)
        {
            //ExStart:Windward XML report generation
            // Initialize the engine
            Report.Init();

            // Open template file and create output file
            FileStream template = File.OpenRead("../../../../Data/Samples/Source/Bulleted List.docx");
            FileStream output = File.Create("../../../../Data/Samples/Destination/Xml Bulleted List.docx");

            // Create report process
            Report myReport = new ReportPdf(template, output);


            // Open an inputfilestream for our data file
            FileStream Xml = File.OpenRead("../../../../Data/Data Source/WW-Products.xml");

            // Open a data object to connect to our xml file
            IReportDataSource data = new XmlDataSourceImpl(Xml, false);

            // Run the report process
            myReport.ProcessSetup();
            // The second parameter is "" to tell the process that our data is the default data source
            myReport.ProcessData(data, "SW");
            myReport.ProcessComplete();

            // Close out of our template file and output
            output.Close();
            template.Close();
            Xml.Close();

            // Opens the finished report
            string fullPath = Path.GetFullPath("../../../../Data/Samples/Destination/Xml Bulleted List.docx");
            System.Diagnostics.Process.Start(fullPath);
            //ExEnd:Windward XML report generation
        }
    }
}

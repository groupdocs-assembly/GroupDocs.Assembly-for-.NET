//ExStart:ReportGenerator
using GroupDocs.Assembly;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;

namespace GroupDocs.ReportGenerator.BusinessLayer
{
    public class ReportGenerator
    {
        /// <summary>
        /// Properties for Report Generator
        /// </summary>
        private String _TemplateSourcePath;

        public String TemplateSourcePath
        {
            get { return _TemplateSourcePath; }
            set { _TemplateSourcePath = value; }
        }
        private String _ReportDestinationPath;

        public String ReportDestinationPath
        {
            get { return _ReportDestinationPath; }
            set { _ReportDestinationPath = value; }
        }

        private String _LicensePath = "D:/ReportGenerator/GroupDocs.ReportGenerator/GroupDocs.Assembly Product Family.lic";
       
        //ExStart:GenerateReportShowcase
        /// <summary>
        /// Generates the report by using dataset.
        /// </summary>
        /// <param name="ds"></param>
        public void GenerateReportDB(DataSet ds)
        {
            try
            {
                License lic = new License();
                //lic.SetLicense(_LicensePath);
                //Instantiate DocumentAssembler class
                DocumentAssembler assembler = new DocumentAssembler();
                //Call AssembleDocument to generate Common List Report in open document format
                assembler.AssembleDocument(_TemplateSourcePath, _ReportDestinationPath, ds);
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }
        //ExEnd:GenerateReportShowcase
    }
}
//ExEnd:ReportGenerator

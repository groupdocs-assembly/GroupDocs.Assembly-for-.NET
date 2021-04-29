using System;
using GroupDocs.Assembly;
using GroupDocs.Assembly.Data;
using NUnit.Framework;

namespace AssemblyExamples
{
    [TestFixture]
    public class TableSet : AssemblyExamplesBase
    {
        /// <summary>
        /// Shows how to change document table column type.
        /// Feature is supported by version 17.01 or greater.
        /// </summary>
        [Test]
        public void ChangingColumnType()
        {
            //ExStart:ChangingColumnType
            DocumentTable table = new DocumentTable(DataSourcesDir + "Managers.docx", 1,
                new DocumentTableOptions {FirstRowContainsColumnNames = true});

            // NOTE: For non-Spreadsheet documents, the type of a document table column is always string by default.
            Assert.AreEqual(typeof(string), table.Columns["Total_Contract_Price"].Type);

            // Change the column's type to double thus enabling to use arithmetic operations on values of the column 
            // such as summing in templates.
            table.Columns["Total_Contract_Price"].Type = typeof(double);

            DocumentAssembler assembler = new DocumentAssembler();

            assembler.AssembleDocument(TemplatesDir + "Changed column type.pptx",
                ArtifactsDir + "TableSet.ChangingColumnType.pptx", new DataSourceInfo(table, "Managers"));
            //ExEnd:ChangingColumnType
        }

        /// <summary>
        /// Shows how to load document table set using default options.
        /// Features is supported by version 17.01 or greater.
        /// </summary>
        [Test]
        public void LoadTableSet()
        {
            //ExStart:LoadTableSet
            // Load all document tables using default options.
            DocumentTableSet tableSet =
                new DocumentTableSet(DataSourcesDir + "Multiple tables.docx");

            Assert.AreEqual(3, tableSet.Tables.Count);
            Assert.AreEqual("Table1", tableSet.Tables[0].Name);
            Assert.AreEqual("Table2", tableSet.Tables[1].Name);
            Assert.AreEqual("Table3", tableSet.Tables[2].Name);
            //ExEnd:LoadTableSet
        }

        /// <summary>
        /// Show how to Load document table set using custom options
        /// Features is supported by version 17.01 or greater
        /// </summary>
        //ExStart:LoadTableSetWithCustomOptions
        [Test] //ExSkip
        public void LoadTableSetWithCustomOptions()
        {
            // Load document tables using custom options.
            DocumentTableSet tableSet = new DocumentTableSet(
                DataSourcesDir + "Multiple tables.docx",
                new CustomLoadHandler());

            // Ensure that the second table is not loaded.
            Assert.AreEqual(2, tableSet.Tables.Count);
            Assert.AreEqual("Table1", tableSet.Tables[0].Name);
            Assert.AreEqual("Table3", tableSet.Tables[1].Name);

            // Ensure that default options are used to load the first table (that is, default column names are used).
            Assert.AreEqual(2, tableSet.Tables[0].Columns.Count);
            Assert.AreEqual("Column1", tableSet.Tables[0].Columns[0].Name);
            Assert.AreEqual("Column2", tableSet.Tables[0].Columns[1].Name);

            // Ensure that custom options are used to load the third table (that is, column names are extracted).
            Assert.AreEqual(2, tableSet.Tables[1].Columns.Count);
            Assert.AreEqual("Name", tableSet.Tables[1].Columns[0].Name);
            Assert.AreEqual("Address", tableSet.Tables[1].Columns[1].Name);
        }

        public class CustomLoadHandler : IDocumentTableLoadHandler
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
        //ExEnd:LoadTableSetWithCustomOptions

        /// <summary>
        /// Shows how to use document TableSet as DataSource.
        /// Features is supported by version 17.01 or greater.
        /// </summary>
        //ExStart:UseTableSetAsDataSource
        [Test] //ExSkip
        public void SeveralTables()
        {
            // Set table column names to be extracted from the document.
            DocumentTableSet tableSet = new DocumentTableSet(
                DataSourcesDir + "Multiple tables.docx",
                new ColumnNameExtractingHandler());

            // Set table names for convenience.
            tableSet.Tables[0].Name = "Planets";
            tableSet.Tables[1].Name = "Persons";
            tableSet.Tables[2].Name = "Companies";

            Assert.AreEqual(3, tableSet.Tables.Count);
            Assert.AreEqual("Planets", tableSet.Tables[0].Name);
            Assert.AreEqual("Persons", tableSet.Tables[1].Name);
            Assert.AreEqual("Companies", tableSet.Tables[2].Name);

            Assert.AreEqual(2, tableSet.Tables[0].Columns.Count);
            Assert.AreEqual("Planet", tableSet.Tables[0].Columns[0].Name);
            Assert.AreEqual("Number", tableSet.Tables[0].Columns[1].Name);

            Assert.AreEqual(2, tableSet.Tables[1].Columns.Count);
            Assert.AreEqual("First_Name", tableSet.Tables[1].Columns[0].Name);
            Assert.AreEqual("Last_Name", tableSet.Tables[1].Columns[1].Name);

            Assert.AreEqual(2, tableSet.Tables[2].Columns.Count);
            Assert.AreEqual("Name", tableSet.Tables[2].Columns[0].Name);
            Assert.AreEqual("Address", tableSet.Tables[2].Columns[1].Name);

            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(
                TemplatesDir + "Several tables.pptx",
                ArtifactsDir + "TableSet.SeveralTables.pptx",
                new DataSourceInfo(tableSet));
        }

        public class ColumnNameExtractingHandler : IDocumentTableLoadHandler
        {
            public void Handle(DocumentTableLoadArgs args)
            {
                args.Options = new DocumentTableOptions { FirstRowContainsColumnNames = true };
            }
        }
        //ExEnd:UseTableSetAsDataSource

        /// <summary>
        /// Shows how to define document table relations.
        /// Feature is supported by version 17.01 or greater.
        /// </summary>
        [Test]
        public void DefiningRelations()
        {
            //ExStart:DefiningRelations
            // Set table column names to be extracted from the document.
            DocumentTableSet tableSet =
                new DocumentTableSet(DataSourcesDir + "Related tables.xlsx", new GetColumnNames());

            // Define relations between tables.
            // NOTE: For Spreadsheet documents, table names are extracted from sheet names.
            tableSet.Relations.Add(
                tableSet.Tables["CLIENT"].Columns["ID"],
                tableSet.Tables["CONTRACT"].Columns["CLIENT_ID"]);
            tableSet.Relations.Add(
                tableSet.Tables["MANAGER"].Columns["ID"],
                tableSet.Tables["CONTRACT"].Columns["MANAGER_ID"]);

            DocumentAssembler assembler = new DocumentAssembler();
            
            assembler.AssembleDocument(TemplatesDir + "Table relations.docx",
                ArtifactsDir + "TableSet.DefiningRelations.docx", new DataSourceInfo(tableSet));
            //ExEnd:DefiningRelations
        }

        public class GetColumnNames : IDocumentTableLoadHandler
        {
            public void Handle(DocumentTableLoadArgs args)
            {
                args.Options = new DocumentTableOptions { FirstRowContainsColumnNames = true };
            }
        }

        /// <summary>
        /// Import Spreadsheet into HTML Document.
        /// </summary>
        [Test]
        public void ImportingSpreadsheetTable()
        {
            //ExStart:ImportingSpreadsheetTable
            // Do not extract column names from the first row, so that the first row to be treated as a data row.
            // Limit the largest row index, so that only the first four data rows to be loaded.
            DocumentTableOptions options = new DocumentTableOptions { MaxRowIndex = 3 };

            // Use data of the _second_ table in the document.
            DocumentTable table = new DocumentTable(DataSourcesDir + "Contracts.xlsx", 0, options);

            // Check column count and names.
            Assert.AreEqual(3, table.Columns.Count);

            // NOTE: Default column names are used, because we do not extract the names from the first row.
            Assert.AreEqual("A", table.Columns[0].Name);
            Assert.AreEqual("B", table.Columns[1].Name);
            Assert.AreEqual("C", table.Columns[2].Name);

            // Assemble a document using the external document table as a data source.
            DocumentAssembler assembler = new DocumentAssembler();
            assembler.AssembleDocument(TemplatesDir + "Importing spreadsheet table.html",
                ArtifactsDir + "TableSet.ImportingSpreadsheetTable.html",
                new DataSourceInfo(table, "table"));
            //ExEnd:ImportingSpreadsheetTable
        }
    }
}

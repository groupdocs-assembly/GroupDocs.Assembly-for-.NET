using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using GroupDocs.Assembly.Data;
using Newtonsoft.Json;
using NUnit.Framework;

namespace AssemblyExamples.Data
{
    //ExStart:DataLayer
    public class DataLayer : AssemblyExamplesBase
    {
        public static string ProductXmlFile = DataSourcesDir + "Products.xml";
        public static string CustomerXmlFile = DataSourcesDir + "Customers.xml";
        public static string OrderXmlFile = DataSourcesDir + "Orders.xml";
        public static string ProductOrderXmlFile = DataSourcesDir + "ProductOrders.xml";
        public static string JsonFile = DataSourcesDir + "Customers.json";
        public static string ExcelDataFile = DataSourcesDir + "Contracts.xlsx";
        
        //ExStart:PopulateData
        /// <summary>
        /// This function initializes/populates the data. 
        /// Initialize Customer information, product information and order information.
        /// </summary>
        /// <returns>Returns customer's complete information.</returns>
        public static IEnumerable<BusinessObjects.Customer> PopulateData()
        {
            BusinessObjects.Customer customer = new BusinessObjects.Customer
            {
                CustomerName = "Jane Doe", CustomerContactNumber = "+9211874",
                ShippingAddress = "Flat # 1, Kiyani Plaza ISB", Barcode = "123456789qwertyu0025"
            };

            customer.Order = new[]
            {
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product {ProductName = "Lumia 525"}, Customer = customer, Price = 170,
                    ProductQuantity = 5, OrderDate = new DateTime(2015, 1, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service {ServiceName = "Regular Cleaning"},
                        new BusinessObjects.Service {ServiceName = "Oven Cleaning"}
                    }
                }

            };
            yield return customer;

            customer = new BusinessObjects.Customer
            {
                CustomerName = "John Smith", CustomerContactNumber = "+458789",
                ShippingAddress = "Quette House, Park Road, ISB", Barcode = "123456789qwertyu0025"
            };
            customer.Order = new[]
            {
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product {ProductName = "Lenovo G50"}, Customer = customer,
                    Price = 480, ProductQuantity = 2, OrderDate = new DateTime(2015, 2, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service {ServiceName = "Regular Cleaning"},
                        new BusinessObjects.Service {ServiceName = "Oven Cleaning"},
                        new BusinessObjects.Service {ServiceName = "Carpet Cleaning"}
                    }
                },
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product {ProductName = "Pavilion G6"}, Customer = customer,
                    Price = 400, ProductQuantity = 1, OrderDate = new DateTime(2015, 10, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service {ServiceName = "Regular Cleaning"},
                        new BusinessObjects.Service {ServiceName = "Carpet Cleaning"}

                    }
                },
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product {ProductName = "Nexus 5"}, Customer = customer, Price = 320,
                    ProductQuantity = 3, OrderDate = new DateTime(2015, 6, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service {ServiceName = "Oven Cleaning"}
                    }
                }
            };
            yield return customer;
        }
        //ExEnd:PopulateData
        
        //ExStart:GetOrdersData
        /// <summary>
        /// Fetches order from PopulateData.
        /// </summary>
        /// <returns>Returns order details, one data at a time.</returns>
        public static IEnumerable<BusinessObjects.Order> GetOrdersData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order;
            }
        }
        //ExEnd:GetOrdersData
        
        //ExStart:GetProductsData
        /// <summary>
        /// Fetches products from PopulateData.
        /// </summary>
        /// <returns>Returns product details, one data at a time.</returns>
        public static IEnumerable<BusinessObjects.Product> GetProductsData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order.Product;
            }
        }
        //ExEnd:GetProductsData
        
        //ExStart:GetSingleCustomerData
        /// <summary>
        /// Fetches customer details of very first customer.
        /// </summary>
        /// <returns>Returns first customer's information.</returns>
        public static BusinessObjects.Customer GetCustomerData()
        {
            using (IEnumerator<BusinessObjects.Customer> customer = PopulateData().GetEnumerator())
            {
                customer.MoveNext();
                customer.MoveNext();

                return customer.Current;
            }
        }
        //ExEnd:GetSingleCustomerData
        
        //ExStart:GetCustomerDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data.
        /// </summary>
        /// <returns>Returns deserialized data.</returns>
        public static IEnumerable<BusinessObjects.Customer> GetCustomerDataFromJson()
        {
            string rawData = File.ReadAllText(JsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                yield return customer;
            }
        }
        //ExEnd:GetCustomerDataJson
        
        //ExStart:GetCustomerOrderDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data.
        /// </summary>
        /// <returns>Returns deserialized data.</returns>
        public static IEnumerable<BusinessObjects.Order> GetCustomerOrderDataFromJson()
        {
            string rawData = File.ReadAllText(JsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                foreach (BusinessObjects.Order order in customer.Order)
                {
                    yield return order;
                }
            }
        }
        //ExEnd:GetCustomerOrderDataJson
        
        //ExStart:GetProductsDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data.
        /// </summary>
        /// <returns>Returns deserialized data.</returns>
        public static IEnumerable<BusinessObjects.Product> GetProductsDataJson()
        {
            string rawData = File.ReadAllText(JsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order.Product;
            }
        }
        //ExEnd:GetProductsDataJson

        //ExStart:GetAllDataXML
        /// <summary>
        /// Fetches the data from the XML files and store data in the DataSet.
        /// </summary>
        /// <returns>Returns the DataSet.</returns>
        public static DataSet GetAllDataFromXml()
        {
            DataSet tempDs = new DataSet();
            DataSet mainDs = new DataSet();
            FileStream fsReadXml;

            using (fsReadXml = new FileStream(CustomerXmlFile, FileMode.Open))
            {
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Customers";

                mainDs.Merge(tempDs.Tables[0]);
            }

            tempDs = new DataSet();

            using (fsReadXml = new FileStream(OrderXmlFile, FileMode.Open))
            {
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Orders";

                mainDs.Merge(tempDs.Tables[0]);
            }

            tempDs = new DataSet();

            using (fsReadXml = new FileStream(ProductOrderXmlFile, FileMode.Open))
            {
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "ProductOrders";

                mainDs.Merge(tempDs.Tables[0]);
            }

            tempDs = new DataSet();

            using (fsReadXml = new FileStream(ProductXmlFile, FileMode.Open))
            {
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Product";

                mainDs.Merge(tempDs.Tables[0]);
            }

            // Defining relations between the tables.
            DataColumn orderColumn = mainDs.Tables["Orders"].Columns["OrderID"];
            DataColumn productColumn = mainDs.Tables["ProductOrders"].Columns["OrderID"];
            DataColumn customerColumn = mainDs.Tables["Customers"].Columns["CustomerID"];
            DataColumn customerOrderColumn = mainDs.Tables["Orders"].Columns["CustomerID"];
            DataColumn productOrderProductIdColumn = mainDs.Tables["ProductOrders"].Columns["ProductID"];
            DataColumn productProductIdColumn = mainDs.Tables["Product"].Columns["ProductID"];

            DataRelation dr = new DataRelation("Order_ProductOrders", orderColumn, productColumn);
            mainDs.Relations.Add(dr);
            DataRelation dr2 = new DataRelation("Customer_Orders", customerColumn, customerOrderColumn);
            mainDs.Relations.Add(dr2);
            DataRelation dr3 = new DataRelation("Product_ProductOrders", productProductIdColumn,
                productOrderProductIdColumn);
            mainDs.Relations.Add(dr3);

            return mainDs;
        }
        //ExEnd:GetAllDataXML

        /// <summary>
        /// Generate report from excel data source.
        /// </summary>
        /// <returns></returns>
        public static DocumentTable ExcelData()
        {
            DocumentTableOptions options = new DocumentTableOptions { FirstRowContainsColumnNames = true };

            // Use data of the _first_ worksheet.
            DocumentTable table = new DocumentTable(ExcelDataFile, 0, options);

            // Check column count, names, and types.
            Assert.AreEqual(3, table.Columns.Count);

            Assert.AreEqual("Client", table.Columns[0].Name);
            Assert.AreEqual(typeof(string), table.Columns[0].Type);

            Assert.AreEqual("Manager", table.Columns[1].Name);
            Assert.AreEqual(typeof(string), table.Columns[1].Type);

            // NOTE: A space is replaced with an underscore, because spaces are not allowed in column names.
            Assert.AreEqual("Price", table.Columns[2].Name);

            // NOTE: The type of the column is double, because all cells in the column contain numeric values.
            Assert.AreEqual(typeof(double), table.Columns[2].Type);

            return table;
        }

        /// <summary>
        /// Creates an Email data source object.
        /// </summary>
        /// <param name="fileName">Name of the template file.</param>
        /// <param name="dataSource">data source.</param>
        /// <returns></returns>
        public static BusinessObjects.EmailDataSourcesObjects EmailDataSourceObject(string fileName, object dataSource)
        {
            //ExStart:EmailDataSourceObject
            List<string> recipients = new List<string>
            {
                "Named Recipient <named@example.com>", "unnamed@example.com"
            };

            BusinessObjects.EmailDataSourcesObjects dataSources = new BusinessObjects.EmailDataSourcesObjects
            {
                DataSource = dataSource,
                Sender = "Example Sender <sender@example.com>",
                Recipients = recipients,
                CC = "cc@example.com",
                Subject = Path.GetFileNameWithoutExtension(fileName),
                Title = "title"
            };

            return dataSources;
            //ExEnd:EmailDataSourceObject
        }

        public static BusinessObjects.EmailDataSourcesNames EmailDataSourceName(string name)
        {
            //ExStart:EmailDataSourceName
            BusinessObjects.EmailDataSourcesNames dataSourceNames = new BusinessObjects.EmailDataSourcesNames
            {
                Name = name,
                Sender = "sender",
                Recipients = "recipients",
                CC = "cc",
                Subject = "subject",
                Title = "title"
            };
            
            return dataSourceNames;
            //ExEnd:EmailDataSourceName
        }
    }
    //ExEnd:DataLayer
}

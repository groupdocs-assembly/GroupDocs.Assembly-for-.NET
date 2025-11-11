using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using GroupDocs.Assembly.Data;
using Newtonsoft.Json;

namespace GroupDocs.Assembly.Examples.CSharp.Data
{
    public class DataLayer
    {
        private static readonly string ProductXmlFile = Path.Combine(Constants.DataSourcesPath, "Products.xml");
        private static readonly string CustomerXmlFile = Path.Combine(Constants.DataSourcesPath, "Customers.xml");
        private static readonly string OrderXmlFile = Path.Combine(Constants.DataSourcesPath, "Orders.xml");
        private static readonly string ProductOrderXmlFile = Path.Combine(Constants.DataSourcesPath, "ProductOrders.xml");
        private static readonly string JsonFile = Path.Combine(Constants.DataSourcesPath, "Customers.json");
        private static readonly string ExcelDataFile = Path.Combine(Constants.DataSourcesPath, "Contracts.xlsx");

        public static IEnumerable<BusinessObjects.Customer> PopulateData()
        {
            BusinessObjects.Customer customer = new BusinessObjects.Customer()
            {
                CustomerName = "Jane Doe",
                CustomerContactNumber = "+9211874",
                ShippingAddress = "Flat # 1, Kiyani Plaza ISB",
                Barcode = "123456789qwertyu0025"
            };

            customer.Order = new[]
            {
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product { ProductName = "Lumia 525" },
                    Customer = customer,
                    Price = 170,
                    ProductQuantity = 5,
                    OrderDate = new DateTime(2015, 1, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service { ServiceName = "Regular Cleaning" },
                        new BusinessObjects.Service { ServiceName = "Oven Cleaning" }
                    }
                }
            };
            yield return customer;

            customer = new BusinessObjects.Customer
            {
                CustomerName = "John Smith",
                CustomerContactNumber = "+458789",
                ShippingAddress = "Quette House, Park Road, ISB",
                Barcode = "123456789qwertyu0025"
            };

            customer.Order = new[]
            {
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product { ProductName = "Lenovo G50" },
                    Customer = customer,
                    Price = 480,
                    ProductQuantity = 2,
                    OrderDate = new DateTime(2015, 2, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service { ServiceName = "Regular Cleaning" },
                        new BusinessObjects.Service { ServiceName = "Oven Cleaning" },
                        new BusinessObjects.Service { ServiceName = "Carpet Cleaning" }
                    }
                },
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product { ProductName = "Pavilion G6" },
                    Customer = customer,
                    Price = 400,
                    ProductQuantity = 1,
                    OrderDate = new DateTime(2015, 10, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service { ServiceName = "Regular Cleaning" },
                        new BusinessObjects.Service { ServiceName = "Carpet Cleaning" }
                    }
                },
                new BusinessObjects.Order
                {
                    Product = new BusinessObjects.Product { ProductName = "Nexus 5" },
                    Customer = customer,
                    Price = 320,
                    ProductQuantity = 3,
                    OrderDate = new DateTime(2015, 6, 1),
                    Services = new[]
                    {
                        new BusinessObjects.Service { ServiceName = "Oven Cleaning" }
                    }
                }
            };
            yield return customer;
        }

        public static IEnumerable<BusinessObjects.Order> GetOrdersData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                {
                    yield return order;
                }
            }
        }

        public static IEnumerable<BusinessObjects.Product> GetProductsData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                {
                    yield return order.Product;
                }
            }
        }

        public static BusinessObjects.Customer GetCustomerData()
        {
            using IEnumerator<BusinessObjects.Customer> customer = PopulateData().GetEnumerator();
            customer.MoveNext();
            customer.MoveNext();
            return customer.Current;
        }

        public static IEnumerable<BusinessObjects.Customer> GetCustomerDataFromJson()
        {
            string rawData = File.ReadAllText(JsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                yield return customer;
            }
        }

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

        public static IEnumerable<BusinessObjects.Product> GetProductsDataJson()
        {
            string rawData = File.ReadAllText(JsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                foreach (BusinessObjects.Order order in customer.Order)
                {
                    yield return order.Product;
                }
            }
        }

        public static DataSet GetAllDataFromXml()
        {
            DataSet mainDs = new DataSet();

            using (FileStream fsReadXml = new FileStream(CustomerXmlFile, FileMode.Open, FileAccess.Read))
            {
                DataSet tempDs = new DataSet();
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Customers";
                mainDs.Merge(tempDs.Tables[0]);
            }

            using (FileStream fsReadXml = new FileStream(OrderXmlFile, FileMode.Open, FileAccess.Read))
            {
                DataSet tempDs = new DataSet();
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Orders";
                mainDs.Merge(tempDs.Tables[0]);
            }

            using (FileStream fsReadXml = new FileStream(ProductOrderXmlFile, FileMode.Open, FileAccess.Read))
            {
                DataSet tempDs = new DataSet();
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "ProductOrders";
                mainDs.Merge(tempDs.Tables[0]);
            }

            using (FileStream fsReadXml = new FileStream(ProductXmlFile, FileMode.Open, FileAccess.Read))
            {
                DataSet tempDs = new DataSet();
                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Product";
                mainDs.Merge(tempDs.Tables[0]);
            }

            DataColumn orderColumn = mainDs.Tables["Orders"].Columns["OrderID"];
            DataColumn productColumn = mainDs.Tables["ProductOrders"].Columns["OrderID"];
            DataColumn customerColumn = mainDs.Tables["Customers"].Columns["CustomerID"];
            DataColumn customerOrderColumn = mainDs.Tables["Orders"].Columns["CustomerID"];
            DataColumn productOrderProductIdColumn = mainDs.Tables["ProductOrders"].Columns["ProductID"];
            DataColumn productProductIdColumn = mainDs.Tables["Product"].Columns["ProductID"];

            mainDs.Relations.Add(new DataRelation("Order_ProductOrders", orderColumn, productColumn));
            mainDs.Relations.Add(new DataRelation("Customer_Orders", customerColumn, customerOrderColumn));
            mainDs.Relations.Add(new DataRelation("Product_ProductOrders", productProductIdColumn, productOrderProductIdColumn));

            // Convert relative Photo paths to absolute paths
            if (mainDs.Tables["Customers"].Columns.Contains("Photo"))
            {
                foreach (DataRow row in mainDs.Tables["Customers"].Rows)
                {
                    if (row["Photo"] != DBNull.Value && !string.IsNullOrEmpty(row["Photo"].ToString()))
                    {
                        string photoPath = row["Photo"].ToString();
                        // If it's a relative path, convert it to absolute
                        if (!Path.IsPathRooted(photoPath))
                        {
                            // Resolve relative path like '../../../../Data/Images/no-photo.jpg'
                            // The path is relative to the DataSources directory, so resolve from there
                            string resolvedPath = Path.GetFullPath(Path.Combine(Constants.DataSourcesPath, photoPath));
                            // If the resolved path doesn't exist, try resolving from the base directory
                            if (!File.Exists(resolvedPath) && photoPath.Contains("Images"))
                            {
                                // Extract just the filename and use Constants.ImagesPath
                                string fileName = Path.GetFileName(photoPath);
                                resolvedPath = Path.Combine(Constants.ImagesPath, fileName);
                            }
                            // Normalize the path
                            row["Photo"] = Path.GetFullPath(resolvedPath);
                        }
                        else
                        {
                            // Already absolute, just normalize it
                            row["Photo"] = Path.GetFullPath(photoPath);
                        }
                    }
                }
            }

            return mainDs;
        }

        public static DocumentTable ExcelData()
        {
            DocumentTableOptions options = new DocumentTableOptions() { FirstRowContainsColumnNames = true };
            DocumentTable table = new DocumentTable(ExcelDataFile, 0, options);
            return table;
        }

        public static BusinessObjects.EmailDataSourcesObjects EmailDataSourceObject(string fileName, object dataSource)
        {
            return new BusinessObjects.EmailDataSourcesObjects
            {
                DataSource = dataSource,
                Sender = "Example Sender <sender@example.com>",
                Recipients = new List<string>
                {
                    "Named Recipient <named@example.com>",
                    "unnamed@example.com"
                },
                CC = "cc@example.com",
                Subject = Path.GetFileNameWithoutExtension(fileName),
                Title = "title"
            };
        }

        public static BusinessObjects.EmailDataSourcesNames EmailDataSourceName(string name)
        {
            return new BusinessObjects.EmailDataSourcesNames
            {
                Name = name,
                Sender = "sender",
                Recipients = "recipients",
                CC = "cc",
                Subject = "subject",
                Title = "title"
            };
        }
    }
}



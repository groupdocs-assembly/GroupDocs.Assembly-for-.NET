using GroupDocs.AssemblyExamples.ProjectEntities;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using Newtonsoft.Json;
using GroupDocs.Assembly.Data;
using System.Diagnostics;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    #region DataLayer
    //ExStart:DataLayer
    public static class DataLayer
    {
        public const string productXMLfile = "../../../../Data/Data Sources/XML DataSource/Products.xml";
        public const string customerXMLfile = "../../../../Data/Data Sources/XML DataSource/Customers.xml";
        public const string orderXMLfile = "../../../../Data/Data Sources/XML DataSource/Orders.xml";
        public const string productOrderXMLfile = "../../../../Data/Data Sources/XML DataSource/ProductOrders.xml";
        public const string jsonFile = "../../../../Data/Data Sources/JSON DataSource/CustomerData-Json.json";
        public const string wordDataFile = "../../../../Data/Data Sources/Word DataSource/Managers Data.docx";
        public const string excelDataFile = "../../../../Data/Data Sources/Excel DataSource/Contracts Data.xlsx";

        #region DataInitialization
        //ExStart:PopulateData
        /// <summary>
        /// This function initializes/populates the data. 
        /// Initialize Customer information, product information and order information
        /// </summary>
        /// <returns>Returns customer's complete information</returns>
        public static IEnumerable<BusinessObjects.Customer> PopulateData()
        {
            BusinessObjects.Customer customer = new BusinessObjects.Customer { CustomerName = "Atir Tahir", CustomerContactNumber = "+9211874", ShippingAddress = "Flat # 1, Kiyani Plaza ISB", Barcode = "123456789qwertyu0025" };
            customer.Order = new BusinessObjects.Order[]
            {
                  new BusinessObjects.Order { Product = new BusinessObjects.Product { ProductName ="Lumia 525" }, Customer = customer, Price= 170, ProductQuantity = 5, OrderDate = new DateTime(2015, 1, 1) }

            };
            yield return customer; //yield return statement will return one data at a time


            customer = new BusinessObjects.Customer { CustomerName = "Usman Aziz", CustomerContactNumber = "+458789", ShippingAddress = "Quette House, Park Road, ISB", Barcode = "123456789qwertyu0025" };
            customer.Order = new BusinessObjects.Order[]
            {
                  new BusinessObjects.Order { Product = new BusinessObjects.Product { ProductName = "Lenovo G50" }, Customer = customer, Price = 480, ProductQuantity = 2, OrderDate = new DateTime(2015, 2, 1) },
                  new BusinessObjects.Order { Product = new BusinessObjects.Product { ProductName = "Pavilion G6"}, Customer = customer, Price = 400, ProductQuantity = 1, OrderDate = new DateTime(2015, 10, 1) },
                  new BusinessObjects.Order { Product = new BusinessObjects.Product { ProductName = "Nexus 5"}, Customer = customer, Price = 320, ProductQuantity = 3, OrderDate = new DateTime(2015, 6, 1) }
            };
            yield return customer; //yield return statement will return one data at a time 
        }
        //ExEnd:PopulateData
        #endregion

        #region GetOrders
        //ExStart:GetOrdersData
        /// <summary>
        /// Fetches order from PopulateData
        /// </summary>
        /// <returns>Returns order details, one data at a time</returns>
        public static IEnumerable<BusinessObjects.Order> GetOrdersData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order; //yield return statement returns one data at a time
            }
        }
        //ExEnd:GetOrdersData
        #endregion

        #region GetProducts
        //ExStart:GetProductsData
        /// <summary>
        /// Fetches products from PopulateData
        /// </summary>
        /// <returns>Returns product details, one data at a time</returns>
        public static IEnumerable<BusinessObjects.Product> GetProductsData()
        {
            foreach (BusinessObjects.Customer customer in PopulateData())
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order.Product;
            }
        }
        //ExEnd:GetProductsData
        #endregion

        #region GetSingleCustomerData
        //ExStart:GetSingleCustomerData
        /// <summary>
        /// Fetches customer details of very first customer
        /// </summary>
        /// <returns>Returns first customer's infromation</returns>
        public static BusinessObjects.Customer GetCustomerData()
        {
            IEnumerator<BusinessObjects.Customer> customer = PopulateData().GetEnumerator();
            customer.MoveNext();
            return customer.Current;
        }
        //ExEnd:GetSingleCustomerData
        #endregion

        #region GetOrdersDataDB
        //ExStart:GetOrdersDataDB
        /// <summary>
        /// Fetches orders from database
        /// </summary>
        /// <returns>Returns order details, one data at a time</returns>
        public static IEnumerable<Order> GetOrdersDataDB()
        {
            //create object of data context
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            return dbEntities.Orders;
        }
        //ExEnd:GetOrdersDataDB
        #endregion

        #region GetProductsDataDB
        //ExStart:GetProductsDataDB
        /// <summary>
        /// Fetches products from database 
        /// </summary>
        /// <returns>Returns products information, one data at a time </returns>
        public static IEnumerable<Product> GetProductsDataDB()
        {
            //create object of data context
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            return dbEntities.Products;
        }
        //ExEnd:GetProductsDataDB
        #endregion

        #region GetCustomersDataDB
        //ExStart:GetCustomersDataDB
        /// <summary>
        /// Fetches customers from database
        /// </summary>
        /// <returns>Returns customers information, one data at a time</returns>
        public static IEnumerable<Customer> GetCustomersDataDB()
        {
            //create object of data context
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            return dbEntities.Customers;
        }
        //ExEnd:GetCustomersDataDB
        #endregion

        #region GetSingleCustomerDataDB
        //ExStart:GetSingleCustomerDataDB
        /// <summary>
        /// Fetches single customer data
        /// </summary>
        /// <returns>Return single, first customer's data</returns>
        public static Customer GetSingleCustomerDataDB()
        {
            //create object of data context
            DatabaseEntitiesDataContext dbEntites = new DatabaseEntitiesDataContext();
            IEnumerator<Customer> customer = GetCustomersDataDB().GetEnumerator();
            customer.MoveNext();
            return customer.Current;
        }
        //ExEnd:GetSingleCustomerDataDB
        #endregion

        #region GetProductsDataDT
        //ExStart:GetProductsDataDT
        /// <summary>
        /// Fetches Products and ProductOrders information, store them in DataTables and load DataTable to DataSet
        /// </summary>
        /// <returns>Returns DataSet</returns>
        public static DataSet GetProductsDT()
        {
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            var products = (from c in dbEntities.Products
                            select c).AsEnumerable();
            var productOrders = (from c in dbEntities.ProductOrders
                                 select c).AsEnumerable();
            DataTable Products = new DataTable();
            //ToADOTable function converts query results into DataTable
            Products = products.ToADOTable(rec => new object[] { products });
            DataTable ProductOrders = new DataTable();
            ProductOrders = productOrders.ToADOTable(rec => new object[] { productOrders });
            ProductOrders.TableName = "ProductOrder";
            Products.TableName = "products";
            DataSet dataSet = new DataSet();
            //Adding DataTable in DataSet
            dataSet.Tables.Add(Products);
            dataSet.Tables.Add(ProductOrders);
            return dataSet;
        }
        //ExEnd:GetProductsDataDT
        #endregion

        #region GetSingleCustomerDataDT
        //ExStart:GetSingleCustomerDT
        /// <summary>
        /// Fetches Customers from database
        /// </summary>
        /// <returns>Returns DataSet, very first record from the table</returns>
        public static DataRow GetSingleCustomerDT()
        {
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            var customers = (from c in dbEntities.Customers
                             select new
                             {
                                 c.CustomerID,
                                 c.CustomerName,
                                 c.ShippingAddress,
                                 c.CustomerContactNumber,
                                 c.Photo
                             }).AsEnumerable();
            DataTable Customers = new DataTable();
            //ToADOTable function converts query results into DataTable
            Customers = customers.ToADOTable(rec => new object[] { customers });
            Customers.TableName = "Customers";
            DataSet dataSet = new DataSet();
            //Adding DataTable in DataSet
            dataSet.Tables.Add(Customers);
            return dataSet.Tables["Customers"].Rows[0];
        }
        //ExEnd:GetSingleCustomerDT
        #endregion

        #region GetCustomersAndOrdersDataDT
        //ExStart:GetCustomersAndOrdersDataDT
        /// <summary>
        /// Fetches Orders, ProductOrders and Customers from database
        /// </summary>
        /// <returns>Returns DataSet</returns>
        public static DataSet GetCustomersAndOrdersDataDT()
        {
            DatabaseEntitiesDataContext dbEntities = new DatabaseEntitiesDataContext();
            var orders = (from c in dbEntities.Orders
                          select c).AsEnumerable();
            var productOrders = (from c in dbEntities.ProductOrders
                                 select c).AsEnumerable();
            var customers = (from c in dbEntities.Customers
                             select c).AsEnumerable();
            DataTable Orders = new DataTable();
            //ToADOTable function converts query results into DataTable
            Orders = orders.ToADOTable(rec => new object[] { orders });
            DataTable ProductOrders = new DataTable();
            ProductOrders = productOrders.ToADOTable(rec => new object[] { productOrders });
            DataTable Customers = new DataTable();
            Customers = customers.ToADOTable(rec => new object[] { customers });
            ProductOrders.TableName = "ProductOrder";
            Orders.TableName = "Orders";
            Customers.TableName = "Customers";
            DataSet dataSet = new DataSet();
            //Adding DataTable in DataSet
            dataSet.Tables.Add(Orders);
            dataSet.Tables.Add(ProductOrders);
            dataSet.Tables.Add(Customers);
            return dataSet;
        }
        //ExEnd:GetCustomersAndOrdersDataDT
        #endregion

        #region GetAllDataFromXML
        //ExStart:GetAllDataXML
        /// <summary>
        /// Fetches the data from the XML files and store data in the DataSet
        /// </summary>
        /// <returns>Returns the DataSet</returns>
        public static DataSet GetAllDataFromXML()
        {
            try
            {
                DataSet tempDs = new DataSet();
                DataSet mainDs = new DataSet();


                System.IO.FileStream fsReadXml = new System.IO.FileStream(customerXMLfile, System.IO.FileMode.Open);

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Customers";

                mainDs.Merge(tempDs.Tables[0]);
                tempDs = new DataSet();

                fsReadXml = new System.IO.FileStream(orderXMLfile, System.IO.FileMode.Open);

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Orders";

                mainDs.Merge(tempDs.Tables[0]);
                tempDs = new DataSet();

                fsReadXml = new System.IO.FileStream(productOrderXMLfile, System.IO.FileMode.Open);

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "ProductOrders";

                mainDs.Merge(tempDs.Tables[0]);
                tempDs = new DataSet();

                fsReadXml = new System.IO.FileStream(productXMLfile, System.IO.FileMode.Open);

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema);
                tempDs.Tables[0].TableName = "Product";

                mainDs.Merge(tempDs.Tables[0]);

                //Defining relations between the tables
                DataColumn productColumn, orderColumn, customerColumn, customerOrderColumn, productProductIDColumn, productOrderProductIDColumn;
                orderColumn = mainDs.Tables["Orders"].Columns["OrderID"];
                productColumn = mainDs.Tables["ProductOrders"].Columns["OrderID"];
                customerColumn = mainDs.Tables["Customers"].Columns["CustomerID"];
                customerOrderColumn = mainDs.Tables["Orders"].Columns["CustomerID"];

                productOrderProductIDColumn = mainDs.Tables["ProductOrders"].Columns["ProductID"];
                productProductIDColumn = mainDs.Tables["Product"].Columns["ProductID"];

                DataRelation dr2 = new DataRelation("Customer_Orders", customerColumn, customerOrderColumn);
                mainDs.Relations.Add(dr2);
                DataRelation dr = new DataRelation("Order_ProductOrders", orderColumn, productColumn);
                mainDs.Relations.Add(dr);
                DataRelation dr3 = new DataRelation("Product_ProductOrders", productProductIDColumn, productOrderProductIDColumn);
                mainDs.Relations.Add(dr3);
                return mainDs;
            }
            catch
            {
                return null;
            }
        }
        //ExEnd:GetAllDataXML
        #endregion

        #region GetSingleRowXML
        //ExStart:GetSingleCustomerXML
        /// <summary>
        /// Fetches information from the xml file and add it to the DataSet
        /// </summary>
        /// <returns>Returns DataSet, first record from the table</returns>
        public static DataRow GetSingleCustomerXML()
        {
            DataSet singleCustomer = new DataSet();

            FileStream readProductXML = new FileStream(customerXMLfile, FileMode.Open);
            singleCustomer.ReadXml(readProductXML, XmlReadMode.ReadSchema);
            singleCustomer.Tables[0].TableName = "Customers";

            return singleCustomer.Tables["Customers"].Rows[0];
        }
        //ExEnd:GetSingleCustomerXML
        #endregion


        #region GetCustomerDataJson
        //ExStart:GetCustomerDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data
        /// </summary>
        /// <returns>Returns deserialized data</returns>
        public static IEnumerable<BusinessObjects.Customer> GetCustomerDataFromJson()
        {
            string rawData = File.ReadAllText(jsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                yield return customer;
            }
        }
        //ExEnd:GetCustomerDataJson
        #endregion

        #region GetCustomerOrderDataJson
        //ExStart:GetCustomerOrderDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data
        /// </summary>
        /// <returns>Returns deserialized data</returns>
        public static IEnumerable<BusinessObjects.Order> GetCustomerOrderDataFromJson()
        {
            string rawData = File.ReadAllText(jsonFile);
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
        #endregion

        #region GetProductsJson
        //ExStart:GetProductsDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data
        /// </summary>
        /// <returns>Returns deserialized data</returns>
        public static IEnumerable<BusinessObjects.Product> GetProductsDataJson()
        {
            string rawData = File.ReadAllText(jsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            foreach (BusinessObjects.Customer customer in customers)
            {
                foreach (BusinessObjects.Order order in customer.Order)
                    yield return order.Product;
            }
        }
        //ExEnd:GetProductsDataJson
        #endregion

        #region GetSingleCustomerDataJson
        //ExStart:GetSingleCustomerDataJson
        /// <summary>
        /// Deserializes the json file, loop over the deserialized data
        /// </summary>
        /// <returns>Returns deserialized data</returns>
        public static BusinessObjects.Customer GetSingleCustomerDataJson()
        {
            string rawData = File.ReadAllText(jsonFile);
            BusinessObjects.Customer[] customers = JsonConvert.DeserializeObject<BusinessObjects.Customer[]>(rawData);

            IEnumerator<BusinessObjects.Customer> customer = GetCustomerDataFromJson().GetEnumerator();
            customer.MoveNext();
            return customer.Current;
        }
        //ExEnd:GetSingleCustomerDataJson
        #endregion
        /// <summary>
        /// Generate report from excel data source
        /// </summary>
        /// <returns></returns>
        public static GroupDocs.Assembly.Data.DocumentTable ExcelData()
        {
            DocumentTableOptions options = new DocumentTableOptions();
            options.FirstRowContainsColumnNames = true;

            // Use data of the _first_ worksheet.
            DocumentTable table = new DocumentTable(excelDataFile, 0, options);

            // Check column count, names, and types.
            Debug.Assert(table.Columns.Count == 3);

            Debug.Assert(table.Columns[0].Name == "Client");
            Debug.Assert(table.Columns[0].Type == typeof(string));

            Debug.Assert(table.Columns[1].Name == "Manager");
            Debug.Assert(table.Columns[1].Type == typeof(string));

            // NOTE: A space is replaced with an underscore, because spaces are not allowed in column names.
            Debug.Assert(table.Columns[2].Name == "Contract_Price");

            // NOTE: The type of the column is double, because all cells in the column contain numeric values.
            Debug.Assert(table.Columns[2].Type == typeof(double));
            return table;
        }
        /// <summary>
        /// Import word doc to presentation
        /// </summary>
        /// <returns></returns>
        public static GroupDocs.Assembly.Data.DocumentTable ImportingWordDocToPresentation()
        {

            // Do not extract column names from the first row, so that the first row to be treated as a data row.
            // Limit the largest row index, so that only the first four data rows to be loaded.
            DocumentTableOptions options = new DocumentTableOptions();
            options.MaxRowIndex = 3;

            // Use data of the _second_ table in the document.
            DocumentTable table = new DocumentTable(wordDataFile, 1, options);

            // Check column count and names.
            Debug.Assert(table.Columns.Count == 2);

            // NOTE: Default column names are used, because we do not extract the names from the first row.
            Debug.Assert(table.Columns[0].Name == "Column1");
            Debug.Assert(table.Columns[1].Name == "Column2");
            return table;
        }
    }
    //ExEnd:DataLayer
    #endregion DataLayer
    #region LazyAndRecursiveAccess
    //ExStart:DynamicEntity
    public interface IPropertyProvider<T>
    {
        T this[string propertyName] { get; }
    }

    public class DynamicEntity : IPropertyProvider<string>
    {
        /// <summary>
        /// Gets a property value by its name.
        /// </summary>
        public string this[string propertyName]
        {
            get
            {
                // In this example, we simply return a property name as its value.
                // In a real-life application, a real property value should be returned.
                // This value can be cached using for example a Dictionary, or fetched
                // every time the property is requested.
                return propertyName + " Value";
            }
        }

        /// <summary>
        /// Provides access to individual related <see cref="DynamicEntity"/> instances
        /// by their names.
        /// </summary>
        public IPropertyProvider<DynamicEntity> Entities
        {
            get { return mEntities; }
        }

        /// <summary>
        /// Provides access to enumerations of related <see cref="DynamicEntity"/> instances
        /// by their names.
        /// </summary>
        public IPropertyProvider<IEnumerable<DynamicEntity>> Children
        {
            get { return mChildren; }
        }

        private class ReferencedEntities : IPropertyProvider<DynamicEntity>
        {
            public DynamicEntity this[string propertyName]
            {
                get
                {
                    // In this example, we simply return the root entity.
                    // In a real-life application, a DynamicEntity instance corresponding
                    // to propertyName for the given root entity should be returned.
                    // This instance can be cached using for example a Dictionary,
                    // or fetched every time the referenced entity is requested.
                    return mRootEntity;
                }
            }

            public ReferencedEntities(DynamicEntity rootEntity)
            {
                // The reference to the root entity allows to access fields of the root entity
                // (such as an identifier) in the above indexer for a real-life application.
                mRootEntity = rootEntity;
            }

            private readonly DynamicEntity mRootEntity;
        }

        private class ChildEntities : IPropertyProvider<IEnumerable<DynamicEntity>>
        {
            public IEnumerable<DynamicEntity> this[string propertyName]
            {
                get
                {
                    // In this example, we simply return the root entity three times.
                    // In a real-life application, an enumeration of DynamicEntity instances
                    // corresponding to propertyName for the given root entity should be returned.
                    // This enumeration can be cached using for example a Dictionary,
                    // or fetched every time the child entities are requested.
                    yield return mRootEntity;
                    yield return mRootEntity;
                    yield return mRootEntity;
                }
            }

            public ChildEntities(DynamicEntity rootEntity)
            {
                // The reference to the root entity allows to access fields of the root entity
                // (such as an identifier) in the above indexer for a real-life application.
                mRootEntity = rootEntity;
            }

            private readonly DynamicEntity mRootEntity;
        }
        public DynamicEntity(Guid id)
        {
            // In this example, we use Guid to represent an entity identifier.
            // In a real-life application, the identifier can be of any type or even missing.
            mId = id;

            // In this example, we simply initialize fields in the constructor.
            // In a real-life application, these fields can be initialized lazily
            // at the corresponding properties, if needed.
            mEntities = new ReferencedEntities(this);
            mChildren = new ChildEntities(this);
        }

        private readonly Guid mId;
        private readonly ReferencedEntities mEntities;
        private readonly ChildEntities mChildren;
    }
    //ExEnd:DynamicEntity
    #endregion

}

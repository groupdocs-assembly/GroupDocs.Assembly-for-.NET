Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects
Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects.GroupDocs.AssemblyExamples.ProjectEntities
Imports System.Reflection
Imports System.IO
Imports Newtonsoft.Json
Imports GroupDocs.Assembly.Data


Namespace GroupDocs.AssemblyExamples.BusinessLayer
#Region "DataLayer"
    'ExStart:DataLayer
    Public NotInheritable Class DataLayer
        Private Sub New()
        End Sub

        Public Const productXMLfile As String = "../../../../Data/XML DataSource/Products.xml"
        Public Const customerXMLfile As String = "../../../../Data/XML DataSource/Customers.xml"
        Public Const orderXMLfile As String = "../../../../Data/XML DataSource/Orders.xml"
        Public Const productOrderXMLfile As String = "../../../../Data/XML DataSource/ProductOrders.xml"
        Public Const jsonFile As String = "../../../../Data/Data Sources/JSON DataSource/CustomerData-Json.json"
        Public Const wordDataFile As String = "../../../../Data/Data Sources/Word DataSource/Managers Data.docx"
        Public Const excelDataFile As String = "../../../../Data/Data Sources/Excel DataSource/Contracts Data.xlsx"
        Public Const presentationDataFile As String = "../../../../Data/Data Sources/Presentation DataSource/Managers Data.pptx"



#Region "DataInitialization"
        'ExStart:PopulateData
        ''' <summary>
        ''' This function initializes/populates the data. 
        ''' Initialize Customer information, product information and order information
        ''' </summary>
        ''' <returns>Returns customer's complete information</returns>
        Public Shared Iterator Function PopulateData() As IEnumerable(Of BusinessObjects.Customer)
            Dim customer As New BusinessObjects.Customer() With {
                 .CustomerName = "Atir Tahir",
                 .CustomerContactNumber = "+9211874",
                 .ShippingAddress = "Flat # 1, Kiyani Plaza ISB",
                 .Barcode = "123456789qwertyu0025"}



            customer.Order = New BusinessObjects.Order() {New BusinessObjects.Order() With {
                 .Product = New BusinessObjects.Product() With {
                     .ProductName = "Lumia 525"
                },
                 .Customer = customer,
                 .Price = 170,
                 .ProductQuantity = 5,
                 .OrderDate = New DateTime(2015, 1, 1)
            }}
            Yield customer
            'yield return statement will return one data at a time

            customer = New BusinessObjects.Customer() With {
                 .CustomerName = "Usman Aziz",
                 .CustomerContactNumber = "+458789",
                 .ShippingAddress = "Quette House, Park Road, ISB",
                 .Barcode = "123456789qwertyu0025"}


            customer.Order = New BusinessObjects.Order() {New BusinessObjects.Order() With {
                 .Product = New BusinessObjects.Product() With {
                     .ProductName = "Lenovo G50"
                },
                 .Customer = customer,
                 .Price = 480,
                 .ProductQuantity = 2,
                 .OrderDate = New DateTime(2015, 2, 1)
            }, New BusinessObjects.Order() With {
                 .Product = New BusinessObjects.Product() With {
                     .ProductName = "Pavilion G6"
                },
                 .Customer = customer,
                 .Price = 400,
                 .ProductQuantity = 1,
                 .OrderDate = New DateTime(2015, 10, 1)
            }, New BusinessObjects.Order() With {
                 .Product = New BusinessObjects.Product() With {
                     .ProductName = "Nexus 5"
                },
                 .Customer = customer,
                 .Price = 320,
                 .ProductQuantity = 3,
                 .OrderDate = New DateTime(2015, 6, 1)
            }}
            Yield customer
            'yield return statement will return one data at a time 
        End Function
        'ExEnd:PopulateData
#End Region

#Region "GetOrders"
        'ExStart:GetOrdersData
        ''' <summary>
        ''' Fetches order from PopulateData
        ''' </summary>
        ''' <returns>Returns order details, one data at a time</returns>
        Public Shared Iterator Function GetOrdersData() As IEnumerable(Of BusinessObjects.Order)
            For Each customer As BusinessObjects.Customer In PopulateData()
                For Each order As BusinessObjects.Order In customer.Order
                    Yield order
                    'yield return statement returns one data at a time
                Next
            Next
        End Function
        'ExEnd:GetOrdersData
#End Region

#Region "GetProducts"
        'ExStart:GetProductsData
        ''' <summary>
        ''' Fetches products from PopulateData
        ''' </summary>
        ''' <returns>Returns product details, one data at a time</returns>
        Public Shared Iterator Function GetProductsData() As IEnumerable(Of BusinessObjects.Product)
            For Each customer As BusinessObjects.Customer In PopulateData()
                For Each order As BusinessObjects.Order In customer.Order
                    Yield order.Product
                Next
            Next
        End Function
        'ExEnd:GetProductsData
#End Region

#Region "GetSingleCustomerData"
        'ExStart:GetSingleCustomerData
        ''' <summary>
        ''' Fetches customer details of very first customer
        ''' </summary>
        ''' <returns>Returns first customer's infromation</returns>
        Public Shared Function GetCustomerData() As BusinessObjects.Customer
            Dim customer As IEnumerator(Of BusinessObjects.Customer) = PopulateData().GetEnumerator()
            customer.MoveNext()

            Return customer.Current
        End Function
        'ExEnd:GetSingleCustomerData
#End Region

#Region "GetOrdersDataDB"
        'ExStart:GetOrdersDataDB
        ''' <summary>
        ''' Fetches orders from database
        ''' </summary>
        ''' <returns>Returns order details, one data at a time</returns>
        Public Shared Function GetOrdersDataDB() As IEnumerable(Of Order)
            'create object of data context
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Return dbEntities.Orders
        End Function
        'ExEnd:GetOrdersDataDB
#End Region

#Region "GetProductsDataDB"
        'ExStart:GetProductsDataDB
        ''' <summary>
        ''' Fetches products from database 
        ''' </summary>
        ''' <returns>Returns products information, one data at a time </returns>
        Public Shared Function GetProductsDataDB() As IEnumerable(Of Product)
            'create object of data context
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Return dbEntities.Products
        End Function
        'ExEnd:GetProductsDataDB
#End Region

#Region "GetCustomersDataDB"
        'ExStart:GetCustomersDataDB
        ''' <summary>
        ''' Fetches customers from database
        ''' </summary>
        ''' <returns>Returns customers information, one data at a time</returns>
        Public Shared Function GetCustomersDataDB() As IEnumerable(Of Customer)
            'create object of data context
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Return dbEntities.Customers
        End Function
        'ExEnd:GetCustomersDataDB
#End Region

#Region "GetSingleCustomerDataDB"
        'ExStart:GetSingleCustomerDataDB
        ''' <summary>
        ''' Fetches single customer data
        ''' </summary>
        ''' <returns>Return single, first customer's data</returns>
        Public Shared Function GetSingleCustomerDataDB() As Customer
            'create object of data context
            Dim dbEntites As New DatabaseEntitiesDataContext()
            Dim customer As IEnumerator(Of Customer) = GetCustomersDataDB().GetEnumerator()
            customer.MoveNext()
            Return customer.Current
        End Function
        'ExEnd:GetSingleCustomerDataDB
#End Region

#Region "GetSingleCustomerDataDT"
        'ExStart:GetSingleCustomerDT
        ''' <summary>
        ''' Fetches Customers from database
        ''' </summary>
        ''' <returns>Returns DataSet, very first record from the table</returns>
        Public Shared Function GetSingleCustomerDT() As DataRow
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Dim customersQuery = (From c In dbEntities.Customers Select New With {
        c.CustomerID,
        c.CustomerName,
        c.ShippingAddress,
        c.CustomerContactNumber,
        c.Photo
    }).AsEnumerable()
            Dim Customers As New DataTable()
            'ToADOTable function converts query results into DataTable
            Customers = customersQuery.ToADOTable(Function(rec) New Object() {customersQuery})
            Customers.TableName = "Customers"
            Dim dataSet As New DataSet()
            'Adding DataTable in DataSet
            dataSet.Tables.Add(Customers)
            Return dataSet.Tables("Customers").Rows(0)
        End Function
        'ExEnd:GetSingleCustomerDT
#End Region

#Region "GetCustomersAndOrdersDataDT"
        'ExStart:GetCustomersAndOrdersDataDT
        ''' <summary>
        ''' Fetches Orders, ProductOrders and Customers from database
        ''' </summary>
        ''' <returns>Returns DataSet</returns>
        Public Shared Function GetCustomersAndOrdersDataDT() As DataSet
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Dim ordersQuery = (From c In dbEntities.Orders).AsEnumerable()
            Dim productOrdersQuery = (From c In dbEntities.ProductOrders).AsEnumerable()
            Dim customersQuery = (From c In dbEntities.Customers).AsEnumerable()
            Dim Orders As New DataTable()
            'ToADOTable function converts query results into DataTable
            Orders = ordersQuery.ToADOTable(Function(rec) New Object() {ordersQuery})
            Dim ProductOrders As New DataTable()
            ProductOrders = productOrdersQuery.ToADOTable(Function(rec) New Object() {productOrdersQuery})
            Dim Customers As New DataTable()
            Customers = customersQuery.ToADOTable(Function(rec) New Object() {customersQuery})
            ProductOrders.TableName = "ProductOrder"
            Orders.TableName = "Orders"
            Customers.TableName = "Customers"
            Dim dataSet As New DataSet()
            'Adding DataTable in DataSet
            dataSet.Tables.Add(Orders)
            dataSet.Tables.Add(ProductOrders)
            dataSet.Tables.Add(Customers)
            Return dataSet
        End Function
        'ExEnd:GetCustomersAndOrdersDataDT
#End Region

#Region "GetProductsDataDT"
        'ExStart:GetProductsDataDT
        ''' <summary>
        ''' Fetches Products and ProductOrders information, store them in DataTables and load DataTable to DataSet
        ''' </summary>
        ''' <returns>Returns DataSet</returns>
        Public Shared Function GetProductsDT() As DataSet
            Dim dbEntities As New DatabaseEntitiesDataContext()
            Dim productsQuery = (From c In dbEntities.Products).AsEnumerable()
            Dim productOrdersQuery = (From c In dbEntities.ProductOrders).AsEnumerable()
            Dim Products As New DataTable()
            'ToADOTable function converts query results into DataTable
            Products = productsQuery.ToADOTable(Function(rec) New Object() {productsQuery})
            Dim ProductOrders As New DataTable()
            ProductOrders = productOrdersQuery.ToADOTable(Function(rec) New Object() {productOrdersQuery})
            ProductOrders.TableName = "ProductOrder"
            Products.TableName = "products"
            Dim dataSet As New DataSet()
            'Adding DataTable in DataSet
            dataSet.Tables.Add(Products)
            dataSet.Tables.Add(ProductOrders)
            Return dataSet
        End Function
        'ExEnd:GetProductsDataDT
#End Region

#Region "GetAllDataFromXML"
        'ExStart:GetAllDataXML
        ''' <summary>
        ''' Fetches the data from the XML files and store data in the DataSet
        ''' </summary>
        ''' <returns>Returns the DataSet</returns>
        Public Shared Function GetAllDataFromXML() As DataSet
            Try
                Dim tempDs As New DataSet()
                Dim mainDs As New DataSet()


                Dim fsReadXml As New System.IO.FileStream(customerXMLfile, System.IO.FileMode.Open)

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
                tempDs.Tables(0).TableName = "Customers"

                mainDs.Merge(tempDs.Tables(0))
                tempDs = New DataSet()

                fsReadXml = New System.IO.FileStream(orderXMLfile, System.IO.FileMode.Open)

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
                tempDs.Tables(0).TableName = "Orders"

                mainDs.Merge(tempDs.Tables(0))
                tempDs = New DataSet()

                fsReadXml = New System.IO.FileStream(productOrderXMLfile, System.IO.FileMode.Open)

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
                tempDs.Tables(0).TableName = "ProductOrders"

                mainDs.Merge(tempDs.Tables(0))
                tempDs = New DataSet()

                fsReadXml = New System.IO.FileStream(productXMLfile, System.IO.FileMode.Open)

                tempDs.ReadXml(fsReadXml, XmlReadMode.ReadSchema)
                tempDs.Tables(0).TableName = "Product"

                mainDs.Merge(tempDs.Tables(0))

                'Defining relations between the tables
                Dim productColumn As DataColumn, orderColumn As DataColumn, customerColumn As DataColumn, customerOrderColumn As DataColumn, productProductIDColumn As DataColumn, productOrderProductIDColumn As DataColumn
                orderColumn = mainDs.Tables("Orders").Columns("OrderID")
                productColumn = mainDs.Tables("ProductOrders").Columns("OrderID")
                customerColumn = mainDs.Tables("Customers").Columns("CustomerID")
                customerOrderColumn = mainDs.Tables("Orders").Columns("CustomerID")

                productOrderProductIDColumn = mainDs.Tables("ProductOrders").Columns("ProductID")
                productProductIDColumn = mainDs.Tables("Product").Columns("ProductID")

                Dim dr2 As New DataRelation("Customer_Orders", customerColumn, customerOrderColumn)
                mainDs.Relations.Add(dr2)
                Dim dr As New DataRelation("Order_ProductOrders", orderColumn, productColumn)
                mainDs.Relations.Add(dr)
                Dim dr3 As New DataRelation("Product_ProductOrders", productProductIDColumn, productOrderProductIDColumn)
                mainDs.Relations.Add(dr3)
                Return mainDs
            Catch
                Return Nothing
            End Try
        End Function
        'ExEnd:GetAllDataXML
#End Region

#Region "GetSingleRowXML"
        'ExStart:GetSingleCustomerXML
        ''' <summary>
        ''' Fetches information from the xml file and add it to the DataSet
        ''' </summary>
        ''' <returns>Returns DataSet, first record from the table</returns>
        Public Shared Function GetSingleCustomerXML() As DataRow
            Dim singleCustomer As New DataSet()

            Dim readProductXML As New FileStream(customerXMLfile, FileMode.Open)
            singleCustomer.ReadXml(readProductXML, XmlReadMode.ReadSchema)
            singleCustomer.Tables(0).TableName = "Customers"

            Return singleCustomer.Tables("Customers").Rows(0)
        End Function
        'ExEnd:GetSingleCustomerXML
#End Region

#Region "GetCustomerDataJson"
        'ExStart:GetCustomerDataJson
        ''' <summary>
        ''' Deserializes the json file, loop over the deserialized data
        ''' </summary>
        ''' <returns>Returns deserialized data</returns>
        Public Shared Iterator Function GetCustomerDataFromJson() As IEnumerable(Of BusinessObjects.Customer)
            Dim rawData As String = File.ReadAllText(jsonFile)
            Dim customers As BusinessObjects.Customer() = JsonConvert.DeserializeObject(Of BusinessObjects.Customer())(rawData)

            For Each customer As BusinessObjects.Customer In customers
                Yield customer
            Next
        End Function
        'ExEnd:GetCustomerDataJson
#End Region

#Region "GetCustomerOrderDataJson"
        'ExStart:GetCustomerOrderDataJson
        ''' <summary>
        ''' Deserializes the json file, loop over the deserialized data
        ''' </summary>
        ''' <returns>Returns deserialized data</returns>
        Public Shared Iterator Function GetCustomerOrderDataFromJson() As IEnumerable(Of BusinessObjects.Order)
            Dim rawData As String = File.ReadAllText(jsonFile)
            Dim customers As BusinessObjects.Customer() = JsonConvert.DeserializeObject(Of BusinessObjects.Customer())(rawData)

            For Each customer As BusinessObjects.Customer In customers
                For Each order As BusinessObjects.Order In customer.Order
                    Yield order
                Next
            Next
        End Function
        'ExEnd:GetCustomerOrderDataJson
#End Region

#Region "GetProductsJson"
        'ExStart:GetProductsDataJson
        ''' <summary>
        ''' Deserializes the json file, loop over the deserialized data
        ''' </summary>
        ''' <returns>Returns deserialized data</returns>
        Public Shared Iterator Function GetProductsDataJson() As IEnumerable(Of BusinessObjects.Product)
            Dim rawData As String = File.ReadAllText(jsonFile)
            Dim customers As BusinessObjects.Customer() = JsonConvert.DeserializeObject(Of BusinessObjects.Customer())(rawData)

            For Each customer As BusinessObjects.Customer In customers
                For Each order As BusinessObjects.Order In customer.Order
                    Yield order.Product
                Next
            Next
        End Function
        'ExEnd:GetProductsDataJson
#End Region

#Region "GetSingleCustomerDataJson"
        'ExStart:GetSingleCustomerDataJson
        ''' <summary>
        ''' Deserializes the json file, loop over the deserialized data
        ''' </summary>
        ''' <returns>Returns deserialized data</returns>
        Public Shared Function GetSingleCustomerDataJson() As BusinessObjects.Customer
            Dim rawData As String = File.ReadAllText(jsonFile)
            Dim customers As BusinessObjects.Customer() = JsonConvert.DeserializeObject(Of BusinessObjects.Customer())(rawData)

            Dim customer As IEnumerator(Of BusinessObjects.Customer) = GetCustomerDataFromJson().GetEnumerator()
            customer.MoveNext()
            Return customer.Current
        End Function
        'ExEnd:GetSingleCustomerDataJson
#End Region

        ''' <summary>
        ''' Generate report from excel data source
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function ExcelData() As Global.GroupDocs.Assembly.Data.DocumentTable
            Dim options As New DocumentTableOptions()
            options.FirstRowContainsColumnNames = True

            ' Use data of the _first_ worksheet.
            Dim table As New DocumentTable(excelDataFile, 0, options)

            ' Check column count, names, and types.
            Debug.Assert(table.Columns.Count = 3)

            Debug.Assert(table.Columns(0).Name = "Client")
            Debug.Assert(table.Columns(0).Type = GetType(String))

            Debug.Assert(table.Columns(1).Name = "Manager")
            Debug.Assert(table.Columns(1).Type = GetType(String))

            ' NOTE: A space is replaced with an underscore, because spaces are not allowed in column names.
            Debug.Assert(table.Columns(2).Name = "Contract_Price")

            ' NOTE: The type of the column is double, because all cells in the column contain numeric values.
            Debug.Assert(table.Columns(2).Type = GetType(Double))
            Return table
        End Function

        ''' <summary>
        ''' Import word doc to presentation
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function ImportingWordDocToPresentation() As Global.GroupDocs.Assembly.Data.DocumentTable

            ' Do not extract column names from the first row, so that the first row to be treated as a data row.
            ' Limit the largest row index, so that only the first four data rows to be loaded.
            Dim options As New DocumentTableOptions()
            options.MaxRowIndex = 3

            ' Use data of the _second_ table in the document.
            Dim table As New DocumentTable(wordDataFile, 1, options)

            ' Check column count and names.
            Debug.Assert(table.Columns.Count = 2)

            ' NOTE: Default column names are used, because we do not extract the names from the first row.
            Debug.Assert(table.Columns(0).Name = "Column1")
            Debug.Assert(table.Columns(1).Name = "Column2")
            Return table
        End Function


        ''' <summary>
        ''' Import spread sheet to html document
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function ImportingSpreadsheetToHtml() As Global.GroupDocs.Assembly.Data.DocumentTable

            ' Do not extract column names from the first row, so that the first row to be treated as a data row.
            ' Limit the largest row index, so that only the first four data rows to be loaded.
            Dim options As New DocumentTableOptions()
            'options.MaxRowIndex = 3;

            ' Use data of the _second_ table in the document.
            Dim table As New DocumentTable(excelDataFile, 0)

            ' Check column count and names.
            Debug.Assert(table.Columns.Count = 3)

            ' NOTE: Default column names are used, because we do not extract the names from the first row.
            Debug.Assert(table.Columns(0).Name = "A")
            Debug.Assert(table.Columns(1).Name = "B")
            Debug.Assert(table.Columns(2).Name = "C")
            Return table
        End Function



        ''' <summary>
        ''' Presentation file data source
        ''' </summary>
        ''' <returns></returns>
        Public Shared Function PresentationData() As Global.GroupDocs.Assembly.Data.DocumentTable

            ' Do not extract column names from the first row, so that the first row to be treated as a data row.
            ' Limit the largest row index, so that only the first four data rows to be loaded.
            Dim options As New DocumentTableOptions()
            options.MaxRowIndex = 3

            ' Use data of the _second_ table in the document.
            Dim table As New DocumentTable(presentationDataFile, 1, options)

            ' Check column count and names.
            Debug.Assert(table.Columns.Count = 2)

            ' NOTE: Default column names are used, because we do not extract the names from the first row.
            Debug.Assert(table.Columns(0).Name = "Column1")
            Debug.Assert(table.Columns(1).Name = "Column2")
            Return table
        End Function



        ''' <summary>
        ''' Creates an Email data source object
        ''' </summary>
        ''' <param name="fileName">Name of the template file</param>
        ''' <param name="dataSource">data source</param>
        ''' <returns></returns>

        Public Shared Function EmailDataSourceObject(fileName As String, dataSource As Object) As Object()
            'ExStart:EmailDataSourceObject
            Dim dataSources As Object()
            Dim extension As String = Path.GetExtension(fileName)

            If (extension = ".msg") OrElse (extension = ".eml") Then
                Dim recipients As New List(Of String)()
                recipients.Add("Named Recipient <named@example.com>")
                recipients.Add("unnamed@example.com")


                dataSources = New Object() {dataSource, "Example Sender <sender@example.com>", recipients, "cc@example.com", Path.GetFileNameWithoutExtension(fileName)}
            Else
                dataSources = New Object() {dataSource}
            End If
            Return dataSources
            'ExEnd:EmailDataSourceObject
        End Function

        Public Shared Function EmailDataSourceName(extension As String, name As String) As String()
            'ExStart:EmailDataSourceName
            Dim dataSourceNames As String()
            If (extension = ".msg") OrElse (extension = ".eml") Then

                dataSourceNames = New String() {name, "sender", "recipients", "cc", "subject"}
            Else
                dataSourceNames = New String() {}
            End If
            Return dataSourceNames
            'ExEnd:EmailDataSourceName
        End Function



    End Class
    'ExEnd:DataLayer
#End Region
#Region "LazyAndRecursiveAccess"
    'ExStart:RecursiveandLazyAccessOfData
    Public Interface IPropertyProvider(Of T)
        Default ReadOnly Property Item(propertyName As String) As T
    End Interface

    Public Class DynamicEntity
        Implements IPropertyProvider(Of String)
        ''' <summary>
        ''' Gets a property value by its name.
        ''' </summary>
        Default Public ReadOnly Property Item(propertyName As String) As String Implements IPropertyProvider(Of String).Item
            Get
                ' In this example, we simply return a property name as its value.
                ' In a real-life application, a real property value should be returned.
                ' This value can be cached using for example a Dictionary, or fetched
                ' every time the property is requested.
                Return propertyName & Convert.ToString(" Value")
            End Get
        End Property

        ''' <summary>
        ''' Provides access to individual related <see cref="DynamicEntity"/> instances
        ''' by their names.
        ''' </summary>
        Public ReadOnly Property Entities() As IPropertyProvider(Of DynamicEntity)
            Get
                Return mEntities
            End Get
        End Property

        ''' <summary>
        ''' Provides access to enumerations of related <see cref="DynamicEntity"/> instances
        ''' by their names.
        ''' </summary>
        Public ReadOnly Property Children() As IPropertyProvider(Of IEnumerable(Of DynamicEntity))
            Get
                Return mChildren
            End Get
        End Property

        Private Class ReferencedEntities
            Implements IPropertyProvider(Of DynamicEntity)
            Default Public ReadOnly Property Item(propertyName As String) As DynamicEntity Implements IPropertyProvider(Of Global.GroupDocs.AssemblyExamples.BusinessLayer.GroupDocs.AssemblyExamples.BusinessLayer.DynamicEntity).Item
                Get
                    ' In this example, we simply return the root entity.
                    ' In a real-life application, a DynamicEntity instance corresponding
                    ' to propertyName for the given root entity should be returned.
                    ' This instance can be cached using for example a Dictionary,
                    ' or fetched every time the referenced entity is requested.
                    Return mRootEntity
                End Get
            End Property

            Public Sub New(rootEntity As DynamicEntity)
                ' The reference to the root entity allows to access fields of the root entity
                ' (such as an identifier) in the above indexer for a real-life application.
                mRootEntity = rootEntity
            End Sub

            Private ReadOnly mRootEntity As DynamicEntity
        End Class

        Private Class ChildEntities
            Implements IPropertyProvider(Of IEnumerable(Of DynamicEntity))
            Public ReadOnly Iterator Property Item(propertyName As String) As IEnumerable(Of DynamicEntity) Implements IPropertyProvider(Of System.Collections.Generic.IEnumerable(Of Global.GroupDocs.AssemblyExamples.BusinessLayer.GroupDocs.AssemblyExamples.BusinessLayer.DynamicEntity)).Item
                Get
                    ' In this example, we simply return the root entity three times.
                    ' In a real-life application, an enumeration of DynamicEntity instances
                    ' corresponding to propertyName for the given root entity should be returned.
                    ' This enumeration can be cached using for example a Dictionary,
                    ' or fetched every time the child entities are requested.
                    Yield mRootEntity
                    Yield mRootEntity
                    Yield mRootEntity
                End Get
            End Property

            Public Sub New(rootEntity As DynamicEntity)
                ' The reference to the root entity allows to access fields of the root entity
                ' (such as an identifier) in the above indexer for a real-life application.
                mRootEntity = rootEntity
            End Sub

            Private ReadOnly mRootEntity As DynamicEntity
        End Class
        Public Sub New(id As Guid)
            ' In this example, we use Guid to represent an entity identifier.
            ' In a real-life application, the identifier can be of any type or even missing.
            mId = id

            ' In this example, we simply initialize fields in the constructor.
            ' In a real-life application, these fields can be initialized lazily
            ' at the corresponding properties, if needed.
            mEntities = New ReferencedEntities(Me)
            mChildren = New ChildEntities(Me)
        End Sub

        Private ReadOnly mId As Guid
        Private ReadOnly mEntities As ReferencedEntities
        Private ReadOnly mChildren As ChildEntities
    End Class
    'ExEnd:RecursiveandLazyAccessOfData
#End Region
End Namespace



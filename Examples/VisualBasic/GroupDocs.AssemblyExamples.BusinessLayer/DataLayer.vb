Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects
Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects.GroupDocs.AssemblyExamples.ProjectEntities
Imports System.Reflection


Namespace GroupDocs.AssemblyExamples.BusinessLayer
    'ExStart:DataLayer
    Public NotInheritable Class DataLayer
        Private Sub New()
        End Sub
#Region "DataInitialization"
        'ExStart:PopulateData
        ''' <summary>
        ''' This function initializes/populates the data. 
        ''' Initialize Customer information, product information and order information
        ''' </summary>
        ''' <returns>Returns customer's complete information</returns>
        Public Shared Iterator Function PopulateData() As IEnumerable(Of BusinessObjects.Customer)
            Dim customer As New BusinessObjects.Customer() With { _
                 .CustomerName = "Atir Tahir", _
                 .CustomerContactNumber = "+9211874", _
                 .ShippingAddress = "Flat # 1, Kiyani Plaza ISB" _
            }

            customer.Order = New BusinessObjects.Order() {New BusinessObjects.Order() With { _
                 .Product = New BusinessObjects.Product() With { _
                     .ProductName = "Lumia 525" _
                }, _
                 .Customer = customer, _
                 .Price = 170, _
                 .ProductQuantity = 5, _
                 .OrderDate = New DateTime(2015, 1, 1) _
            }}
            Yield customer
            'yield return statement will return one data at a time

            customer = New BusinessObjects.Customer() With { _
                 .CustomerName = "Usman Aziz", _
                 .CustomerContactNumber = "+458789", _
                 .ShippingAddress = "Quette House, Park Road, ISB" _
            }
            customer.Order = New BusinessObjects.Order() {New BusinessObjects.Order() With { _
                 .Product = New BusinessObjects.Product() With { _
                     .ProductName = "Lenovo G50" _
                }, _
                 .Customer = customer, _
                 .Price = 480, _
                 .ProductQuantity = 2, _
                 .OrderDate = New DateTime(2015, 2, 1) _
            }, New BusinessObjects.Order() With { _
                 .Product = New BusinessObjects.Product() With { _
                     .ProductName = "Pavilion G6" _
                }, _
                 .Customer = customer, _
                 .Price = 400, _
                 .ProductQuantity = 1, _
                 .OrderDate = New DateTime(2015, 10, 1) _
            }, New BusinessObjects.Order() With { _
                 .Product = New BusinessObjects.Product() With { _
                     .ProductName = "Nexus 5" _
                }, _
                 .Customer = customer, _
                 .Price = 320, _
                 .ProductQuantity = 3, _
                 .OrderDate = New DateTime(2015, 6, 1) _
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
            Dim customersQuery = (From c In dbEntities.Customers Select New With { _
        c.CustomerID, _
        c.CustomerName, _
        c.ShippingAddress, _
        c.CustomerContactNumber, _
        c.Photo _
    }).AsEnumerable()
            Dim Customers As New DataTable()
            'ToADOTable function converts DataBase table into DataTable
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
            'ToADOTable function converts DataBase table into DataTable
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
            'ToADOTable function converts DataBase table into DataTable
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

    End Class
   
    'ExEnd:DataLayer
End Namespace




Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.IO

Namespace GroupDocs.AssemblyExamples.ProjectBusinessObjects
    Public Class BusinessObjects
        Public Shared imagePath As String = "../../../../Data/Images/"

        'ExStart:ProjectEntities
        Public Class Customer
            Public Property CustomerName() As [String]
                Get
                    Return m_CustomerName
                End Get
                Set(value As [String])
                    m_CustomerName = Value
                End Set
            End Property
            Private m_CustomerName As [String]
            Public Property ShippingAddress() As String
                Get
                    Return m_ShippingAddress
                End Get
                Set(value As String)
                    m_ShippingAddress = Value
                End Set
            End Property
            Private m_ShippingAddress As String
            Public Property CustomerContactNumber() As String
                Get
                    Return m_CustomerContactNumber
                End Get
                Set(value As String)
                    m_CustomerContactNumber = Value
                End Set
            End Property
            Private m_CustomerContactNumber As String
            Public Property Order() As IEnumerable(Of Order)
                Get
                    Return m_Order
                End Get
                Set(value As IEnumerable(Of Order))
                    m_Order = Value
                End Set
            End Property
            Private m_Order As IEnumerable(Of Order)
            Public ReadOnly Property Photo() As [String]
                Get
                    Return Path.Combine(Path.GetFullPath(imagePath), "no-photo.jpg")
                End Get
            End Property

        End Class

        Public Class Order
            Public Property Customer() As Customer
                Get
                    Return m_Customer
                End Get
                Set(value As Customer)
                    m_Customer = Value
                End Set
            End Property
            Private m_Customer As Customer
            Public Property Product() As Product
                Get
                    Return m_Product
                End Get
                Set(value As Product)
                    m_Product = Value
                End Set
            End Property
            Private m_Product As Product
            Public Property ProductQuantity() As Integer
                Get
                    Return m_ProductQuantity
                End Get
                Set(value As Integer)
                    m_ProductQuantity = Value
                End Set
            End Property
            Private m_ProductQuantity As Integer
            Public Property Price() As Integer
                Get
                    Return m_Price
                End Get
                Set(value As Integer)
                    m_Price = Value
                End Set
            End Property
            Private m_Price As Integer
            Public Property OrderDate() As DateTime
                Get
                    Return m_OrderDate
                End Get
                Set(value As DateTime)
                    m_OrderDate = Value
                End Set
            End Property
            Private m_OrderDate As DateTime
            Public Property OrderNumber() As Integer
                Get
                    Return m_OrderNumber
                End Get
                Set(value As Integer)
                    m_OrderNumber = Value
                End Set
            End Property
            Private m_OrderNumber As Integer
            Public Property ShippingDate() As DateTime
                Get
                    Return m_ShippingDate
                End Get
                Set(value As DateTime)
                    m_ShippingDate = Value
                End Set
            End Property
            Private m_ShippingDate As DateTime
        End Class

        Public Class Product
            Public Property ProductName() As String
                Get
                    Return m_ProductName
                End Get
                Set(value As String)
                    m_ProductName = Value
                End Set
            End Property
            Private m_ProductName As String
            Public Property UnitInStock() As Integer
                Get
                    Return m_UnitInStock
                End Get
                Set(value As Integer)
                    m_UnitInStock = Value
                End Set
            End Property
            Private m_UnitInStock As Integer
            Public Property Discount() As Integer
                Get
                    Return m_Discount
                End Get
                Set(value As Integer)
                    m_Discount = Value
                End Set
            End Property
            Private m_Discount As Integer
            Public Property ProductPrice() As String
                Get
                    Return m_ProductPrice
                End Get
                Set(value As String)
                    m_ProductPrice = Value
                End Set
            End Property
            Private m_ProductPrice As String
        End Class
        'ExEnd:ProjectEntities
    End Class
End Namespace

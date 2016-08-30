
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.IO

Namespace GroupDocs.AssemblyExamples.ProjectEntities
    Public Class BusinessObjects
        Public Shared imagePath As String = "../../../../Data/Images/"
        Public Shared docPath As String = "../../../../Data/OuterDocuments/"

        'ExStart:ProjectEntities
        Public Class Customer
            Public Property CustomerName() As [String]
                Get
                    Return m_CustomerName
                End Get
                Set(value As [String])
                    m_CustomerName = value
                End Set
            End Property
            Private m_CustomerName As [String]
            Public Property ShippingAddress() As String
                Get
                    Return m_ShippingAddress
                End Get
                Set(value As String)
                    m_ShippingAddress = value
                End Set
            End Property
            Private m_ShippingAddress As String
            Public Property CustomerContactNumber() As String
                Get
                    Return m_CustomerContactNumber
                End Get
                Set(value As String)
                    m_CustomerContactNumber = value
                End Set
            End Property
            Private m_CustomerContactNumber As String
            Public Property Order() As IEnumerable(Of Order)
                Get
                    Return m_Order
                End Get
                Set(value As IEnumerable(Of Order))
                    m_Order = value
                End Set
            End Property
            Private m_Order As IEnumerable(Of Order)

            Public Property Barcode() As String
                Get
                    Return m_Barcode
                End Get
                Set(value As String)
                    m_Barcode = Value
                End Set
            End Property
            Private m_Barcode As String

            Public ReadOnly Property Photo() As [String]
                Get
                    Return Path.Combine(Path.GetFullPath(imagePath), "no-photo.jpg")
                End Get
            End Property

            Public ReadOnly Property Document() As [String]
                Get
                    Return Path.Combine(Path.GetFullPath(docPath), "outerDoc.odt")
                End Get
            End Property

        End Class

        Public Class Order
            Public Property Customer() As Customer
                Get
                    Return m_Customer
                End Get
                Set(value As Customer)
                    m_Customer = value
                End Set
            End Property
            Private m_Customer As Customer
            Public Property Product() As Product
                Get
                    Return m_Product
                End Get
                Set(value As Product)
                    m_Product = value
                End Set
            End Property
            Private m_Product As Product
            Public Property ProductQuantity() As Integer
                Get
                    Return m_ProductQuantity
                End Get
                Set(value As Integer)
                    m_ProductQuantity = value
                End Set
            End Property
            Private m_ProductQuantity As Integer
            Public Property Price() As Integer
                Get
                    Return m_Price
                End Get
                Set(value As Integer)
                    m_Price = value
                End Set
            End Property
            Private m_Price As Integer

            Public Property Barcode() As String
                Get
                    Return m_Barcode
                End Get
                Set(value As String)
                    m_Barcode = Value
                End Set
            End Property
            Private m_Barcode As String

            Public Property OrderDate() As DateTime
                Get
                    Return m_OrderDate
                End Get
                Set(value As DateTime)
                    m_OrderDate = value
                End Set
            End Property
            Private m_OrderDate As DateTime
            Public Property OrderNumber() As Integer
                Get
                    Return m_OrderNumber
                End Get
                Set(value As Integer)
                    m_OrderNumber = value
                End Set
            End Property
            Private m_OrderNumber As Integer
            Public Property ShippingDate() As DateTime
                Get
                    Return m_ShippingDate
                End Get
                Set(value As DateTime)
                    m_ShippingDate = value
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
                    m_ProductName = value
                End Set
            End Property
            Private m_ProductName As String
            Public Property UnitInStock() As Integer
                Get
                    Return m_UnitInStock
                End Get
                Set(value As Integer)
                    m_UnitInStock = value
                End Set
            End Property
            Private m_UnitInStock As Integer
            Public Property Discount() As Integer
                Get
                    Return m_Discount
                End Get
                Set(value As Integer)
                    m_Discount = value
                End Set
            End Property
            Private m_Discount As Integer
            Public Property ProductPrice() As String
                Get
                    Return m_ProductPrice
                End Get
                Set(value As String)
                    m_ProductPrice = value
                End Set
            End Property
            Private m_ProductPrice As String
        End Class
        'ExEnd:ProjectEntities
    End Class
End Namespace

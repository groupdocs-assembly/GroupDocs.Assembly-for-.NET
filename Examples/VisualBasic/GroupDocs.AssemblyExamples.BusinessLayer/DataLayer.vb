﻿Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports GroupDocs.AssemblyExamples.ProjectBusinessObjects.GroupDocs.AssemblyExamples.ProjectBusinessObjects

Namespace GroupDocs.AssemblyExamples.BusinessLayer
    Public NotInheritable Class DataLayer
        Private Sub New()
        End Sub


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

#Region "GetOrders"
        'ExStart:GetOrdersData
        ''' <summary>
        ''' Fetches order details from PopulateData
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
        ''' Fetches product details from PopulateData
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

#Region "GetSingleCustomer"
        'ExStart:GetCustomerData
        ''' <summary>
        ''' Fetches customer details of very first customer
        ''' </summary>
        ''' <returns>Returns first customer's infromation</returns>
        Public Shared Function GetCustomerData() As BusinessObjects.Customer
            Dim customer As IEnumerator(Of BusinessObjects.Customer) = PopulateData().GetEnumerator()
            customer.MoveNext()

            Return customer.Current
        End Function
        'ExEnd:GetCustomerData
#End Region

        Private Shared Sub [yield](p1 As Object)
            Throw New NotImplementedException
        End Sub

    End Class
End Namespace

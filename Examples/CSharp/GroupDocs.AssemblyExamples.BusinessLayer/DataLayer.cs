using GroupDocs.AssemblyExamples.ProjectBusinessObjects;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace GroupDocs.AssemblyExamples.BusinessLayer
{
    public static class DataLayer
    {

        #region DataInitialization
        //ExStart:PopulateData
        /// <summary>
        /// This function initializes/populates the data. 
        /// Initialize Customer information, product information and order information
        /// </summary>
        /// <returns>Returns customer's complete information</returns>
        public static IEnumerable<BusinessObjects.Customer> PopulateData()
        {
            BusinessObjects.Customer customer = new BusinessObjects.Customer { CustomerName = "Atir Tahir", CustomerContactNumber = "+9211874", ShippingAddress = "Flat # 1, Kiyani Plaza ISB" };
            customer.Order = new BusinessObjects.Order[]
            {
                  new BusinessObjects.Order { Product = new BusinessObjects.Product { ProductName ="Lumia 525" }, Customer = customer, Price= 170, ProductQuantity = 5, OrderDate = new DateTime(2015, 1, 1) }

            };
            yield return customer; //yield return statement will return one data at a time


            customer = new BusinessObjects.Customer { CustomerName = "Usman Aziz", CustomerContactNumber = "+458789", ShippingAddress = "Quette House, Park Road, ISB" };
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
        /// Fetches order details from PopulateData
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
        /// Fetches product details from PopulateData
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

        #region GetSingleCustomer
        //ExStart:GetCustomerData
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
        //ExEnd:GetCustomerData
        #endregion
    }
}

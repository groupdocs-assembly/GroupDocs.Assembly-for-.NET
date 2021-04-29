using System;
using System.Collections.Generic;
using System.IO;

namespace AssemblyExamples.Data
{
    public class BusinessObjects : AssemblyExamplesBase
    {
        //ExStart:ProjectEntities
        public class Customer
        {
            public string CustomerName { get; set; }
            public string ShippingAddress { get; set; }
            public string CustomerContactNumber { get; set; }
            public IEnumerable<Order> Order { get; set; }
            public string Barcode { get; set; }
            public string Photo => Path.Combine(Path.GetFullPath(ImagesDir), "no-photo.jpg");
            public string Document => Path.Combine(Path.GetFullPath(TemplatesDir), "Outer document.docx");
            public string Color { get; set; }
        }

        public class Order
        {
            public Customer Customer { get; set; }
            public Product Product { get; set; }
            public int ProductQuantity { get; set; }
            public int Price { get; set; }
            public string Barcode { get; set; }
            public DateTime OrderDate { get; set; }
            public int OrderNumber { get; set; }
            public DateTime ShippingDate { get; set; }
            public IEnumerable<Service> Services { get; set; }
        }

        public class Product
        {
            public string ProductName { get; set; }
            public int UnitInStock { get; set; }
            public int Discount { get; set; }
            public string ProductPrice { get; set; }
        }

        public class Service
        {
            public string ServiceName { get; set; }
        }
        
        public class EmailDataSourcesObjects
        {
            public object DataSource { get; set; }
            public string Sender { get; set; }
            public List<string> Recipients { get; set; }
            public string CC { get; set; }
            public string Subject { get; set; }
            public string Title { get; set; }
        }

        public class EmailDataSourcesNames
        {
            public string Name { get; set; }
            public string Sender { get; set; }
            public string Recipients { get; set; }
            public string CC { get; set; }
            public string Subject { get; set; }
            public string Title { get; set; }
        }
        //ExEnd:ProjectEntities
    }
}

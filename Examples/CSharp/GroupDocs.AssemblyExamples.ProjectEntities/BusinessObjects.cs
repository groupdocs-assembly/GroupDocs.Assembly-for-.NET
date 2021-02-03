using System;
using System.Collections.Generic;
using System.IO;

namespace GroupDocs.AssemblyExamples.ProjectEntities
{
    public class BusinessObjects
    {
        public static string ImagePath = "../../../../Data/Images/";
        public static string DocPath = "../../../../Data/OuterDocuments/";

        //ExStart:ProjectEntities
        public class Customer
        {
            public string CustomerName { get; set; }
            public string ShippingAddress { get; set; }
            public string CustomerContactNumber { get; set; }
            public IEnumerable<Order> Order { get; set; }
            public string Barcode { get; set; }
            public string Photo => Path.Combine(Path.GetFullPath(ImagePath), "no-photo.jpg");
            public string Document => Path.Combine(Path.GetFullPath(DocPath), "OuterDocument.docx");
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
        //ExEnd:ProjectEntities
    }
}

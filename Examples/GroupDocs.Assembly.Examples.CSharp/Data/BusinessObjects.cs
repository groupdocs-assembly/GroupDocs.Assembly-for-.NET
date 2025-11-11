using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace GroupDocs.Assembly.Examples.CSharp.Data
{
    public class BusinessObjects
    {
        public class Customer
        {
            public string CustomerName { get; set; }
            public string ShippingAddress { get; set; }
            public string CustomerContactNumber { get; set; }
            public IEnumerable<Order> Order { get; set; } = new List<Order>();
            public string Barcode { get; set; }
            public string Photo
            {
                get
                {
                    string imagePath = Path.Combine(Constants.ImagesPath, "no-photo.jpg");
                    // Ensure we return an absolute path that's properly normalized
                    string fullPath = Path.GetFullPath(imagePath);
                    // Normalize the path to use forward slashes for URI compatibility
                    return new Uri(fullPath, UriKind.Absolute).LocalPath;
                }
            }
            public string Document
            {
                get
                {
                    string docPath = Path.Combine(Constants.TemplatesPath, "Outer document.docx");
                    return Path.GetFullPath(docPath);
                }
            }
            public string Color { get; set; }
        }

        public class Order
        {
            public Customer Customer { get; set; }
            public Product Product { get; set; }
            public int ProductQuantity { get; set; }
            public int Price { get; set; }
            public string Barcode { get; set; }
            public System.DateTime OrderDate { get; set; }
            public int OrderNumber { get; set; }
            public System.DateTime ShippingDate { get; set; }
            public IEnumerable<Service> Services { get; set; } = new List<Service>();
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
            public List<string> Recipients { get; set; } = new List<string>();
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
    }
}



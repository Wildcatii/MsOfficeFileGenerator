using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using Synvata.MsOfficeFileGenerator.Excel;

namespace Demo
{
	class Program
	{
		static void Main(string[] args)
		{
			List<Product> products = new List<Product>(3);
			products.Add(new Product()
			{
				Code = "P_A",
				Discount = 0.1F,
				Id = 1,
				Name = "F-15",
				Price = 6000.85m,
				AvailableDate = DateTime.Now
			});
			products.Add(new Product()
			{
				Code = "P_B",
				Discount = 0.2F,
				Id = 2,
				Name = "F-16",
				Price = 5500.50m,
				AvailableDate = DateTime.Now.AddDays(-3),
				IsOffline = true,
				NullableBool = true
			});
			products.Add(new Product()
			{
				Code = "P_C",
				Discount = 0.35F,
				Id = 3,
				Name = "F-18",
				Price = 5800.0m,
				AvailableDate = DateTime.Now.AddDays(-5)
			});

			string fileName = Path.Combine(Environment.CurrentDirectory, "Test.xlsx");

			// create .xlsx file as MemoryStream
			using (MemoryStream ms = ExcelGenerator.CreateStream(products, "Products"))
			{
				using (FileStream fs = File.Create(fileName))
				{
					ms.WriteTo(fs);
				}
			}

			// create .xlsx file
			//ExcelGenerator.CreateFile(products, "Products", fileName);
		}
	}
}

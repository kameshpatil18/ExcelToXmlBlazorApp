using System;
using System.IO;
using System.Linq;
using System.Collections.Generic;
using System.Xml.Linq;
using System.Xml.Schema;
using OfficeOpenXml;
using Microsoft.AspNetCore.Components;
using Microsoft.AspNetCore.Components.Forms;

namespace ExcelToXmlBlazorApp.Pages
{
    public partial class ExcelToXmlComponent : ComponentBase
    {
        private string statusMessage = "Select an Excel file to convert to XML.";
        private IBrowserFile selectedFile;
        private string outputXmlDirectory = "wwwroot/xml_output/";
        private List<ShipOrder> shipOrders = new List<ShipOrder>();
        private List<string> generatedXmlFiles = new List<string>();

        private void OnFileSelected(InputFileChangeEventArgs e)
        {
            selectedFile = e.File;
            statusMessage = $"Selected file: {selectedFile.Name}";
        }

        private async Task GenerateXml()
        {
            if (selectedFile == null)
            {
                statusMessage = "No file selected.";
                return;
            }

            if (!Directory.Exists(outputXmlDirectory))
            {
                Directory.CreateDirectory(outputXmlDirectory);
            }

            string filePath = Path.Combine(outputXmlDirectory, selectedFile.Name);
            await using FileStream fs = new FileStream(filePath, FileMode.Create);
            await selectedFile.OpenReadStream().CopyToAsync(fs);

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;
                int fileCount = 0;
                shipOrders.Clear();
                generatedXmlFiles.Clear();

                for (int row = 2; row <= rowCount; row++)
                {
                    var cell = worksheet.Cells[row, 3]; // Column index of Address
                    var fillColor = cell.Style.Fill.BackgroundColor.Rgb;

                    if (string.IsNullOrEmpty(fillColor) || fillColor == "FFFFFF" || fillColor == "000000")
                    {
                        continue; // Skip non-shaded rows
                    }

                    string orderId = worksheet.Cells[row, 1].Text;
                    DateTime orderDate = worksheet.Cells[row, 2].GetValue<DateTime>();
                    string address = worksheet.Cells[row, 7].Text;
                    string city = worksheet.Cells[row, 3].Text;
                    string region = ExtractRegion(address);
                    string productCategory = worksheet.Cells[row, 4].Text;
                    string productName = worksheet.Cells[row, 5].Text;
                    int quantity = int.Parse(worksheet.Cells[row, 6].Text);
                    decimal price = decimal.Parse(worksheet.Cells[row, 8].Text);
                    decimal total = quantity * price;

                    shipOrders.Add(new ShipOrder
                    {
                        OrderId = orderId,
                        OrderDate = orderDate,
                        Address = address,
                        City = city,
                        Region = region,
                        ProductCategory = productCategory,
                        ProductName = productName,
                        Quantity = quantity,
                        Price = price,
                        Total = total
                    });

                    fileCount++;
                    string xmlFilePath = Path.Combine(outputXmlDirectory, $"record_{fileCount}.xml");
                    generatedXmlFiles.Add($"xml_output/record_{fileCount}.xml");
                    XDocument xmlDoc = new XDocument(
                        new XElement("shiporder",
                            new XAttribute("orderid", orderId),
                            new XAttribute("orderdate", orderDate.ToString("yyyy-MM-dd")),
                            new XElement("shipto",
                                new XElement("orderid", orderId),
                                new XElement("address", address),
                                new XElement("city", city),
                                new XElement("region", region)
                            ),
                            new XElement("item",
                                new XElement("productCategory", productCategory),
                                new XElement("productName", productName),
                                new XElement("quantity", quantity),
                                new XElement("price", price),
                                new XElement("total", total)
                            )
                        )
                    );
                    xmlDoc.Save(xmlFilePath);
                }
                statusMessage = shipOrders.Count > 0 ? $"{shipOrders.Count} shaded rows loaded and XMLs generated." : "No shaded rows found.";
            }
        }

        private string ExtractRegion(string address)
        {
            string[] parts = address.Split(',');
            return parts.Length >= 2 ? parts[parts.Length - 2].Trim() : "Unknown";
        }
    }

    public class ShipOrder
    {
        public string OrderId { get; set; }
        public DateTime OrderDate { get; set; }
        public string Address { get; set; }
        public string City { get; set; }
        public string Region { get; set; }
        public string ProductCategory { get; set; }
        public string ProductName { get; set; }
        public int Quantity { get; set; }
        public decimal Price { get; set; }
        public decimal Total { get; set; }
    }
}
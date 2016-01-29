using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Validation;

namespace TestXLS {
    class Program {
        static void Main(string[] args) {

            try {
                OpenXmlValidator validator = new OpenXmlValidator();
                int count = 0;
                string file = "";
                foreach(ValidationErrorInfo error in validator.Validate(SpreadsheetDocument.Open(@"C:\Users\joao.delgado\Desktop\Reports\" + file + ".xlsx", true))) {
                    count++;
                    Console.WriteLine("Error " + count);
                    Console.WriteLine("Description: " + error.Description);
                    Console.WriteLine("Path: " + error.Path.XPath);
                    Console.WriteLine("Part: " + error.Part.Uri);
                    Console.WriteLine("-------------------------------------------");
                }
                Console.ReadKey();
            } catch(Exception ex) {
                Console.WriteLine(ex.Message);
                Console.ReadKey();
            }
        }
    }
}
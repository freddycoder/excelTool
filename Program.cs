// This is a console app that list the sheet of a workbook and the name of the workbook
// It is using the package OpenXML

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

// Define the namespace and main class
class Program
{
    // Define the main function
    static void Main(string[] args)
    {
        // Define the path of the file
        string path = @"C:\Users\jacqu\Downloads\Default14_07_2023_11_27_06.xlsx";

        Console.WriteLine("Enter the name of the sheet you want to select: ");
        string? sheetName = Console.ReadLine();

        // Open the file
        using (SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(path, false))
        {
            var feuille = SeletionneFeuille(spreadsheetDocument, sheetName);

            // From the sheet loop line by line and print the value of the cell
            foreach (Row row in feuille.Descendants<Row>())
            {
                foreach (Cell cell in row.Descendants<Cell>())
                {
                    Console.WriteLine(cell.CellReference + " " + cell.CellValue.Text);
                }
            }
        }
    }

    private static Sheet SeletionneFeuille(SpreadsheetDocument spreadsheetDocument, string sheetName)
    {
        // Get the workbook
        WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;

        // Get the sheets
        Sheets sheets = workbookPart.Workbook.Sheets;

        // Loop through the sheets
        foreach (Sheet sheet in sheets)
        {
            // Print the name of the sheet
            Console.WriteLine("Sheet name: " + sheetName);

            if (sheetName == sheet.Name) {
                return sheet;
            }
        }

        throw new KeyNotFoundException($"Sheet {sheetName} not found");
    }
}
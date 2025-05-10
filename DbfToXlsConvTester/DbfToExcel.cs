using System.Data;
using System.IO;
using Microsoft.VisualBasic.FileIO; // Para ler DBF
using OfficeOpenXml; // Para criar Excel XLSX

namespace DbfToXlsConvTester
{
    internal class DbfToExcel
    {
        // Para converter de DBF para XLSX (Excel 2007+) utilizando OfficeOpenXml
        public void ConvertDbfToExcel(string dbfFilePath, string excelOutputPath)
        {
            // 1. Ler o arquivo DBF para um DataTable
            DataTable dbfData = ReadDbfFile(dbfFilePath);

            // 2. Criar um arquivo Excel (XLSX) com EPPlus
            using (ExcelPackage excelPackage = new ExcelPackage())
            {
                ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add("Sheet1");

                // 2.1. Escrever cabeçalhos
                for (int col = 0; col < dbfData.Columns.Count; col++)
                {
                    worksheet.Cells[1, col + 1].Value = dbfData.Columns[col].ColumnName;
                }

                // 2.2. Escrever dados
                for (int row = 0; row < dbfData.Rows.Count; row++)
                {
                    for (int col = 0; col < dbfData.Columns.Count; col++)
                    {
                        worksheet.Cells[row + 2, col + 1].Value = dbfData.Rows[row][col];
                    }
                }

                // 3. Salvar o arquivo Excel
                FileInfo excelFile = new FileInfo(excelOutputPath);
                excelPackage.SaveAs(excelFile);
            }

        }

        // Método para ler DBF usando Microsoft.VisualBasic.FileIO
        private DataTable ReadDbfFile(string dbfFilePath)
        {
            DataTable table = new DataTable();

            using (TextFieldParser parser = new TextFieldParser(dbfFilePath))
            {
                parser.TextFieldType = FieldType.FixedWidth; // Ou Delimited, dependendo do formato
                parser.SetDelimiters(","); // Se for delimitado por vírgula

                string[] headers = parser.ReadFields();

                foreach (string header in headers)
                {
                    table.Columns.Add(header);
                }

                while (!parser.EndOfData)
                {
                    string[] fields = parser.ReadFields();
                    table.Rows.Add(fields);
                }
            }

            return table;
        }
    }
}

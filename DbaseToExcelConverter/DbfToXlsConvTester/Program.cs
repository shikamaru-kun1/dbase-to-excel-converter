using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbfToXlsConvTester
{
    internal class Program
    {
        private static Resources resources = new Resources();

        static void Main(string[] args)
        {
            try
            {
                // 1. Obter diretório do projeto
                string projectPath = resources.GetProjectDirectory();
                Console.WriteLine($"caminho: {projectPath}");

                // 2. Criar estrutura de pastas de importação
                string targetPath = Path.Combine(projectPath, "DataRepo", "In");
                resources.CreateDirectoryIfNotExists(targetPath);

                // 3. Copiar arquivos
                string sourcePath = "C:/Users/Administrator/Projectos EMJIT/Integration/DbfToExcelConverter/DbfToExcelConverter/Tests/TestFiles/Repo/DbfSamples";
                resources.CopySpecificFiles(sourcePath, targetPath);

                // 4. Criar a pasta de destino (para os fcheiros convertidos)
                string outputPath = Path.Combine(projectPath, "DataRepo", "Out");
                resources.CreateDirectoryIfNotExists(outputPath);

                // 5. Converter DBase para Excel
                //new DbfToExcel().ConvertDbfToExcel

                //Console.WriteLine("Operação concluída com sucesso!");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Erro: {ex.Message}");
            }
            //DbfToExcel dbfToExcel = new DbfToExcel();
            //string dbfFile = @"C:\dados\arquivo.dbf";
            //string excelFile = @"C:\export\dados.xls";

            //dbfToExcel.ConvertDbfToExcel(dbfFile, excelFile);
        }
    }
}

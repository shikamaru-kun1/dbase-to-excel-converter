using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DbfToXlsConvTester
{
    public class Resources
    {
        public Resources() { }

        public string GetProjectDirectory()
        {
            // Para projectos .NET Core (bin/Debug/netX.Y)
            string currentDir = Directory.GetCurrentDirectory();
            Console.WriteLine($"path: {Directory.GetParent(currentDir).Parent.Parent.FullName}");
            return Directory.GetParent(currentDir).Parent.Parent.FullName;
            //return currentDir;
        }

        public void CreateDirectoryIfNotExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
                Console.WriteLine($"Diretório criado: {path}");
            }
        }

        public void CopyFiles(string sourceDir, string targetDir)
        {
            var files = Directory.EnumerateFiles(sourceDir);

            foreach (var file in files)
            {
                string destFile = Path.Combine(targetDir, Path.GetFileName(file));
                File.Copy(file, destFile, overwrite: true);
                Console.WriteLine($"Copiado: {Path.GetFileName(file)}");
            }
        }
    }
}

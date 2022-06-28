using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ArcelorFileHandler.Services
{
    public static class FileManager
    {
        public static void DirSearch(string path)
        {
            try
            {
                foreach (string dir in Directory.GetDirectories(path))
                {
                    foreach (string file in Directory.GetFiles(dir))
                    {
                        Console.WriteLine(file);
                    }
                    DirSearch(dir);
                }
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);
            }
        }
    }
}

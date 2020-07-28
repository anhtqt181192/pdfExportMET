using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace T2P.Export_HTML_To_PDF.Extentions
{
    public static class ByteHelpers
    {
        public static void ToFile(this byte[] bytes, string fileName)
        {
            try
            {
                using (var fs = new FileStream(fileName, FileMode.Create, FileAccess.Write))
                {
                    fs.Write(bytes, 0, bytes.Length);
                    fs.Close();
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception caught in process: {0}", ex);
            }
        }

        public static byte[] GetByteFromFile(string fileUrl)
        {
            try
            {
                return File.ReadAllBytes(fileUrl);
            }
            catch (Exception ex)
            {
                Console.WriteLine("File not found");
            }
            return null;
        }
    }
}

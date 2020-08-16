using System;
using System.IO;
using System.Security;


namespace DTE_ComparaArco
{
    public class DTE_log
    {
        public DTE_log(string Ruta)
        {
            this.Ruta = Ruta;
        }

        public void Add(string level, string sLog)
        {


            CreateDirectory();
            string nombre = GetNameFile();

            string cadena = String.Format("[{0}] {1} - {2}{3}", DateTime.Now, level, sLog, Environment.NewLine);

            if (Transporte == "All" || Transporte == "F")
            {
                StreamWriter sw = new StreamWriter(Path.Combine(Ruta, nombre), true);
                sw.Write(cadena);
                sw.Close();
            }
            if (Transporte == "All" || Transporte == "C")
            {
                Console.WriteLine(cadena);
            }

        }

        #region HELPER
        private string GetNameFile()
        {
            string nombre = "DTEComparalog_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + ".log";
            return nombre;
        }

        private void CreateDirectory()
        {
            try
            {
                if (!Directory.Exists(Ruta))
                    Directory.CreateDirectory(Ruta);
            }
            catch (Exception ex)
            {
                if
                  (
                      ex is UnauthorizedAccessException
                      || ex is ArgumentNullException
                      || ex is PathTooLongException
                      || ex is DirectoryNotFoundException
                      || ex is NotSupportedException
                      || ex is ArgumentException
                      || ex is SecurityException
                      || ex is IOException
                  )
                {
                    throw new Exception(ex.Message);
                }
            }
        }
        #endregion

        private string Ruta = "";
        private string Transporte = "All";  // C: Consola  F: Archivo All: Ambos
    }
}

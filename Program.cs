using NPOI.XSSF.UserModel;  // Para trabajar con archivos .xlsx
using NPOI.SS.UserModel;
using System.Globalization;    // Interfaz común para manipular hojas de cálculo

namespace BuildF394
{
    /// <summary>
    ///
    /// </summary>
    internal class Program
    {
        /// <summary>
        ///
        /// </summary>
        private static void Main()
        {
            // Obtener la ruta del directorio del ejecutable
            string rutaEjecutable = AppDomain.CurrentDomain.BaseDirectory;

            // Construir la ruta completa al archivo plantilla.xlsx
            string rutaArchivo = Path.Combine(rutaEjecutable, "plantilla.xls");

            // Verificar si el archivo existe
            if (!File.Exists(rutaArchivo))
            {
                Console.WriteLine("El archivo no se encuentra en la ruta especificada.");
                return;
            }

            // Abrir el archivo Excel (Formato .xlsx)
            using (FileStream file = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
            {
                XSSFWorkbook workbook = new XSSFWorkbook(file);  // Para archivos .xlsx

                // Obtener las hojas de cálculo
                ISheet hoja = workbook.GetSheetAt(0);

                Console.WriteLine("ObtenerRegistros");
                int j = ObtenerRegistros(hoja);

                Console.WriteLine("ProcesarFilas");
                List<string[]> datos = ProcesarFilas(hoja, j - 1);

                int secuencia = j - 1;

                string nceros_sec = new string('0', 8 - (secuencia + 1).ToString().Length);
                int Total = j - secuencia;

                string date = hoja.GetRow(0).GetCell(5).ToString();
                string day = date.Substring(0, 2);
                string month = date.Substring(3, 3).ToUpper();
                string year = date.Substring(date.Length - 4);

                Console.WriteLine("Escribiendo encabezados");
                string tipo1 = $"0000001114000025{day}{month}{year}{datos.Count().ToString().PadLeft(8, '0')}SVIDCOLMENA0907";
                string tipo3 = "00000023000000000000000000000000";
                string tipo4 = "00000034000000000000000000000002";
                string tipo6 = $"{datos.Count().PadZerosLeft()}6";

                datos.Insert(0, new string[] { tipo1 });
                datos.Insert(1, new string[] { tipo3 });
                datos.Insert(2, new string[] { tipo4 });

                // Registro de cierre al final
                datos.Add(new string[] { tipo6 });

                Console.WriteLine("EscribirEnArchivoPlano");
                EscribirEnArchivoPlano(datos, month, year, rutaEjecutable);

                Console.WriteLine("Proceso completado.");
            }
        }

        /// <summary>
        /// Función para escribir registros directamente en archivo de texto
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="rutaEjecutable"></param>
        /// <param name="hoja"></param>
        private static void EscribirEnArchivoPlano(List<string[]> datos, string mes, string año, string rutaEjecutable)
        {
            // Unir mes y año en el formato deseado
            string fecha = $"{mes}-{año}";

            // Construir nombre del archivo de salida
            string nombre = $"{fecha}-fto394.txt";

            string rutaGuardado = Path.Combine(rutaEjecutable, nombre);

            // Borra el archivo si existe
            if (File.Exists(rutaGuardado)) File.Delete(rutaGuardado);

            Console.WriteLine($"Archivo creado {rutaGuardado}");

            using (StreamWriter writer = new StreamWriter(rutaGuardado))
            {
                foreach (var fila in datos)
                {
                    // Escribir cada fila concatenando todos los valores
                    writer.WriteLine(string.Join("", fila));
                }
            }
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="secuencia"></param>
        /// <param name="j"></param>
        /// <param name="hojaRow"></param>
        /// <param name="xlsHoja"></param>
        /// <returns></returns>
        private static string[] CrearFilaEnMemoria(int secuencia, int j, int hojaRow, ISheet xlsHoja)
        {
            string nceros1 = new string('0', 8 - secuencia.ToString().Length);
            string nceros4 = j < 10 ? "0" : "";
            string nceros6 = new string('0', 6 - hojaRow.ToString().Length);

            // Almacenar todos los datos en un array de strings
            string[] filaDatos = new string[8];
            filaDatos[0] = nceros1 + secuencia;
            filaDatos[1] = "5";
            filaDatos[2] = "394";
            filaDatos[3] = nceros4 + j;
            filaDatos[4] = "01";
            filaDatos[5] = nceros6 + hojaRow;
            filaDatos[6] = "+";
            filaDatos[7] = ObtenerValorCelda(xlsHoja, hojaRow, j);

            return filaDatos;
        }

        private static string ObtenerValorCelda(ISheet xlHoja1, int hoja1Row, int j)
        {
            if (Extensions.EsTexto(j))
            {
                return xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).ToString().PadRight(50);
            }
            else if (Extensions.EsFecha(j))
            {
                DateTime fecha = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).DateCellValue.Value;
                string diaStr = fecha.Day < 10 ? "0" + fecha.Day : fecha.Day.ToString();
                string mesStr = fecha.Month < 10 ? "0" + fecha.Month : fecha.Month.ToString();
                string añoStr = fecha.Year.ToString();
                return $"{diaStr}{mesStr}{añoStr}";
            }
            else if (Extensions.EsNumero(j))
            {
                double valor = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).NumericCellValue;
                return Math.Round(valor, 2).ToString("0.00").Replace(",", ".");
            }
            return "";
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="hoja"></param>
        /// <param name="registros"></param>
        /// <returns></returns>
        private static List<string[]> ProcesarFilas(ISheet hoja, int registros)
        {
            List<string[]> dataEnMemoria = new List<string[]>();
            int k = 0;
            int secuencia = 0;

            for (int col = 1; col <= 84; col++)
            {
                for (int fila = 1; fila <= registros; fila++)
                {
                    if (hoja.CeldaNoVacia(fila, col))
                    {
                        k++;
                        secuencia = 3 + k;
                        string[] filaDatos = CrearFilaEnMemoria(secuencia, col, fila, hoja);  // Almacenar en memoria
                        dataEnMemoria.Add(filaDatos);
                    }
                }
            }

            return dataEnMemoria;
        }

        private static int ObtenerRegistros(ISheet sheet)
        {
            int j = 1;
            while (sheet.GetRow(5 + j) != null && sheet.GetRow(5 + j).GetCell(0) != null &&
                   sheet.GetRow(5 + j).GetCell(0).CellType != CellType.Blank)
            {
                j++;
            }
            return j - 1;
        }
    }
}
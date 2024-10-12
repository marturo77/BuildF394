using NPOI.XSSF.UserModel;  // Para trabajar con archivos .xlsx
using NPOI.SS.UserModel;    // Interfaz común para manipular hojas de cálculo

namespace BuildF394
{
    internal class Program
    {
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
                ISheet xlHoja1 = workbook.GetSheetAt(0);
                ISheet xlHoja2 = workbook.GetSheetAt(1);
                ISheet xlHoja3 = workbook.GetSheetAt(2);

                string ruta = xlHoja1.GetRow(1).GetCell(5).ToString();

                Console.WriteLine("ObtenerRegistros");
                int j = ObtenerRegistros(xlHoja1);

                Console.WriteLine("ObtenerRegistros");
                List<string[]> datos = ProcesarFilas(xlHoja1, j - 1);

                int secuencia = j - 1;

                string nceros_sec = new string('0', 8 - (secuencia + 1).ToString().Length);
                int Total = j - secuencia;
                string total_reg = nceros_sec + Total;

                string fecha = xlHoja1.GetRow(0).GetCell(5).ToString();

                string reg_tipo_1 = "0000001114000025" + fecha.Substring(0, 2) + fecha.Substring(3, 2) + total_reg + "SVIDCOLMENA0907";
                string reg_tipo_3 = "00000023000000000000000000000000";
                string reg_tipo_4 = "00000034000000000000000000000002";
                string reg_tipo_6 = total_reg + "6";

                IRow row3 = xlHoja3.CreateRow(0);
                row3.CreateCell(0).SetCellValue(reg_tipo_1);
                xlHoja3.CreateRow(1).CreateCell(0).SetCellValue(reg_tipo_3);
                xlHoja3.CreateRow(2).CreateCell(0).SetCellValue(reg_tipo_4);

                for (int m = 4; m <= secuencia; m++)
                {
                    IRow row3M = xlHoja3.CreateRow(m);
                    string value =
                        xlHoja2.GetRow(m - 2).GetCell(0).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(1).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(2).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(3).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(4).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(5).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(6).ToString() +
                        xlHoja2.GetRow(m - 2).GetCell(7).ToString();

                    row3M.CreateCell(0).SetCellValue(value);
                }

                xlHoja3.CreateRow(secuencia + 1).CreateCell(0).SetCellValue(reg_tipo_6);

                Console.WriteLine("EscribirEnArchivoPlano");
                EscribirEnArchivoPlano(datos, rutaEjecutable, xlHoja1);

                Console.WriteLine("Proceso completado.");
            }
        }

        /// <summary>
        /// Función para escribir registros directamente en archivo de texto
        /// </summary>
        /// <param name="dataEnMemoria"></param>
        /// <param name="rutaEjecutable"></param>
        /// <param name="hoja"></param>
        private static void EscribirEnArchivoPlano(List<string[]> dataEnMemoria, string rutaEjecutable, ISheet hoja)
        {
            string fechaOriginal = hoja.GetRow(0).GetCell(5).ToString();

            // Extraer el mes y el año manualmente
            string mes = fechaOriginal.Substring(3, 3).ToUpper();  // Obtiene "DIC"
            string año = fechaOriginal.Substring(fechaOriginal.Length - 4);  // Obtiene "2014"

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
                foreach (var fila in dataEnMemoria)
                {
                    // Escribir cada fila concatenando todos los valores
                    writer.WriteLine(string.Join("", fila));
                }
            }
        }

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

        private static List<string[]> ProcesarFilas(ISheet xlHoja1, int registros)
        {
            List<string[]> dataEnMemoria = new List<string[]>();  // Estructura para almacenar los datos en memoria
            int k = 0;
            int secuencia = 0;

            for (int j = 1; j <= 84; j++)
            {
                for (int hoja1Row = 1; hoja1Row <= registros; hoja1Row++)
                {
                    if (xlHoja1.CeldaNoVacia(hoja1Row, j))
                    {
                        k++;
                        secuencia = 3 + k;
                        string[] filaDatos = CrearFilaEnMemoria(secuencia, j, hoja1Row, xlHoja1);  // Almacenar en memoria
                        dataEnMemoria.Add(filaDatos);  // Añadir a la lista
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
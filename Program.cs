using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace BuildF394
{
    /// <summary>
    ///
    /// </summary>
    internal class Program
    {
        /// <summary>
        /// Offset en las filas para la pagina 1
        /// </summary>
        private const int ROW_OFFSET = 5;

        /// <summary>
        ///
        /// </summary>
        private class HeaderData
        {
            /// <summary>
            ///
            /// </summary>
            public string EntityType = "14";

            /// <summary>
            ///
            /// </summary>
            public string EntityCode = "000025";

            /// <summary>
            ///
            /// </summary>
            public string KeyWord = "SVIDCOLMENA";

            /// <summary>
            ///
            /// </summary>
            public string Area = "09";

            /// <summary>
            ///
            /// </summary>
            public string ReportType = "07";
        }

        /// <summary>
        ///
        /// </summary>
        private static void Main()
        {
            try
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
                    XSSFWorkbook workbook = new XSSFWorkbook(file);

                    // Obtener las hojas de cálculo
                    ISheet hoja = workbook.GetSheetAt(0);

                    Console.WriteLine("ObtenerRegistros");

                    Console.WriteLine("ProcesarFilas");
                    List<string[]> datos = ProcesarFilas(hoja, ContarRegistros(hoja));

                    // Fecha del reporte
                    string date = hoja.GetRow(0).GetCell(5).ToString();
                    string day = date.Substring(0, 2);
                    string month = "12";//date.Substring(3, 3).ToUpper();
                    string year = date.Substring(date.Length - 4);

                    Console.WriteLine("Escribiendo encabezados");

                    // Entidad con los datos del encabezado
                    HeaderData d = new HeaderData();

                    //REGISTRO TIPO 1
                    //la secuencia es 1;debe tener 48 caracteres
                    string tipo1 = $"{1.PadZerosLeft()}{1}{d.EntityType}{d.EntityCode}{day}{month}{year}{datos.Count().PadZerosLeft()}{d.KeyWord}{d.Area}{d.ReportType}";
                    if (tipo1.Length != 48) throw new Exception("Registro tipo 1 invalido");

                    // REGISTRO TIPO 3
                    // la secuencia es 2; debe tener 43 caracteres
                    string evaluationType = "0";
                    int fideicomiso = 0;
                    string tipo3 = $"{2.PadZerosLeft()}{3}{evaluationType}{fideicomiso.PadZerosLeft(17)}0000000000000000";
                    if (tipo3.Length != 43) throw new Exception("Registro tipo 3 invalido");

                    // REGISTRO TIPO 4
                    //la secuencia es 3; debe tener 31 caracteres
                    string tipo4 = $"{3.PadZerosLeft()}{4}0000000000000000000002";
                    if (tipo4.Length != 31) throw new Exception("Registro tipo 4 invalido");

                    // REGISTRO TIPO 6
                    //la secuencia va de ultimo; debe tener 31 caracteres
                    int lastSequence = datos.Count() + 3;
                    string tipo6 = $"{lastSequence.PadZerosLeft()}{6}";
                    if (tipo6.Length != 9) throw new Exception("Registro tipo 6 invalido");

                    // tipo1, tipo3, tipo 4 van al principio
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
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        /// <summary>
        /// Función para escribir registros directamente en archivo de texto
        /// </summary>
        /// <param name="datos"></param>
        /// <param name="mes"></param>
        /// <param name="año"></param>
        /// <param name="rutaEjecutable"></param>
        private static void EscribirEnArchivoPlano(List<string[]> datos, string mes, string año, string rutaEjecutable)
        {
            // Construir nombre del archivo de salida
            string nombre = $"{mes}-{año}-fto394.txt";

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
        /// Crea registros tipo 5
        /// </summary>
        /// <param name="secuencia"></param>
        /// <param name="fila"></param>
        /// <param name="columna"></param>
        /// <param name="hoja"></param>
        /// <returns></returns>
        private static string[] CrearFila(int secuencia, int fila, int columna, ISheet hoja)
        {
            string nceros6 = new string('0', 6 - fila.ToString().Length);

            // Almacenar todos los datos en un array de strings
            string[] result = new string[8];

            //Secuencia 8
            result[0] = secuencia.PadZerosLeft();

            //Tipo de registro 5
            result[1] = "5";

            //Codigo de formato
            result[2] = "394";

            //Codigo de Columna
            result[3] = columna.PadZerosLeft(2);

            // Codigo de unidad de captura
            result[4] = "01";

            // Codigo de la subcuenta
            result[5] = nceros6 + fila;

            //Signo
            result[6] = "+";

            //Valor longitud 17 o 50
            result[7] = ObtenerValorCelda(hoja, fila, columna);

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="hoja"></param>
        /// <param name="fila"></param>
        /// <param name="columna"></param>
        /// <returns></returns>
        private static string ObtenerValorCelda(ISheet hoja, int fila, int columna)
        {
            if (columna.EsTexto())
            {
                return hoja.GetRow(ROW_OFFSET + fila).GetCell(columna - 1).ToString().PadRight(50);
            }
            else if (columna.EsFecha())
            {
                DateTime fecha = hoja.GetRow(ROW_OFFSET + fila).GetCell(columna - 1).DateCellValue.Value;
                string day = fecha.Day.PadZerosLeft(2);
                string month = fecha.Month.PadZerosLeft(2);
                string year = fecha.Year.ToString();
                return $"{day}{month}{year}";
            }
            else if (columna.EsNumero())
            {
                double valor = hoja.GetRow(ROW_OFFSET + fila).GetCell(columna - 1).NumericCellValue;
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
            List<string[]> result = new List<string[]>();
            int counter = 0;
            int secuencia = 0;

            for (int col = 1; col <= 84; col++)
            {
                for (int fila = 1; fila <= registros; fila++)
                {
                    if (hoja.CeldaNoVacia(fila, col))
                    {
                        counter++;
                        secuencia = 3 + counter;
                        string[] filaDatos = CrearFila(secuencia, fila, col, hoja);
                        result.Add(filaDatos);
                    }
                }
            }

            return result;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sheet"></param>
        /// <returns></returns>
        private static int ContarRegistros(ISheet sheet)
        {
            int j = 1;
            while (sheet.GetRow(5 + j) != null && sheet.GetRow(ROW_OFFSET + j).GetCell(0) != null &&
                   sheet.GetRow(5 + j).GetCell(0).CellType != CellType.Blank)
            {
                j++;
            }
            return j - 1;
        }
    }
}
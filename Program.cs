using NPOI.SS.Formula.Functions;
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
        public const int ROW_OFFSET = 6;

        /// <summary>
        /// Codigo de formato de document
        /// </summary>
        public const string FORMAT_CODE = "394";

        /// <summary>
        ///
        /// </summary>
        public const string CAPTURE_CODE = "01";

        /// <summary>
        /// Total de columnas en la pagina 1
        /// </summary>
        public const int TOTAL_COLUMNS = 84;

        /// <summary>
        /// Numero de filas que tendra el archivo plano en el encabezado
        /// </summary>
        public const int HEADER_ROWS = 3;

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

                    Console.WriteLine("ProcesarFilas");
                    List<string[]> datos = ProcesarFilas(hoja, ContarRegistros(hoja));

                    // Resolver Fecha del reporte
                    // TODO: Revisar el tema del mes que presento un problema en el parsing
                    string date = hoja.GetRow(0).GetCell(5).ToString();
                    string day = date.Substring(0, 2);
                    string month = "12";//date.Substring(3, 3).ToUpper();
                    string year = date.Substring(date.Length - 4);

                    Console.WriteLine("Escribiendo encabezados");

                    // Entidad con los datos del encabezado
                    HeaderData headerData = new HeaderData();

                    #region Registro Tipo 1 - Registro de control y encabezado

                    //la secuencia es 1;debe tener 48 caracteres
                    string tipo1 = $"{1.PadZerosLeft()}{1}{headerData.EntityType}{headerData.EntityCode}{day}{month}{year}{datos.Count().PadZerosLeft()}{headerData.KeyWord}{headerData.Area}{headerData.ReportType}";
                    if (tipo1.Length != 48) throw new Exception("Registro tipo 1 invalido");

                    #endregion Registro Tipo 1 - Registro de control y encabezado

                    #region RegistroTipo 2 - Tipo de identificacion

                    // la secuencia es 2; debe tener 43 caracteres
                    string evaluationType = "0";
                    int fideicomiso = 0;
                    string tipo3 = $"{2.PadZerosLeft()}{3}{evaluationType}{fideicomiso.PadZerosLeft(17)}0000000000000000";
                    if (tipo3.Length != 43) throw new Exception("Registro tipo 3 invalido");

                    #endregion RegistroTipo 2 - Tipo de identificacion

                    #region Registro Tipo 4 - Formatos

                    //la secuencia es 3; debe tener 31 caracteres
                    string tipo4 = $"{3.PadZerosLeft()}{4}0000000000000000000002";
                    if (tipo4.Length != 31) throw new Exception("Registro tipo 4 invalido");

                    #endregion Registro Tipo 4 - Formatos

                    #region Registro tipo 6 - Cierre o fin de archivo

                    //la secuencia va de ultimo; debe tener 31 caracteres
                    int lastSequence = datos.Count() + 3;
                    string tipo6 = $"{lastSequence.PadZerosLeft()}{6}";
                    if (tipo6.Length != 9) throw new Exception("Registro tipo 6 invalido");

                    #endregion Registro tipo 6 - Cierre o fin de archivo

                    #region Colocar registros al principio y el tipo6 al final

                    // tipo1, tipo3, tipo 4 van al principio
                    datos.Insert(0, new string[] { tipo1 });
                    datos.Insert(1, new string[] { tipo3 });
                    datos.Insert(2, new string[] { tipo4 });

                    // Registro de cierre al final
                    datos.Add(new string[] { tipo6 });

                    #endregion Colocar registros al principio y el tipo6 al final

                    #region Escribir en disco

                    Console.WriteLine("Escribir Archivo Plano");
                    WriteTextPlain(datos, month, year, rutaEjecutable);

                    #endregion Escribir en disco

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
        private static void WriteTextPlain(List<string[]> datos, string mes, string año, string rutaEjecutable)
        {
            // Construir nombre del archivo de salida
            string nombre = $"{mes}-{año}-fto394.txt";

            string rutaGuardado = Path.Combine(rutaEjecutable, nombre);

            // Borra el archivo si existe
            if (File.Exists(rutaGuardado)) File.Delete(rutaGuardado);

            using (StreamWriter writer = new StreamWriter(rutaGuardado))
            {
                foreach (var fila in datos)
                {
                    // Escribir cada fila concatenando todos los valores
                    writer.WriteLine(string.Join("", fila));
                }
            }

            Console.WriteLine($"Archivo creado {rutaGuardado}");
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="secuence"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string[] GetConsolidationRecord(int secuence, int col, double total)
        {
            // Almacenar todos los datos en un array de strings
            string[] record = new string[8];

            //Secuencia 8 caracteres
            record[0] = secuence.PadZerosLeft();

            //Tipo de registro 5
            record[1] = "5";

            //Codigo de formato
            record[2] = Program.FORMAT_CODE;

            //Codigo de Columna/es la columna
            record[3] = (col + 1).PadZerosLeft(2);

            //Codigo de unidad de captura
            record[4] = "02";

            //Codigo de la subcuenta
            record[5] = "000001";

            //Signo
            record[6] = "+";

            //Valor longitud 17
            record[7] = ((int)total).PadZerosLeft(17);

            return record;
        }

        /// <summary>
        /// Crea registros tipo 5
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="secuence"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        private static string[] GetRecord(ISheet sheet, int secuence, int row, int col)
        {
            // Almacenar todos los datos en un array de strings
            string[] record = new string[8];

            //Secuencia 8 caracteres
            record[0] = secuence.PadZerosLeft();

            //Tipo de registro 5
            record[1] = "5";

            //Codigo de formato
            record[2] = Program.FORMAT_CODE;

            //Codigo de Columna/Cuenta
            record[3] = (col + 1).PadZerosLeft(2);

            //Codigo de unidad de captura
            record[4] = Program.CAPTURE_CODE;

            //Codigo de la subcuenta la fila
            record[5] = (row + 1).PadZerosLeft(6);

            //Signo
            record[6] = "+";

            record[7] = GetValue(sheet, row, col);

            return record;
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="columna"></param>
        /// <returns></returns>
        private static string GetValue(ISheet sheet, int row, int col)
        {
            ICell cell = sheet.GetRow(ROW_OFFSET + row).GetCell(col);

            if (col.IsRamo())
            {
                return cell.ToString().PadLeft(17, '0');
            }
            else if (col.IsInsuranceCode())
            {
                return cell.NumericCellValue.PadZerosLeft(17);
            }
            else if (col.IsText())
            {
                return cell.ToString().PadRight(50);
            }
            else if (col.IsDate())
            {
                DateTime? date = cell.DateCellValue;

                if (date.HasValue)
                {
                    string day = date.Value.Day.PadZerosLeft(2);
                    string month = date.Value.Month.PadZerosLeft(2);
                    string year = date.Value.Year.ToString();
                    return $"{day}{month}{year}";
                }
                else throw new Exception($"{row}{col} No tiene un campo de fecha");
            }
            else if (col.IsNumber())
            {
                return cell.NumericCellValue.PadZerosLeft(17);
            }

            // Si no devuelve cadena vacha
            return string.Empty;
        }

        /// <summary>
        /// Obtiene un array de registros tipo 5
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="totalRows"></param>
        /// <returns></returns>
        private static List<string[]> ProcesarFilas(ISheet sheet, int totalRows)
        {
            List<string[]> result = new List<string[]>();

            int secuencia = Program.HEADER_ROWS + 1;

            // Recorre las columnas
            for (int col = 0; col < Program.TOTAL_COLUMNS; col++)
            {
                double consolidate = 0;

                // Para todas las filas
                for (int row = 0; row < totalRows; row++)
                {
                    // Si hay datos
                    if (sheet.IsNotEmpty(row, col))
                    {
                        // Obtiene un registro tipo 5 para la cuenta
                        string[] record = GetRecord(sheet, secuencia, row, col);

                        // Pruebas
                        //if (!string.IsNullOrEmpty(record[7]))
                        {
                            result.Add(record);
                            secuencia++;
                        }

                        // Si es una columna de consolidacion va sumando
                        if (col.IsConsolidation())
                        {
                            ICell cell = sheet.GetRow(ROW_OFFSET + row).GetCell(col);
                            consolidate += cell.NumericCellValue;
                        }
                    }
                }

                // Es el final de la fila y hay que consolidar
                if (consolidate > 0)
                {
                    Console.WriteLine($"{secuencia.PadZerosLeft()} Columna consolidada {col}={consolidate.PadZerosLeft(17)}");

                    // Obtiene un registro tipo 5 para la cuenta
                    result.Add(GetConsolidationRecord(secuencia, col, consolidate));

                    secuencia++;
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
            int total = 0;

            IRow row = sheet.GetRow(ROW_OFFSET + total);

            while (row != null && row.GetCell(0) != null && row.GetCell(0).CellType != CellType.Blank)
            {
                total++;
                row = sheet.GetRow(ROW_OFFSET + total);
            }
            return total;
        }
    }
}
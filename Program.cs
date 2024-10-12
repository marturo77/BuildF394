using BuildF394;
using NPOI.SS.UserModel;    // Interfaz común para manipular hojas de cálculo
using NPOI.XSSF.UserModel;
using System.Globalization;

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
        string rutaArchivo = Path.Combine(rutaEjecutable, "plantilla.xls");

        // Verificar si el archivo existe
        if (!File.Exists(rutaArchivo))
        {
            Console.WriteLine("El archivo no se encuentra en la ruta especificada.");
            return;
        }

        // Abrir y procesar el archivo Excel
        using (FileStream stream = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
        {
            DateTime start = DateTime.Now;
            Console.WriteLine($"Abriendo {rutaArchivo}");
            XSSFWorkbook workbook = new XSSFWorkbook(stream);  // Para archivos .xlsx
            ISheet hoja = workbook.GetSheetAt(0);

            // Obtener ruta de celda
            string ruta = hoja.GetRow(1).GetCell(5).ToString();

            Console.WriteLine("ObtenerRegistros");
            int registros = ObtenerRegistros(hoja);

            Console.WriteLine("ProcesarFilas");
            // Procesar filas y almacenar en memoria
            List<string[]> dataEnMemoria = ProcesarFilas(hoja, registros);

            // Escribir registros directamente en archivo de texto
            Console.WriteLine("Escribir en archivo de texto");
            EscribirEnArchivoPlano(dataEnMemoria, rutaEjecutable, hoja);

            TimeSpan ts = DateTime.Now - start;
            Console.WriteLine($"Proceso completado. {ts.TotalSeconds} seg");
        }
    }

    /// <summary>
    /// Función para obtener el número de registros de la hoja 1
    /// </summary>
    /// <param name="sheet"></param>
    /// <returns></returns>
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

    /// <summary>
    /// Función para procesar las filas de la hoja 1 y almacenar en memoria
    /// </summary>
    /// <param name="xlHoja1"></param>
    /// <param name="registros"></param>
    /// <returns></returns>
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

    /// <summary>
    /// Función para crear una nueva fila en memoria
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

    /// <summary>
    /// Función para obtener el valor procesado de una celda
    /// </summary>
    /// <param name="xlHoja1"></param>
    /// <param name="hoja1Row"></param>
    /// <param name="j"></param>
    /// <returns></returns>
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
}
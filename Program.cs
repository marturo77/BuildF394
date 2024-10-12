using NPOI.SS.UserModel;    // Interfaz común para manipular hojas de cálculo
using NPOI.XSSF.UserModel;  // Para trabajar con archivos .xlsx

internal class Program
{
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
            Console.WriteLine($"Abriendo {rutaArchivo}");
            XSSFWorkbook workbook = new XSSFWorkbook(stream);  // Para archivos .xlsx
            ISheet xlHoja1 = workbook.GetSheetAt(0);

            // Obtener ruta de celda
            string ruta = xlHoja1.GetRow(1).GetCell(5).ToString();

            Console.WriteLine("ObtenerRegistros");
            int registros = ObtenerRegistros(xlHoja1);

            Console.WriteLine("ProcesarFilas");
            // Procesar filas y almacenar en memoria
            List<string[]> dataEnMemoria = ProcesarFilas(xlHoja1, registros);

            Console.WriteLine("Escribir en archivo de texto");
            // Escribir registros directamente en archivo de texto
            EscribirEnArchivoPlano(dataEnMemoria, rutaEjecutable, xlHoja1);

            Console.WriteLine("Proceso completado.");
        }
    }

    // Función para obtener el número de registros de la hoja 1
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

    // Función para procesar las filas de la hoja 1 y almacenar en memoria
    private static List<string[]> ProcesarFilas(ISheet xlHoja1, int registros)
    {
        List<string[]> dataEnMemoria = new List<string[]>();  // Estructura para almacenar los datos en memoria
        int k = 0;
        int secuencia = 0;

        for (int j = 1; j <= 84; j++)
        {
            for (int hoja1Row = 1; hoja1Row <= registros; hoja1Row++)
            {
                if (CeldaNoVacia(xlHoja1, hoja1Row, j))
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

    // Función para verificar si la celda no está vacía
    private static bool CeldaNoVacia(ISheet sheet, int fila, int columna)
    {
        return sheet.GetRow(5 + fila)?.GetCell(columna - 1) != null &&
               sheet.GetRow(5 + fila).GetCell(columna - 1).ToString() != "";
    }

    // Función para crear una nueva fila en memoria
    private static string[] CrearFilaEnMemoria(int secuencia, int j, int hoja1Row, ISheet xlHoja1)
    {
        string nceros1 = new string('0', 8 - secuencia.ToString().Length);
        string nceros4 = j < 10 ? "0" : "";
        string nceros6 = new string('0', 6 - hoja1Row.ToString().Length);

        // Almacenar todos los datos en un array de strings
        string[] filaDatos = new string[8];
        filaDatos[0] = nceros1 + secuencia;
        filaDatos[1] = "5";
        filaDatos[2] = "394";
        filaDatos[3] = nceros4 + j;
        filaDatos[4] = "01";
        filaDatos[5] = nceros6 + hoja1Row;
        filaDatos[6] = "+";
        filaDatos[7] = ObtenerValorCelda(xlHoja1, hoja1Row, j);

        return filaDatos;
    }

    // Función para obtener el valor procesado de una celda
    private static string ObtenerValorCelda(ISheet xlHoja1, int hoja1Row, int j)
    {
        if (EsTexto(j))
        {
            return xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).ToString().PadRight(50);
        }
        else if (EsFecha(j))
        {
            DateTime fecha = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).DateCellValue.Value;
            string diaStr = fecha.Day < 10 ? "0" + fecha.Day : fecha.Day.ToString();
            string mesStr = fecha.Month < 10 ? "0" + fecha.Month : fecha.Month.ToString();
            string añoStr = fecha.Year.ToString();
            return $"{diaStr}{mesStr}{añoStr}";
        }
        else if (EsNumero(j))
        {
            double valor = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).NumericCellValue;
            return Math.Round(valor, 2).ToString("0.00").Replace(",", ".");
        }
        return "";
    }

    // Funciones para identificar tipo de dato
    private static bool EsTexto(int columna) => new[] { 8, 10, 21, 47, 55, 63, 66, 76 }.Contains(columna);

    private static bool EsFecha(int columna) => new[] { 9, 11, 20, 46, 54, 62, 83 }.Contains(columna);

    private static bool EsNumero(int columna) => new[] { 4, 6, 7, 30, 32, 37, 79, 80, 81, 82 }.Contains(columna);

    // Función para escribir registros directamente en archivo de texto
    private static void EscribirEnArchivoPlano(List<string[]> dataEnMemoria, string rutaEjecutable, ISheet xlHoja1)
    {
        // Construir nombre del archivo de salida
        string nombre = xlHoja1.GetRow(0).GetCell(5).ToString().Substring(6, 4) +
                        xlHoja1.GetRow(0).GetCell(5).ToString().Substring(3, 2) +
                        "fto394.txt";
        string rutaGuardado = Path.Combine(rutaEjecutable, nombre);

        if (File.Exists(rutaGuardado)) File.Delete(rutaGuardado);

        Console.WriteLine($"Archivo creado {rutaGuardado}");

        using (StreamWriter writer = new StreamWriter(rutaGuardado))
        {
            foreach (var fila in dataEnMemoria)
            {
                writer.WriteLine(string.Join("", fila));  // Escribir cada fila concatenando todos los valores
            }
        }
    }
}
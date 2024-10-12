using NPOI.SS.UserModel;    // Interfaz común para manipular hojas de cálculo
using NPOI.XSSF.UserModel;  // Para trabajar con archivos .xlsx
using System;
using System.IO;

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
        using (FileStream file = new FileStream(rutaArchivo, FileMode.Open, FileAccess.Read))
        {
            XSSFWorkbook workbook = new XSSFWorkbook(file);  // Para archivos .xlsx
            ISheet xlHoja1 = workbook.GetSheetAt(0);
            ISheet xlHoja2 = workbook.GetSheetAt(1);
            ISheet xlHoja3 = workbook.GetSheetAt(2);

            // Limpiar hoja 2 antes de iniciar
            ClearSheet(xlHoja2);

            // Obtener ruta de celda
            string ruta = xlHoja1.GetRow(1).GetCell(5).ToString();
            int registros = ObtenerRegistros(xlHoja1);
            ProcesarFilas(xlHoja1, xlHoja2, registros);

            // Crear registros finales en la hoja 3
            CrearRegistrosFinales(xlHoja2, xlHoja3, registros);

            // Guardar el archivo procesado
            GuardarArchivo(workbook, rutaEjecutable, xlHoja1);
            Console.WriteLine("Proceso completado.");
        }
    }

    // Función para limpiar una hoja
    private static void ClearSheet(ISheet sheet)
    {
        for (int i = sheet.LastRowNum; i >= 0; i--)
        {
            IRow row = sheet.GetRow(i);
            if (row != null)
            {
                sheet.RemoveRow(row);
            }
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

    // Función para procesar las filas de la hoja 1 y escribir en la hoja 2
    private static void ProcesarFilas(ISheet xlHoja1, ISheet xlHoja2, int registros)
    {
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
                    IRow row2 = CrearFilaHoja2(xlHoja2, secuencia, j, hoja1Row);
                    ProcesarCeldas(xlHoja1, row2, hoja1Row, j);
                }
            }
        }
    }

    // Función para verificar si la celda no está vacía
    private static bool CeldaNoVacia(ISheet sheet, int fila, int columna)
    {
        return sheet.GetRow(5 + fila)?.GetCell(columna - 1) != null &&
               sheet.GetRow(5 + fila).GetCell(columna - 1).ToString() != "";
    }

    // Función para crear una nueva fila en la hoja 2
    private static IRow CrearFilaHoja2(ISheet sheet, int secuencia, int j, int hoja1Row)
    {
        int lon_secuencia = secuencia.ToString().Length;
        string nceros1 = new string('0', 8 - lon_secuencia);

        IRow row = sheet.CreateRow(1 + hoja1Row);
        row.CreateCell(0).SetCellValue(nceros1 + secuencia);
        row.CreateCell(1).SetCellValue(5);
        row.CreateCell(2).SetCellValue(394);

        string nceros4 = j < 10 ? "0" : "";
        row.CreateCell(3).SetCellValue(nceros4 + j);
        row.CreateCell(4).SetCellValue("01");

        int lon_reg = hoja1Row.ToString().Length;
        string nceros6 = new string('0', 6 - lon_reg);
        row.CreateCell(5).SetCellValue(nceros6 + hoja1Row);
        row.CreateCell(6).SetCellValue("+");

        return row;
    }

    // Función para procesar celdas y llenar la hoja 2
    private static void ProcesarCeldas(ISheet xlHoja1, IRow row, int hoja1Row, int j)
    {
        if (EsTexto(j))
        {
            string valorCelda = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).ToString();
            row.CreateCell(7).SetCellValue(valorCelda.PadRight(50));
        }
        else if (EsFecha(j))
        {
            DateTime fecha = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).DateCellValue.Value;
            string diaStr = fecha.Day < 10 ? "0" + fecha.Day : fecha.Day.ToString();
            string mesStr = fecha.Month < 10 ? "0" + fecha.Month : fecha.Month.ToString();
            string añoStr = fecha.Year.ToString();
            row.CreateCell(7).SetCellValue($"{diaStr}{mesStr}{añoStr}");
        }
        else if (EsNumero(j))
        {
            double valor = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).NumericCellValue;
            string campo_num = Math.Round(valor, 2).ToString("0.00").Replace(",", ".");
            row.CreateCell(7).SetCellValue(campo_num);
        }
    }

    // Funciones para identificar tipo de dato
    private static bool EsTexto(int columna) => new[] { 8, 10, 21, 47, 55, 63, 66, 76 }.Contains(columna);

    private static bool EsFecha(int columna) => new[] { 9, 11, 20, 46, 54, 62, 83 }.Contains(columna);

    private static bool EsNumero(int columna)
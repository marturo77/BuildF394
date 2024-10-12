using NPOI.SS.UserModel;    // Interfaz común para manipular hojas de cálculo
using NPOI.XSSF.UserModel;  // Para trabajar con archivos .xlsx

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

            // **Borrar el contenido de la hoja 2**
            // Eliminar todas las filas de la hoja 2 antes de empezar el proceso
            for (int i = xlHoja2.LastRowNum; i >= 0; i--)
            {
                IRow row = xlHoja2.GetRow(i);
                if (row != null)
                {
                    xlHoja2.RemoveRow(row);  // Eliminar la fila si no es nula
                }
            }

            string ruta = xlHoja1.GetRow(1).GetCell(5).ToString();
            int j = 1;

            // Verificar el valor en la primera columna hasta encontrar una celda vacía
            while (xlHoja1.GetRow(5 + j) != null && xlHoja1.GetRow(5 + j).GetCell(0) != null &&
                   xlHoja1.GetRow(5 + j).GetCell(0).CellType != CellType.Blank)
            {
                j++;
            }

            int registros = j - 1;
            int k = 0;
            int secuencia = 0;  // Declaración de secuencia fuera del ciclo

            for (j = 1; j <= 84; j++)
            {
                for (int hoja1Row = 1; hoja1Row <= registros; hoja1Row++)
                {
                    Console.WriteLine($"procesando fila {hoja1Row}");

                    if (xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1) != null && xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).ToString() != "")
                    {
                        k++;
                        secuencia = 3 + k;
                        int lon_secuencia = secuencia.ToString().Length;
                        string nceros1 = new string('0', 8 - lon_secuencia);

                        IRow row2 = xlHoja2.CreateRow(1 + k); // Crear fila en la hoja 2
                        row2.CreateCell(0).SetCellValue(nceros1 + secuencia);
                        row2.CreateCell(1).SetCellValue(5);
                        row2.CreateCell(2).SetCellValue(394);

                        if (j < 10)
                        {
                            string nceros4 = "0";
                            row2.CreateCell(3).SetCellValue(nceros4 + j);
                        }
                        else
                        {
                            row2.CreateCell(3).SetCellValue(j);
                        }

                        row2.CreateCell(4).SetCellValue("01");
                        int reg = hoja1Row;
                        int lon_reg = reg.ToString().Length;
                        string nceros6 = new string('0', 6 - lon_reg);
                        row2.CreateCell(5).SetCellValue(nceros6 + hoja1Row);
                        row2.CreateCell(6).SetCellValue("+");

                        if (j == 8 || j == 10 || j == 21 || j == 47 || j == 55 || j == 63 || j == 66 || j == 76)
                        {
                            string valorCelda = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).ToString();
                            int lon_camp = valorCelda.Length;
                            string nespacios = new string(' ', 50 - lon_camp);
                            row2.CreateCell(7).SetCellValue(valorCelda + nespacios);
                        }
                        else if (j == 9 || j == 11 || j == 20 || j == 46 || j == 54 || j == 62 || j == 83)
                        {
                            string nceros = new string('0', 17 - 8);
                            DateTime fecha = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).DateCellValue.Value;
                            string diaStr = fecha.Day < 10 ? "0" + fecha.Day : fecha.Day.ToString();
                            string mesStr = fecha.Month < 10 ? "0" + fecha.Month : fecha.Month.ToString();
                            string añoStr = fecha.Year.ToString();
                            row2.CreateCell(7).SetCellValue(nceros + diaStr + mesStr + añoStr);
                        }
                        else if (j == 4 || j == 6 || j == 7 || j == 30 || j == 32 || j == 37 || j == 79 || j == 80 || j == 81 || j == 82)
                        {
                            double valor = xlHoja1.GetRow(5 + hoja1Row).GetCell(j - 1).NumericCellValue;
                            double decimales = Math.Round(valor - Math.Floor(valor), 2);
                            string campo_num = decimales == 0 ? valor.ToString("0.00") : valor.ToString().Replace(",", ".");
                            row2.CreateCell(7).SetCellValue(campo_num);
                        }
                    }
                }
            }

            string nceros_sec = new string('0', 8 - (secuencia + 1).ToString().Length);
            int Total = secuencia + 1;
            string total_reg = nceros_sec + Total;

            string reg_tipo_1 = "0000001114000025" + xlHoja1.GetRow(0).GetCell(5).ToString().Substring(0, 2) + xlHoja1.GetRow(0).GetCell(5).ToString().Substring(3, 2) + total_reg + "SVIDCOLMENA0907";
            string reg_tipo_3 = "00000023000000000000000000000000";
            string reg_tipo_4 = "00000034000000000000000000000002";
            string reg_tipo_6 = total_reg + "6";

            IRow row3 = xlHoja3.CreateRow(0);
            row3.CreateCell(0).SetCellValue(reg_tipo_1);
            xlHoja3.CreateRow(1).CreateCell(0).SetCellValue(reg_tipo_3);
            xlHoja3.CreateRow(2).CreateCell(0).SetCellValue(reg_tipo_4);

            for (int row = 4; row <= secuencia; row++)
            {
                IRow row3M = xlHoja3.CreateRow(row);
                row3M.CreateCell(0).SetCellValue(
                    xlHoja2.GetRow(row - 2).GetCell(0).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(1).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(2).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(3).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(4).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(5).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(6).ToString() +
                    xlHoja2.GetRow(row - 2).GetCell(7).ToString());
            }

            xlHoja3.CreateRow(secuencia + 1).CreateCell(0).SetCellValue(reg_tipo_6);
            string nombre = xlHoja1.GetRow(0).GetCell(5).ToString().Substring(6, 4) + xlHoja1.GetRow(0).GetCell(5).ToString().Substring(3, 2) + "fto394.txt";

            // Guardar el archivo en formato .xlsx
            string rutaGuardado = Path.Combine(rutaEjecutable, nombre);
            using (FileStream output = new FileStream(rutaGuardado, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(output);
            }

            Console.WriteLine("Proceso completado.");
        }
    }
}
using NPOI.SS.UserModel;

namespace BuildF394
{
    /// <summary>
    /// Metodos de extension
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="pad"></param>
        /// <returns></returns>
        public static string PadZerosLeft(this int value, int pad = 8)
        {
            return value.ToString().PadLeft(pad, '0');
        }

        /// <summary>
        ///
        /// </summary>
        /// <param name="value"></param>
        /// <param name="pad"></param>
        /// <returns></returns>
        public static string PadZerosLeft(this double value, int pad = 8)
        {
            value = Math.Round(value, 2);
            return value.ToString("0.00").Replace(",", ".").PadLeft(pad, '0');
        }

        /// <summary>
        /// Funciones para identificar tipo de dato
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsText(this int index) => new[] { 20, 46, 54, 62, 65, 75 }.Contains(index);

        /// <summary>
        /// /
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        // Es la columna del codigo poliza
        public static bool IsInsuranceCode(this int index) => new[] { 7, 9 }.Contains(index);

        /// <summary>
        ///
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsRamo(this int index) => new[] { 0 }.Contains(index);

        /// <summary>
        ///
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsDate(this int index) => new[] { 8, 10, 19, 45, 53, 61, 82 }.Contains(index);

        /// <summary>
        ///
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsNumber(this int index) => new[] { 3, 5, 6, 29, 31, 36, 78, 79, 80, 81 }.Contains(index);

        /// <summary>
        /// Es una columna para consolidar
        /// </summary>
        /// <param name="index"></param>
        /// <returns></returns>
        public static bool IsConsolidation(this int index) => ConsolidationIndex().Contains(index);

        /// <summary>
        /// Indices de columnas que hay que consolidar o sumar
        /// </summary>
        /// <returns></returns>
        public static int[] ConsolidationIndex() => new[] { 26, 32, 33, 34, 37, 76, 77, 83 }.ToArray();

        /// <summary>
        /// Función para verificar si la celda no está vacía
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="row"></param>
        /// <param name="col"></param>
        /// <returns></returns>
        public static bool IsNotEmpty(this ISheet sheet, int row, int col)
        {
            object? value = sheet.GetRow(Program.ROW_OFFSET + row)?.GetCell(col);

            return value != null && value.ToString() != "";
        }
    }
}
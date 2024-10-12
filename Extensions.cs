using NPOI.SS.UserModel;

namespace BuildF394
{
    /// <summary>
    ///
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
        /// Funciones para identificar tipo de dato
        /// </summary>
        /// <param name="columna"></param>
        /// <returns></returns>
        public static bool EsTexto(int columna) => new[] { 8, 10, 21, 47, 55, 63, 66, 76 }.Contains(columna);

        /// <summary>
        ///
        /// </summary>
        /// <param name="columna"></param>
        /// <returns></returns>
        public static bool EsFecha(int columna) => new[] { 9, 11, 20, 46, 54, 62, 83 }.Contains(columna);

        /// <summary>
        ///
        /// </summary>
        /// <param name="columna"></param>
        /// <returns></returns>
        public static bool EsNumero(int columna) => new[] { 4, 6, 7, 30, 32, 37, 79, 80, 81, 82 }.Contains(columna);

        /// <summary>
        /// Función para verificar si la celda no está vacía
        /// </summary>
        /// <param name="sheet"></param>
        /// <param name="fila"></param>
        /// <param name="columna"></param>
        /// <returns></returns>
        public static bool CeldaNoVacia(this ISheet sheet, int fila, int columna)
        {
            return sheet.GetRow(5 + fila)?.GetCell(columna - 1) != null &&
                   sheet.GetRow(5 + fila).GetCell(columna - 1).ToString() != "";
        }
    }
}
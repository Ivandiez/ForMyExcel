using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using _Excel = Microsoft.Office.Interop.Excel;

namespace ForMyExcel
{
    class ExcelClass
    {
        readonly string path = "";   // путь к файлу
        readonly _Application excel = new _Excel.Application();
        Workbook wb;    // переменная для файла xlsx
        Worksheet ws;   // переменная для листа excel

        /// <summary>
        /// Конструктор для класса Excel без параметров
        /// </summary>
        public ExcelClass()
        {

        }

        /// <summary>
        /// Конструктор для класса Excel с параметрами
        /// </summary>
        /// <param name="path">путь к файлу</param>
        /// <param name="Sheet">лист</param>
        public ExcelClass(string path, int Sheet)
        {
            this.path = path;
            wb = excel.Workbooks.Open(path);
            ws = wb.Worksheets[Sheet];
        }

        /// <summary>
        /// Чтение нескольких ячеек из диапазона
        /// </summary>
        /// <param name="starti">начальное значение строки</param>
        /// <param name="starty">начальное значение столбца</param>
        /// <param name="endi">конечное значение строки</param>
        /// <param name="endy">конечное значение столбца</param>
        /// <returns>содержимое ячеек из диапазона</returns>
        public double[,] ReadRange(int starti, int starty, int endi, int endy)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            var holder = range.Value2;
            double[,] returndouble = new double[endi - starti + 1, endy - starty + 1];
            for (int p = 1; p <= endi - starti + 1; p++)
                for (int q = 1; q <= endy - starty + 1; q++)
                    returndouble[p - 1, q - 1] = holder[p, q];
            return returndouble;
        }

        /// <summary>
        /// Запись в диапазон ячеек
        /// </summary>
        /// <param name="starti">начальное значение строки</param>
        /// <param name="starty">начальное значение столбца</param>
        /// <param name="endi">конечное значение строки</param>
        /// <param name="endy">конечное значение столбца</param>
        /// <param name="writestring">Содержимое ячеек, которое будет записано</param>
        public void WriteRange(int starti, int starty, int endi, int endy, double[,] writestring)
        {
            Range range = (Range)ws.Range[ws.Cells[starti, starty], ws.Cells[endi, endy]];
            range.Value2 = writestring;
        }

        public void ClearRange()
        {
            Range range = ws.Range[ws.Cells[7, 3], ws.Cells[37, 5]];
            range.ClearContents();
        }

        /// <summary>
        /// Закрытие процесса excel
        /// </summary>
        public void Close()
        {
            wb.Close();
        }

        /// <summary>
        /// Сохранение файла excel по умолчанию в том же файле
        /// </summary>
        public void Save()
        {
            wb.Save();
        }

        /// <summary>
        /// Сохранение файла excel в какой-то новый файл
        /// </summary>
        /// <param name="path">путь к новому файлу</param>
        public void SaveAs(string path)
        {
            wb.SaveAs(path);
        }

    }
}

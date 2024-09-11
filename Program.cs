using System;
using System.IO;
using OfficeOpenXml; // Подключение библиотеки EPPlus

namespace First_Lab;
struct TableCorners{
    public int row {  get; set; }
    public int col { get; set; }
    public TableCorners(int x, int y){row = x;col = y;}
}

class Solution // Решение
{
    private double Qtotal, Qfact, Qost, Sfact;
    private double Xcp, Sost, Fpast, Fkrit = 3.88;
    private object[,] array;
    int N, K;
    private TableCorners[] Corners = new TableCorners[4];
    public void function(object [,] values, int maxRows, int maxCols) // функция для поиска границ таблицы
    {
        int currRow = 0;
        int currCol = 0;
        bool flag = false;
        for (int row = 0; row < maxRows; row++){
            for (int col = 0; col < maxCols; col++){   
                if (values[row, col] != null && !flag){
                    currRow = row;
                    currCol = col;
                    flag = true;
                }   
            }
        }
        Corners[0] = new TableCorners(currRow, currCol);
        flag = false;
        for (int row = Corners[0].row + 2; row < maxRows; row++){
            if (values[row, currCol] == null && !flag){
                currRow = row;
                flag = true;
            }
        }
        Corners[1] = new TableCorners(currRow - 1, currCol);
        N = Convert.ToInt32(values[Corners[1].row - 1, Corners[1].col]);
        flag = false;
        for(int col = Corners[1].col; col < maxCols;col++){
            if (values[Corners[1].row, col] == null && !flag)
            {
                currCol = col;
                flag = true;
            }
        }
        Corners[2] = new TableCorners(currRow - 1, currCol - 1);
        K = Corners[2].col - Corners[1].col;
        int u = Corners[2].row;
        int c = Corners[0].row;
        int j = Corners[2].col;
        Corners[3] = new TableCorners(Corners[0].row + 1, Corners[2].col);
        for (int i = Corners[1].col + 1; i <= Corners[2].col; i++){
            Xcp += Convert.ToDouble(values[Corners[1].row, i]);
            if (i == Corners[2].col) { Xcp /= K;}
        }
        array = values.Clone() as object[,];  
    }
    public void Compute(){ // Вычисления
        int count = 0, currRow, currCol;
        double summ = 0;
        double[] amounts = new double[K];
        while (count != K){
            currRow = Corners[0].row + 2;
            currCol = Corners[0].col + 1 + count;
            for(int i = currRow;i < currRow + N;i++){
                summ += Math.Pow((Convert.ToDouble(array[i, currCol]) - Xcp), 2);
            }
            amounts[count] = summ;
            count += 1;
        }
        Qtotal = amounts[K - 1];
        summ = 0;
        for (int col = Corners[1].col + 1;col <= Corners[2].col;col++){
            summ += Math.Pow((Convert.ToDouble(array[Corners[1].row, col]) - Xcp), 2);
        }
        Qfact = 5 * summ;
        Qost = Qtotal - Qfact;
        Sfact = Qfact/(K - 1);
        Sost = Qost /(K * (N - 1));
        Fpast = Sfact / Sost;
    }
    public void check_null_hypotheses() // Проверка на нулевую гипотезу
    {
        if (Fpast < Fkrit)
        {
            Console.WriteLine($"Т.к. Fpast = {Fpast} < Fkrit = {Fkrit} значит");
            Console.WriteLine("Фактор(участок) существенно влияет на качество продукции.");
        }
        else
        {
            Console.WriteLine($"Т.к. Fpast = {Fpast} > Fkrit = {Fkrit} значит");
            Console.WriteLine("Фактор(участок) существенно не влияет на качество продукции.");
        }
    }
}
class Run
{
    static void Main(string[] args)
    {
        // Указываем путь к файлу Excel
        var FilePath = Path.Combine(Directory.GetCurrentDirectory(), "Data.xlsx");
        var File = new FileInfo(FilePath);
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        if(!File.Exists) // Проверка на наличие файла
                {
            Console.WriteLine($"Файл не найден: {FilePath}");
            return;
        }
        // Открываем файл Excel
        using (var package = new ExcelPackage(File))
        {
            var List_Exel = package.Workbook.Worksheets[1];
            int maxRows = 20, maxCols = 20;
            object[,] values = new object[maxRows, maxCols];
            for (int row = 0; row < maxRows; row++){
                for(int col = 0; col < maxCols; col++){
                    values[row, col] = List_Exel.Cells[row + 1, col + 1].Value; // Сохранение значения в массив
                }
            }
            Solution solution = new Solution(); // Создаю экземпляр класса Solution
            solution.function(values,maxRows, maxCols);
            solution.Compute();
            solution.check_null_hypotheses();
        }
    }
}
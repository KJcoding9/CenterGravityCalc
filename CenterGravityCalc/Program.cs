using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Xls;
using DocumentFormat.OpenXml.Drawing;
using System.Numerics;
using DocumentFormat.OpenXml.InkML;

namespace CenterGravityCalc
{
    internal class Program
    {
        List<double> mG = new List<double>(); //圧力の合計値
        List<double> xG = new List<double>(); 
        List<double> yG = new List<double>();

        double[] offset = new double[32];

        List<double[]> digitData = new List<double[]>();
        static void Main(string[] args)
        {
            Program p = new Program();

            Console.WriteLine("ファイルの総数を入力してください");

            var fileCount = int.Parse(Console.ReadLine());

            Console.WriteLine("テキストファイルが全て入ったフォルダパスを入力してください");

            string textFilesPath = Console.ReadLine();

            string[] textFiles = Directory.GetFiles(textFilesPath).
                OrderBy(f => f).ToArray();

            Console.WriteLine("csvファイルを生成するフォルダパスを入力してください");

            string csvFilesPath = Console.ReadLine();

            for(int i=0; i<fileCount; i++)
            {
                string fileName = csvFilesPath + "\\t" + i+".csv";
                p.CreateEmptyCsvFile(fileName);
            }

            string[] csvFiles = Directory.GetFiles(csvFilesPath).
            OrderBy(f => f).ToArray();

            int count = 0;
            foreach (string textFile in textFiles)
            {
                p.FileRead(textFile,csvFiles[count]);
                string excelFileName = "t" + count+".xlsx";
                Workbook workBook = new Workbook();
                p.ExcelConvert(csvFiles[count], workBook, excelFileName);

                count++;

            }

            Console.WriteLine("重心位置算出結果用のエクセルファイルパスを入力してください");

            var result = Console.ReadLine();

            Console.WriteLine("生成されたファイルをまとめたフォルダのパスを入力してください");

            string excelPath = Console.ReadLine();

            string[] files = Directory.GetFiles(excelPath).
                OrderBy(f => f).ToArray();

            foreach(string file in files)
            {
                var workBook = new XLWorkbook(file);

                var workSheet = workBook.Worksheets.Worksheet(1);

                p.Calc(workSheet); //これが毎回今までの分加算されている。
            }

            p.ExcelFillinData(result);

            //var workBook2 = new XLWorkbook(result);

            //Console.WriteLine("座ってないデータが入っていましたか？ はい:1いいえ:2");

            //int answer = int.Parse(Console.ReadLine());

            //if (answer == 1)
            //{
            //    p.OffsetAdaptation(workBook2, result);
            //}
        }

        //テキストファイルからCSVファイルに変換する
        void FileRead(string textPath,string csvPath)
        {
            string[] lines = File.ReadAllLines(textPath);

            using(StreamWriter writer = new StreamWriter(csvPath))
            {
                foreach(string line in lines)
                {
                    string[] values = line.Split('\t');

                    writer.WriteLine(string.Join(",", values));
                }
            }
        }

        //指定されたxとy座標で重心位置を算出する（全部で32CH)
        void Calc(IXLWorksheet worksheet)
        {
            var lastRow = worksheet.LastRowUsed().RowNumber();

            //var pressureData = new List<double[]>();

            int a=0;
            for (int i = 7; i <= lastRow; i++)
            {
                digitData.Add(new double[33]);
                //pressureData.Add(new double[33]);
                string[] t = new string[33];
                for (int j = 1; j <= 32; j++)
                {
                    t[j-1] = worksheet.Cell(i, j + 2).Value.ToString();

                    digitData[a][j-1] = double.Parse(t[j-1]);

                    if (a == 0)
                    {
                        offset[j - 1] = digitData[a][j - 1];
                    }
                }
                a++;
            }
            var coordinateX = new int[32] { 0, 0, 0, 0, 0, 0, 0, 0, 24, 24, 24, 24, 24, 24, 24, 24, 48, 48, 48, 48, 48, 48, 48, 48,
            72,72,72,72,72,72,72,72};

            var coordinateY = new int[32] { 0,24,48,72,96,120,144,168, 0, 24, 48, 72, 96, 120, 144, 168 , 0, 24, 48, 72, 96, 120, 144, 168,
            0,24,48,72,96,120,144,168};

            for (int i=0; i<lastRow-6; i++)
            {
                double sum = 0;
                double moluculeX = 0;
                double moluculeY = 0;
                for(int j=0; j<32; j++)
                {
                    moluculeX += coordinateX[j] * digitData[i][j];
                    moluculeY += coordinateY[j] * digitData[i][j];
                    sum += digitData[i][j];
                }
                xG.Add(moluculeX / sum);
                yG.Add(moluculeY / sum);
                mG.Add(sum);
            }
        }

        //CSVからエクセルに変換する
        void ExcelConvert(string csvPath,Workbook book,string fileName)
        {
            try
            {
                //CSVファイルをロード
                book.LoadFromFile(@csvPath, ",", 1, 1);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message + "正しくファイルが読み込まれませんでした。");
                book.Save();
            }
            Worksheet sheet = book.Worksheets[0];
            //ワークシート使用範囲にアクセス
            CellRange usedRange = sheet.AllocatedRange;

            //数値をテキストとして保存するときにエラーを無視
            usedRange.IgnoreErrorOptions = IgnoreErrorType.NumberAsText;

            //行の高さと列の幅を自動調整
            usedRange.AutoFitColumns();
            usedRange.AutoFitRows();

            try
            {
                book.SaveToFile(fileName, ExcelVersion.Version2016);
            }
            catch (Exception e)
            {
                Console.WriteLine(e + "正しくエクセルファイルに変換できませんでした");
                book.Save();
            }
        }

        //算出したデータをエクセルに記入する
        void ExcelFillinData(string excelPath)
        {
            var workbook = new XLWorkbook(excelPath);

            var worksheet = workbook.Worksheets.Worksheet(1);
            double timeStamp=0;
            // var worksheet = workbook.AddWorksheet("重心");

            worksheet.Cell(2, 1).Value = "TimeStamp";
            worksheet.Cell(2,2).Value = "X";
            worksheet.Cell(2,3).Value = "Y";
            worksheet.Cell(2,4).Value = "MG";

            //var lastRow = worksheet.LastRowUsed().RowNumber();


            for (int i = 3; i < mG.Count - 3; i++)
            {
                worksheet.Cell(i, 1).Value = timeStamp;
                worksheet.Cell(i, 2).Value = xG[i];
                worksheet.Cell(i, 3).Value = yG[i];
                worksheet.Cell(i, 4).Value = mG[i];

                timeStamp += 0.016;
            }

            try
            {
                workbook.SaveAs(excelPath);
            }
            catch(Exception e)
            {
                Console.WriteLine(e+"正しく保存できませんでした。");
            }
        }

        void OffsetAdaptation(XLWorkbook workbook, string excelPath)
        {
            Console.WriteLine("座っていないデータ数を入力してください");
            var notSitting = int.Parse(Console.ReadLine());
            var worksheet = workbook.Worksheet("重心");
            List<double[]> offsetAdaptationData = new List<double[]>();

            List<double> adxG = new List<double>();
            List<double> adyG = new List<double>();

            for (int i = 3; i < notSitting - 3; i++)
            {
                string[] t = new string[33];
                for (int j = 1; j <= 32; j++)
                {
                    t[j - 1] = worksheet.Cell(i, j + 2).Value.ToString();

                    offsetAdaptationData[i][j - 1] = digitData[i][j - 1] - offset[j];
                }
            }
            var coordinateX = new int[32] { 0, 0, 0, 0, 0, 0, 0, 0, 24, 24, 24, 24, 24, 24, 24, 24, 48, 48, 48, 48, 48, 48, 48, 48,
            72,72,72,72,72,72,72,72};

            var coordinateY = new int[32] { 0,24,48,72,96,120,144,168, 0, 24, 48, 72, 96, 120, 144, 168 , 0, 24, 48, 72, 96, 120, 144, 168,
            0,24,48,72,96,120,144,168};

            for (int i = 0; i < notSitting - 3; i++)
            {
                double sum = 0;
                double moluculeX = 0;
                double moluculeY = 0;

                for (int j = 0; j < 32; j++)
                {
                    moluculeX += coordinateX[j] * offsetAdaptationData[i][j];
                    moluculeY += coordinateY[j] * offsetAdaptationData[i][j];
                    sum += offsetAdaptationData[i][j];
                }
                adxG.Add(moluculeX / sum);
                adyG.Add(moluculeY / sum);

                worksheet.Cell(i + 3, 2).Value = adxG[i];
                worksheet.Cell(i + 3, 3).Value = adyG[i];
            }

            try
            {
                workbook.SaveAs(excelPath);
            }
            catch (Exception e)
            {
                Console.WriteLine(e + "正しく保存できませんでした。");
            }
        }

        void CreateEmptyCsvFile(string filePath)
        {
            using (StreamWriter sw = new StreamWriter(filePath, false))
            {

            }
        }
    }
}

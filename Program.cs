using System;
using System.Text.RegularExpressions;
using System.Drawing;
using Excel = Microsoft.Office.Interop.Excel;

namespace MonorisuApp
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("y→スタート，exit→終了");

            string image_file = "";
            string dangan_file = "Danganronpa V3_ Killing Harmony 20[0-9]*_[0-9]*_[0-9]* [0-9]*_[0-9]*_[0-9]*.png";
            int image_file_char = 0;
            var danron_ss = new List<string>() { };
            string new_danron_ss;

            int[,] monoris_block = new int[22, 11];
            int[,] monoris_block_Hanten = new int[11, 22];
            Bitmap bitmap;
            byte[,,] monoRGB;

            try
            {
                //image_file = Directory.GetCurrentDirectory() + @"\image\";
                image_file = System.Environment.GetFolderPath(Environment.SpecialFolder.MyVideos) + @"\Captures\";
                image_file_char = image_file.Length;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            while (true)
            {
                string line = Console.ReadLine();
                Console.WriteLine(line);

                if (line == "y")
                {
                    try
                    {
                        string[] image_name = Directory.GetFiles(image_file, "*.png");
                        foreach (string a in image_name)
                        {
                            string x = a.Remove(0, image_file_char);
                            if (Regex.IsMatch(x, @dangan_file))
                            {
                                danron_ss.Add(x);
                            }
                        }
                        int new_ss_count = danron_ss.Count - 1;
                        new_danron_ss = danron_ss[new_ss_count];
                        Console.WriteLine("最新は→" + new_danron_ss);

                        bitmap = new Bitmap(image_file + new_danron_ss);

                        monoRGB = new byte[22, 11, 3];
                        int monoBLK_x = 0;
                        int monoBLK_y = 0;

                        for (int x = 120; x <= 1800; x += 80)
                        {
                            if (x == 120)
                            {
                                monoBLK_x = 0;
                            }
                            else
                            {
                                monoBLK_x++;
                            }
                            for (int y = 120; y <= 920; y += 80)
                            {
                                if (y == 120)
                                {
                                    monoBLK_y = 0;
                                }
                                else { monoBLK_y++; }


                                Color pixel = bitmap.GetPixel(x, y);
                                byte R = pixel.R;
                                byte G = pixel.G;
                                byte B = pixel.B;
                                monoRGB[monoBLK_x, monoBLK_y, 0] = R;
                                monoRGB[monoBLK_x, monoBLK_y, 1] = G;
                                monoRGB[monoBLK_x, monoBLK_y, 2] = B;
                                //Console.Write(monoBLK_x + " " + monoBLK_y + " " +monoRGB[monoBLK_x, monoBLK_y, 0] + " " + monoRGB[monoBLK_x, monoBLK_y, 1] + " " + monoRGB[monoBLK_x, monoBLK_y, 2] + "\n");

                                switch (monoRGB[monoBLK_x, monoBLK_y, 0])
                                {
                                    case 208:
                                        if (monoRGB[monoBLK_x, monoBLK_y, 1] == 208 && monoRGB[monoBLK_x, monoBLK_y, 2] == 208)
                                        {
                                            monoris_block[monoBLK_x, monoBLK_y] = 1;
                                            break;
                                        }
                                        monoris_block[monoBLK_x, monoBLK_y] = 0;
                                        break;

                                    case 250:
                                        if (monoRGB[monoBLK_x, monoBLK_y, 1] == 162 && monoRGB[monoBLK_x, monoBLK_y, 2] == 208)
                                        {
                                            monoris_block[monoBLK_x, monoBLK_y] = 2;
                                            break;
                                        }
                                        monoris_block[monoBLK_x, monoBLK_y] = 0;
                                        break;

                                    case 229:
                                        if (monoRGB[monoBLK_x, monoBLK_y, 1] == 192 && monoRGB[monoBLK_x, monoBLK_y, 2] == 113)
                                        {
                                            monoris_block[monoBLK_x, monoBLK_y] = 3;
                                            break;
                                        }
                                        monoris_block[monoBLK_x, monoBLK_y] = 0;
                                        break;

                                    case 97:
                                        if (monoRGB[monoBLK_x, monoBLK_y, 1] == 186 && monoRGB[monoBLK_x, monoBLK_y, 2] == 205)
                                        {
                                            monoris_block[monoBLK_x, monoBLK_y] = 4;
                                            break;
                                        }
                                        monoris_block[monoBLK_x, monoBLK_y] = 0;
                                        break;

                                    default:
                                        monoris_block[monoBLK_x, monoBLK_y] = 0;
                                        break;
                                }

                            }
                        }

                        for (int i = 0; i < monoBLK_y + 1; i++)
                        {

                            for (int j = 0; j < monoBLK_x + 1; j++)
                            {
                                monoris_block_Hanten[i, monoBLK_x - j] = monoris_block[monoBLK_x - j, i];
                            }
                        }
                        for (int i = 0; i < monoBLK_y + 1; i++)
                        {

                            for (int j = 0; j < monoBLK_x + 1; j++)
                            {
                                Console.Write(monoris_block_Hanten[i, j] + " ");
                            }
                            Console.Write("\n");
                        }

                        try
                        {
                            // Excel操作用オブジェクト
                            Microsoft.Office.Interop.Excel.Application xlApp = null;
                            Microsoft.Office.Interop.Excel.Workbooks xlBooks = null;
                            Microsoft.Office.Interop.Excel.Workbook xlBook = null;
                            Microsoft.Office.Interop.Excel.Sheets xlSheets = null;
                            Microsoft.Office.Interop.Excel.Worksheet xlSheet = null;

                            // Excelアプリケーション生成
                            xlApp = new Microsoft.Office.Interop.Excel.Application();

                            // 既存のBookを開く
                            xlBooks = xlApp.Workbooks;
                            xlBook = xlBooks.Open(Directory.GetCurrentDirectory() + @"/monoautosimu13.xlsm");

                            // シートを選択する
                            xlSheets = xlBook.Worksheets;
                            // 1シート目を操作対象に設定する
                            // ※Worksheets[n]はオブジェクト型を返すため、Worksheet型にキャスト
                            xlSheet = xlSheets[1] as Microsoft.Office.Interop.Excel.Worksheet;

                            // 表示
                            xlApp.Visible = true;

                            Microsoft.Office.Interop.Excel.Range xlCellsFrom = null;   //セル始点（中継用）
                            Microsoft.Office.Interop.Excel.Range xlRangeFrom = null;   //セル始点
                            Microsoft.Office.Interop.Excel.Range xlCellsTo = null;     //セル終点（中継用）
                            Microsoft.Office.Interop.Excel.Range xlRangeTo = null;     //セル終点   
                            Microsoft.Office.Interop.Excel.Range xlTarget = null;      //配列設定対象レンジ

                            //始点セル取得
                            xlCellsFrom = xlSheet.Cells;
                            xlRangeFrom = xlCellsFrom[2, 2] as Microsoft.Office.Interop.Excel.Range;

                            //終点セル取得
                            xlCellsTo = xlSheet.Cells;
                            xlRangeTo = xlCellsTo[12, 23] as Microsoft.Office.Interop.Excel.Range;

                            //貼り付け対象レンジ
                            xlTarget = xlSheet.Range[xlRangeFrom, xlRangeTo];

                            //◆値設定◆
                            xlTarget.Value = monoris_block_Hanten;

                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlTarget);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRangeTo);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellsTo);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlRangeFrom);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlCellsFrom);

                            // Sheet解放
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheet);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlSheets);

                            // Book解放
                            //xlBook.Close();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBook);
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlBooks);

                            // Excelアプリケーションを解放
                            //xlApp.Quit();
                            System.Runtime.InteropServices.Marshal.ReleaseComObject(xlApp);

                            //https://excelcsharp.lance40.com/post-4.html
                        }
                        catch (Exception e)
                        {
                            Console.WriteLine(e.Message);
                            throw;
                        }

                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.ToString());
                    }
                }
                else
                {
                    if(line == "exit")
                    {
                        break;
                    }
                }
            }
        }
    }
}

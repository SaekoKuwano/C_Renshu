using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OfficeOpenXml;
using System.IO;
using System.Text.RegularExpressions;

namespace ConsoleApp1
{
    internal class Program
    {
        private static void Main(string[] args)
        {
            // ファイルをすべて取得する
            DirectoryInfo di = new DirectoryInfo(@"C:\Users\ethna\OneDrive\デスクトップ\EPPlus_Test");

            IEnumerable<FileInfo> files = di.EnumerateFiles("*", SearchOption.AllDirectories);

            // 削除対象のシート名
            string DelName = "DellTests";

            // 要素数分処理を繰り返す
            foreach (FileInfo f in files)
            {
                Console.WriteLine(f.FullName);

                // インスタンス化する
                FileInfo fileInfo = new FileInfo(f.FullName);

                // Excelファイル作成
                using (ExcelPackage package = new ExcelPackage(fileInfo))
                {
                    // シートが存在するか確認
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Where(s => s.Name == DelName).FirstOrDefault();

                    if (worksheet != null)
                    {
                        // シートを削除
                        package.Workbook.Worksheets.Delete(DelName);

                        // 保存
                        package.Save();
                    }
                    else
                    {
                        // コメント
                        Console.WriteLine("シート存在なし。次のループへ");
                    }
                }

                // 正規表現
                Regex reg = new Regex("(result)");

                // リネーム名作成
                string result = reg.Replace(f.FullName, "チェンジ");

                // ファイルが存在しているか確認
                if (File.Exists(result))
                {
                    Console.WriteLine("リネーム名存在あり。次のループへ");

                    continue;
                }

                // リネームする
                fileInfo.MoveTo(result);

                // 読取り専用にする
                File.SetAttributes(result, FileAttributes.ReadOnly);
            }

            #region Excel操作

            //// 出力ファイルの準備（実行ファイルと同じフォルダに出力される）
            //FileInfo newFile = new FileInfo(Fn);

            //if (newFile.Exists)
            //{
            //    newFile.Delete();
            //    newFile = new FileInfo(Fn);
            //}

            //// Excelファイル作成
            //using (ExcelPackage package = new ExcelPackage(newFile))
            //{
            //    // ワークシートを一枚追加
            //    ExcelWorksheet sheet = package.Workbook.Worksheets.Add("testSeets");
            //    ExcelWorksheet sheetA = package.Workbook.Worksheets.Add("DellTests");

            //    // A1セルに書き込み
            //    sheet.Cells["A1"].Value = "Hello World";

            //    // セルはR1C1形式でも指定可
            //    sheet.Cells[2, 1].Value = 27;

            //    // シートを削除
            //    package.Workbook.Worksheets.Delete("DellTests");

            //    // 保存
            //    package.Save();
            //}

            #endregion Excel操作

            Console.WriteLine("続行するには何かキーを押してください．．．");
            Console.ReadKey();
        }
    }
}
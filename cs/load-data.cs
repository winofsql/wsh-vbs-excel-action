using System;
using System.Windows.Forms;
using System.Diagnostics;

namespace load_data
{
    class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            var path = $@"{Environment.CurrentDirectory}\syain.xlsx";

            // Excel アプリケーション
            dynamic excelApp =
                Activator
                    .CreateInstance(Type
                        .GetTypeFromProgID("Excel.Application"));

            // Excel を表示( 完成したらコメント化 )
            excelApp.Visible = true;

            // 警告を出さない
            excelApp.DisplayAlerts = false;

            // Excel ブック( 既存 )
            dynamic workBook = excelApp.Workbooks.Open(path);

            // 最大化
            excelApp.ActiveWindow.WindowState = -4137;

            MessageBox.Show("STOP");

            // 保存した事にする
            workBook.Saved = true;

            // 終了
            excelApp.Quit();

            // 解放
            System.Runtime.InteropServices.Marshal.ReleaseComObject (excelApp);

            // C# ではほぼ完全解放無理なので強制終了させる
            foreach (var p in Process.GetProcessesByName("EXCEL"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }
            // dotnet.exe が常駐すると他のコマンドアプリの excel 解放処理を邪魔する
            foreach (var p in Process.GetProcessesByName("dotnet"))
            {
                if (p.MainWindowTitle == "")
                {
                    p.Kill();
                }
            }

        }
    }
}

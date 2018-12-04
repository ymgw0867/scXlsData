using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using scXlsData.common;
using Excel = Microsoft.Office.Interop.Excel;

namespace scXlsData
{
    public partial class frmMenu : Form
    {
        public frmMenu()
        {
            InitializeComponent();
        }

        string xlsFname = string.Empty;
        string xlsPass = string.Empty;
        int xlsJyokenFormat = 0;

        private void button4_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void frmMenu_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (xlsFname == string.Empty)
            {
                MessageBox.Show("解約管理表ファイルが本システムに設定されていません。「環境設定」より使用する解約管理表ファイルを登録してください。", "ファイル未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (!isXlsFileExists(xlsFname))
            {
                MessageBox.Show("設定されている解約管理表ファイルが存在しません。「環境設定」より使用する解約管理表ファイルを登録してください。", "ファイル未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            string sPath = System.IO.Path.GetDirectoryName(xlsFname);

            // 自らのロックファイルを削除する
            Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());

            // 他のPCで処理中の場合、続行不可
            if (Utility.existsLockFile(sPath))
            {
                MessageBox.Show("他のPCが解約管理表に接続しています。再度実行してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            Hide();
            Form1 frm = new Form1(xlsFname, xlsPass, xlsJyokenFormat);
            frm.ShowDialog();
            Show();
        }

        private bool isXlsFileExists(string sFile)
        {
            if (!System.IO.File.Exists(sFile))
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        private void frmMenu_Load(object sender, EventArgs e)
        {
            // 環境設定読み込み
            loadConfig();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Hide();
            config.frmConfig frm = new config.frmConfig();
            frm.ShowDialog();
            Show();

            // 環境設定読み込み
            loadConfig();
        }

        ///-----------------------------------------------------
        /// <summary>
        ///     環境設定情報読み込み </summary>
        ///-----------------------------------------------------
        private void loadConfig()
        {
            DataSet1 dts = new DataSet1();
            DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();

            adp.FillByID(dts.環境設定);

            var s = dts.環境設定.Single(a => a.ID == 1);

            if (s.IstargetXlsFileNull())
            {
                xlsFname = string.Empty;
            }
            else
            {
                xlsFname = s.targetXlsFile;
            }

            if (s.IssheetPasswordNull())
            {
                xlsPass = string.Empty;
            }
            else
            {
                xlsPass = s.sheetPassword;
            }

            if (s.Is新規条件付き書式設定Null())
            {
                xlsJyokenFormat = 0;
            }
            else
            {
                xlsJyokenFormat = s.新規条件付き書式設定;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (xlsFname == string.Empty)
            {
                MessageBox.Show("解約管理表ファイルが本システムに設定されていません。「環境設定」より使用する解約管理表ファイルを登録してください。", "ファイル未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            if (!isXlsFileExists(xlsFname))
            {
                MessageBox.Show("設定されている解約管理表ファイルが存在しません。「環境設定」より使用する解約管理表ファイルを登録してください。", "ファイル未登録", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                return;
            }

            // ファイルを指定したファイル名・パスにコピー
            string fName = copyXlsFile(xlsFname);

            this.Cursor = Cursors.WaitCursor;

            // コピーしたエクセルファイルの読み取りパスワードを解除する
            if (fName != string.Empty)
            {
                xlsPassUnLock(fName, xlsPass, "");

                MessageBox.Show("解約管理表ファイルの出力が終了しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }

            this.Cursor = Cursors.Default;
        }

        ///-----------------------------------------------------------
        /// <summary>
        ///     ファイルを指定のパス名前で保存する </summary>
        /// <param name="sFile">
        ///     入力元ファイル</param>
        /// <returns>
        ///     出力ファイル名</returns>
        ///-----------------------------------------------------------
        private string copyXlsFile(string sFile)
        {
            try
            {
                //はじめのファイル名を指定する
                //はじめに「ファイル名」で表示される文字列を指定する
                //saveFileDialog1.FileName = "新しいファイル.html";
                saveFileDialog1.FileName = System.IO.Path.GetFileName(xlsFname);

                ////はじめに表示されるフォルダを指定する
                //saveFileDialog1.InitialDirectory = @"C:\";

                //[ファイルの種類]に表示される選択肢を指定する
                //指定しない（空の文字列）の時は、現在のディレクトリが表示される
                saveFileDialog1.Filter = "エクセルファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

                //[ファイルの種類]ではじめに選択されるものを指定する
                //2番目の「すべてのファイル」が選択されているようにする
                saveFileDialog1.FilterIndex = 2;

                //タイトルを設定する
                saveFileDialog1.Title = "解約管理表の保存先のファイルを選択してください";

                //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                saveFileDialog1.RestoreDirectory = true;

                //既に存在するファイル名を指定したとき警告する
                //デフォルトでTrueなので指定する必要はない
                saveFileDialog1.OverwritePrompt = true;

                //存在しないパスが指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                saveFileDialog1.CheckPathExists = true;

                //ダイアログを表示する
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    // 存在するファイルのとき削除する
                    if (System.IO.File.Exists(saveFileDialog1.FileName))
                    {
                        System.IO.File.Delete(saveFileDialog1.FileName);
                    }

                    // 指定ファイルにコピー
                    System.IO.File.Copy(sFile, saveFileDialog1.FileName);

                    // コピー先ファイル名を返す
                    return saveFileDialog1.FileName;
                }
                else
                {
                    return string.Empty;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return string.Empty;
            }
            finally
            {
                // 不要になった時点で破棄する
                saveFileDialog1.Dispose();
            }
        }

        ///--------------------------------------------------------------------------------
        /// <summary>
        ///     指定のエクセルファイルを読み取りパスワードを解除して保存する </summary>
        /// <param name="sFile">
        ///     パスを含む指定エクセルファイル名</param>
        /// <param name="rPw">
        ///     読み取りパスワード</param>
        /// <param name="wPw">
        ///     書き込みパスワード</param>
        ///--------------------------------------------------------------------------------
        private void xlsPassUnLock(string sFile, string rPw, string wPw)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = null;

            try
            {
                // Excelファイルを開く（ファイルパスワード付き）
                oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sFile, Type.Missing, Type.Missing, Type.Missing,
                    rPw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                oXls.DisplayAlerts = false;

                // Excelファイル書き込み（ファイルパスワード解除）
                oXlsBook.SaveAs(sFile, Type.Missing, wPw, Type.Missing, Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                // Excelを終了
                oXls.Quit();

                // COM オブジェクトの参照カウントを解放する 
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);

                if (oXlsBook != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                }

                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;

                GC.Collect();
            }
        }
    }
}

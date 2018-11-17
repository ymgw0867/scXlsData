using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using ClosedXML.Excel;
using scXlsData.common;

namespace scXlsData
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();


        #region グリッドカラム定義
        string colBuCode = "c1";        // 物件ＣＤ
        string colBuName = "c2";        // 物件名
        string colGou = "c3";           // 号室
        string colNewStayDate = "c4";   // 新規入居開始日

        // 解約申し込み
        string col_KaiyakuContact_01 = "c5";    // 解約申込日
        string col_KaiyakuContact_02 = "c6";    // 店舗からの鍵受取
        string col_KaiyakuContact_03 = "c7";    // 解約日
        string col_KaiyakuContact_04 = "c8";    // 立会日
        string col_KaiyakuContact_05 = "c9";    // 立会時間
        string col_KaiyakuContact_06 = "c10";   // 立会費用請求
        string col_KaiyakuContact_07 = "c11";   // 立会費用入金
        string col_KaiyakuContact_08 = "c12";   // 解約申し込み担当

        // 解約
        string col_Kaiyaku_01 = "c13";  // 鍵返却日  
        string col_Kaiyaku_02 = "c14";  // 鍵受取場所
        string col_Kaiyaku_03 = "c15";  // 退去確認
        string col_Kaiyaku_04 = "c16";  // ＲＣ依頼書

        // ルームチェック
        string col_RoomCheck_01 = "c17";    // ルームチェック 
        string col_RoomCheck_02 = "c18";    // ＲＣ依頼→ルームチェックまでの日数
        string col_RoomCheck_03 = "c19";    // 鍵交換日
        string col_RoomCheck_04 = "c20";    // 鍵置き場
        string col_RoomCheck_05 = "c21";    // ルームチェック担当

        // 書類作成
        string col_Shorui_01 = "c22"; // レジ→CS書類提出
        string col_Shorui_02 = "c23"; // 工事見積書精査済み
        string col_Shorui_03 = "c24"; // そなえーる
        string col_Shorui_04 = "c25"; // 書類作成担当

        // 手続き
        string col_Tetsu_01 = "c26"; // 営業担当確認
        string col_Tetsu_02 = "c27"; // オーナー見積書送付
        string col_Tetsu_03 = "c28"; // ②見積書送付→本日までの日数
        string col_Tetsu_04 = "c29"; // オーナー承諾日
        string col_Tetsu_05 = "c30"; // テナント明細発送

        // 発注
        string col_Hacchu_01 = "c31"; // 発注
        string col_Hacchu_02 = "c32"; // ルームチェック→発注までの日数
        string col_Hacchu_03 = "c33"; // 発注担当

        // 工事着工
        string col_Kouji_01 = "c34"; // 業者名
        string col_Kouji_02 = "c35"; // 工事代発注金額
        string col_Kouji_03 = "c36"; // 工事着工
        string col_Kouji_04 = "c37"; // 工事終了予定日
        string col_Kouji_05 = "c38"; // 工事終了
        string col_Kouji_06 = "c39"; // 発注→工事終了までの日数
        string col_Kouji_07 = "c40"; // ＲＣ依頼→工事終了までの日数

        // 完了検査
        string col_Kanryo_01 = "c41"; // 検査依頼日
        string col_Kanryo_02 = "c42"; // 検査日
        string col_Kanryo_03 = "c43"; // 検査完了
        string col_Kanryo_04 = "c44"; // 完了検査担当

        // 備考
        string col_Bikou_01 = "c45"; // 備考

        // スカイワン
        string col_SkyOne_01 = "c46"; // 金額
        string col_SkyOne_02 = "c47"; // 保証会社

        // 行番号
        string col_xlsRowNum = "c48";



        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            Dispose();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            adp.FillByID(dts.環境設定);

            var s = dts.環境設定.Single(a => a.ID == 1);
            string xlsFname = s.targetXlsFile;
            string xlsPass = s.sheetPassword;

            impXlsSheet(xlsFname, xlsPass);

            GridViewSetting(dataGridView1);

            gridViewShowData(dataGridView1);

            dispInitial();
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
        }

        private void impXlsSheet(string sPath, string sPassWord)
        {
            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = null;
            Excel.Worksheet oxlsSheet = null;
            //Excel.Worksheet sheet = null;
            
            try
            {
                // Excelファイルを開く（ファイルパスワード付き）
                oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sPath, Type.Missing, Type.Missing, Type.Missing,
                    "abcd", Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                oxlsSheet = (Excel.Worksheet)oXlsBook.Sheets[Properties.Settings.Default.xlsSheetName];
                
                // シート保護解除
                //oxlsSheet.Unprotect(sPassWord);

                // ローカルフォルダのExcelファイル書き込み済みのとき削除する
                if (System.IO.File.Exists(Properties.Settings.Default.imPortPath))
                {
                    System.IO.File.Delete(Properties.Settings.Default.imPortPath);
                }

                // ローカルフォルダへExcelファイル書き込み（ファイルパスワード解除）
                oXlsBook.SaveAs(Properties.Settings.Default.imPortPath, Type.Missing, "", Type.Missing, Type.Missing, Type.Missing,
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
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oxlsSheet);
                //System.Runtime.InteropServices.Marshal.ReleaseComObject(sheet);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXlsBook);
                System.Runtime.InteropServices.Marshal.ReleaseComObject(oXls);

                oXls = null;
                oXlsBook = null;
                oxlsSheet = null;
                //sheet = null;

                GC.Collect();
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューの定義を行います </summary>
        /// <param name="tempDGV">
        ///     データグリッドビューオブジェクト</param>
        ///---------------------------------------------------------------------
        public void GridViewSetting(DataGridView tempDGV)
        {
            try
            {
                //フォームサイズ定義

                // 列スタイルを変更するe

                tempDGV.EnableHeadersVisualStyles = false;

                // 列ヘッダー表示位置指定
                tempDGV.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                // 列幅自動調整
                //tempDGV.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                // 列ヘッダーフォント指定
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("ＭＳ ゴシック", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 237;

                // 奇数行の色
                tempDGV.AlternatingRowsDefaultCellStyle.BackColor = SystemColors.Control;

                // 行ヘッダを表示しない
                tempDGV.RowHeadersVisible = false;

                // 選択モード
                tempDGV.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
                tempDGV.MultiSelect = false;

                // カラム定義
                tempDGV.Columns.Add(colBuCode, "物件ＣＤ");
                tempDGV.Columns.Add(colBuName, "物件名");
                tempDGV.Columns.Add(colGou, "号室");
                tempDGV.Columns.Add(colNewStayDate, "新規入居開始日");

                // 解約申し込み
                tempDGV.Columns.Add(col_KaiyakuContact_01, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_02, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_03, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_04, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_05, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_06, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_07, "解約申し込み");
                tempDGV.Columns.Add(col_KaiyakuContact_08, "解約申し込み");

                // 解約
                tempDGV.Columns.Add(col_Kaiyaku_01, "解約");
                tempDGV.Columns.Add(col_Kaiyaku_02, "解約");
                tempDGV.Columns.Add(col_Kaiyaku_03, "解約");
                tempDGV.Columns.Add(col_Kaiyaku_04, "解約");

                // ルームチェック
                tempDGV.Columns.Add(col_RoomCheck_01, "ルームチェック");
                tempDGV.Columns.Add(col_RoomCheck_02, "ルームチェック");
                tempDGV.Columns.Add(col_RoomCheck_03, "ルームチェック");
                tempDGV.Columns.Add(col_RoomCheck_04, "ルームチェック");
                tempDGV.Columns.Add(col_RoomCheck_05, "ルームチェック");

                // 書類作成
                tempDGV.Columns.Add(col_Shorui_01, "書類作成");
                tempDGV.Columns.Add(col_Shorui_02, "書類作成");
                tempDGV.Columns.Add(col_Shorui_03, "書類作成");
                tempDGV.Columns.Add(col_Shorui_04, "書類作成");

                // 手続き
                tempDGV.Columns.Add(col_Tetsu_01, "手続き");
                tempDGV.Columns.Add(col_Tetsu_02, "手続き");
                tempDGV.Columns.Add(col_Tetsu_03, "手続き");
                tempDGV.Columns.Add(col_Tetsu_04, "手続き");
                tempDGV.Columns.Add(col_Tetsu_05, "手続き");

                // 発注
                tempDGV.Columns.Add(col_Hacchu_01, "超過時間");
                tempDGV.Columns.Add(col_Hacchu_02, "超過時間");
                tempDGV.Columns.Add(col_Hacchu_03, "超過時間");

                // 工事着工
                tempDGV.Columns.Add(col_Kouji_01, "発注");
                tempDGV.Columns.Add(col_Kouji_02, "発注");
                tempDGV.Columns.Add(col_Kouji_03, "発注");
                tempDGV.Columns.Add(col_Kouji_04, "発注");
                tempDGV.Columns.Add(col_Kouji_05, "発注");
                tempDGV.Columns.Add(col_Kouji_06, "発注");
                tempDGV.Columns.Add(col_Kouji_07, "発注");

                // 完了検査
                tempDGV.Columns.Add(col_Kanryo_01, "完了検査");
                tempDGV.Columns.Add(col_Kanryo_02, "完了検査");
                tempDGV.Columns.Add(col_Kanryo_03, "完了検査");
                tempDGV.Columns.Add(col_Kanryo_04, "完了検査");

                // 備考
                tempDGV.Columns.Add(col_Bikou_01, "備考");

                // スカイワン
                tempDGV.Columns.Add(col_SkyOne_01, "スカイワン");
                tempDGV.Columns.Add(col_SkyOne_02, "スカイワン");

                // 行番号
                tempDGV.Columns.Add(col_xlsRowNum, "行番号");
                tempDGV.Columns[col_xlsRowNum].Visible = false;

                // 各列幅指定
                tempDGV.Columns[colBuCode].Width = 70;
                tempDGV.Columns[colBuName].Width = 200;
                tempDGV.Columns[colGou].Width = 60;
                tempDGV.Columns[colNewStayDate].Width = 120;

                tempDGV.Columns[col_KaiyakuContact_01].Width = 100;
                tempDGV.Columns[col_KaiyakuContact_02].Width = 126;
                tempDGV.Columns[col_KaiyakuContact_03].Width = 100;
                tempDGV.Columns[col_KaiyakuContact_04].Width = 100;
                tempDGV.Columns[col_KaiyakuContact_05].Width = 100;
                tempDGV.Columns[col_KaiyakuContact_06].Width = 120;
                tempDGV.Columns[col_KaiyakuContact_07].Width = 120;
                tempDGV.Columns[col_KaiyakuContact_08].Width = 130;

                tempDGV.Columns[col_Kaiyaku_01].Width = 100;
                tempDGV.Columns[col_Kaiyaku_02].Width = 100;
                tempDGV.Columns[col_Kaiyaku_03].Width = 100;
                tempDGV.Columns[col_Kaiyaku_04].Width = 100;

                tempDGV.Columns[col_RoomCheck_01].Width = 100;
                tempDGV.Columns[col_RoomCheck_02].Width = 210;
                tempDGV.Columns[col_RoomCheck_03].Width = 100;
                tempDGV.Columns[col_RoomCheck_04].Width = 100;
                tempDGV.Columns[col_RoomCheck_05].Width = 100;

                tempDGV.Columns[col_Shorui_01].Width = 120;
                tempDGV.Columns[col_Shorui_02].Width = 126;
                tempDGV.Columns[col_Shorui_03].Width = 120;
                tempDGV.Columns[col_Shorui_04].Width = 120;

                tempDGV.Columns[col_Tetsu_01].Width = 110;
                tempDGV.Columns[col_Tetsu_02].Width = 120;
                tempDGV.Columns[col_Tetsu_03].Width = 200;
                tempDGV.Columns[col_Tetsu_04].Width = 100;
                tempDGV.Columns[col_Tetsu_05].Width = 120;

                tempDGV.Columns[col_Hacchu_01].Width = 100;
                tempDGV.Columns[col_Hacchu_02].Width = 190;
                tempDGV.Columns[col_Hacchu_03].Width = 100;

                tempDGV.Columns[col_Kouji_01].Width = 200;
                tempDGV.Columns[col_Kouji_02].Width = 120;
                tempDGV.Columns[col_Kouji_03].Width = 100;
                tempDGV.Columns[col_Kouji_04].Width = 120;
                tempDGV.Columns[col_Kouji_05].Width = 100;
                tempDGV.Columns[col_Kouji_06].Width = 190;
                tempDGV.Columns[col_Kouji_07].Width = 200;

                tempDGV.Columns[col_Kanryo_01].Width = 100;
                tempDGV.Columns[col_Kanryo_02].Width = 100;
                tempDGV.Columns[col_Kanryo_03].Width = 100;
                tempDGV.Columns[col_Kanryo_04].Width = 110;

                tempDGV.Columns[col_Bikou_01].Width = 400;

                tempDGV.Columns[col_SkyOne_01].Width = 100;
                tempDGV.Columns[col_SkyOne_02].Width = 100;

                // 表示位置
                for (int i = 0; i < tempDGV.Columns.Count; i++)
                {
                    string cName = tempDGV.Columns[i].Name;
                    if (cName == colBuName || cName == col_Kouji_01 || cName == col_Bikou_01)
                    {
                        tempDGV.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleLeft;
                    }
                    else
                    {
                        tempDGV.Columns[i].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                    }
                }

                // 追加行表示しない
                tempDGV.AllowUserToAddRows = false;

                // データグリッドビューから行削除を禁止する
                tempDGV.AllowUserToDeleteRows = false;

                // 手動による列移動の禁止
                tempDGV.AllowUserToOrderColumns = false;

                // 列サイズ変更禁止
                tempDGV.AllowUserToResizeColumns = true;

                // 行サイズ変更禁止
                tempDGV.AllowUserToResizeRows = false;

                // 行ヘッダーの自動調節
                //tempDGV.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;

                // 編集可否
                tempDGV.ReadOnly = true;
            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message, "エラーメッセージ", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void gridViewShowData(DataGridView g)
        {
            using (var bk = new XLWorkbook(Properties.Settings.Default.imPortPath, XLEventTracking.Disabled))
            {
                var sheet1 = bk.Worksheet(Properties.Settings.Default.xlsSheetName);
                var tbl = sheet1.RangeUsed().AsTable();

                foreach (var t in tbl.DataRange.Rows())
                {
                    if (t.RowNumber() < 5)
                    {
                        continue;
                    }

                    if (t.RowNumber() == 5)
                    {
                        for (int i = 0; i < tbl.DataRange.ColumnCount(); i++)
                        {
                            g.Columns[i].HeaderText = Utility.NulltoStr(t.Cell(i + 1).Value).Replace("\n", "").Replace(" ", "").Replace("　", "");
                        }
                    }
                    else
                    {
                        g.Rows.Add();

                        for (int i = 0; i < tbl.DataRange.ColumnCount(); i++)
                        {
                            DateTime dt;

                            if (i == 8)
                            {
                                // 立会時間
                                if (DateTime.TryParse(Utility.NulltoStr(t.Cell(i + 1).Value), out dt))
                                {
                                    g[i, g.Rows.Count - 1].Value = dt.Hour + ":" + dt.Minute.ToString().PadLeft(2, '0');
                                }
                                else
                                {
                                    g[i, g.Rows.Count - 1].Value = Utility.NulltoStr(t.Cell(i + 1).Value);
                                }
                            }
                            else
                            {
                                // 日付形式か？
                                if (DateTime.TryParse(Utility.NulltoStr(t.Cell(i + 1).Value), out dt))
                                {
                                    // 日付情報
                                    g[i, g.Rows.Count - 1].Value = dt.ToShortDateString();
                                }
                                else
                                {
                                    // 文字列情報
                                    g[i, g.Rows.Count - 1].Value = Utility.NulltoStr(t.Cell(i + 1).Value);
                                }
                            }
                        }

                        g[col_xlsRowNum, g.Rows.Count - 1].Value = t.RowNumber();
                    }
                }
                sheet1.Dispose();
            }
            g.CurrentCell = null;
        } 

        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0)
            {
                return;
            }

            string cData = Utility.NulltoStr(dataGridView1[colBuCode, e.RowIndex].Value) + " : " +
                           Utility.NulltoStr(dataGridView1[colBuName, e.RowIndex].Value) + " " +
                           Utility.NulltoStr(dataGridView1[colGou, e.RowIndex].Value) + "号室";

            if (MessageBox.Show(cData + "が選択されました。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            txtBuCode.Text = Utility.NulltoStr(dataGridView1[colBuCode, e.RowIndex].Value);
            txtBuName.Text = Utility.NulltoStr(dataGridView1[colBuName, e.RowIndex].Value);
            txtGou.Text = Utility.NulltoStr(dataGridView1[colGou, e.RowIndex].Value);
            txtNewStayDate.Text = Utility.NulltoStr(dataGridView1[colNewStayDate, e.RowIndex].Value);

            txtKaiyakuContact01.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_01, e.RowIndex].Value);
            txtKaiyakuContact02.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_02, e.RowIndex].Value);
            txtKaiyakuContact03.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_03, e.RowIndex].Value);
            txtKaiyakuContact04.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_04, e.RowIndex].Value);
            txtKaiyakuContact05.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_05, e.RowIndex].Value);
            txtKaiyakuContact06.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_06, e.RowIndex].Value);
            txtKaiyakuContact07.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_07, e.RowIndex].Value);
            txtKaiyakuContact08.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_08, e.RowIndex].Value);
            
            txtKaiyaku01.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_01, e.RowIndex].Value);
            txtKaiyaku02.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_02, e.RowIndex].Value);
            txtKaiyaku03.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_03, e.RowIndex].Value);
            txtKaiyaku04.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_04, e.RowIndex].Value);

            txtRoomCheck01.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_01, e.RowIndex].Value);
            txtRoomCheck02.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_02, e.RowIndex].Value);
            txtRoomCheck03.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_03, e.RowIndex].Value);
            txtRoomCheck04.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_04, e.RowIndex].Value);
            txtRoomCheck05.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_05, e.RowIndex].Value);

            txtShorui01.Text = Utility.NulltoStr(dataGridView1[col_Shorui_01, e.RowIndex].Value);
            txtShorui02.Text = Utility.NulltoStr(dataGridView1[col_Shorui_02, e.RowIndex].Value);
            txtShorui03.Text = Utility.NulltoStr(dataGridView1[col_Shorui_03, e.RowIndex].Value);
            txtShorui04.Text = Utility.NulltoStr(dataGridView1[col_Shorui_04, e.RowIndex].Value);

            txtTetsu01.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_01, e.RowIndex].Value);
            txtTetsu02.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_02, e.RowIndex].Value);
            txtTetsu03.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_03, e.RowIndex].Value);
            txtTetsu04.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_04, e.RowIndex].Value);
            txtTetsu05.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_05, e.RowIndex].Value);

            txtHacchu01.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_01, e.RowIndex].Value);
            txtHacchu02.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_02, e.RowIndex].Value);
            txtHacchu03.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_03, e.RowIndex].Value);

            txtKouji01.Text = Utility.NulltoStr(dataGridView1[col_Kouji_01, e.RowIndex].Value);
            txtKouji02.Text = Utility.NulltoStr(dataGridView1[col_Kouji_02, e.RowIndex].Value);
            txtKouji03.Text = Utility.NulltoStr(dataGridView1[col_Kouji_03, e.RowIndex].Value);
            txtKouji04.Text = Utility.NulltoStr(dataGridView1[col_Kouji_04, e.RowIndex].Value);
            txtKouji05.Text = Utility.NulltoStr(dataGridView1[col_Kouji_05, e.RowIndex].Value);
            txtKouji06.Text = Utility.NulltoStr(dataGridView1[col_Kouji_06, e.RowIndex].Value);
            txtKouji07.Text = Utility.NulltoStr(dataGridView1[col_Kouji_07, e.RowIndex].Value);

            txtKanryo01.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_01, e.RowIndex].Value);
            txtKanryo02.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_02, e.RowIndex].Value);
            txtKanryo03.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_03, e.RowIndex].Value);
            txtKanryo04.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_04, e.RowIndex].Value);

            txtSkyOne01.Text = Utility.NulltoStr(dataGridView1[col_SkyOne_01, e.RowIndex].Value);
            txtSkyOne02.Text = Utility.NulltoStr(dataGridView1[col_SkyOne_02, e.RowIndex].Value);

            txtBikou.Text = Utility.NulltoStr(dataGridView1[col_Bikou_01, e.RowIndex].Value);
        }

        private void dispInitial()
        {
            txtBuCode.Text = string.Empty;
            txtBuName.Text = string.Empty;
            txtGou.Text = string.Empty;
            txtNewStayDate.Text = string.Empty;

            txtKaiyakuContact01.Text = string.Empty;
            txtKaiyakuContact02.Text = string.Empty;
            txtKaiyakuContact03.Text = string.Empty;
            txtKaiyakuContact04.Text = string.Empty;
            txtKaiyakuContact05.Text = string.Empty;
            txtKaiyakuContact06.Text = string.Empty;
            txtKaiyakuContact07.Text = string.Empty;
            txtKaiyakuContact08.Text = string.Empty;

            txtKaiyaku01.Text = string.Empty;
            txtKaiyaku02.Text = string.Empty;
            txtKaiyaku03.Text = string.Empty;
            txtKaiyaku04.Text = string.Empty;

            txtRoomCheck01.Text = string.Empty;
            txtRoomCheck02.Text = string.Empty;
            txtRoomCheck03.Text = string.Empty;
            txtRoomCheck04.Text = string.Empty;
            txtRoomCheck05.Text = string.Empty;

            txtShorui01.Text = string.Empty;
            txtShorui02.Text = string.Empty;
            txtShorui03.Text = string.Empty;
            txtShorui04.Text = string.Empty;

            txtTetsu01.Text = string.Empty;
            txtTetsu02.Text = string.Empty;
            txtTetsu03.Text = string.Empty;
            txtTetsu04.Text = string.Empty;
            txtTetsu05.Text = string.Empty;

            txtHacchu01.Text = string.Empty;
            txtHacchu02.Text = string.Empty;
            txtHacchu03.Text = string.Empty;

            txtKouji01.Text = string.Empty;
            txtKouji02.Text = string.Empty;
            txtKouji03.Text = string.Empty;
            txtKouji04.Text = string.Empty;
            txtKouji05.Text = string.Empty;
            txtKouji06.Text = string.Empty;
            txtKouji07.Text = string.Empty;

            txtKanryo01.Text = string.Empty;
            txtKanryo02.Text = string.Empty;
            txtKanryo03.Text = string.Empty;
            txtKanryo04.Text = string.Empty;

            txtSkyOne01.Text = string.Empty;
            txtSkyOne02.Text = string.Empty;

            txtBikou.Text = string.Empty;
        }
    }
}

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
        public Form1(string _xlsFName, string _xlsPass, int _xlsJyokenFormat)
        {
            InitializeComponent();

            xlsFname = _xlsFName;
            xlsPass = _xlsPass;
            xlsJyokenFormat = _xlsJyokenFormat;

            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);
        }

        //DataSet1 dts = new DataSet1();
        //DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();

        string tFile = string.Empty;
        string xlsFname = string.Empty;
        string xlsPass = string.Empty;
        int xlsJyokenFormat = 0;

        string upFlg = "1";     // 更新フラグ
        string addFlg = "2";    // 追加フラグ
        int uCnt = 0;           // 更新カウント

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

        // 更新
        string col_upFlg = "c49";

        #endregion

        private void button3_Click(object sender, EventArgs e)
        {
            //if (uCnt > 0)
            //{
            //    if (MessageBox.Show(uCnt + "件の更新を保存して終了しますか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            //    {
            //        string sPath = System.IO.Path.GetDirectoryName(xlsFname);

            //        // 他のPCで処理中の場合、続行不可
            //        //if (Utility.existsLockFile(sPath))
            //        //{
            //        //    MessageBox.Show("他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。再度実行してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //        //    return;
            //        //}

            //        // 他のPCで処理中の場合、続行不可
            //        while (Utility.existsLockFile(sPath))
            //        {
            //            Cursor = Cursors.WaitCursor;
            //            pictureBox1.Visible = true;
            //            lblMsg.Text = "他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。少々おまちください...";
            //            System.Threading.Thread.Sleep(100);
            //            Application.DoEvents();
            //        }

            //        Cursor = Cursors.Default;
            //        pictureBox1.Visible = false;
            //        lblMsg.Text = "";

            //        dataUpdate(dataGridView1, tFile, xlsPass, string.Empty);
            //    }
            //}

            // フォームを閉じる
            Close();
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (uCnt > 0)
            {
                if (MessageBox.Show(uCnt + "件の更新を保存して終了しますか？", "更新確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    string sPath = System.IO.Path.GetDirectoryName(xlsFname);

                    // 他のPCで処理中の場合、続行不可
                    //if (Utility.existsLockFile(sPath))
                    //{
                    //    MessageBox.Show("他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。再度実行してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    //    return;
                    //}

                    // 他のPCで処理中の場合、続行不可
                    while (Utility.existsLockFile(sPath))
                    {
                        Cursor = Cursors.WaitCursor;
                        pictureBox1.Visible = true;
                        lblMsg.Text = "他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。少々おまちください...";
                        System.Threading.Thread.Sleep(100);
                        Application.DoEvents();
                    }

                    Cursor = Cursors.Default;
                    pictureBox1.Visible = false;
                    lblMsg.Text = "";

                    dataUpdate(dataGridView1, tFile, xlsPass, string.Empty);
                }

                // 後片付け
                Dispose();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtBuCode.AutoSize = false;
            txtBuCode.Height = 20;

            txtBuName.AutoSize = false;
            txtBuName.Height = 20;

            txtGou.AutoSize = false;
            txtGou.Height = 20;

            txtNewStayDate.AutoSize = false;
            txtNewStayDate.Height = 20;

            txtKaiyakuContact01.AutoSize = false;
            txtKaiyakuContact01.Height = 20;

            txtKaiyakuContact02.AutoSize = false;
            txtKaiyakuContact02.Height = 20;

            txtKaiyakuContact03.AutoSize = false;
            txtKaiyakuContact03.Height = 20;

            txtKaiyakuContact04.AutoSize = false;
            txtKaiyakuContact04.Height = 20;

            txtKaiyakuContact05.AutoSize = false;
            txtKaiyakuContact05.Height = 20;

            txtKaiyakuContact06.AutoSize = false;
            txtKaiyakuContact06.Height = 20;

            txtKaiyakuContact07.AutoSize = false;
            txtKaiyakuContact07.Height = 20;

            //// オーナードローを指定
            //cmbKaiyakuContact08.DrawMode = DrawMode.OwnerDrawFixed;

            //// 項目の高さを設定
            //cmbKaiyakuContact08.ItemHeight = 18;

            //// DrawItemイベントハンドラの追加
            //cmbKaiyakuContact08.DrawItem += new DrawItemEventHandler(ComboBox1_DrawItem);
            ////ComboBox1のリストに項目を追加する
            //cmbKaiyakuContact08.Items.Add("おはよう。");
            //cmbKaiyakuContact08.Items.Add("こんにちは。");
            //cmbKaiyakuContact08.Items.Add("こんばんは。");

            txtKaiyaku01.AutoSize = false;
            txtKaiyaku01.Height = 20;

            txtKaiyaku02.AutoSize = false;
            txtKaiyaku02.Height = 20;

            txtKaiyaku03.AutoSize = false;
            txtKaiyaku03.Height = 20;

            txtKaiyaku04.AutoSize = false;
            txtKaiyaku04.Height = 20;

            txtRoomCheck01.AutoSize = false;
            txtRoomCheck01.Height = 20;

            txtRoomCheck02.AutoSize = false;
            txtRoomCheck02.Height = 20;

            txtRoomCheck03.AutoSize = false;
            txtRoomCheck03.Height = 20;

            cmbKeyOkiba.AutoSize = false;
            cmbKeyOkiba.Height = 20;

            cmbRoomCheck05.AutoSize = false;
            cmbRoomCheck05.Height = 20;

            txtShorui01.AutoSize = false;
            txtShorui01.Height = 20;

            txtShorui02.AutoSize = false;
            txtShorui02.Height = 20;

            txtShorui03.AutoSize = false;
            txtShorui03.Height = 20;

            cmbShorui04.AutoSize = false;
            cmbShorui04.Height = 20;

            txtTetsu01.AutoSize = false;
            txtTetsu01.Height = 20;

            txtTetsu02.AutoSize = false;
            txtTetsu02.Height = 20;

            txtTetsu03.AutoSize = false;
            txtTetsu03.Height = 20;

            txtTetsu04.AutoSize = false;
            txtTetsu04.Height = 20;

            txtTetsu05.AutoSize = false;
            txtTetsu05.Height = 20;

            txtHacchu01.AutoSize = false;
            txtHacchu01.Height = 20;

            txtHacchu02.AutoSize = false;
            txtHacchu02.Height = 20;

            txtKouji02.AutoSize = false;
            txtKouji02.Height = 20;

            txtKouji03.AutoSize = false;
            txtKouji03.Height = 20;

            txtKouji04.AutoSize = false;
            txtKouji04.Height = 20;

            txtKouji05.AutoSize = false;
            txtKouji05.Height = 20;

            txtKouji06.AutoSize = false;
            txtKouji06.Height = 20;

            txtKouji07.AutoSize = false;
            txtKouji07.Height = 20;

            txtKanryo01.AutoSize = false;
            txtKanryo01.Height = 20;

            txtKanryo02.AutoSize = false;
            txtKanryo02.Height = 20;

            txtKanryo03.AutoSize = false;
            txtKanryo03.Height = 20;

            txtSkyOne01.AutoSize = false;
            txtSkyOne01.Height = 20;

            txtBikou.AutoSize = false;
            txtBikou.Height = 20;

            // データグリッドビュー定義
            GridViewSetting(dataGridView1);

            panel1.Enabled = false;
            button1.Enabled = false;
            button2.Enabled = false;

            txtsBuName.Enabled = false;
            txtsBuName.Text = string.Empty;
            button5.Enabled = false;

            pictureBox1.Visible = false;
        }

        private void Form1_Shown(object sender, EventArgs e)
        {
            dataGridView1.CurrentCell = null;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     Excelファイルをパスワード付きでオープン・クローズする </summary>
        /// <param name="sPath">
        ///     Excelファイルパス</param>
        /// <param name="rPw">
        ///     読み込みパスワード</param>
        /// <param name="wPw">
        ///     書き込みパスワード</param>
        /// <returns>
        ///     成功：true, 失敗：false</returns>
        ///-------------------------------------------------------------------------
        private bool impXlsSheet(string sPath, string rPw, string wPw)
        {
            lblMsg.Text = "Excelを起動しています...";
            System.Threading.Thread.Sleep(100);
            Application.DoEvents();

            // エクセルオブジェクト
            Excel.Application oXls = new Excel.Application();
            Excel.Workbook oXlsBook = null;
            
            try
            {
                if (rPw != string.Empty)
                {
                    lblMsg.Text = sPath + " のパスワードを解除しています...";
                }
                else
                {
                    lblMsg.Text = sPath + " を開いています...";
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Excelファイルを開く（ファイルパスワード付き）
                oXlsBook = (Excel.Workbook)(oXls.Workbooks.Open(sPath, Type.Missing, Type.Missing, Type.Missing,
                    rPw, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing, Type.Missing, Type.Missing));

                if (rPw != string.Empty)
                {
                    lblMsg.Text = sPath + " のパスワードが解除されました...";
                }
                else
                {
                    lblMsg.Text = sPath + " を開きました...";
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                oXls.DisplayAlerts = false;

                if (wPw != string.Empty)
                {
                    lblMsg.Text = sPath + " をパスワード付きで保存しています...";
                }
                else
                {
                    lblMsg.Text = sPath + " を保存しています...";
                }

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Excelファイル書き込み（ファイルパスワード解除）
                oXlsBook.SaveAs(sPath, Type.Missing, wPw, Type.Missing, Type.Missing, Type.Missing,
                                Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                                Type.Missing, Type.Missing);

                lblMsg.Text = sPath + " を保存しました...";

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                // Bookをクローズ
                oXlsBook.Close(Type.Missing, Type.Missing, Type.Missing);

                lblMsg.Text = "Excelを終了しました...";

                System.Threading.Thread.Sleep(100);
                Application.DoEvents();

                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return false;
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
                tempDGV.ColumnHeadersDefaultCellStyle.Font = new Font("MS UI Gothic", 8, FontStyle.Regular);

                // データフォント指定
                tempDGV.DefaultCellStyle.Font = new Font("MS UI Gothic", 9, FontStyle.Regular);

                // 行の高さ
                tempDGV.ColumnHeadersHeight = 20;
                tempDGV.RowTemplate.Height = 20;

                // 全体の高さ
                tempDGV.Height = 236;

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
                tempDGV.Columns.Add(col_xlsRowNum, "行");
                //tempDGV.Columns[col_xlsRowNum].Visible = false;

                // 更新フラグ
                tempDGV.Columns.Add(col_upFlg, "update");
                tempDGV.Columns[col_upFlg].Visible = false;

                // 各列幅指定
                tempDGV.Columns[colBuCode].Width = 70;
                tempDGV.Columns[colBuName].Width = 200;
                tempDGV.Columns[colGou].Width = 60;
                tempDGV.Columns[colNewStayDate].Width = 120;

                tempDGV.Columns[colNewStayDate].Frozen = true;

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

                tempDGV.Columns[col_xlsRowNum].Width = 50;

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

        ///---------------------------------------------------------------------------------
        /// <summary>
        ///     エクセルシートの内容をデータグリッドビューに表示する </summary>
        /// <param name="g">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sFile">
        ///     エクセルファイルパス</param>
        /// <param name="rPass">
        ///     エクセルファイル読み込みパスワード</param>
        /// <param name="wPass">
        ///     エクセルファイル書き込みパスワード</param>
        ///---------------------------------------------------------------------------------
        private void gridViewShowData(DataGridView g, string sFile, string rPass, string wPass)
        {
            string msg = "";

            Cursor = Cursors.WaitCursor;

            try
            {
                string sPath = System.IO.Path.GetDirectoryName(xlsFname);

                //LOCKファイル作成
                Utility.makeLockFile(sPath, System.Net.Dns.GetHostName());

                // 対象エクセルファイルのパスワードを解除する
                if (impXlsSheet(sFile, rPass, wPass))
                {
                    lblMsg.Text = "Excelブックを取得しています...";
                    System.Threading.Thread.Sleep(100);
                    Application.DoEvents();

                    using (var bk = new XLWorkbook(sFile, XLEventTracking.Disabled))
                    {
                        // 対象エクセルファイルのパスワード付きで書き込み
                        if (impXlsSheet(sFile, wPass, rPass))
                        {
                            // ロックファイルを削除する
                            Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());

                            msg = "解約管理表を読み込んでいます...";

                            lblMsg.Text = msg;
                            //System.Threading.Thread.Sleep(100);
                            //Application.DoEvents();

                            var sheet1 = bk.Worksheet(Properties.Settings.Default.xlsSheetName);
                            var tbl = sheet1.RangeUsed().AsTable();

                            g.Rows.Clear();

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
                                        g.Columns[i].HeaderText = Utility.NulltoStr(t.Cell(i + 1).Value).Replace("\n", "").Replace("\r", "").Replace(" ", "").Replace("　", "");
                                    }
                                }
                                else
                                {
                                    lblMsg.Text = msg + Utility.NulltoStr(t.Cell(1).Value) + ":" + Utility.NulltoStr(t.Cell(2).Value);
                                    System.Threading.Thread.Sleep(10);
                                    Application.DoEvents();

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

                            lblMsg.Text = "解約管理表の読み込みが終了しました...";
                            System.Threading.Thread.Sleep(30);
                            Application.DoEvents();
                        }
                        else
                        {
                            // Excelファイルのパスワード付きで書き込みに失敗
                            // ：ロックファイルを削除する
                            Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());
                        }
                    }
                }
                else
                {
                    // Excelファイルのパスワードを解除してオープンに失敗
                    // ：ロックファイルを削除する
                    Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                g.CurrentCell = null;
                Cursor = Cursors.Default;
            }
        }

        ///----------------------------------------------------------------------
        /// <summary>
        ///     指定のグリッドビュー行のデータを編集画面に表示する </summary>
        /// <param name="sender">
        ///     </param>
        /// <param name="e">
        ///     </param>
        ///----------------------------------------------------------------------
        private void dataGridView1_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (radioButton1.Checked)
                {
                    // 新規登録モードでは不可
                    return;
                }

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

                // 指定のグリッドビュー行のデータを編集画面に表示する
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
                //txtKaiyakuContact08.Text = Utility.NulltoStr(dataGridView1[col_KaiyakuContact_08, e.RowIndex].Value);

                txtKaiyaku01.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_01, e.RowIndex].Value);
                txtKaiyaku02.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_02, e.RowIndex].Value);
                txtKaiyaku03.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_03, e.RowIndex].Value);
                txtKaiyaku04.Text = Utility.NulltoStr(dataGridView1[col_Kaiyaku_04, e.RowIndex].Value);

                txtRoomCheck01.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_01, e.RowIndex].Value);
                //txtRoomCheck02.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_02, e.RowIndex].Value);
                txtRoomCheck03.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_03, e.RowIndex].Value);
                //txtRoomCheck04.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_04, e.RowIndex].Value);
                //txtRoomCheck05.Text = Utility.NulltoStr(dataGridView1[col_RoomCheck_05, e.RowIndex].Value);

                txtShorui01.Text = Utility.NulltoStr(dataGridView1[col_Shorui_01, e.RowIndex].Value);
                txtShorui02.Text = Utility.NulltoStr(dataGridView1[col_Shorui_02, e.RowIndex].Value);
                txtShorui03.Text = Utility.NulltoStr(dataGridView1[col_Shorui_03, e.RowIndex].Value);
                //txtShorui04.Text = Utility.NulltoStr(dataGridView1[col_Shorui_04, e.RowIndex].Value);

                txtTetsu01.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_01, e.RowIndex].Value);
                txtTetsu02.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_02, e.RowIndex].Value);
                //txtTetsu03.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_03, e.RowIndex].Value);
                txtTetsu04.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_04, e.RowIndex].Value);
                txtTetsu05.Text = Utility.NulltoStr(dataGridView1[col_Tetsu_05, e.RowIndex].Value);

                txtHacchu01.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_01, e.RowIndex].Value);
                //txtHacchu02.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_02, e.RowIndex].Value);
                //txtHacchu03.Text = Utility.NulltoStr(dataGridView1[col_Hacchu_03, e.RowIndex].Value);

                //txtKouji01.Text = Utility.NulltoStr(dataGridView1[col_Kouji_01, e.RowIndex].Value);
                txtKouji02.Text = Utility.NulltoStr(dataGridView1[col_Kouji_02, e.RowIndex].Value);
                txtKouji03.Text = Utility.NulltoStr(dataGridView1[col_Kouji_03, e.RowIndex].Value);
                txtKouji04.Text = Utility.NulltoStr(dataGridView1[col_Kouji_04, e.RowIndex].Value);
                txtKouji05.Text = Utility.NulltoStr(dataGridView1[col_Kouji_05, e.RowIndex].Value);
                //txtKouji06.Text = Utility.NulltoStr(dataGridView1[col_Kouji_06, e.RowIndex].Value);
                //txtKouji07.Text = Utility.NulltoStr(dataGridView1[col_Kouji_07, e.RowIndex].Value);

                txtKanryo01.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_01, e.RowIndex].Value);
                txtKanryo02.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_02, e.RowIndex].Value);
                txtKanryo03.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_03, e.RowIndex].Value);
                //txtKanryo04.Text = Utility.NulltoStr(dataGridView1[col_Kanryo_04, e.RowIndex].Value);

                txtSkyOne01.Text = Utility.NulltoStr(dataGridView1[col_SkyOne_01, e.RowIndex].Value);
                //txtSkyOne02.Text = Utility.NulltoStr(dataGridView1[col_SkyOne_02, e.RowIndex].Value);

                txtBikou.Text = Utility.NulltoStr(dataGridView1[col_Bikou_01, e.RowIndex].Value);

                // 行番号
                lblRow.Text = Utility.NulltoStr(dataGridView1[col_xlsRowNum, e.RowIndex].Value);

                // 以下、コンボボックス
                selCmbItem(cmbGyousha, col_Kouji_01, e.RowIndex);       // 業者名コンボボックス
                selCmbItem(cmbHosho, col_SkyOne_02, e.RowIndex);        // 保証会社       
                selCmbItem(cmbKeyOkiba, col_RoomCheck_04, e.RowIndex);  // 鍵置場
                selCmbItem(cmbKaiyakuContact08, col_KaiyakuContact_08, e.RowIndex); // 解約申し込み担当
                selCmbItem(cmbRoomCheck05, col_RoomCheck_05, e.RowIndex);   // ルームチェック担当
                selCmbItem(cmbShorui04, col_Shorui_04, e.RowIndex);     // 書類担当
                selCmbItem(cmbHacchu03, col_Hacchu_03, e.RowIndex);     // 発注担当
                selCmbItem(cmbKanryo04, col_Kanryo_04, e.RowIndex);     // 完了担当

                button2.Enabled = true;     // 更新ボタン
                button1.Enabled = true;     // 取消ボタン
                panel1.Enabled = false;

                if (radioButton2.Checked)
                {
                    txtBuCode.Enabled = false;
                }
                else
                {
                    txtBuCode.Enabled = true;
                }
        }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        
        ///-----------------------------------------------------
        /// <summary>
        ///     コンボボックスアイテム表示 </summary>
        /// <param name="cb">
        ///     コンボボックスオブジェクト</param>
        /// <param name="dgCol">
        ///     データグリッドビューカラム</param>
        /// <param name="r">
        ///     データグリッドビューrowindex</param>
        ///-----------------------------------------------------
        private void selCmbItem(ComboBox cb, string dgCol, int r)
        {
            string gName = Utility.NulltoStr(dataGridView1[dgCol, r].Value);

            if (gName == string.Empty)
            {
                cb.SelectedIndex = -1;
                cb.Text = string.Empty;
                return;
            }

            for (int i = 0; i < cb.Items.Count; i++)
            {
                if (cb.Items[i].ToString() == gName)
                {
                    cb.SelectedIndex = i;
                    break;
                }
            }
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
            //txtKaiyakuContact08.Text = string.Empty;

            txtKaiyaku01.Text = string.Empty;
            txtKaiyaku02.Text = string.Empty;
            txtKaiyaku03.Text = string.Empty;
            txtKaiyaku04.Text = string.Empty;

            txtRoomCheck01.Text = string.Empty;
            txtRoomCheck02.Text = string.Empty;
            txtRoomCheck03.Text = string.Empty;
            //txtRoomCheck04.Text = string.Empty;
            //txtRoomCheck05.Text = string.Empty;

            txtShorui01.Text = string.Empty;
            txtShorui02.Text = string.Empty;
            txtShorui03.Text = string.Empty;
            //txtShorui04.Text = string.Empty;

            txtTetsu01.Text = string.Empty;
            txtTetsu02.Text = string.Empty;
            txtTetsu03.Text = string.Empty;
            txtTetsu04.Text = string.Empty;
            txtTetsu05.Text = string.Empty;

            txtHacchu01.Text = string.Empty;
            txtHacchu02.Text = string.Empty;
            //txtHacchu03.Text = string.Empty;

            //txtKouji01.Text = string.Empty;
            txtKouji02.Text = string.Empty;
            txtKouji03.Text = string.Empty;
            txtKouji04.Text = string.Empty;
            txtKouji05.Text = string.Empty;
            txtKouji06.Text = string.Empty;
            txtKouji07.Text = string.Empty;

            txtKanryo01.Text = string.Empty;
            txtKanryo02.Text = string.Empty;
            txtKanryo03.Text = string.Empty;
            //txtKanryo04.Text = string.Empty;

            txtSkyOne01.Text = string.Empty;
            //txtSkyOne02.Text = string.Empty;

            txtBikou.Text = string.Empty;

            radioButton1.Checked = false;
            radioButton2.Checked = true;

            lblRow.Text = string.Empty;

            setCmbItems(dataGridView1, cmbGyousha, col_Kouji_01);       // 工事業者
            setCmbItems(dataGridView1, cmbHosho, col_SkyOne_02);        // 保証会社
            setCmbItems(dataGridView1, cmbKeyOkiba, col_RoomCheck_04);  // 鍵置場
            setCmbItems(dataGridView1, cmbKaiyakuContact08, col_KaiyakuContact_08); // 解約申し込み担当
            setCmbItems(dataGridView1, cmbRoomCheck05, col_RoomCheck_05);   // ルームチェック担当
            setCmbItems(dataGridView1, cmbShorui04, col_Shorui_04);         // 書類担当
            setCmbItems(dataGridView1, cmbHacchu03, col_Hacchu_03);         // 発注担当
            setCmbItems(dataGridView1, cmbKanryo04, col_Kanryo_04);         // 完了担当

            button2.Enabled = false;    // 更新ボタン
            button1.Enabled = false;    // 取消ボタン
            panel1.Enabled = true;

            txtsBuName.Enabled = true;
            txtsBuName.Text = string.Empty;

            //button5.Enabled = true;

            dataGridView1.CurrentCell = null;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            string msg = string.Empty;

            if (radioButton1.Checked)
            {
                msg = "新規登録";
            }
            else
            {
                msg = "更新";
            }

            if (MessageBox.Show("表示中のデータを" + msg + "します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
            {
                return;
            }

            if (radioButton1.Checked)
            {
                if (Utility.NulltoStr(txtBuCode.Text) == string.Empty)
                {
                    MessageBox.Show("物件ＣＤが未入力です", "入力確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                    txtBuCode.Focus();
                    return;
                }

                if (!chkBuCode(dataGridView1, Utility.NulltoStr(txtBuCode.Text)))
                {
                    if (MessageBox.Show("既に登録済みの物件ＣＤです。続行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.No)
                    {
                        txtBuCode.Focus();
                        return;
                    }
                }

                // 新規登録
                addDataGrid(dataGridView1);

                dataGridView1.CurrentCell = dataGridView1[1, dataGridView1.RowCount - 1];
                dataGridView1.CurrentCell = null;

                // 画面初期化
                dispInitial();
            }
            else
            {
                // dataGridView 更新
                updateFlg(dataGridView1, Utility.StrtoInt(lblRow.Text));

                // 画面初期化
                dispInitial();
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     物件ＣＤ登録済みチェック </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <returns>
        ///     true:登録なし、false:登録有り</returns>
        ///----------------------------------------------------------
        private bool chkBuCode(DataGridView dg, string bCode)
        {
            bool rtn = true;

            for (int i = 0; i < dg.RowCount; i++)
            {
                if (Utility.NulltoStr(dg[colBuCode, i].Value) == bCode)
                {
                    rtn = false;
                    break;
                }
            }

            return rtn;
        }

        private void addDataGrid(DataGridView dg)
        {
            // データグリッドビューに追加登録
            dg.Rows.Add();
            dataGridUpdate(dg, dg.RowCount - 1, addFlg);

            uCnt++; // 更新数カウント
        }

        private void updateFlg(DataGridView dg, int rNum)
        {
            for (int i = 0; i < dg.RowCount; i++)
            {
                if (Utility.StrtoInt(Utility.NulltoStr(dg[col_xlsRowNum, i].Value)) == rNum)
                {
                    // データグリッドビュー更新
                    dataGridUpdate(dataGridView1, i, upFlg);

                    uCnt++; // 更新数カウント

                    break;
                }
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     データグリッドビューのデータを更新する </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     rowIndex</param>
        /// <param name="uFlg">
        ///     更新フラグ</param>
        ///--------------------------------------------------------------
        private void dataGridUpdate(DataGridView dg, int i, string uFlg)
        {
            dg[colBuCode, i].Value = txtBuCode.Text;
            dg[colBuName, i].Value = txtBuName.Text;
            dg[colGou, i].Value = txtGou.Text;

            dg[colNewStayDate, i].Value = txtNewStayDate.Text;

            dg[col_KaiyakuContact_01, i].Value = txtKaiyakuContact01.Text;
            dg[col_KaiyakuContact_02, i].Value = txtKaiyakuContact02.Text;
            dg[col_KaiyakuContact_03, i].Value = txtKaiyakuContact03.Text;
            dg[col_KaiyakuContact_04, i].Value = txtKaiyakuContact04.Text;
            dg[col_KaiyakuContact_05, i].Value = txtKaiyakuContact05.Text;
            dg[col_KaiyakuContact_06, i].Value = txtKaiyakuContact06.Text;
            dg[col_KaiyakuContact_07, i].Value = txtKaiyakuContact07.Text;
            //dg[col_KaiyakuContact_08, i].Value = txtKaiyakuContact08.Text;
            dg[col_KaiyakuContact_08, i].Value = Utility.NulltoStr(cmbKaiyakuContact08.Text);

            dg[col_Kaiyaku_01, i].Value = txtKaiyaku01.Text;
            dg[col_Kaiyaku_02, i].Value = txtKaiyaku02.Text;
            dg[col_Kaiyaku_03, i].Value = txtKaiyaku03.Text;
            dg[col_Kaiyaku_04, i].Value = txtKaiyaku04.Text;

            dg[col_RoomCheck_01, i].Value = txtRoomCheck01.Text;
            dg[col_RoomCheck_02, i].Value = txtRoomCheck02.Text;
            dg[col_RoomCheck_03, i].Value = txtRoomCheck03.Text;
            //dg[col_RoomCheck_04, i].Value = txtRoomCheck04.Text;
            dg[col_RoomCheck_04, i].Value = Utility.NulltoStr(cmbKeyOkiba.Text);
            //dg[col_RoomCheck_05, i].Value = txtRoomCheck05.Text;
            dg[col_RoomCheck_05, i].Value = Utility.NulltoStr(cmbRoomCheck05.Text);

            dg[col_Shorui_01, i].Value = txtShorui01.Text;
            dg[col_Shorui_02, i].Value = txtShorui02.Text;
            dg[col_Shorui_03, i].Value = txtShorui03.Text;
            //dg[col_Shorui_04, i].Value = txtShorui04.Text;
            dg[col_Shorui_04, i].Value = Utility.NulltoStr(cmbShorui04.Text);

            dg[col_Tetsu_01, i].Value = txtTetsu01.Text;
            dg[col_Tetsu_02, i].Value = txtTetsu02.Text;
            dg[col_Tetsu_03, i].Value = txtTetsu03.Text;
            dg[col_Tetsu_04, i].Value = txtTetsu04.Text;
            dg[col_Tetsu_05, i].Value = txtTetsu05.Text;

            dg[col_Hacchu_01, i].Value = txtHacchu01.Text;
            dg[col_Hacchu_02, i].Value = txtHacchu02.Text;
            //dg[col_Hacchu_03, i].Value = txtHacchu03.Text;
            dg[col_Hacchu_03, i].Value = Utility.NulltoStr(cmbHacchu03.Text);

            //dg[col_Kouji_01, i].Value = txtKouji01.Text;
            dg[col_Kouji_01, i].Value = Utility.NulltoStr(cmbGyousha.Text);
            dg[col_Kouji_02, i].Value = txtKouji02.Text;
            dg[col_Kouji_03, i].Value = txtKouji03.Text;
            dg[col_Kouji_04, i].Value = txtKouji04.Text;
            dg[col_Kouji_05, i].Value = txtKouji05.Text;
            dg[col_Kouji_06, i].Value = txtKouji06.Text;
            dg[col_Kouji_07, i].Value = txtKouji07.Text;

            dg[col_Kanryo_01, i].Value = txtKanryo01.Text;
            dg[col_Kanryo_02, i].Value = txtKanryo02.Text;
            dg[col_Kanryo_03, i].Value = txtKanryo03.Text;
            //dg[col_Kanryo_04, i].Value = txtKanryo04.Text;
            dg[col_Kanryo_04, i].Value = Utility.NulltoStr(cmbKanryo04.Text);

            dg[col_Bikou_01, i].Value = txtBikou.Text;

            dg[col_SkyOne_01, i].Value = txtSkyOne01.Text;
            //dg[col_SkyOne_02, i].Value = txtSkyOne02.Text;
            dg[col_SkyOne_02, i].Value = Utility.NulltoStr(cmbHosho.Text);

            dg[col_upFlg, i].Value = uFlg;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     エクセルシート更新 </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sFile">
        ///     対象エクセルファイル</param>
        /// <param name="rPass">
        ///     読み込みパスワード</param>
        /// <param name="wPass">
        ///     書き込みパスワード</param>
        ///---------------------------------------------------------------------
        private void dataUpdate(DataGridView dg, string sFile, string rPass, string wPass)
        {
            bool uStatus = false;

            int uCnt = 0;
            Cursor = Cursors.WaitCursor;

            string sPath = System.IO.Path.GetDirectoryName(sFile);

            //LOCKファイル作成
            Utility.makeLockFile(sPath, System.Net.Dns.GetHostName());

            // 対象エクセルファイルのパスワードを解除する
            impXlsSheet(sFile, rPass, wPass);

            // ブック取得
            using (var bk = new XLWorkbook(sFile, XLEventTracking.Disabled))
            {
                var sheet1 = bk.Worksheet(Properties.Settings.Default.xlsSheetName);

                for (int i = 0; i < dg.RowCount; i++)
                {
                    // Excel行更新処理
                    if (Utility.NulltoStr(dg[col_upFlg, i].Value) == upFlg)
                    {
                        int bCode = Utility.StrtoInt(Utility.NulltoStr(dg[colBuCode, i].Value));
                        int rowNum = Utility.StrtoInt(Utility.NulltoStr(dg[col_xlsRowNum, i].Value));

                        // 対象行を取得
                        var row = sheet1.Row(rowNum);

                        if (Utility.StrtoInt(Utility.NulltoStr(row.Cell(1).Value)) == bCode)
                        {
                            // 物件ＣＤが一致したら更新
                            setXlsRowData(row, dg, i);
                            uCnt++;
                            uStatus = true;
                        }
                    }
                    else if (Utility.NulltoStr(dg[col_upFlg, i].Value) == addFlg)
                    {
                        // Excel行追加処理
                        var tbl = sheet1.RangeUsed().AsTable();
                        var row = sheet1.Row(tbl.RowCount());

                        // 現最下行の新規入居開始日セルの下罫線を変更
                        row.Cell(4).Style.Border.SetBottomBorder(XLBorderStyleValues.None);
                        row.Cell(4).Style.Border.SetBottomBorder(XLBorderStyleValues.Thin);

                        // 追加行
                        row = sheet1.Row(tbl.RowCount() + 1);
                        row.Height = 50.25;

                        // 追加行罫線を引く
                        sheet1.Range(row.Cell(1), row.Cell(tbl.ColumnCount())).Style
                            .Border.SetTopBorder(XLBorderStyleValues.Thin)
                            .Border.SetBottomBorder(XLBorderStyleValues.Thin)
                            .Border.SetLeftBorder(XLBorderStyleValues.Thin)
                            .Border.SetRightBorder(XLBorderStyleValues.Thin);

                        // 追加行書式設定
                        xlsNewRowStyleSet(sheet1, row, tbl.ColumnCount());

                        // 物件ＣＤ
                        row.Cell(1).Value = Utility.NulltoStr(dg[colBuCode, i].Value);

                        // 物件ＣＤ以外の項目セット
                        setXlsRowData(row, dg, i);

                        uCnt++;
                        uStatus = true;
                    }
                }

                // 更新
                if (uStatus)
                {
                    bk.SaveAs(sFile);
                }
            }

            // 対象エクセルファイルのパスワード付きで書き込み
            impXlsSheet(sFile, wPass, rPass);

            // Lockファイル削除
            Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());

            Cursor = Cursors.Default;

            if (uStatus)
            {
                MessageBox.Show(uCnt + "件、更新しました", "確認", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     新規登録行数式登録 </summary>
        /// <param name="sheet">
        ///     シートオブジェクト</param>
        /// <param name="row">
        ///     Rowオブジェクト</param>
        /// <param name="cLen">
        ///     </param>
        ///--------------------------------------------------------------
        private void xlsNewRowStyleSet(IXLWorksheet sheet, IXLRow row, int cLen)
        {
            for (int i = 1; i <= cLen; i++)
            {
                row.Cell(i).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center)
                    .Alignment.SetVertical(XLAlignmentVerticalValues.Center)
                    .Font.SetFontName("游ゴシック")
                    .Font.SetFontSize(12);

                // 物件名、業者名
                if (i == 2 || i == 34)
                {
                    row.Cell(i).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Left);
                }

                // 表示形式：月日
                if (i == 4 || i == 5 || i == 6 || i == 7 || i == 8 || i == 10 || i == 11 ||
                    i == 13 || i == 14 || i == 16 || i == 17 || i == 19 || i == 22 || 
                    i == 23 || i == 24 || i == 26 || i == 27 || i == 29 || i == 30 || i == 31 || 
                    i == 36 || i == 37 || i == 38 || i == 41 || i == 42 || i == 43)
                {
                    row.Cell(i).Style.NumberFormat.SetFormat("m/d");
                }

                // 立会時間　表示形式：時分
                if (i == 9)
                {
                    row.Cell(i).Style.NumberFormat.SetFormat("HH:mm");
                }

                // 工事代発注金額、スカイワン・金額
                if (i == 35 || i == 46)
                {
                    row.Cell(i).Style.NumberFormat.SetFormat("¥#,##0");
                }

                // 担当・背景色と右罫線
                if (i == 12 || i == 21 || i ==25 || i == 33 || i == 44)
                {
                    row.Cell(i).Style.Fill.BackgroundColor = XLColor.LightGray;
                    row.Cell(i).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                }

                // 新規入居開始日
                if (i == 4)
                {
                    // 赤の罫線
                    row.Cell(i).Style.Border.BottomBorder = XLBorderStyleValues.Thick;
                    row.Cell(i).Style.Border.LeftBorder = XLBorderStyleValues.Thick;
                    row.Cell(i).Style.Border.RightBorder = XLBorderStyleValues.Thick;
                    row.Cell(i).Style.Border.BottomBorderColor = XLColor.Red;
                    row.Cell(i).Style.Border.LeftBorderColor = XLColor.Red;
                    row.Cell(i).Style.Border.RightBorderColor = XLColor.Red;

                    // 太字
                    row.Cell(i).Style.Font.SetBold(true);

                    // 折り返して全体を表示
                    row.Cell(i).Style.Alignment.WrapText = true;
                }

                if (i == 45 || i == 47)
                {
                    row.Cell(i).Style.Border.RightBorder = XLBorderStyleValues.Medium;
                }

                // RC依頼からルームチェックまで
                if (i == 18)
                {
                    // 日数計算：数式
                    string formula = "=IF(AND(ISNUMBER(Q" + row.RowNumber() + "), ISNUMBER(P" + row.RowNumber() + ")), Q" + row.RowNumber() + " - P" + row.RowNumber() + ", " + @"""" + @"""" + ")";
                    row.Cell(i).FormulaA1 = formula;

                    if (xlsJyokenFormat == 1)
                    {
                        // 条件付き書式を追加
                        var cell1 = sheet.Cell("R6");
                        var cell2 = row.Cell(i);
                        IXLRange xLRange = sheet.Range(cell1, cell2);

                        // 空白セル
                        xLRange.AddConditionalFormat()
                            .WhenIsBlank()
                            .Fill.SetBackgroundColor(XLColor.White);

                        // ０～２日　空色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(0, 2)
                            .Fill.SetBackgroundColor(XLColor.SkyBlue);

                        // ３～５日　黄色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(3, 5)
                            .Fill.SetBackgroundColor(XLColor.Yellow);

                        // ６日～　赤色
                        xLRange.AddConditionalFormat()
                            .WhenEqualOrGreaterThan(6)
                            .Fill.SetBackgroundColor(XLColor.Red);
                    }
                }

                // 見積書送付から本日まで
                if (i == 28)
                {
                    // 日数計算：数式
                    string formula = "=IF(ISNUMBER(AA" + row.RowNumber() + "), TODAY()- AA" + row.RowNumber() + ", " + @"""" + @"""" + ")";
                    row.Cell(i).FormulaA1 = formula;

                    if (xlsJyokenFormat == 1)
                    {
                        // 条件付き書式を追加
                        var cell1 = sheet.Cell("AB6");
                        var cell2 = row.Cell(i);
                        IXLRange xLRange = sheet.Range(cell1, cell2);

                        // 空白セル
                        xLRange.AddConditionalFormat()
                            .WhenIsBlank()
                            .Fill.SetBackgroundColor(XLColor.White);

                        // ０～14日　空色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(0, 14)
                            .Fill.SetBackgroundColor(XLColor.SkyBlue);

                        // 15～30日　黄色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(15, 30)
                            .Fill.SetBackgroundColor(XLColor.Yellow);

                        // 31日～　赤色
                        xLRange.AddConditionalFormat()
                            .WhenEqualOrGreaterThan(31)
                            .Fill.SetBackgroundColor(XLColor.Red);
                    }
                }

                // ルームチェックから発注までの日数計算：数式
                if (i == 32)
                {
                    string formula = "=IF(AND(ISNUMBER(AE" + row.RowNumber() + "), ISNUMBER(Q" + row.RowNumber() + ")), AE" + row.RowNumber() + " - Q" + row.RowNumber() + ", " + @"""" + @"""" + ")";
                    row.Cell(i).FormulaA1 = formula;

                    if (xlsJyokenFormat == 1)
                    {

                        // 条件付き書式を追加
                        var cell1 = sheet.Cell("AF6");
                        var cell2 = row.Cell(i);
                        IXLRange xLRange = sheet.Range(cell1, cell2);

                        // 空白セル
                        xLRange.AddConditionalFormat()
                            .WhenIsBlank()
                            .Fill.SetBackgroundColor(XLColor.White);

                        // ０～5日　空色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(0, 5)
                            .Fill.SetBackgroundColor(XLColor.SkyBlue);

                        // 6～9日　黄色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(6, 9)
                            .Fill.SetBackgroundColor(XLColor.Yellow);

                        // 10日～　赤色
                        xLRange.AddConditionalFormat()
                            .WhenEqualOrGreaterThan(10)
                            .Fill.SetBackgroundColor(XLColor.Red);
                    }
                }

                // 発注から工事終了までの日数計算：数式
                if (i == 39)
                {
                    string formula = "=IF(AND(ISNUMBER(AL" + row.RowNumber() + "), ISNUMBER(AE" + row.RowNumber() + ")), AL" + row.RowNumber() + " - AE" + row.RowNumber() + ", " + @"""" + @"""" + ")";
                    row.Cell(i).FormulaA1 = formula;

                    if (xlsJyokenFormat == 1)
                    {
                        // 条件付き書式を追加
                        var cell1 = sheet.Cell("AM6");
                        var cell2 = row.Cell(i);
                        IXLRange xLRange = sheet.Range(cell1, cell2);

                        // 空白セル
                        xLRange.AddConditionalFormat()
                            .WhenIsBlank()
                            .Fill.SetBackgroundColor(XLColor.White);

                        // ０～9日　空色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(0, 9)
                            .Fill.SetBackgroundColor(XLColor.SkyBlue);

                        // 10～14日　黄色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(10, 14)
                            .Fill.SetBackgroundColor(XLColor.Yellow);

                        // 15日～　赤色
                        xLRange.AddConditionalFormat()
                            .WhenEqualOrGreaterThan(15)
                            .Fill.SetBackgroundColor(XLColor.Red);
                    }
                }

                // RC依頼から工事終了までの日数計算：数式
                if (i == 40)
                {
                    string formula = "=IF(AND(ISNUMBER(AL" + row.RowNumber() + "), ISNUMBER(P" + row.RowNumber() + ")), AL" + row.RowNumber() + " - P" + row.RowNumber() + " + 1, " + @"""" + @"""" + ")";
                    row.Cell(i).FormulaA1 = formula;

                    if (xlsJyokenFormat == 1)
                    {
                        // 条件付き書式を追加
                        var cell1 = sheet.Cell("AN6");
                        var cell2 = row.Cell(i);
                        IXLRange xLRange = sheet.Range(cell1, cell2);

                        // 空白セル
                        xLRange.AddConditionalFormat()
                            .WhenIsBlank()
                            .Fill.SetBackgroundColor(XLColor.White);

                        // ０～14日　空色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(0, 14)
                            .Fill.SetBackgroundColor(XLColor.SkyBlue);

                        // 15～21日　黄色
                        xLRange.AddConditionalFormat()
                            .WhenBetween(15, 21)
                            .Fill.SetBackgroundColor(XLColor.Yellow);

                        // 22日～　赤色
                        xLRange.AddConditionalFormat()
                            .WhenEqualOrGreaterThan(22)
                            .Fill.SetBackgroundColor(XLColor.Red);
                    }
                }
            }
        }


        ///--------------------------------------------------------------------
        /// <summary>
        ///     エクセルのセルにデータをセットする </summary>
        /// <param name="row">
        ///     エクセル行</param>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="i">
        ///     データグリッドビュー行index</param>
        ///--------------------------------------------------------------------
        private void setXlsRowData(IXLRow row, DataGridView dg, int i)
        {
            row.Cell(2).Value = Utility.NulltoStr(dg[colBuName, i].Value);
            row.Cell(3).Value = Utility.NulltoStr(dg[colGou, i].Value);
            row.Cell(4).Value = Utility.NulltoStr(dg[colNewStayDate, i].Value);

            row.Cell(5).Value = Utility.NulltoStr(dg[col_KaiyakuContact_01, i].Value);
            row.Cell(6).Value = Utility.NulltoStr(dg[col_KaiyakuContact_02, i].Value);
            row.Cell(7).Value = Utility.NulltoStr(dg[col_KaiyakuContact_03, i].Value);
            row.Cell(8).Value = Utility.NulltoStr(dg[col_KaiyakuContact_04, i].Value);
            row.Cell(9).Value = Utility.NulltoStr(dg[col_KaiyakuContact_05, i].Value);
            row.Cell(10).Value = Utility.NulltoStr(dg[col_KaiyakuContact_06, i].Value);
            row.Cell(11).Value = Utility.NulltoStr(dg[col_KaiyakuContact_07, i].Value);
            row.Cell(12).Value = Utility.NulltoStr(dg[col_KaiyakuContact_08, i].Value);

            row.Cell(13).Value = Utility.NulltoStr(dg[col_Kaiyaku_01, i].Value);
            row.Cell(14).Value = Utility.NulltoStr(dg[col_Kaiyaku_02, i].Value);
            row.Cell(15).Value = Utility.NulltoStr(dg[col_Kaiyaku_03, i].Value);
            row.Cell(16).Value = Utility.NulltoStr(dg[col_Kaiyaku_04, i].Value);

            row.Cell(17).Value = Utility.NulltoStr(dg[col_RoomCheck_01, i].Value);
            //row.Cell(18).Value = Utility.NulltoStr(dg[col_RoomCheck_02, i].Value);
            row.Cell(19).Value = Utility.NulltoStr(dg[col_RoomCheck_03, i].Value);
            row.Cell(20).Value = Utility.NulltoStr(dg[col_RoomCheck_04, i].Value);
            row.Cell(21).Value = Utility.NulltoStr(dg[col_RoomCheck_05, i].Value);

            row.Cell(22).Value = Utility.NulltoStr(dg[col_Shorui_01, i].Value);
            row.Cell(23).Value = Utility.NulltoStr(dg[col_Shorui_02, i].Value);
            row.Cell(24).Value = Utility.NulltoStr(dg[col_Shorui_03, i].Value);
            row.Cell(25).Value = Utility.NulltoStr(dg[col_Shorui_04, i].Value);

            row.Cell(26).Value = Utility.NulltoStr(dg[col_Tetsu_01, i].Value);
            row.Cell(27).Value = Utility.NulltoStr(dg[col_Tetsu_02, i].Value);
            //row.Cell(28).Value = Utility.NulltoStr(dg[col_Tetsu_03, i].Value);
            row.Cell(29).Value = Utility.NulltoStr(dg[col_Tetsu_04, i].Value);
            row.Cell(30).Value = Utility.NulltoStr(dg[col_Tetsu_05, i].Value);

            row.Cell(31).Value = Utility.NulltoStr(dg[col_Hacchu_01, i].Value);
            //row.Cell(32).Value = Utility.NulltoStr(dg[col_Hacchu_02, i].Value);
            row.Cell(33).Value = Utility.NulltoStr(dg[col_Hacchu_03, i].Value);

            row.Cell(34).Value = Utility.NulltoStr(dg[col_Kouji_01, i].Value);
            row.Cell(35).Value = Utility.NulltoStr(dg[col_Kouji_02, i].Value);
            row.Cell(36).Value = Utility.NulltoStr(dg[col_Kouji_03, i].Value);
            row.Cell(37).Value = Utility.NulltoStr(dg[col_Kouji_04, i].Value);
            row.Cell(38).Value = Utility.NulltoStr(dg[col_Kouji_05, i].Value);
            //row.Cell(39).Value = Utility.NulltoStr(dg[col_Kouji_06, i].Value);
            //row.Cell(40).Value = Utility.NulltoStr(dg[col_Kouji_07, i].Value);

            row.Cell(41).Value = Utility.NulltoStr(dg[col_Kanryo_01, i].Value);
            row.Cell(42).Value = Utility.NulltoStr(dg[col_Kanryo_02, i].Value);
            row.Cell(43).Value = Utility.NulltoStr(dg[col_Kanryo_03, i].Value);
            row.Cell(44).Value = Utility.NulltoStr(dg[col_Kanryo_04, i].Value);

            row.Cell(45).Value = Utility.NulltoStr(dg[col_Bikou_01, i].Value);

            row.Cell(46).Value = Utility.NulltoStr(dg[col_SkyOne_01, i].Value);
            row.Cell(47).Value = Utility.NulltoStr(dg[col_SkyOne_02, i].Value);
        }




        private bool XlsUpdate(string sFile, int r)
        {
            //try
            //{
            //    using (var bk = new XLWorkbook(sFile, XLEventTracking.Disabled))
            //    {
            //        var sheet1 = bk.Worksheet(Properties.Settings.Default.xlsSheetName);

            //        // 対象行を取得
            //        var row = sheet1.Row(r);

            //        row.Cell(1).Value = txtBuCode.Text;
            //        row.Cell(2).Value = txtBuName.Text;
            //        row.Cell(3).Value = txtGou.Text;
            //        row.Cell(4).Value = txtNewStayDate.Text;

            //        row.Cell(5).Value = txtKaiyakuContact01.Text;
            //        row.Cell(6).Value = txtKaiyakuContact02.Text;
            //        row.Cell(7).Value = txtKaiyakuContact03.Text;
            //        row.Cell(8).Value = txtKaiyakuContact04.Text;
            //        row.Cell(9).Value = txtKaiyakuContact05.Text;
            //        row.Cell(10).Value = txtKaiyakuContact06.Text;
            //        row.Cell(11).Value = txtKaiyakuContact07.Text;
            //        //row.Cell(12).Value = txtKaiyakuContact08.Text;
            //        row.Cell(12).Value = Utility.NulltoStr(cmbKaiyakuContact08.Text);

            //        row.Cell(13).Value = txtKaiyaku01.Text;
            //        row.Cell(14).Value = txtKaiyaku02.Text;
            //        row.Cell(15).Value = txtKaiyaku03.Text;
            //        row.Cell(16).Value = txtKaiyaku04.Text;

            //        row.Cell(17).Value = txtRoomCheck01.Text;
            //        row.Cell(18).Value = txtRoomCheck02.Text;
            //        row.Cell(19).Value = txtRoomCheck03.Text;
            //        //row.Cell(20).Value = txtRoomCheck04.Text;
            //        row.Cell(20).Value = Utility.NulltoStr(cmbKeyOkiba.Text);
            //        row.Cell(21).Value = txtRoomCheck05.Text;

            //        row.Cell(22).Value = txtShorui01.Text;
            //        row.Cell(23).Value = txtShorui02.Text;
            //        row.Cell(24).Value = txtShorui03.Text;
            //        row.Cell(25).Value = txtShorui04.Text;

            //        row.Cell(26).Value = txtTetsu01.Text;
            //        row.Cell(27).Value = txtTetsu02.Text;
            //        row.Cell(28).Value = txtTetsu03.Text;
            //        row.Cell(29).Value = txtTetsu04.Text;
            //        row.Cell(30).Value = txtTetsu05.Text;

            //        row.Cell(31).Value = txtHacchu01.Text;
            //        row.Cell(32).Value = txtHacchu02.Text;
            //        row.Cell(33).Value = txtHacchu03.Text;

            //        //row.Cell(34).Value = txtKouji01.Text;
            //        row.Cell(34).Value = Utility.NulltoStr(cmbGyousha.Text);
            //        row.Cell(35).Value = txtKouji02.Text;
            //        row.Cell(36).Value = txtKouji03.Text;
            //        row.Cell(37).Value = txtKouji04.Text;
            //        row.Cell(38).Value = txtKouji05.Text;
            //        row.Cell(39).Value = txtKouji06.Text;
            //        row.Cell(40).Value = txtKouji07.Text;

            //        row.Cell(41).Value = txtKanryo01.Text;
            //        row.Cell(42).Value = txtKanryo02.Text;
            //        row.Cell(43).Value = txtKanryo03.Text;
            //        row.Cell(44).Value = txtKanryo04.Text;

            //        row.Cell(45).Value = txtBikou.Text;

            //        row.Cell(46).Value = txtSkyOne01.Text;
            //        //row.Cell(47).Value = txtSkyOne02.Text;
            //        row.Cell(47).Value = Utility.NulltoStr(cmbHosho.Text);

            //        // 更新
            //        bk.SaveAs(sFile);

            //        return true;
            //    }
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show(ex.Message);
            //    return false;
            //}

            return true;
        }

        private void txtRoomCheck01_TextChanged(object sender, EventArgs e)
        {
            // RC依頼→ルームチェックまでの日数と背景色
            txtRoomCheck02.Text = getDaySpan(txtKaiyaku04.Text, txtRoomCheck01.Text);
            backColorUpdate_01(txtRoomCheck02.Text);

            // ルームチェックから発注までの日数と背景色
            txtHacchu02.Text = getDaySpan(txtRoomCheck01.Text, txtHacchu01.Text);
            backColorUpdate_03(txtHacchu02.Text);
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     経過日数を計算する </summary>
        /// <param name="sDt">
        ///     開始日付</param>
        /// <param name="eDt">
        ///     終了日付</param>
        /// <returns>
        ///     日数</returns>
        ///--------------------------------------------------------------
        private string getDaySpan(string sDt, string eDt)
        {
            DateTime fromDt;
            DateTime ToDt;

            if ((DateTime.TryParse(sDt, out fromDt)) && (DateTime.TryParse(eDt, out ToDt)))
            {
                if (fromDt > ToDt)
                {
                    return string.Empty;
                }
                else
                {
                    return Utility.GetTimeSpan(fromDt, ToDt).TotalDays.ToString();
                }
            }
            else
            {
                return string.Empty;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     RC依頼→ルームチェックまでの日数の背景色 </summary>
        /// <param name="n">
        ///     日数 </param>
        ///--------------------------------------------------------------
        private void backColorUpdate_01(string n)
        {
            if (n == string.Empty)
            {
                txtRoomCheck02.BackColor = Color.Empty;
                return;
            }

            int nDays = Utility.StrtoInt(n);

            if (nDays <= 2)
            {
                txtRoomCheck02.BackColor = Color.SkyBlue;
            }
            else if (nDays <= 5)
            {
                txtRoomCheck02.BackColor = Color.Yellow;
            }
            else
            {
                txtRoomCheck02.BackColor = Color.Red;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     オーナー見積書送付から本日までの日数の背景色 </summary>
        /// <param name="n">
        ///     日数 </param>
        ///--------------------------------------------------------------
        private void backColorUpdate_02(string n)
        {
            if (n == string.Empty)
            {
                txtTetsu03.BackColor = Color.Empty;
                return;
            }

            int nDays = Utility.StrtoInt(n);

            if (nDays <= 14)
            {
                txtTetsu03.BackColor = Color.SkyBlue;
            }
            else if (nDays <= 30)
            {
                txtTetsu03.BackColor = Color.Yellow;
            }
            else
            {
                txtTetsu03.BackColor = Color.Red;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     ルームチェックから発注までの日数の背景色 </summary>
        /// <param name="n">
        ///     日数 </param>
        ///--------------------------------------------------------------
        private void backColorUpdate_03(string n)
        {
            if (n == string.Empty)
            {
                txtHacchu02.BackColor = Color.Empty;
                return;
            }

            int nDays = Utility.StrtoInt(n);

            if (nDays <= 5)
            {
                txtHacchu02.BackColor = Color.SkyBlue;
            }
            else if (nDays <= 9)
            {
                txtHacchu02.BackColor = Color.Yellow;
            }
            else
            {
                txtHacchu02.BackColor = Color.Red;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     発注から工事終了までの日数の背景色 </summary>
        /// <param name="n">
        ///     日数 </param>
        ///--------------------------------------------------------------
        private void backColorUpdate_04(string n)
        {
            if (n == string.Empty)
            {
                txtKouji06.BackColor = Color.Empty;
                return;
            }

            int nDays = Utility.StrtoInt(n);

            if (nDays <= 9)
            {
                txtKouji06.BackColor = Color.SkyBlue;
            }
            else if (nDays <= 14)
            {
                txtKouji06.BackColor = Color.Yellow;
            }
            else
            {
                txtKouji06.BackColor = Color.Red;
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     RC依頼から工事終了までの日数の背景色 </summary>
        /// <param name="n">
        ///     日数 </param>
        ///--------------------------------------------------------------
        private void backColorUpdate_05(string n)
        {
            if (n == string.Empty)
            {
                txtKouji07.BackColor = Color.Empty;
                return;
            }

            int nDays = Utility.StrtoInt(n);

            if (nDays <= 14)
            {
                txtKouji07.BackColor = Color.SkyBlue;
            }
            else if (nDays <= 21)
            {
                txtKouji07.BackColor = Color.Yellow;
            }
            else
            {
                txtKouji07.BackColor = Color.Red;
            }
        }

        private void txtKaiyaku04_TextChanged(object sender, EventArgs e)
        {
            // RC依頼→ルームチェックまでの日数と背景色
            txtRoomCheck02.Text = getDaySpan(txtKaiyaku04.Text, txtRoomCheck01.Text);
            backColorUpdate_01(txtRoomCheck02.Text);

            // RC依頼→工事終了までの日数と背景色
            string d = getDaySpan(txtKaiyaku04.Text, txtKouji05.Text);

            if (d != string.Empty)
            {
                txtKouji07.Text = (Utility.StrtoInt(d) + 1).ToString();
            }
            else
            {
                txtKouji07.Text = d;
            }

            backColorUpdate_05(txtKouji07.Text);
        }

        private void txtTetsu02_TextChanged(object sender, EventArgs e)
        {
            // 見積書送付から本日までの日数と背景色
            txtTetsu03.Text = getDaySpan(txtTetsu02.Text, DateTime.Today.ToShortDateString());
            backColorUpdate_02(txtTetsu03.Text);
        }

        private void txtHacchu01_TextChanged(object sender, EventArgs e)
        {
            // ルームチェックから発注までの日数と背景色
            txtHacchu02.Text = getDaySpan(txtRoomCheck01.Text, txtHacchu01.Text);
            backColorUpdate_03(txtHacchu02.Text);

            // 発注から工事終了までの日数と背景色
            txtKouji06.Text = getDaySpan(txtHacchu01.Text, txtKouji05.Text);
            backColorUpdate_04(txtKouji06.Text);
        }

        private void txtKouji05_TextChanged(object sender, EventArgs e)
        {
            // 発注から工事終了までの日数と背景色
            txtKouji06.Text = getDaySpan(txtHacchu01.Text, txtKouji05.Text);
            backColorUpdate_04(txtKouji06.Text);

            // RC依頼→工事終了までの日数と背景色
            string d = getDaySpan(txtKaiyaku04.Text, txtKouji05.Text);

            if (d != string.Empty)
            {
                txtKouji07.Text = (Utility.StrtoInt(d) + 1).ToString();
            }
            else
            {
                txtKouji07.Text = d;
            }

            backColorUpdate_05(txtKouji07.Text);
        }

        private void setCmbGyousha(DataGridView dg)
        {
            for (int i = 0; i < dg.RowCount; i++)
            {
                string gName = Utility.NulltoStr(dg[col_Kouji_01, i].Value);

                if (gName == string.Empty)
                {
                    continue;
                }

                if (cmbGyousha.Items.Count < 1)
                {
                    cmbGyousha.Items.Add(gName);
                }
                else
                {
                    bool isItem = false;

                    for (int iX = 0; iX < cmbGyousha.Items.Count; iX++)
                    {
                        if (cmbGyousha.Items[iX].ToString() == gName)
                        {
                            isItem = true;  // 追加済みの業者名
                            break;
                        }
                    }

                    // 未追加なら追加する
                    if (!isItem)
                    {
                        cmbGyousha.Items.Add(gName);
                    }
                }
            }
        }

        ///----------------------------------------------------------
        /// <summary>
        ///     コンボボックスアイテムセット </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="cb">
        ///     コンボボックスオブジェクト</param>
        /// <param name="dgCol">
        ///     データグリッドカラム</param>
        ///----------------------------------------------------------
        private void setCmbItems(DataGridView dg, ComboBox cb, string dgCol)
        {
            for (int i = 0; i < dg.RowCount; i++)
            {
                string gName = Utility.NulltoStr(dg[dgCol, i].Value);

                if (gName == string.Empty)
                {
                    continue;
                }

                if (cb.Items.Count < 1)
                {
                    cb.Items.Add(gName);
                }
                else
                {
                    bool isItem = false;

                    for (int iX = 0; iX < cb.Items.Count; iX++)
                    {
                        if (cb.Items[iX].ToString() == gName)
                        {
                            isItem = true;  // 追加済みの業者名
                            break;
                        }
                    }

                    // 未追加なら追加する
                    if (!isItem)
                    {
                        cb.Items.Add(gName);
                    }
                }
            }

            cb.SelectedIndex = -1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show("表示中データを登録しないで画面を戻します。よろしいですか", "取消確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question)== DialogResult.No)
            {
                return;
            }

            // 画面初期化
            dispInitial();
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
            {
                // 新規登録
                //dataGridView1.Enabled = false;
                //dataGridView1.CurrentCell = null;
                txtBuCode.Enabled = true;
                button2.Enabled = true;
                txtBuCode.Focus();

                dataGridView1.CurrentCell = dataGridView1[1, dataGridView1.RowCount - 1];
                dataGridView1.CurrentCell = null;
            }
            else
            {
                // 編集
                dataGridView1.Enabled = true;
                txtBuCode.Enabled = false;
                button2.Enabled = false;
            }
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
            {
                // 編集
                dataGridView1.Enabled = true;
                txtBuCode.Enabled = false;
                button2.Enabled = false;
            }
            else
            {
                // 新規登録
                //dataGridView1.Enabled = false;
                dataGridView1.CurrentCell = null;
                txtBuCode.Enabled = true;
                button2.Enabled = true;
                txtBuCode.Focus();
            }
        }

        private void txtBuCode_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string sPath = System.IO.Path.GetDirectoryName(xlsFname);

            // 自らのロックファイルを削除する
            Utility.deleteLockFile(sPath, System.Net.Dns.GetHostName());

            // 他のPCで処理中の場合、続行不可
            //if (Utility.existsLockFile(sPath))
            //{
            //    MessageBox.Show("他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。再度実行してください。", "確認", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
            //    return;
            //}
            
            while (Utility.existsLockFile(sPath))
            {
                Cursor = Cursors.WaitCursor;
                pictureBox1.Visible = true;
                lblMsg.Text = "他のPCが解約管理表エクセルファイルをオープンまたはクローズ中です。少々おまちください...";
                System.Threading.Thread.Sleep(100);
                Application.DoEvents();
            }

            lblMsg.Text = "";
            Cursor = Cursors.Default;
            pictureBox1.Visible = false;

            // Excelファイルを開く
            tFile = xlsFname;
            gridViewShowData(dataGridView1, tFile, xlsPass, string.Empty);

            // 画面初期化
            dispInitial();

            // ボタン
            button4.Enabled = false;

            dataGridView1.CurrentCell = null;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            int sr = 0;

            if (dataGridView1.CurrentCell == null)
            {
                sr = 0;
            }
            else
            {
                DataGridViewCell cell = dataGridView1.CurrentCell;
                sr = cell.RowIndex + 1;
            }

            int r = searchBuName(dataGridView1, txtsBuName.Text, sr);
            if (r == -1)
            {
                MessageBox.Show("該当する物件はありませんでした", "検索結果", MessageBoxButtons.OK, MessageBoxIcon.Information);
                dataGridView1.CurrentCell = null;
            }
            else
            {
                dataGridView1.CurrentCell = dataGridView1[1, r];
            }
        }

        ///--------------------------------------------------------------
        /// <summary>
        ///     データグリッドビュー物件名検索 </summary>
        /// <param name="dg">
        ///     データグリッドビューオブジェクト</param>
        /// <param name="sBuName">
        ///     検索物件名</param>
        /// <param name="r">
        ///     開始行Index</param>
        /// <returns>
        ///     検索行Index</returns>
        ///--------------------------------------------------------------
        private int searchBuName(DataGridView dg, string sBuName, int r)
        {
            for (int i = r; i < dg.RowCount; i++)
            {
                if (Utility.NulltoStr(dg[colBuName, i].Value).Contains(sBuName))
                {
                    return i;
                }
            }

            return -1;
        }

        private void txtsBuName_TextChanged(object sender, EventArgs e)
        {
            if (txtsBuName.Text == string.Empty)
            {
                button5.Enabled = false;
            }
            else
            {
                button5.Enabled = true;
            }
        }
    }
}

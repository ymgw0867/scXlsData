using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using scXlsData.common;
using System.Data.OleDb;

namespace scXlsData.config
{
    public partial class frmConfig : Form
    {
        public frmConfig()
        {
            InitializeComponent();
        }

        DataSet1 dts = new DataSet1();
        DataSet1TableAdapters.環境設定TableAdapter adp = new DataSet1TableAdapters.環境設定TableAdapter();

        private void frmConfig_Load(object sender, EventArgs e)
        {
            Utility.WindowsMaxSize(this, this.Width, this.Height);
            Utility.WindowsMinSize(this, this.Width, this.Height);

            adp.FillByID(dts.環境設定);

            if (dts.環境設定.Any(a => a.ID == 1))
            {
                var s = dts.環境設定.Single(a => a.ID == 1);

                if (s.IstargetXlsFileNull())
                {
                    txtFilePath.Text = string.Empty;
                }
                else
                {
                    txtFilePath.Text = s.targetXlsFile;
                }

                if (s.IssheetPasswordNull())
                {
                    txtPassword.Text = string.Empty;
                }
                else
                {
                    txtPassword.Text = s.sheetPassword;
                }
            }
            else
            {
                txtFilePath.Text = string.Empty;
                txtPassword.Text = string.Empty;
            }
        }

        /// -----------------------------------------------------------------
        /// <summary>
        ///     フォルダダイアログ選択 </summary>
        /// <returns>
        ///     フォルダー名</returns>
        /// -----------------------------------------------------------------
        private void userFolderSelect()
        {
            openFileDialog1.Title = "解約管理表ファイル選択";
            openFileDialog1.FileName = string.Empty;
            openFileDialog1.Filter = "エクセルファイル(*.xlsx)|*.xlsx|全てのファイル(*.*)|*.*";

            //ダイアログボックスを表示し「保存」ボタンが選択されたらファイル名を表示
            //string fileName;
            DialogResult ret = openFileDialog1.ShowDialog();

            if (ret == System.Windows.Forms.DialogResult.OK)
            {
                txtFilePath.Text = openFileDialog1.FileName;
            }

            // 不要になった時点で破棄する
            openFileDialog1.Dispose();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            userFolderSelect();
        }

        private void txtYear_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar < '0' || e.KeyChar > '9') && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            // データ更新
            DataUpdate();
        }

        private void DataUpdate()
        {
            if (MessageBox.Show("データを更新してよろしいですか","確認",MessageBoxButtons.YesNo,MessageBoxIcon.Question) == System.Windows.Forms.DialogResult.No) return;

            // エラーチェック
            if (!errCheck())
            {
                return;
            }

            adp.UpdateQuery(txtFilePath.Text, txtPassword.Text, 1);
 
            // 終了
            this.Close();
        }

        private bool errCheck()
        {
            // ファイルパス
            if (txtFilePath.Text.Trim() == string.Empty)
            {
                MessageBox.Show("エクセルファイルパスが指定されていません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtFilePath.Focus();
                return false;
            }

            if (!System.IO.File.Exists(txtFilePath.Text))
            {
                MessageBox.Show("指定されたエクセルファイルが見つかりません", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                txtFilePath.Focus();
                return false;
            }

            return true;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void frmConfig_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 後片付け
            this.Dispose();
        }
    }
}

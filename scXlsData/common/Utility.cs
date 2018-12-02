using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Drawing;
using System.Windows;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;

namespace scXlsData.common
{
    class Utility
    {
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     Height</param>
        /// ---------------------------------------------------------------------
        public static void WindowsMinSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MinimumSize = new System.Drawing.Size(wSize, hSize);
        }
        /// ---------------------------------------------------------------------
        /// <summary>
        ///     ウィンドウ最小サイズの設定 </summary>
        /// <param name="tempFrm">
        ///     対象とするウィンドウオブジェクト</param>
        /// <param name="wSize">
        ///     width</param>
        /// <param name="hSize">
        ///     height</param>
        /// --------------------------------------------------------------------
        public static void WindowsMaxSize(Form tempFrm, int wSize, int hSize)
        {
            tempFrm.MaximumSize = new System.Drawing.Size(wSize, hSize);
        }

        /// <summary>
        /// フォームのデータ登録モード
        /// </summary>
        public class frmMode
        {
            public int Mode { get; set; }
            public string ID { get; set; }
            public int rowIndex { get; set; }
            public int closeMode { get; set; }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     文字列の値が数字かチェックする </summary>
        /// <param name="tempStr">
        ///     検証する文字列</param>
        /// <returns>
        ///     数字:true,数字でない:false</returns>
        /// ------------------------------------------------------------------------------
        public static bool NumericCheck(string tempStr)
        {
            double d;

            if (tempStr == null) return false;

            if (double.TryParse(tempStr, System.Globalization.NumberStyles.Any, System.Globalization.NumberFormatInfo.InvariantInfo, out d) == false)
                return false;

            return true;
        }
        
        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     emptyを"0"に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string EmptytoZero(string tempStr)
        {
            if (tempStr == string.Empty)
            {
                return "0";
            }
            else
            {
                return tempStr;
            }
        }

        /// ------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのとき文字型値を返す</returns>
        /// ------------------------------------------------------------------------------
        public static string NulltoStr(string tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                return tempStr;
            }
        }

        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string NulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }
        
        /// -------------------------------------------------------------------------------
        /// <summary>
        ///     Nullをstring.Empty("")に置き換える </summary>
        /// <param name="tempStr">
        ///     stringオブジェクト</param>
        /// <returns>
        ///     nullのときstring.Empty、not nullのときそのまま値を返す</returns>
        /// -------------------------------------------------------------------------------
        public static string dbNulltoStr(object tempStr)
        {
            if (tempStr == null)
            {
                return string.Empty;
            }
            else
            {
                if (tempStr == DBNull.Value)
                {
                    return string.Empty;
                }
                else
                {
                    return (string)tempStr.ToString();
                }
            }
        }
        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をIntへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     Int型の値</returns>
        /// --------------------------------------------------------------------------------
        public static int StrtoInt(string tempStr)
        {
            if (NumericCheck(tempStr)) return int.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     文字型をDoubleへ変換して返す（数値でないときは０を返す）</summary>
        /// <param name="tempStr">
        ///     文字型の値</param>
        /// <returns>
        ///     double型の値</returns>
        /// --------------------------------------------------------------------------------
        public static double StrtoDouble(string tempStr)
        {
            if (NumericCheck(tempStr)) return double.Parse(tempStr);
            else return 0;
        }

        /// --------------------------------------------------------------------------------
        /// <summary>
        ///     経過時間を返す </summary>
        /// <param name="s">
        ///     開始時間</param>
        /// <param name="e">
        ///     終了時間</param>
        /// <returns>
        ///     経過時間</returns>
        /// --------------------------------------------------------------------------------
        public static TimeSpan GetTimeSpan(DateTime s, DateTime e)
        {
            TimeSpan ts;
            if (s > e)
            {
                TimeSpan j = new TimeSpan(24, 0, 0);
                ts = e + j - s;
            }
            else
            {
                ts = e - s;
            }

            return ts;
        }

        /// ------------------------------------------------------------------------
        /// <summary>
        ///     指定した精度の数値に切り捨てます。</summary>
        /// <param name="dValue">
        ///     丸め対象の倍精度浮動小数点数。</param>
        /// <param name="iDigits">
        ///     戻り値の有効桁数の精度。</param>
        /// <returns>
        ///     iDigits に等しい精度の数値に切り捨てられた数値。</returns>
        /// ------------------------------------------------------------------------
        public static double ToRoundDown(double dValue, int iDigits)
        {
            double dCoef = System.Math.Pow(10, iDigits);

            return dValue > 0 ? System.Math.Floor(dValue * dCoef) / dCoef :
                                System.Math.Ceiling(dValue * dCoef) / dCoef;
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     数値(分合計）から時間：分形式の文字列を求める</summary>
        /// <param name="tm">
        ///     数値</param>
        /// <returns>
        ///     hh:mm形式の文字列</returns>
        ///-------------------------------------------------------------------------
        public static string dblToHHMM(double tm)
        {
            int hor = (int)(System.Math.Floor(tm / 60));
            int min = (int)(tm % 60);
            string hm = hor.ToString() + ":" + min.ToString().PadLeft(2, '0');
            return hm;
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     任意のディレクトリのファイルを削除する </summary>
        /// <param name="sPath">
        ///     指定するディレクトリ</param>
        /// <param name="sFileType">
        ///     ファイル名及び形式</param>
        /// --------------------------------------------------------------------
        public static void FileDelete(string sPath, string sFileType)
        {
            //sFileTypeワイルドカード"*"は、すべてのファイルを意味する
            foreach (string files in System.IO.Directory.GetFiles(sPath, sFileType))
            {
                // ファイルを削除する
                System.IO.File.Delete(files);
            }
        }

        ///---------------------------------------------------------------------
        /// <summary>
        ///     文字列を指定文字数をＭＡＸとして返します</summary>
        /// <param name="s">
        ///     文字列</param>
        /// <param name="n">
        ///     文字数</param>
        /// <returns>
        ///     文字数範囲内の文字列</returns>
        /// --------------------------------------------------------------------
        public static string GetStringSubMax(string s, int n)
        {
            string val = string.Empty;
            s = s.Replace(" ","");
            if (s.Length > n) val = s.Substring(0, n);
            else val = s;

            return val;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     8ケタ左詰め右空白埋めの給与大臣検索用の社員コード文字列を返す
        /// </summary>
        /// <param name="sCode">
        ///     コード</param>
        /// <returns>
        ///     給与大臣検索用の社員コード文字列</returns>
        /// --------------------------------------------------------------------
        public static string bldShainCode(string sCode)
        {
            return sCode.PadLeft(4, '0').PadRight(8, ' ').Substring(0, 8);
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     チェックボックスのステータスを数値で返す </summary>
        /// <param name="chk">
        ///     チェックボックスオブジェクト</param>
        /// <returns>
        ///     チェックのとき１、未チェックのとき０</returns>
        /// --------------------------------------------------------------------
        public static int checkToInt(CheckBox chk)
        {
            int rtn = 0;
            if (chk.CheckState == CheckState.Checked)
            {
                rtn = 1;
            }
            else
            {
                rtn = 0;
            }

            return rtn;
        }

        /// --------------------------------------------------------------------
        /// <summary>
        ///     数値をBool値で返す </summary>
        /// <param name="chk">
        ///     数値</param>
        /// <returns>
        ///     １のときTrue、0のときFalse</returns>
        /// --------------------------------------------------------------------
        public static bool intToCheck(int chk)
        {
            bool rtn = false;
            if (chk == 0)
            {
                rtn = false;
            }
            else if (chk == 1)
            {
                rtn = true;
            }

            return rtn;
        }

        public static bool getHHMM(string val, out int hh, out int mm)
        {
            hh = 0;
            mm = 0;

            // 文字列が空のとき
            if (val.Trim() == string.Empty)
            {
                hh = 0;
                mm = 0;
                return false;
            }

            DateTime eDate;
            if (DateTime.TryParse(val, out eDate))
            {
                hh = eDate.Hour;
                mm = eDate.Minute;
            }
            else
            {
                string[] z = val.Split(':');
                for (int i = 0; i < z.Length; i++)
                {
                    if (i == 0) hh = StrtoInt(z[i]);
                    if (i == 1) mm = StrtoInt(z[i]);
                }
            }

            return true;
        }
        
        ///----------------------------------------------------------------------
        /// <summary>
        ///     分から時：分(hh：mm)形式の文字列に変換して返す </summary>
        /// <param name="minu">
        ///     分</param>
        /// <returns>
        ///     hh：mm形式文字列</returns>
        ///----------------------------------------------------------------------
        public static string getHHMM(int minu)
        {
            int hh = minu / 60;
            int mm = minu - (hh * 60);

            return hh + ":" + mm.ToString().PadLeft(2, '0');
        }
               
        public static OleDbConnection dbConnect()
        {
            // データベース接続文字列
            OleDbConnection Cn = new OleDbConnection();
            StringBuilder sb = new StringBuilder();
            sb.Clear();
            sb.Append("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=");
            //sb.Append(Properties.Settings.Default.ocrMdbPath);
            Cn.ConnectionString = sb.ToString();
            Cn.Open();

            return Cn;
        }

        ///----------------------------------------------------------------------------
        /// <summary>
        ///     CSVファイルを出力する</summary>
        /// <param name="sPath">
        ///     出力するパス</param>
        /// <param name="arrayData">
        ///     書き込む配列データ</param>
        /// <param name="sFileName">
        ///     CSVファイル名</param>
        ///----------------------------------------------------------------------------
        public static void csvFileWrite(string sPath, string[] arrayData, string sFileName)
        {
            // ファイル名
            string outFileName = sPath + sFileName + ".csv";

            // 出力ファイルが存在するとき
            if (System.IO.File.Exists(outFileName))
            {
                // リネーム付加文字列（タイムスタンプ）
                string newFileName = DateTime.Now.Year.ToString() + DateTime.Now.Month.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Day.ToString().PadLeft(2, '0') + DateTime.Now.Hour.ToString().PadLeft(2, '0') +
                                     DateTime.Now.Minute.ToString().PadLeft(2, '0') + DateTime.Now.Second.ToString().PadLeft(2, '0');

                // リネーム後ファイル名
                string reFileName = sPath + sFileName + newFileName + ".csv";

                // 既存のファイルをリネーム
                File.Move(outFileName, reFileName);
            }

            // CSVファイル出力 : 2015/08/25 UTF-8で出力
            //File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.GetEncoding(65001));
            //File.WriteAllLines(outFileName, arrayData, System.Text.Encoding.UTF8);

            // BOM無しで出力：2015/09/24
            System.Text.Encoding enc = new System.Text.UTF8Encoding(false);
            File.WriteAllLines(outFileName, arrayData, enc);
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     自らのロックファイルが存在したら削除する </summary>
        /// <param name="fPath">
        ///     パス</param>
        /// <param name="PcK">
        ///     自分のロックファイル文字列</param>
        ///-------------------------------------------------------------------------
        public static void deleteLockFile(string fPath, string PcK)
        {
            string FileName = fPath + @"\" + global.LOCK_FILEHEAD + PcK + ".loc";

            if (System.IO.File.Exists(FileName))
            {
                System.IO.File.Delete(FileName);
            }
        }

        ///-------------------------------------------------------------------------
        /// <summary>
        ///     データフォルダにロックファイルが存在するか調べる </summary>
        /// <param name="fPath">
        ///     データフォルダパス</param>
        /// <returns>
        ///     true:ロックファイルあり、false:ロックファイルなし</returns>
        ///-------------------------------------------------------------------------
        public static Boolean existsLockFile(string fPath)
        {
            int s = System.IO.Directory.GetFiles(fPath, global.LOCK_FILEHEAD + "*.*", System.IO.SearchOption.TopDirectoryOnly).Count();

            if (s == 0)
            {
                return false; //LOCKファイルが存在しない
            }
            else
            {
                return true;   //存在する
            }
        }

        ///----------------------------------------------------------------
        /// <summary>
        ///     ロックファイルを登録する </summary>
        /// <param name="fPath">
        ///     書き込み先フォルダパス</param>
        /// <param name="LocName">
        ///     ファイル名</param>
        ///----------------------------------------------------------------
        public static void makeLockFile(string fPath, string LocName)
        {
            string FileName = fPath + @"\" + global.LOCK_FILEHEAD + LocName + ".loc";

            //存在する場合は、処理なし
            if (System.IO.File.Exists(FileName))
            {
                return;
            }

            // ロックファイルを登録する
            try
            {
                System.IO.StreamWriter outFile = new System.IO.StreamWriter(FileName, false, System.Text.Encoding.GetEncoding(932));
                outFile.Close();
            }
            catch
            {
            }

            return;
        }
    }
}

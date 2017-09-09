using ExcelGreper.classes;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;
using static System.Windows.Forms.ListViewItem;

namespace ExcelGreper
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            // 検索結果一覧の初期化
            lvResult.View = View.Details;
            int w10p = (lvResult.Width - 5) / 10;
            int w5p = w10p / 2;
            lvResult.Columns.Add("No", w5p, HorizontalAlignment.Right);
            lvResult.Columns.Add("ファイルパス", w10p * 3, HorizontalAlignment.Left);
            lvResult.Columns.Add("オブジェクト", w10p * 1, HorizontalAlignment.Left);
            lvResult.Columns.Add("アドレス", w10p * 1, HorizontalAlignment.Left);
            lvResult.Columns.Add("テキスト", w10p * 4 + w5p, HorizontalAlignment.Left);
        }

        private void btnSearch_Click(object sender, EventArgs e)
        {
            string keyword = txtKeyWord.Text;
            string filePath = txtFilePath.Text;

            ExcelGreper.classes.ExcelGreper greper =
                new ExcelGreper.classes.ExcelGreper(keyword, filePath);

            IList<GrepResult> resultList = greper.Execute();

            if (resultList.Count == 0)
            {
                lbStatus.Text = "該当のキーワードは見つかりませんでした。";
                return;
            }

            lvResult.Items.Clear();

            int noCnt = 1;

            foreach (GrepResult result in resultList)
            {
                string viewType = result.Type == GrepResult.Types.Cell ? "セル" : "シェイプ";

                ListViewItem item = lvResult.Items.Add(noCnt.ToString());
                item.SubItems.Add(result.FilePath);
                item.SubItems.Add(viewType);
                item.SubItems.Add(result.CellAddress);
                item.SubItems.Add(result.Text);

                noCnt++;
            }

            lbStatus.Text = resultList.Count + "件見つかりました。";
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            lvResult.Items.Clear();
            lbStatus.Text = "";
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            if (lvResult.Items.Count == 0)
            {
                lbStatus.Text = "エクスポート対象がありません。";
                return;
            }

            //-------------------------------------------------------------------------
            // ファイルパスの入力
            //-------------------------------------------------------------------------
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.FileName =
                DateTime.Now.ToString("yyyyMMdd") + "_grepresult.csv";
            //sfd.InitialDirectory = @"C:\";
            sfd.Filter = "CSVファイル(*.csv)|*.csv";
            sfd.FilterIndex = 1;
            sfd.Title = "保存先のファイルを選択してください";
            sfd.RestoreDirectory = false;
            sfd.OverwritePrompt = true;
            sfd.CheckPathExists = true;

            if (sfd.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            //-------------------------------------------------------------------------
            // Grep結果をファイル出力
            //-------------------------------------------------------------------------
            try
            {
                ExportToCSV(sfd.FileName);
            }
            catch (Exception ex)
            {
                MessageBox.Show("エクスポートに失敗しました。\r\n" + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            MessageBox.Show("エクスポートが完了しました。", "完了", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
        }

        private void ExportToCSV(string exportFilePath)
        {
            StringBuilder buffer = new StringBuilder();

            // 列番号
            // No列は出力しない
            int col = 1; 

            // ヘッダー部の生成
            buffer.Append("ファイルパス,オブジェクト,アドレス,テキスト\r\n");

            // ボディ部の生成
            foreach (ListViewItem item in lvResult.Items)
            {
                foreach (ListViewSubItem subItem in item.SubItems)
                {
                    // 2列目以降はカンマを付加
                    if (col >= 2)
                    {
                        buffer.Append(",");
                    }

                    // 4列目(テキスト)は改行を取り除く
                    if (col == 4)
                    {
                        buffer.Append(subItem.Text.Replace("\n", ""));
                    }
                    else
                    {
                        buffer.Append(subItem.Text);
                    }

                    col++;
                }

                buffer.Append("\r\n");
                col = 1;
            }

            // ファイル出力
            using (StreamWriter writer = new StreamWriter(exportFilePath, false, Encoding.GetEncoding("shift_jis")))
            {
                writer.Write(buffer.ToString());
            }
        }
    }
}

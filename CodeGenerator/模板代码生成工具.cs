using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace CodeGenerator
{
    public partial class 模板代码生成工具 : Form
    {
        public 模板代码生成工具()
        {
            InitializeComponent();
        }

        private void txtExcelTemplatePath_DragDrop(object sender, DragEventArgs e)
        {
            string path = ((System.Array)e.Data.GetData(DataFormats.FileDrop)).GetValue(0).ToString();
            txtFilePath.Text = path;
        }

        private void txtExcelTemplatePath_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Link;
            }
            else
            {
                e.Effect = DragDropEffects.None;
            }
        }

        private void btnChoose_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDlg = new OpenFileDialog();
            openFileDlg.Filter = "Excel 2003文件|*.xls|Excel 2007文件|*.xlsx|所有文件|*";
            //openFileDlg.InitialDirectory = AppDomain.CurrentDomain.BaseDirectory;
            if (DialogResult.OK.Equals(openFileDlg.ShowDialog()))
            {
                txtFilePath.Text = openFileDlg.FileName;
            }
        }

        private void btnCreate_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(txtFilePath.Text.Trim()))
            {
                MessageBox.Show("请选择Excel模板文件！", "提示");
                return;
            }

            TemplateHandler.Create(txtFilePath.Text);

            Process.Start(Path.GetDirectoryName(txtFilePath.Text));

            this.Close();
        }
    }
}

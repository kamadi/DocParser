using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Collections.Generic;
using Microsoft.Office.Interop.Word;
using System.Threading;
using Newtonsoft.Json;

namespace DocParser
{
    public partial class Form1 : Form
    {
        WordDocumentHelper wordDocumentHelper;

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            string filePath = null;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "Word  документы|*.doc;*.docx";
            dialog.Title = "Выберите файл";
            dialog.InitialDirectory = "c:\\";
            dialog.RestoreDirectory = true;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                // store selected file path
                filePath = dialog.FileName.ToString();
                wordDocumentHelper = new WordDocumentHelper(filePath);
                wordDocumentHelper.parse();
                //read(filePath);
            }

        }
    }
}

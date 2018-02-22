using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.IO;
using System.Reflection;

namespace Tools_micro
{
    public partial class Result : Form
    {
        private readonly string TemplaterFileName = Application.StartupPath + @"\Звіт.doc";
        public Result()
        {
            InitializeComponent();
        }

        private void Result_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void button1_Click(object sender, EventArgs e)
        {
            SaveToDoc();
        }
        private void SaveToDoc()
        {
            var wordApp = new Word.Application();
            try
            {
                wordApp.Visible = false;
                var wordDocument = wordApp.Documents.Open(TemplaterFileName);
                //5.12
                ReplaceWordStub("<f51>", label35.Text, wordDocument);
                ReplaceWordStub("<f52>", label37.Text, wordDocument);
                ReplaceWordStub("<f53>", label39.Text, wordDocument);
                ReplaceWordStub("<f54>", label34.Text, wordDocument);
                ReplaceWordStub("<f55>", label32.Text, wordDocument);
                ReplaceWordStub("<f56>", label36.Text, wordDocument);
                ReplaceWordStub("<f57>", label38.Text, wordDocument);
                ReplaceWordStub("<f58>", label40.Text, wordDocument);
                ReplaceWordStub("<f59>", label41.Text, wordDocument);
                ReplaceWordStub("<f60>", label42.Text, wordDocument);
                ReplaceWordStub("<f00>", label211.Text, wordDocument);

                //10.24
                ReplaceWordStub("<f61>", label50.Text, wordDocument);
                ReplaceWordStub("<f62>", label48.Text, wordDocument);
                ReplaceWordStub("<f63>", label46.Text, wordDocument);
                ReplaceWordStub("<f64>", label51.Text, wordDocument);
                ReplaceWordStub("<f65>", label53.Text, wordDocument);
                ReplaceWordStub("<f66>", label49.Text, wordDocument);
                ReplaceWordStub("<f67>", label47.Text, wordDocument);
                ReplaceWordStub("<f68>", label45.Text, wordDocument);
                ReplaceWordStub("<f69>", label44.Text, wordDocument);
                ReplaceWordStub("<f70>", label43.Text, wordDocument);
                ReplaceWordStub("<f01>", label212.Text, wordDocument);

                //15.36
                ReplaceWordStub("<f71>", label92.Text, wordDocument);
                ReplaceWordStub("<f72>", label90.Text, wordDocument);
                ReplaceWordStub("<f73>", label88.Text, wordDocument);
                ReplaceWordStub("<f74>", label93.Text, wordDocument);
                ReplaceWordStub("<f75>", label95.Text, wordDocument);
                ReplaceWordStub("<f76>", label91.Text, wordDocument);
                ReplaceWordStub("<f77>", label89.Text, wordDocument);
                ReplaceWordStub("<f78>", label87.Text, wordDocument);
                ReplaceWordStub("<f79>", label86.Text, wordDocument);
                ReplaceWordStub("<f80>", label85.Text, wordDocument);
                ReplaceWordStub("<f02>", label213.Text, wordDocument);

                //21.5
                ReplaceWordStub("<f81>", label134.Text, wordDocument);
                ReplaceWordStub("<f82>", label132.Text, wordDocument);
                ReplaceWordStub("<f83>", label130.Text, wordDocument);
                ReplaceWordStub("<f84>", label135.Text, wordDocument);
                ReplaceWordStub("<f85>", label137.Text, wordDocument);
                ReplaceWordStub("<f86>", label133.Text, wordDocument);
                ReplaceWordStub("<f87>", label131.Text, wordDocument);
                ReplaceWordStub("<f88>", label129.Text, wordDocument);
                ReplaceWordStub("<f89>", label128.Text, wordDocument);
                ReplaceWordStub("<f90>", label127.Text, wordDocument);
                ReplaceWordStub("<f03>", label214.Text, wordDocument);


                //25
                ReplaceWordStub("<f91>", label176.Text, wordDocument);
                ReplaceWordStub("<f92>", label174.Text, wordDocument);
                ReplaceWordStub("<f93>", label172.Text, wordDocument);
                ReplaceWordStub("<f94>", label177.Text, wordDocument);
                ReplaceWordStub("<f95>", label179.Text, wordDocument);
                ReplaceWordStub("<f96>", label175.Text, wordDocument);
                ReplaceWordStub("<f97>", label173.Text, wordDocument);
                ReplaceWordStub("<f98>", label171.Text, wordDocument);
                ReplaceWordStub("<f99>", label170.Text, wordDocument);
                ReplaceWordStub("<f100>",label169.Text, wordDocument);
                ReplaceWordStub("<f04>", label215.Text, wordDocument);


                wordDocument.Save();
               // wordDocument.Close();
                wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Не підтримується встановлена версія Microsoft Word!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            finally
            {
                //wordApp.Quit();
            }
        }

        private void ReplaceWordStub(string stubToReplace, string text, Word.Document wordDocument)
        {
            var range = wordDocument.Content;
            range.Find.ClearFormatting();
            range.Find.Execute(FindText: stubToReplace, ReplaceWith: text);
        }

    }
}

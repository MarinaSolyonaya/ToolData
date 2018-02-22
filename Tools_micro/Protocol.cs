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
    public partial class Protocol : Form
    {
        private readonly string TemplaterFileName = Application.StartupPath + @"\Shablon.doc";
        public Protocol()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            EnterData form = new EnterData();
            SaveToDoc();
            form.ShowDialog();
        }

        private void SaveToDoc()
        {
            var wordApp = new Word.Application();
            try
            {
                wordApp.Visible = false;
                var wordDocument = wordApp.Documents.Open(TemplaterFileName);
                ReplaceWordStub("<numberP>", textBox1.Text, wordDocument);
                ReplaceWordStub("<dateP>", maskedTextBox1.Text, wordDocument);
                ReplaceWordStub("<f1>", textBox2.Text, wordDocument);
                ReplaceWordStub("<f2>", textBox3.Text, wordDocument);
                ReplaceWordStub("<f3>", textBox4.Text, wordDocument);
                ReplaceWordStub("<f4>", textBox5.Text, wordDocument);
                ReplaceWordStub("<f5>", textBox6.Text, wordDocument);
                ReplaceWordStub("<f6>", textBox7.Text, wordDocument);
                ReplaceWordStub("<f7>", textBox8.Text, wordDocument);
                ReplaceWordStub("<f8>", textBox9.Text, wordDocument);
                ReplaceWordStub("<f9>", textBox10.Text, wordDocument);
                ReplaceWordStub("<f10>", textBox12.Text, wordDocument);
                ReplaceWordStub("<f11>", textBox11.Text, wordDocument);

                wordDocument.SaveAs(Application.StartupPath + @"\Звіт.doc");
                wordDocument.Close();
                //wordApp.Visible = true;
            }
            catch
            {
                MessageBox.Show("Не підтримується встановлена версія Microsoft Word!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            finally
            {
                wordApp.Quit();
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

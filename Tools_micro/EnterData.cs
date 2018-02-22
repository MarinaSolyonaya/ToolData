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
    public partial class EnterData : Form
    {
        double[,] data = new double[7, 5];
        double delT, Ls, tetaM, alpha, q;
        private readonly string TemplaterFileName = Application.StartupPath + @"\Звіт.doc";
        public EnterData()
        {
            InitializeComponent();
            textBox5.TextChanged += textBox32_TextChanged;
            textBox12.TextChanged += textBox32_TextChanged;
            textBox17.TextChanged += textBox32_TextChanged;
            textBox22.TextChanged += textBox32_TextChanged;
            textBox27.TextChanged += textBox32_TextChanged;
            textBox8.TextChanged += textBox33_TextChanged;
            textBox13.TextChanged += textBox33_TextChanged;
            textBox18.TextChanged += textBox33_TextChanged;
            textBox23.TextChanged += textBox33_TextChanged;
            textBox28.TextChanged += textBox33_TextChanged;
            textBox9.TextChanged += textBox34_TextChanged;
            textBox14.TextChanged += textBox34_TextChanged;
            textBox19.TextChanged += textBox34_TextChanged;
            textBox24.TextChanged += textBox34_TextChanged;
            textBox29.TextChanged += textBox34_TextChanged;
            textBox10.TextChanged += textBox35_TextChanged;
            textBox15.TextChanged += textBox35_TextChanged;
            textBox20.TextChanged += textBox35_TextChanged;
            textBox25.TextChanged += textBox35_TextChanged;
            textBox30.TextChanged += textBox35_TextChanged;
            textBox11.TextChanged += textBox36_TextChanged;
            textBox16.TextChanged += textBox36_TextChanged;
            textBox21.TextChanged += textBox36_TextChanged;
            textBox26.TextChanged += textBox36_TextChanged;
            textBox31.TextChanged += textBox36_TextChanged;
        }

        private void textBox36_TextChanged(object sender, EventArgs e)
        {
            double avg = 0.0;
            try
            {
                avg = (Convert.ToDouble(textBox11.Text.ToString()) + Convert.ToDouble(textBox16.Text.ToString()) + Convert.ToDouble(textBox21.Text.ToString()) + Convert.ToDouble(textBox26.Text.ToString()) + Convert.ToDouble(textBox31.Text.ToString())) / 5.0;
            }
            catch { }
            avg = Math.Round(avg, 4);
            textBox36.Text = avg.ToString();
        }

        private void textBox35_TextChanged(object sender, EventArgs e)
        {
            double avg = 0.0;
            try
            {
                avg = (Convert.ToDouble(textBox10.Text.ToString()) + Convert.ToDouble(textBox15.Text.ToString()) + Convert.ToDouble(textBox20.Text.ToString()) + Convert.ToDouble(textBox25.Text.ToString()) + Convert.ToDouble(textBox30.Text.ToString())) / 5.0;
            }
            catch { }
            avg = Math.Round(avg, 4);
            textBox35.Text = avg.ToString();
        }

        private void textBox34_TextChanged(object sender, EventArgs e)
        {
            double avg = 0.0;
            try
            {
                avg = (Convert.ToDouble(textBox9.Text.ToString()) + Convert.ToDouble(textBox14.Text.ToString()) + Convert.ToDouble(textBox19.Text.ToString()) + Convert.ToDouble(textBox24.Text.ToString()) + Convert.ToDouble(textBox29.Text.ToString())) / 5.0;
            }
            catch { }
            avg = Math.Round(avg, 4);
            textBox34.Text = avg.ToString();
        }

        private void textBox33_TextChanged(object sender, EventArgs e)
        {
            double avg = 0.0;
            try
            {
                avg = (Convert.ToDouble(textBox8.Text.ToString()) + Convert.ToDouble(textBox13.Text.ToString()) + Convert.ToDouble(textBox18.Text.ToString()) + Convert.ToDouble(textBox23.Text.ToString()) + Convert.ToDouble(textBox28.Text.ToString())) / 5.0;
            }
            catch { }
            avg = Math.Round(avg, 4);
            textBox33.Text = avg.ToString();
        }

        private void textBox32_TextChanged(object sender, EventArgs e)
        {
            double avg = 0.0;
            try
            {
                avg = (Convert.ToDouble(textBox5.Text.ToString()) + Convert.ToDouble(textBox12.Text.ToString()) + Convert.ToDouble(textBox17.Text.ToString()) + Convert.ToDouble(textBox22.Text.ToString()) + Convert.ToDouble(textBox27.Text.ToString())) / 5.0;
            }
            catch { }
            avg = Math.Round(avg, 4);
            textBox32.Text = avg.ToString();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            GetDate();
            SaveToDoc();
            Calculate();
        }

        private void SaveToDoc()
        {
            var wordApp = new Word.Application();
            try
            {
                wordApp.Visible = false;
                var wordDocument = wordApp.Documents.Open(TemplaterFileName);

                ReplaceWordStub("<f12>", textBox7.Text, wordDocument);
                ReplaceWordStub("<f13>", textBox2.Text, wordDocument);
                ReplaceWordStub("<f14>", textBox1.Text, wordDocument);
                ReplaceWordStub("<f15>", textBox6.Text, wordDocument);
                ReplaceWordStub("<f17>", textBox3.Text, wordDocument);
                ReplaceWordStub("<f18>", textBox4.Text, wordDocument);
                ReplaceWordStub("<f19>", textBox39.Text, wordDocument);
                ReplaceWordStub("<f20>", textBox40.Text, wordDocument);


                ReplaceWordStub("<f21>", data[0, 0].ToString(), wordDocument);
                ReplaceWordStub("<f22>", data[0, 1].ToString(), wordDocument);
                ReplaceWordStub("<f23>", data[0, 2].ToString(), wordDocument);
                ReplaceWordStub("<f24>", data[0, 3].ToString(), wordDocument);
                ReplaceWordStub("<f25>", data[0, 4].ToString(), wordDocument);
                ReplaceWordStub("<f26>", data[1, 0].ToString(), wordDocument);
                ReplaceWordStub("<f27>", data[1, 1].ToString(), wordDocument);
                ReplaceWordStub("<f28>", data[1, 2].ToString(), wordDocument);
                ReplaceWordStub("<f29>", data[1, 3].ToString(), wordDocument);
                ReplaceWordStub("<f30>", data[1, 4].ToString(), wordDocument);
                ReplaceWordStub("<f31>", data[2, 0].ToString(), wordDocument);
                ReplaceWordStub("<f32>", data[2, 1].ToString(), wordDocument);
                ReplaceWordStub("<f33>", data[2, 2].ToString(), wordDocument);
                ReplaceWordStub("<f34>", data[2, 3].ToString(), wordDocument);
                ReplaceWordStub("<f35>", data[2, 4].ToString(), wordDocument);
                ReplaceWordStub("<f36>", data[3, 0].ToString(), wordDocument);
                ReplaceWordStub("<f37>", data[3, 1].ToString(), wordDocument);
                ReplaceWordStub("<f38>", data[3, 2].ToString(), wordDocument);
                ReplaceWordStub("<f39>", data[3, 3].ToString(), wordDocument);
                ReplaceWordStub("<f40>", data[3, 4].ToString(), wordDocument);
                ReplaceWordStub("<f41>", data[4, 0].ToString(), wordDocument);
                ReplaceWordStub("<f42>", data[4, 1].ToString(), wordDocument);
                ReplaceWordStub("<f43>", data[4, 2].ToString(), wordDocument);
                ReplaceWordStub("<f44>", data[4, 3].ToString(), wordDocument);
                ReplaceWordStub("<f45>", data[4, 4].ToString(), wordDocument);
                ReplaceWordStub("<f46>", data[5, 0].ToString(), wordDocument);
                ReplaceWordStub("<f47>", data[5, 1].ToString(), wordDocument);
                ReplaceWordStub("<f48>", data[5, 2].ToString(), wordDocument);
                ReplaceWordStub("<f49>", data[5, 3].ToString(), wordDocument);
                ReplaceWordStub("<f50>", data[5, 4].ToString(), wordDocument);

                wordDocument.Save();
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

        private void Calculate()
        {
            Result form = new Result();
            double ua = 0.0, ub=0.0, uc=0.0;
            //5.12
            for (int i = 0; i < 5; i++)
            {
                ua += Math.Pow(data[i, 0] - data[5, 0],4.0); 
            }
            ua = Math.Sqrt(ua / 20.0) * 1.4;
            form.label35.Text = "5,12";
            form.label37.Text = data[6, 0].ToString();
            form.label36.Text = Math.Round(ua, 4).ToString();
            form.label38.Text = form.label37.Text;
            form.label39.Text = Math.Round(5.12 * alpha * delT,4).ToString();
            form.label40.Text = Math.Round((5.12 * alpha * delT) / Math.Sqrt(3), 4).ToString();
            form.label34.Text = tetaM.ToString();
            form.label41.Text = Math.Round(tetaM / Math.Sqrt(3),4).ToString();
            form.label32.Text = q.ToString();
            form.label42.Text = Math.Round(q/ Math.Sqrt(3), 4).ToString();
            ub = Math.Sqrt(Math.Pow(data[6, 0], 2) + (Ls * alpha * delT) / Math.Sqrt(3)) + tetaM / Math.Sqrt(3) + q / Math.Sqrt(3);
            uc = Math.Round(Math.Sqrt(Math.Pow(ua,2)+Math.Pow(ub,2)),4);
            uc *= 2 * 0.001;
            uc = Math.Round(uc, 4);
            form.label211.Text += " 5,12 ± " + uc.ToString() + " мм";

            //10.24
            for (int i = 0; i < 5; i++)
            {
                ua += Math.Pow(data[i, 1] - data[5, 1], 2.0);
            }
            ua = Math.Sqrt(ua / 20.0) * 1.4;
            form.label50.Text = "10,24";
            form.label48.Text = data[6, 1].ToString();
            form.label49.Text = Math.Round(ua, 4).ToString();
            form.label47.Text = form.label48.Text;
            form.label46.Text = Math.Round((10.24 * alpha * delT), 4).ToString();
            form.label45.Text = Math.Round((10.24 * alpha * delT) / Math.Sqrt(3), 4).ToString();
            form.label51.Text = tetaM.ToString();
            form.label44.Text = Math.Round(tetaM / Math.Sqrt(3), 4).ToString();
            form.label53.Text = q.ToString();
            form.label43.Text = Math.Round(q / Math.Sqrt(3), 4).ToString();
            ub = Math.Sqrt(Math.Pow(data[6, 1], 2) + (Ls * alpha * delT) / Math.Sqrt(3)) + tetaM / Math.Sqrt(3) + q / Math.Sqrt(3);
            uc = Math.Round(Math.Sqrt(Math.Pow(ua, 2) + Math.Pow(ub, 2)), 4);
            uc *= 2 * 0.001;
            uc = Math.Round(uc, 4);
            form.label212.Text += " 10,24 ± " + uc.ToString() + " мм";

            //15.36
            for (int i = 0; i < 5; i++)
            {
                ua += Math.Pow(data[i, 2] - data[5, 2], 2.0);
            }
            ua = Math.Sqrt(ua / 20.0) * 1.4;
            form.label92.Text = "15,36"; 
            form.label90.Text = data[6, 2].ToString();
            form.label91.Text = Math.Round(ua, 4).ToString();
            form.label89.Text = form.label90.Text;
            form.label88.Text = Math.Round((15.36 * alpha * delT), 4).ToString();
            form.label87.Text = Math.Round((15.36 * alpha * delT) / Math.Sqrt(3), 4).ToString();
            form.label93.Text = tetaM.ToString();
            form.label86.Text = Math.Round(tetaM / Math.Sqrt(3), 4).ToString();
            form.label95.Text = q.ToString();
            form.label85.Text = Math.Round(q / Math.Sqrt(3), 4).ToString();
            ub = Math.Sqrt(Math.Pow(data[6, 2], 2) + (Ls * alpha * delT) / Math.Sqrt(3)) + tetaM / Math.Sqrt(3) + q / Math.Sqrt(3);
            uc = Math.Round(Math.Sqrt(Math.Pow(ua, 2) + Math.Pow(ub, 2)), 4);
            uc *= 2 * 0.001;
            uc = Math.Round(uc, 4);
            form.label213.Text += " 15,36 ± " + uc.ToString() + " мм";

            //21.5
            for (int i = 0; i < 5; i++)
            {
                ua += Math.Pow(data[i, 3] - data[5, 3], 2.0);
            }
            ua = Math.Sqrt(ua / 20.0) * 1.4;
            form.label134.Text = "21,5";
            form.label132.Text = data[6, 3].ToString();
            form.label133.Text = Math.Round(ua, 4).ToString();
            form.label131.Text = form.label132.Text;
            form.label130.Text = Math.Round((21.5 * alpha * delT), 4).ToString();
            form.label129.Text = Math.Round((21.5 * alpha * delT) / Math.Sqrt(3), 4).ToString();
            form.label135.Text = tetaM.ToString();
            form.label128.Text = Math.Round(tetaM / Math.Sqrt(3), 4).ToString();
            form.label137.Text = q.ToString();
            form.label127.Text = Math.Round(q / Math.Sqrt(3), 4).ToString();
            ub = Math.Sqrt(Math.Pow(data[6, 3], 2) + (Ls * alpha * delT) / Math.Sqrt(3)) + tetaM / Math.Sqrt(3) + q / Math.Sqrt(3);
            uc = Math.Round(Math.Sqrt(Math.Pow(ua, 2) + Math.Pow(ub, 2)), 4);
            uc *= 2 * 0.001;
            uc = Math.Round(uc, 4);
            form.label214.Text += " 21,5 ± " + uc.ToString() + " мм";

            //25
            for (int i = 0; i < 5; i++)
            {
                ua += Math.Pow(data[i, 4] - data[5, 4], 2.0);
            }
            ua = Math.Sqrt(ua / 20.0) * 1.4;
            form.label176.Text = "25";
            form.label174.Text = data[6, 4].ToString();
            form.label175.Text = Math.Round(ua, 4).ToString();
            form.label173.Text = form.label174.Text;
            form.label172.Text = Math.Round((25 * alpha * delT), 4).ToString();
            form.label171.Text = Math.Round((25 * alpha * delT) / Math.Sqrt(3), 4).ToString();
            form.label177.Text = tetaM.ToString();
            form.label170.Text = Math.Round(tetaM / Math.Sqrt(3), 4).ToString();
            form.label179.Text = q.ToString();
            form.label169.Text = Math.Round(q / Math.Sqrt(3), 4).ToString();
            ub = Math.Sqrt(Math.Pow(data[6, 4], 2) + (Ls * alpha * delT) / Math.Sqrt(3)) + tetaM / Math.Sqrt(3) + q / Math.Sqrt(3);
            uc = Math.Round(Math.Sqrt(Math.Pow(ua, 2) + Math.Pow(ub, 2)), 4);
            uc *= 2*0.001;
            uc = Math.Round(uc, 4);
            form.label215.Text += " 25 ± " + uc.ToString() + " мм";
            //this.Hide();
            form.ShowDialog();
        }

        private void GetDate()
        {
            try
            {
                delT = Math.Abs(Convert.ToDouble(textBox1.Text) - Convert.ToDouble(textBox6.Text));
                alpha = Convert.ToDouble(textBox38.Text);
                tetaM = Math.Round(Math.Sqrt(Math.Pow(Convert.ToDouble(textBox3.Text),2) + Math.Pow(Convert.ToDouble(textBox4.Text),2) + Math.Pow(Convert.ToDouble(textBox39.Text),2)), 4);
                q = Convert.ToDouble(textBox40.Text);
                data[0, 0] = Convert.ToDouble(textBox5.Text);
                data[0, 1] = Convert.ToDouble(textBox8.Text);
                data[0, 2] = Convert.ToDouble(textBox9.Text);
                data[0, 3] = Convert.ToDouble(textBox10.Text);
                data[0, 4] = Convert.ToDouble(textBox11.Text);
                data[1, 0] = Convert.ToDouble(textBox12.Text);
                data[1, 1] = Convert.ToDouble(textBox13.Text);
                data[1, 2] = Convert.ToDouble(textBox14.Text);
                data[1, 3] = Convert.ToDouble(textBox15.Text);
                data[1, 4] = Convert.ToDouble(textBox16.Text);
                data[2, 0] = Convert.ToDouble(textBox17.Text);
                data[2, 1] = Convert.ToDouble(textBox18.Text);
                data[2, 2] = Convert.ToDouble(textBox19.Text);
                data[2, 3] = Convert.ToDouble(textBox20.Text);
                data[2, 4] = Convert.ToDouble(textBox21.Text);
                data[3, 0] = Convert.ToDouble(textBox22.Text);
                data[3, 1] = Convert.ToDouble(textBox23.Text);
                data[3, 2] = Convert.ToDouble(textBox24.Text);
                data[3, 3] = Convert.ToDouble(textBox25.Text);
                data[3, 4] = Convert.ToDouble(textBox26.Text);
                data[4, 0] = Convert.ToDouble(textBox27.Text);
                data[4, 1] = Convert.ToDouble(textBox28.Text);
                data[4, 2] = Convert.ToDouble(textBox29.Text);
                data[4, 3] = Convert.ToDouble(textBox30.Text);
                data[4, 4] = Convert.ToDouble(textBox31.Text);
                data[5, 0] = Convert.ToDouble(textBox32.Text);
                data[5, 1] = Convert.ToDouble(textBox33.Text);
                data[5, 2] = Convert.ToDouble(textBox34.Text);
                data[5, 3] = Convert.ToDouble(textBox35.Text);
                data[5, 4] = Convert.ToDouble(textBox36.Text);
                data[6, 0] = Convert.ToDouble(textBox41.Text);
                data[6, 1] = Convert.ToDouble(textBox42.Text);
                data[6, 2] = Convert.ToDouble(textBox43.Text);
                data[6, 3] = Convert.ToDouble(textBox44.Text);
                data[6, 4] = Convert.ToDouble(textBox45.Text);
            }
            catch
            {
                MessageBox.Show("Невірний формат даних!","Помилка!",MessageBoxButtons.OK,MessageBoxIcon.Error);
            }
         
        }
    }
}

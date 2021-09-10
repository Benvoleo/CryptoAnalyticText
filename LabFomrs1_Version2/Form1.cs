using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Office.Interop.Word;


namespace LabFomrs1_Version2
{

    public partial class Form1 : Form
    {
        

        public Form1()
        {
            InitializeComponent();

            string inputedStr = "";
            try
            {
                using (WordprocessingDocument document = WordprocessingDocument.Open("refer.docx",true))
                {
                    inputedStr = document.MainDocumentPart.Document.Body.InnerText;
                    textBox1.Text += inputedStr.ToLower();


                }
            }
            catch { }

            string rez = inputedStr.ToLower();

            
            rez = rez.Replace(" ", string.Empty);
            rez = rez.Replace(".", string.Empty);
            rez = rez.Replace("?", string.Empty);
            rez = rez.Replace(",", string.Empty);
            rez = rez.Replace("-", string.Empty);
            rez = rez.Replace("=", string.Empty);
            rez = rez.Replace("+", string.Empty);
            rez = rez.Replace("`", string.Empty);
            rez = rez.Replace("—", string.Empty);
            rez = rez.Replace("\n", string.Empty);
            rez = rez.Replace("(", string.Empty);
            rez = rez.Replace(")", string.Empty);
            rez = rez.Replace("/", string.Empty);
            rez = rez.Replace(@"\", string.Empty);
            rez = rez.Replace("'", string.Empty);
            rez = rez.Replace("\"", string.Empty);
            rez = rez.Replace(@"&", string.Empty);
            rez = rez.Replace(";", string.Empty);
            rez = rez.Replace(":", string.Empty);
            rez = rez.Replace("/", string.Empty);

            textBox2.Text += rez;




            Dictionary<string, int> countOfCoupe = new Dictionary<string, int>();
            Dictionary<char, int> countOfSingle = new Dictionary<char, int>();

            List<char> charStr = rez.ToCharArray().ToList();
            int len = charStr.Count;


            //ЗАДАНИЕ 1
            textBox3.Text += "\r\nЗадание 1:\r\n\r\n";



            string path = @"C:\Users\Tibalt\Desktop\Криптология прикладная 3.5\laba1\LabFomrs1_Version3\LabFomrs1_Version2\Task1.txt";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                string createText = "\t\t###################################" + Environment.NewLine;
                File.WriteAllText(path, createText);
            }


            for (int i = 0; i < len; i++)
            {
                if (!countOfSingle.Keys.Contains(charStr[i]))
                {
                    countOfSingle.Add(charStr[i], 1);
                }
                else
                {
                    countOfSingle[charStr[i]] = countOfSingle[charStr[i]] + 1;
                }
            }



            double sum = 0;
            foreach (var elem in countOfSingle)
            {
                double probably = Convert.ToDouble(elem.Value) * 100 / len;
                sum += probably;
                textBox3.Text += $" \t({elem.Key}) встречался {elem.Value} раз. Вероятность: {probably.ToString("0.00")}%\r\n";
                
                File.AppendAllText(path, $" \t|({elem.Key})| |{probably.ToString("0.00")}|%\r\n");

            }

            
            //ЗАДАНИЕ 2
            textBox4.Text += "Задание 2\r\n\r\n";

            path = @"C:\Users\Tibalt\Desktop\Криптология прикладная 3.5\laba1\LabFomrs1_Version3\LabFomrs1_Version2\Task2.txt";
            if (!File.Exists(path))
            {
                // Create a file to write to.
                string createText = "\t\t###################################" + Environment.NewLine;
                File.WriteAllText(path, createText);
            }



            for (int i = 0; i < len - 1; i++)
            {
                string coupe = "" + charStr[i] + charStr[i + 1];
                if (!countOfCoupe.Keys.Contains(coupe))
                {
                    countOfCoupe.Add(coupe, 1);
                }
                else
                {
                    countOfCoupe[coupe] = countOfCoupe[coupe] + 1;
                }
            }

            foreach (var elem in countOfCoupe)
            {
                textBox4.Text += $"Пара: {elem.Key} = {elem.Value} повторений\r\n";
                File.AppendAllText(path, $"Пара: {elem.Key} = {elem.Value} повторений\r\n");
            }



           

        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

         void button1_Click(object sender, EventArgs e)
        {
            string path = @"C:\Users\Tibalt\Desktop\Криптология прикладная 3.5\laba1\LabFomrs1_Version3\LabFomrs1_Version2\jojo.txt";

            // This text is added only once to the file.
            if (!File.Exists(path))
            {
                // Create a file to write to.
                string createText = textBox1 + Environment.NewLine + textBox2 + Environment.NewLine + textBox3 + Environment.NewLine + textBox4 + Environment.NewLine;
                File.WriteAllText(path, createText);
            }
            else
            {
                string createText = textBox1 + Environment.NewLine + textBox2 + Environment.NewLine + textBox3 + Environment.NewLine + textBox4 + Environment.NewLine;
                File.WriteAllText(path, createText);
            }
            

        }
        
    }

    

}


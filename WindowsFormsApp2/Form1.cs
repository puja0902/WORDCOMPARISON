using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            {
                DialogResult firstDoc = openFileDialog1.ShowDialog();
                DialogResult secondDoc = openFileDialog2.ShowDialog();
                if (firstDoc == DialogResult.OK && secondDoc == DialogResult.OK)
                {
                    //string filePath1 = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("bin")) + "Files\\FirstLection.docx";
                    //string filePath2 = Environment.CurrentDirectory.Substring(0, Environment.CurrentDirectory.IndexOf("bin")) + "Files\\SecondLection.docx";
                    string filePath1 = openFileDialog1.FileName;
                    string filePath2 = openFileDialog2.FileName;
                    CompareWordFile(filePath1, filePath2);
                }
            }
        }

        public void CompareWordFile(string fileToCompare, string fileToChange)
        {

            Microsoft.Office.Interop.Word.Application word1 = new Microsoft.Office.Interop.Word.Application();
            Microsoft.Office.Interop.Word.Application word2 = new Microsoft.Office.Interop.Word.Application();
            Document doc1 = new Document();
            Document doc2 = new Document();
            object missing = System.Type.Missing;
            object fileName1 = fileToCompare;
            object fileName2 = fileToChange;
            doc1 = word1.Documents.Open(ref fileName1, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);
            doc2 = word2.Documents.Open(ref fileName2, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing);

            string[] firstFile = doc1.Paragraphs[1].Range.Text.Split('.');
            for (int i = 0; i < firstFile.Length - 1; i++)
            {
                string item = firstFile[i];
                foreach (Microsoft.Office.Interop.Word.Range docRange in doc2.Sentences)
                {
                    if (docRange.Text.Split('.')[0].Trim().Equals(item.Trim(), StringComparison.CurrentCultureIgnoreCase))
                    {
                        docRange.Font.ColorIndex = Microsoft.Office.Interop.Word.WdColorIndex.wdDarkRed;
                    }
                }
            }
            doc2.Save();
            ((_Application)word1).Quit();
            ((_Application)word2).Quit();
        }
    }
    }


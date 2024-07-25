using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;

namespace WinFormsApp2
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {

        }

        private async void button1_Click(object sender, EventArgs e)
        {

            OpenFileDialog ofd = new OpenFileDialog();
            FolderBrowserDialog dialog = new FolderBrowserDialog();
            ofd.ShowDialog();
            dialog.ShowDialog();



            string sourcePath = ofd.FileName; // переменная которя харнит схронемую папку дляя последующиго копривания
            string targetPath = dialog.SelectedPath;// переменная которя харнит путь к папке дляя последующиго копривания
            string test = sourcePath.Insert(sourcePath., "SDERG.CSV");
            //if ()
            string newfile = ofd.FileName;              
            Random random = new Random(19);
            //{
            textBox1.Text = ofd.FileName;
            textBox2 .Text = dialog.SelectedPath; ;
            //}
            IWorkbook workbook;
            Console.WriteLine("Введите путь к файлу");
            sourcePath = sourcePath.Trim();
            sourcePath = sourcePath.Trim(new char[] { '"' }); // результат "ello worl"
            using (FileStream fileStream = new FileStream(sourcePath, FileMode.Open, FileAccess.Read))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            string filePaths2 = "C:\\Users\\user\\Desktop\\asdasdsadasa.txt";
            // Получение листа
            ISheet sheet = workbook.GetSheetAt(0);
            int a = sheet.LastRowNum;
            string[] strings = new string[a + 1];
            StringBuilder csv = new StringBuilder();
            using (StreamWriter writer = new StreamWriter(filePaths2, true))
            {
                await writer.WriteAsync("1,КИЗ,");
                //File.AppendAllText(filePaths2, csv.ToString(), Encoding.Default);
                //await writer.WriteLineAsync("Addition");
                //await writer.WriteAsync("4,5");
            }
            for (int j = 1; j <= a; j++)
            {


                // Чтение данных из ячейки
                IRow row = sheet.GetRow(j);
                string cellValue = row.GetCell(0).StringCellValue;
                strings[j] = cellValue;
                using (StreamWriter writer = new StreamWriter(filePaths2, true))
                {
                    await writer.WriteAsync(strings[j] + "," + '"');
                }
            }
            FileStream fs = new FileStream(filePaths2, FileMode.Open, FileAccess.Read);
            StreamReader sr = new(fs);
            string t;
            t = sr.ReadToEnd();
            sr.Close();
            Console.WriteLine("Укажите куда сохранить ");
          
            System.IO.StreamWriter sw = new System.IO.StreamWriter(test, false, System.Text.Encoding.Unicode);
            sw.Write(t);
            File.Delete(filePaths2);
            sw.Close();
            void asd(int n)
            {


                csv.Append(strings[n] + ", " + '"');
                File.AppendAllText(filePaths2, csv.ToString(), Encoding.Default);

                ///strings[i] = cellValue;
                // Вывод данных ячейки
            }

        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void process1_Exited(object sender, EventArgs e)
        {

        }
    }
}

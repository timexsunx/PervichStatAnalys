using System;
using System.Data;
using System.Linq;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.IO;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Windows.Forms.DataVisualization.Charting;
using System.Diagnostics;

namespace PervichStatAnalys
{
    public partial class Form1 : Form
    {

        bool viborRyada = true;//Переменная для определения Дискретный или Интервальный ряд
        string[] peremRanga2 = new string[] { "ha","hb","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
        "aa","ab","ac","ad","ae","af","ag","ah","ai","aj","ak","al","am","an","ao","ap","aq","ar","as","at","au","av","aw","ax","ay","az",
        "ba","bb","bc","bd","be","bf","bg","bh","bi","bj","bk","bl","bm","bn","bo","bp","bq","br","bs","bt","bu","bv","bw","bx","by","bz",
        "ca","cb","cc","cd","ce","cf","cg","ch","ci","cj","ck","cl","cm","cn","co","cp","cq","cr","cs","ct","cu","cv","cw","cx","cy","cz",
        "da","db","dc","dd","de","df","dg","dh","di","dj","dk","dl","dm","dn","do","dp","dq","dr","ds","dt","du","dv","dw","dx","dy","dz"};
        public Form1()
        {
            InitializeComponent();
            panel9.Visible = false;
            dataGridView1.Visible = false;
            pictureBox6.Visible = false;
            pictureBox7.Visible = false;
            label3.Text = "Сейчас выбран: Дискретный ряд";

        }
        int Error;

        private void VibratFail()
        {
            dataGridView1.Rows.Clear();
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "CSV файл (*.csv)|*.csv";
            ofd.FileName = "";
            ofd.Title = "Открыть";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    using (var Reader = new StreamReader(ofd.FileName))
                    {
                        string stroka;
                        int j = 0;
                        Error = 0;
                        while ((stroka = Reader.ReadLine()) != null)
                        {
                            String[] array = stroka.Split(new char[] { ';' });
                            dataGridView1.RowCount = j + 2;
                            dataGridView1.ColumnCount = array.Length;
                            for (int i = 0; i < array.Length; i++)
                            {
                                try
                                {
                                    if (array[i] != "")
                                    {
                                        dataGridView1.Rows[j].Cells[i].Value = Convert.ToDouble(array[i]);
                                    }
                                }
                                catch
                                {
                                    Error++;
                                }
                            }
                            j++;
                        }
                        dataGridView1.RowCount--;
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                if (Error > 0)
                {
                    DialogResult result = MessageBox.Show("В " + Error + " Ячейке(ах) есть ошибки.\nВы хотите их исправить?", "Сообщение",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No)
                    {
                        VibratFail();
                        return;
                    }
                    if (result == DialogResult.Yes)
                    {
                        Process.Start(ofd.FileName);
                        return;
                    }
                }
                //заполнение массивва данными из datagridview1
                string[,] matrix1 = new string[dataGridView1.RowCount, dataGridView1.ColumnCount];
                for (int j = 0; j < matrix1.GetLength(0); j++)
                {
                    for (int i = 0; i < matrix1.GetLength(1); i++)
                    {
                        matrix1[j, i] = Convert.ToString(dataGridView1.Rows[j].Cells[i].Value);

                    }
                }
                //Конвертация двумерного массива в одномерный
                string[] matrix2 = new string[dataGridView1.RowCount * dataGridView1.ColumnCount];
                int z = 0;
                for (int j = 0; j < matrix1.GetLength(0); j++)
                {
                    for (int i = 0; i < matrix1.GetLength(1); i++)
                    {
                        if (matrix1[j, i] != "")
                        {
                            matrix2[z] = Convert.ToString(matrix1[j, i]);
                            z++;
                        }
                    }
                }
                matrix2 = matrix2.Where(x => x != null).ToArray();//Удаление пустых ячеек из массива
                //Изменение типа массива для сортировки по числам
                double[] matrix5 = new double[matrix2.Length];
                for (int i = 0; i < matrix2.Length; i++)
                {
                    matrix5[i] = Convert.ToDouble(matrix2[i]);
                }
                Array.Sort(matrix5);//Сортировка массива от мала до велика
                //Ряд не может состоять из одного значения и меньше
                if (matrix5.Length < 2)
                {
                    MessageBox.Show("Недопустимая длина ряда", "Сообщение",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }
                //Изменение типа массива для критерия Ирвина
                string[] matrix4 = new string[matrix5.Length];
                for (int i = 0; i < matrix5.Length; i++)
                {
                    matrix4[i] = Convert.ToString(matrix5[i]);
                }
                //Критерий Ирвина Среднее значение
                double sredZnach = 0;
                double summaSredZnach = 0;
                for (int i = 0; i < matrix5.Length; i++)
                {
                    summaSredZnach += matrix5[i];
                }
                sredZnach = summaSredZnach / matrix5.Length;

                double peremDispp = 0;
                double dispersiyaa = 0;
                for (int i = 0; i < matrix5.Length; i++)
                {
                    peremDispp += Math.Pow(matrix5[i] - sredZnach, 2);
                }
                dispersiyaa = peremDispp / matrix5.Length;
                //Среднее квадратическоеое отклонение
                double sredKvadratichh = 0;
                sredKvadratichh = Math.Sqrt(dispersiyaa);

                double kriteriyIrvina = 0;
                int indexL = 1;
                int indexM = 0;
                bool estilinet = false;
                bool estilida = false;
                int promah = 0;
                do
                {
                    kriteriyIrvina = (matrix5[matrix5.Length - indexL] - matrix5[matrix5.Length - indexL - 1]) / sredKvadratichh;
                    if (matrix5.Length == 2 && kriteriyIrvina > 2.8) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length == 3 && kriteriyIrvina > 2.2) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 3 && matrix5.Length <= 10 && kriteriyIrvina > 1.5) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 10 && matrix5.Length <= 20 && kriteriyIrvina > 1.3) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 20 && matrix5.Length <= 30 && kriteriyIrvina > 1.2) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 30 && matrix5.Length <= 50 && kriteriyIrvina > 1.1) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 50 && matrix5.Length <= 100 && kriteriyIrvina > 1.0) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 100 && matrix5.Length <= 400 && kriteriyIrvina > 0.9) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 400 && matrix5.Length <= 1000 && kriteriyIrvina > 0.8) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    if (matrix5.Length > 1000 && matrix5.Length <= 10000 && kriteriyIrvina > 0.7) { matrix4[matrix5.Length - indexL] = null; estilinet = true; promah++; }
                    else { estilinet = false; }

                    summaSredZnach = 0;
                    for (int i = 0; i < matrix5.Length - indexL - 1; i++)
                    {
                        summaSredZnach += matrix5[i];
                    }
                    sredZnach = summaSredZnach / (matrix5.Length - indexL - 1);
                    peremDispp = 0;
                    for (int i = 0; i < matrix5.Length - indexL - 1; i++)
                    {
                        peremDispp += Math.Pow(matrix5[i] - sredZnach, 2);
                    }
                    dispersiyaa = peremDispp / (matrix5.Length - indexL - 1);
                    sredKvadratichh = Math.Sqrt(dispersiyaa);

                    indexL++;
                }
                while (estilinet == true && matrix5.Length - indexL - 1 >= 0);
                do
                {
                    kriteriyIrvina = (matrix5[indexM + 1] - matrix5[indexM]) / sredKvadratichh;
                    if (matrix5.Length == 2 && kriteriyIrvina > 2.8) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length == 3 && kriteriyIrvina > 2.2) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 3 && matrix5.Length <= 10 && kriteriyIrvina > 1.5) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 10 && matrix5.Length <= 20 && kriteriyIrvina > 1.3) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 20 && matrix5.Length <= 30 && kriteriyIrvina > 1.2) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 30 && matrix5.Length <= 50 && kriteriyIrvina > 1.1) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 50 && matrix5.Length <= 100 && kriteriyIrvina > 1.0) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 100 && matrix5.Length <= 400 && kriteriyIrvina > 0.9) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 400 && matrix5.Length <= 1000 && kriteriyIrvina > 0.8) { matrix4[indexM] = null; estilida = true; promah++; }
                    if (matrix5.Length > 1000 && matrix5.Length <= 10000 && kriteriyIrvina > 0.7) { matrix4[indexM] = null; estilida = true; promah++; }
                    else { estilinet = false; }

                    summaSredZnach = 0;
                    for (int i = 0; i < matrix5.Length - indexL - 1; i++)
                    {
                        summaSredZnach += matrix5[i];
                    }
                    sredZnach = summaSredZnach / (matrix5.Length - indexL - 1);
                    peremDispp = 0;
                    for (int i = 0; i < matrix5.Length - indexL - 1; i++)
                    {
                        peremDispp += Math.Pow(matrix5[i] - sredZnach, 2);
                    }
                    dispersiyaa = peremDispp / (matrix5.Length - indexL - 1);
                    sredKvadratichh = Math.Sqrt(dispersiyaa);
                    indexL++;
                    indexM++;
                }
                while (estilida == true && indexM + 1 < matrix5.Length);
                if (promah > 0)
                {
                    DialogResult result = MessageBox.Show("В " + promah + " Ячейке(ах) есть промахи.\nПромахи будут удалены.\nПродолжить?", "Сообщение",
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (result == DialogResult.No)
                    {
                        VibratFail();
                    }
                }
                matrix4 = matrix4.Where(x => x != null).ToArray();//Удаление пустых ячеек из массива
                //Изменение типа массива
                double[] matrix3 = new double[matrix4.Length];
                for (int i = 0; i < matrix4.Length; i++)
                {
                    matrix3[i] = Convert.ToDouble(matrix4[i]);
                }
                //Вычисления дискретного ряда
                if (viborRyada == true)
                {
                    // Группировка и подсчет частоты
                    int a = 0;
                    double sum = 0;
                    double sumchastost = 0;
                    dataGridView3.Rows.Clear();
                    var g = matrix3.GroupBy(i => i);
                    foreach (var k in g)
                    {
                        dataGridView3.Rows.Add();
                        dataGridView3.Rows[a].Cells[0].Value = k.Key; // число
                        dataGridView3.Rows[a].Cells[1].Value = k.Count(); // частота
                        sum += k.Count();
                        dataGridView3.Rows[a].Cells[2].Value = sum; // Накопленная частота
                        dataGridView3.Rows[a].Cells[3].Value = Math.Round(k.Count() / Convert.ToDouble(matrix3.Length) * 100, 2); // Частота
                        sumchastost += k.Count() / Convert.ToDouble(matrix3.Length) * 100;
                        dataGridView3.Rows[a].Cells[4].Value = Math.Round(sumchastost, 2); // Накопленная частота
                        a++;
                    }
                    dataGridView3.Rows.RemoveAt(dataGridView3.Rows.Count - 1);
                    // Построение графика для дискретного ряда
                    chart1.Series[0].Points.Clear();
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        double x = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                        double y = Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value);
                        chart1.Series[0].Points.AddXY(x, y);
                    }
                    // Расчеты для дискретного ряда
                    dataGridView4.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dataGridView4.RowCount = 12;
                    //Средняя величина
                    double sredVelich = 0;
                    double peremSumDiskr = 0;
                    for (int i = 0; i < matrix3.Length; i++)
                    {
                        peremSumDiskr += matrix3[i];
                    }
                    sredVelich = peremSumDiskr / matrix3.Length;
                    dataGridView4.Rows[0].Cells[0].Value = Convert.ToString("Средняя арифметическая взвешенная: ") + Math.Round(Convert.ToDouble(sredVelich), 2);
                    //мода
                    double max = Convert.ToDouble(dataGridView3.Rows[0].Cells[1].Value);
                    double moda = 0;
                    int nalichiye_modi = 0;
                    for (int i = 1; i < dataGridView3.Rows.Count; i++)
                    {
                        if (max < Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value))
                        {
                            max = Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value);
                        }
                    }
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        if (max == Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value))
                        {
                            moda = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                            nalichiye_modi++;
                        }
                    }
                    if (nalichiye_modi == 1)
                    {
                        dataGridView4.Rows[1].Cells[0].Value = Convert.ToString("Мода: ") + Math.Round(moda, 2);
                    }
                    else
                    {
                        dataGridView4.Rows[1].Cells[0].Value = Convert.ToString("Моды нет ");
                    }
                    //Медиана
                    double peremMediana = 0;
                    if (matrix3.Length % 2 == 0)
                    {
                        for (int i = 0; i < matrix3.Length; i++)
                        {
                            if (matrix3.Length / 2 < Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value))
                            {
                                peremMediana = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                                break;
                            }
                        }
                        dataGridView4.Rows[2].Cells[0].Value = Convert.ToString("Медиана: ") + peremMediana;
                    }
                    else
                    {
                        for (int i = 0; i < matrix3.Length; i++)
                        {
                            if ((matrix3.Length + 1) / 2 < Convert.ToDouble(dataGridView3.Rows[i].Cells[2].Value))
                            {
                                peremMediana = Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value);
                                break;
                            }
                        }
                        dataGridView4.Rows[2].Cells[0].Value = Convert.ToString("Медиана: ") + peremMediana;
                    }
                    //Размах вариации
                    dataGridView4.Rows[3].Cells[0].Value = Convert.ToString("Размах вариации: ") + Math.Round(matrix3[matrix3.Length - 1] - matrix3[0], 2);
                    //Среднее линейное отклонение
                    double lineyn = 0;
                    for (int i = 0; i < dataGridView3.Rows.Count; i++)
                    {
                        lineyn += Math.Abs(Convert.ToDouble(dataGridView3.Rows[i].Cells[0].Value) - Math.Round(sredVelich, 2))
                            * Convert.ToDouble(dataGridView3.Rows[i].Cells[1].Value);
                    }
                    dataGridView4.Rows[4].Cells[0].Value = Convert.ToString("Среднее линейное отклонение: ") + Math.Round(lineyn / matrix3.Length, 2);
                    //Дисперсия
                    double peremDisp = 0;
                    double dispersiya = 0;
                    for (int i = 0; i < matrix3.Length; i++)
                    {
                        peremDisp += Math.Pow(matrix3[i] - sredVelich, 2);
                    }
                    dispersiya = peremDisp / matrix3.Length;
                    dataGridView4.Rows[5].Cells[0].Value = Convert.ToString("Дисперсия: ") + Math.Round(dispersiya, 2);
                    //Среднее квадратическоеое отклонение
                    double sredKvadratich = 0;
                    sredKvadratich = Math.Sqrt(dispersiya);
                    dataGridView4.Rows[6].Cells[0].Value = Convert.ToString("Среднее квадратичное отклонение: ") + Math.Round(sredKvadratich, 2);
                    //Коэффициент вариации
                    dataGridView4.Rows[7].Cells[0].Value = Convert.ToString("Коэффициент вариации: ") + Math.Round(sredKvadratich / sredVelich * 100, 2)
                        + Convert.ToString("%");
                    double moment = 0;
                    double normMom = 0;
                    for (int i = 0; i < matrix3.Length; i++)
                    {
                        moment += Math.Pow(matrix3[i] - sredVelich, 3);
                    }
                    moment = moment / matrix3.Length;
                    normMom = moment / Math.Pow(sredKvadratich, 3);
                    //Асимметрия
                    if (normMom == 0)
                    {
                        dataGridView4.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ") + Math.Round(normMom, 2)
                            + Convert.ToString(" - симметричное распределение ряда ");
                    }
                    if (normMom < 0)
                    {
                        dataGridView4.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ")
                            + Math.Round(normMom, 2) + Convert.ToString(" - левосторонняя асимметрия ");
                    }
                    if (normMom > 0)
                    {
                        dataGridView4.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ") + Math.Round(normMom, 2)
                            + Convert.ToString(" - правосторонняя асимметрия ");
                    }
                    //Степень существенности асимметрии
                    double stepSush;
                    double sravnen;
                    double N = dataGridView3.Rows.Count;
                    stepSush = Math.Sqrt(6 * (N - 2) / ((N + 1) * (N + 3)));
                    sravnen = Math.Abs(normMom) / stepSush;
                    if (sravnen > 3)
                    {
                        dataGridView4.Rows[9].Cells[0].Value = Convert.ToString("Оценка существенности асимметрии: ")
                            + Convert.ToString("существенная асимметрия");
                    }
                    if (sravnen < 3)
                    {
                        dataGridView4.Rows[9].Cells[0].Value = Convert.ToString("Оценка существенности асимметрии: ")
                            + Convert.ToString("несущественная асимметрия");
                    }
                    //Эксцесс
                    double peremMoment = 0;
                    double ekscess = 0;
                    for (int i = 0; i < matrix3.Length; i++)
                    {
                        peremMoment += Math.Pow(matrix3[i] - sredVelich, 4);
                    }
                    peremMoment = peremMoment / matrix3.Length;
                    ekscess = Math.Round(peremMoment / Math.Pow(sredKvadratich, 4) - 3, 2);
                    if (ekscess == 0)
                    {
                        dataGridView4.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - нормальное распределение");
                    }
                    if (ekscess > 0)
                    {
                        dataGridView4.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - островершинное распределение");
                    }
                    if (ekscess < 0)
                    {
                        dataGridView4.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - плосковершинное распределение");
                    }
                    //Оценка существенности эксцесса распределения
                    double sushestEksces = 0;
                    double sravnen2 = 0;
                    sushestEksces = Math.Sqrt(24 * N * (N - 2) * (N - 3) / (Math.Pow(N + 1, 2) * (N + 3) * (N + 5)));
                    sravnen2 = Math.Abs(ekscess) / sushestEksces;
                    if (sravnen2 > 3)
                    {
                        dataGridView4.Rows[11].Cells[0].Value = Convert.ToString("Cущественность эксцесса распределения: ")
                            + Convert.ToString("отклонение существенно");
                    }
                    else
                    {
                        dataGridView4.Rows[11].Cells[0].Value = Convert.ToString("Cущественность эксцесса распределения: ")
                            + Convert.ToString("отклонение несущественно");
                    }
                    dataGridView3.ClearSelection();
                    dataGridView4.ClearSelection();
                }
                //Вычисления интервального ряда
                if (viborRyada == false)
                {
                    double c1 = 0;
                    double h = 0;
                    c1 = Math.Floor(1 + 3.322 * Math.Log10(matrix3.Length));//Число интервалов, округление вниз до целого
                    h = Math.Round((Convert.ToDouble(matrix3[matrix3.Length - 1]) - Convert.ToDouble(matrix3[0])) / c1, 1, MidpointRounding.AwayFromZero);//Шаг ряда, округление вверх до целого

                    //Левая и правая границы
                    int f = 0;
                    int chisloshagov = 0;
                    double perem1 = matrix3[0];
                    while (chisloshagov < c1)
                    {
                        dataGridView2.RowCount = f + 2;
                        dataGridView2.Rows[f].Cells[0].Value = perem1;//Левая
                        dataGridView2.Rows[f].Cells[1].Value = perem1 + h;//Правая
                        perem1 += h;
                        chisloshagov++;
                        f++;
                    }
                    dataGridView2.RowCount--;
                    //Нахождение частот интервального ряда
                    int a = 0;
                    int chastota = 0;
                    int nakoplenChastota = 0;
                    double nakoplenChastost = 0;
                    int shag = 0;
                    double perem2 = matrix3[0];
                    while (shag < c1)
                    {
                        for (int i = 0; i < matrix3.Length; i++)
                        {
                            if (perem2 <= matrix3[i] && matrix3[i] < perem2 + h || shag == c1 - 1 && matrix3[i] == matrix3[matrix3.Length - 1])
                            {
                                chastota++;
                                nakoplenChastota++;
                            }
                        }
                        dataGridView2.Rows[a].Cells[2].Value = chastota;
                        dataGridView2.Rows[a].Cells[3].Value = nakoplenChastota;
                        dataGridView2.Rows[a].Cells[4].Value = Math.Round(chastota / Convert.ToDouble(matrix3.Length) * 100, 2);
                        nakoplenChastost += chastota / Convert.ToDouble(matrix3.Length) * 100;
                        dataGridView2.Rows[a].Cells[5].Value = Math.Round(nakoplenChastost, 2);
                        chastota = 0;
                        perem2 += h;
                        a++;
                        shag++;
                    }
                    //Построение графика для интервального ряда
                    chart2.Series[0].Points.Clear();
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        string x = Convert.ToString(dataGridView2.Rows[i].Cells[0].Value) + " - " + Convert.ToString(dataGridView2.Rows[i].Cells[1].Value);
                        double y = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                        chart2.Series[0].Points.AddXY(x, y);
                    }
                    //Расчеты для интервального ряда
                    dataGridView5.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dataGridView5.RowCount = 12;
                    //Средняя величина
                    double sredVelich = 0;
                    double peremSumInter = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        peremSumInter += (Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value))
                            / 2 * Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    }
                    sredVelich = peremSumInter / matrix3.Length;
                    dataGridView5.Rows[0].Cells[0].Value = Convert.ToString("Средняя арифметическая взвешенная: ") + Math.Round(Convert.ToDouble(sredVelich), 2);
                    //мода
                    double max = Convert.ToDouble(dataGridView2.Rows[0].Cells[2].Value);
                    double moda = 0;
                    double predModChastota = 0;
                    double mposleModChastota = 0;
                    double maxChastota = 0;
                    double nachaloModInter = 0;
                    int nalichiye_modi = 0;
                    for (int i = 1; i < dataGridView2.Rows.Count; i++)
                    {
                        if (max < Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value))
                        {
                            max = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                        }
                    }
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        if (max == Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value))
                        {
                            nachaloModInter = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                            maxChastota = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                            if (max == Convert.ToDouble(dataGridView2.Rows[0].Cells[2].Value))
                            {
                                predModChastota = 0;
                            }
                            else
                            {
                                predModChastota = Convert.ToDouble(dataGridView2.Rows[i - 1].Cells[2].Value);
                            }
                            if (max == Convert.ToDouble(dataGridView2.Rows[dataGridView2.Rows.Count - 1].Cells[2].Value))
                            {
                                mposleModChastota = 0;
                            }
                            else
                            {
                                mposleModChastota = Convert.ToDouble(dataGridView2.Rows[i + 1].Cells[2].Value);
                            }
                            nalichiye_modi++;
                        }
                    }
                    if (nalichiye_modi == 1)
                    {
                        moda = nachaloModInter + h * ((maxChastota - predModChastota) / ((maxChastota - predModChastota) + (maxChastota - mposleModChastota)));
                        dataGridView5.Rows[1].Cells[0].Value = Convert.ToString("Мода: ") + Math.Round(moda, 2);
                    }
                    else
                    {
                        dataGridView5.Rows[1].Cells[0].Value = Convert.ToString("Моды нет ");
                    }
                    //Медиана
                    double maxCh = 0;
                    double peremMediana = 0;
                    double nachaloInter = 0;
                    double chastotaMedian = 0;
                    double nakopChastotDoInter = 0;
                    if (matrix3.Length % 2 == 0)
                    {
                        for (int i = 0; i < dataGridView2.Rows.Count; i++)
                        {
                            if (matrix3.Length / 2 < Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))
                            {
                                maxCh = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                                nachaloInter = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                                chastotaMedian = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                                if (maxCh == Convert.ToDouble(dataGridView2.Rows[0].Cells[3].Value))
                                {
                                    nakopChastotDoInter = 0;
                                }
                                else
                                {
                                    nakopChastotDoInter = Convert.ToDouble(dataGridView2.Rows[i - 1].Cells[3].Value);
                                }
                                peremMediana = nachaloInter + (h / chastotaMedian) * (matrix3.Length / 2 - nakopChastotDoInter);
                                break;
                            }
                        }
                        dataGridView5.Rows[2].Cells[0].Value = Convert.ToString("Медиана: ") + Math.Round(peremMediana, 2);
                    }
                    else
                    {
                        for (int i = 0; i < dataGridView2.Rows.Count; i++)
                        {
                            if ((matrix3.Length + 1) / 2 < Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value))
                            {
                                maxCh = Convert.ToDouble(dataGridView2.Rows[i].Cells[3].Value);
                                nachaloInter = Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value);
                                chastotaMedian = Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                                if (maxCh == Convert.ToDouble(dataGridView2.Rows[0].Cells[3].Value))
                                {
                                    nakopChastotDoInter = 0;
                                }
                                else
                                {
                                    nakopChastotDoInter = Convert.ToDouble(dataGridView2.Rows[i - 1].Cells[3].Value);
                                }
                                peremMediana = nachaloInter + (h / chastotaMedian) * (matrix3.Length / 2 - nakopChastotDoInter);
                                break;
                            }
                        }
                        dataGridView5.Rows[2].Cells[0].Value = Convert.ToString("Медиана: ") + Math.Round(peremMediana, 2);
                    }
                    //Размах вариации
                    dataGridView5.Rows[3].Cells[0].Value = Convert.ToString("Размах вариации: ") + Math.Round(matrix3[matrix3.Length - 1] - matrix3[0], 2);
                    //Среднее линейное отклонение
                    double lineyn = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        lineyn += Math.Abs((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value))
                            / 2 - Math.Round(sredVelich, 2))
                            * Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    }
                    dataGridView5.Rows[4].Cells[0].Value = Convert.ToString("Среднее линейное отклонение: ") + Math.Round(lineyn / matrix3.Length, 2);
                    //Дисперсия
                    double peremDisp = 0;
                    double dispersiya = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        peremDisp += Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value))
                            / 2 - Math.Round(sredVelich, 2), 2)
                            * Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    }
                    dispersiya = peremDisp / matrix3.Length;
                    dataGridView5.Rows[5].Cells[0].Value = Convert.ToString("Дисперсия: ") + Math.Round(dispersiya, 2);
                    //Среднее квадратическоеое отклонение
                    double sredKvadratich = 0;
                    sredKvadratich = Math.Sqrt(dispersiya);
                    dataGridView5.Rows[6].Cells[0].Value = Convert.ToString("Среднее квадратичное отклонение: ") + Math.Round(sredKvadratich, 2);
                    //Коэффициент вариации
                    dataGridView5.Rows[7].Cells[0].Value = Convert.ToString("Коэффициент вариации: ") + Math.Round(sredKvadratich / sredVelich * 100, 2)
                        + Convert.ToString("%");
                    double moment = 0;
                    double normMom = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        moment += Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value))
                            / 2 - sredVelich, 3)
                            * Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    }
                    moment = moment / matrix3.Length;
                    normMom = moment / Math.Pow(sredKvadratich, 3);
                    //Асимметрия
                    if (normMom == 0)
                    {
                        dataGridView5.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ") + Math.Round(normMom, 2)
                            + Convert.ToString(" - симметричное распределение ряда ");
                    }
                    if (normMom < 0)
                    {
                        dataGridView5.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ")
                            + Math.Round(normMom, 2) + Convert.ToString(" - левосторонняя асимметрия ");
                    }
                    if (normMom > 0)
                    {
                        dataGridView5.Rows[8].Cells[0].Value = Convert.ToString("Моментный коэффициент асимметрии: ") + Math.Round(normMom, 2)
                            + Convert.ToString(" - правосторонняя асимметрия ");
                    }
                    //Степень существенности асимметрии
                    double stepSush;
                    double sravnen;
                    double N = dataGridView2.Rows.Count;
                    stepSush = Math.Sqrt(6 * (N - 2) / ((N + 1) * (N + 3)));
                    sravnen = Math.Abs(normMom) / stepSush;
                    if (sravnen > 3)
                    {
                        dataGridView5.Rows[9].Cells[0].Value = Convert.ToString("Оценка существенности асимметрии: ")
                            + Convert.ToString("существенная асимметрия");
                    }
                    if (sravnen < 3)
                    {
                        dataGridView5.Rows[9].Cells[0].Value = Convert.ToString("Оценка существенности асимметрии: ")
                            + Convert.ToString("несущественная асимметрия");
                    }
                    //Эксцесс
                    double peremMoment = 0;
                    double ekscess = 0;
                    for (int i = 0; i < dataGridView2.Rows.Count; i++)
                    {
                        peremMoment += Math.Pow((Convert.ToDouble(dataGridView2.Rows[i].Cells[0].Value) + Convert.ToDouble(dataGridView2.Rows[i].Cells[1].Value))
                            / 2 - sredVelich, 4)
                            * Convert.ToDouble(dataGridView2.Rows[i].Cells[2].Value);
                    }
                    peremMoment = peremMoment / matrix3.Length;
                    ekscess = Math.Round(peremMoment / Math.Pow(sredKvadratich, 4) - 3, 2);
                    if (ekscess == 0)
                    {
                        dataGridView5.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - нормальное распределение");
                    }
                    if (ekscess > 0)
                    {
                        dataGridView5.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - островершинное распределение");
                    }
                    if (ekscess < 0)
                    {
                        dataGridView5.Rows[10].Cells[0].Value = Convert.ToString("Эксцесс: ") + ekscess
                            + Convert.ToString(" - плосковершинное распределение");
                    }

                    //Оценка существенности эксцесса распределения
                    double sushestEksces = 0;
                    double sravnen2 = 0;
                    sushestEksces = Math.Sqrt(24 * N * (N - 2) * (N - 3) / (Math.Pow(N + 1, 2) * (N + 3) * (N + 5)));
                    sravnen2 = Math.Abs(ekscess) / sushestEksces;
                    if (sravnen2 > 3)
                    {
                        dataGridView5.Rows[11].Cells[0].Value = Convert.ToString("Cущественность эксцесса распределения: ")
                            + Convert.ToString("отклонение от нормального распределения существенно");
                    }
                    else
                    {
                        dataGridView5.Rows[11].Cells[0].Value = Convert.ToString("Cущественность эксцесса распределения: ")
                            + Convert.ToString("отклонение от нормального распределения несущественно");
                    }
                    dataGridView2.ClearSelection();
                    dataGridView5.ClearSelection();
                }
            }
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox5.Visible == true && dataGridView4.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            if (pictureBox5.Visible == true && dataGridView4.RowCount == 0)
            {
                VibratFail();
            }
            if (pictureBox7.Visible == true && dataGridView5.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            if (pictureBox7.Visible == true && dataGridView5.RowCount == 0)
            {
                VibratFail();
            }
        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }
        private void дискретныйРядToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.pdf)|*.pdf|Документ(*.xlsx)|*.xlsx";
            if (dataGridView4.Rows.Count != 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            //Создаем объект записи пдф-документа в файл
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                            var table1 = new PdfPTable(dataGridView3.ColumnCount);
                            var table2 = new PdfPTable(dataGridView4.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для дискретного ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            //Заголовки
                            for (int j = 0; j < dataGridView3.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView3.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            string frst = "0";
                            string lastt = "0";
                            //Добавляем ячейки исходных данных
                            for (int j = 0; j <= dataGridView3.RowCount - 1; j++)
                            {

                                for (int k = 0; k <= dataGridView3.ColumnCount - 1; k++)
                                {
                                    if (k == 0)
                                    {
                                        frst = dataGridView3.Rows[j].Cells[k].Value.ToString();
                                    }
                                    if (k == dataGridView3.ColumnCount - 1)
                                    {
                                        lastt = dataGridView3.Rows[j].Cells[k].Value.ToString();
                                    }
                                    table1.AddCell(new Phrase(dataGridView3.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);
                            var par05 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par05);
                            //Добавляем ячейки расчетов
                            for (int j = 0; j <= dataGridView4.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView4.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView4.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart1.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Полигон:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            var par07 = new iTextSharp.text.Paragraph("Вывод: общая тенденция на рынке отрицательная, зарплаты снижаются", font);
                            doc.Add(par07);

                            doc.Close();

                        }
                        if (saveFileDialog1.FilterIndex.ToString() == "2")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView4.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView4[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView3.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView3[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = Convert.ToString(dataGridView3[0, i].Value);
                                wsh.Cells[2, h + 12] = dataGridView3[1, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlLine;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void интервальныйРядToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.pdf)|*.pdf|Документ(*.xlsx)|*.xlsx";
            if (dataGridView5.Rows.Count != 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            //Определение шрифта
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                            var table1 = new PdfPTable(dataGridView2.ColumnCount);
                            var table2 = new PdfPTable(dataGridView5.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для интервального ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            //Заголовки
                            for (int j = 0; j < dataGridView2.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView2.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            //Добавляем ячейки исходных данных
                            for (int j = 0; j <= dataGridView2.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView2.ColumnCount - 1; k++)
                                {
                                    table1.AddCell(new Phrase(dataGridView2.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);
                            var par05 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par05);

                            //Добавляем ячейки расчетов
                            for (int j = 0; j <= dataGridView5.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView5.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView5.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            //Добавляем таблицу в документ
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart2.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Гистограмма:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            var par07 = new iTextSharp.text.Paragraph("Вывод: общая тенденция на рынке отрицательная, зарплаты снижаются", font);
                            doc.Add(par07);
                            doc.Close();

                        }
                        if (saveFileDialog1.FilterIndex.ToString() == "2")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView5.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView5.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView5[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView2.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView2[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = "(" + Convert.ToString(dataGridView2[0, i].Value) + ") - (" + Convert.ToString(dataGridView2[1, i].Value) + ")";
                                wsh.Cells[2, h + 12] = dataGridView2[2, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void pictureBox1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.pdf)|*.pdf";
            if (dataGridView4.Rows.Count != 0 && pictureBox5.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                            var table1 = new PdfPTable(dataGridView3.ColumnCount);
                            var table2 = new PdfPTable(dataGridView4.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для дискретного ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            for (int j = 0; j < dataGridView3.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView3.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            for (int j = 0; j <= dataGridView3.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView3.ColumnCount - 1; k++)
                                {
                                    table1.AddCell(new Phrase(dataGridView3.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);

                            //Добавляем ячейки расчетов
                            for (int j = 0; j <= dataGridView4.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView4.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView4.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart1.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Полигон:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            var par07 = new iTextSharp.text.Paragraph("Вывод: зарплаты держаться меньше среднего значения", font);
                            doc.Add(par07);
                            doc.Close();

                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            if (dataGridView5.Rows.Count != 0 && pictureBox7.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);
                            var table1 = new PdfPTable(dataGridView2.ColumnCount);
                            var table2 = new PdfPTable(dataGridView5.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для интервального ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            for (int j = 0; j < dataGridView2.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView2.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            for (int j = 0; j <= dataGridView2.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView2.ColumnCount - 1; k++)
                                {
                                    table1.AddCell(new Phrase(dataGridView2.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);
                            var par05 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par05);
                            for (int j = 0; j <= dataGridView5.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView5.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView5.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            //Добавляем таблицу в документ
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart2.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Гистограмма:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            doc.Close();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else if (dataGridView4.Rows.Count == 0 && pictureBox5.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dataGridView5.Rows.Count == 0 && pictureBox7.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void pictureBox2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.xlsx)|*.xlsx";
            if (dataGridView4.Rows.Count != 0 && pictureBox5.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView4.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView4[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView3.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView3[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = Convert.ToString(dataGridView3[0, i].Value);
                                wsh.Cells[2, h + 12] = dataGridView3[1, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlLine;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            if (dataGridView5.Rows.Count != 0 && pictureBox7.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView5.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView5.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView5[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView2.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView2[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = "(" + Convert.ToString(dataGridView2[0, i].Value) + ") - (" + Convert.ToString(dataGridView2[1, i].Value) + ")";
                                wsh.Cells[2, h + 12] = dataGridView2[2, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else if (dataGridView4.Rows.Count == 0 && pictureBox5.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dataGridView5.Rows.Count == 0 && pictureBox7.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
        private void pictureBox6_Click(object sender, EventArgs e)
        {
            viborRyada = true;
            panel10.Visible = true;
            panel9.Visible = false;

            pictureBox5.Visible = true;
            pictureBox6.Visible = false;

            pictureBox7.Visible = false;
            pictureBox8.Visible = true;
        }
        private void pictureBox8_Click(object sender, EventArgs e)
        {
            viborRyada = false;
            panel9.Visible = true;
            panel10.Visible = false;

            pictureBox5.Visible = false;
            pictureBox6.Visible = true;

            pictureBox7.Visible = true;
            pictureBox8.Visible = false;
        }
        private void pictureBox4_Click(object sender, EventArgs e)
        {
            if (pictureBox5.Visible == true && dataGridView4.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            if (pictureBox5.Visible == true && dataGridView4.RowCount == 0)
            {
                VibratFail();
            }
            if (pictureBox7.Visible == true && dataGridView5.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            if (pictureBox7.Visible == true && dataGridView5.RowCount == 0)
            {
                VibratFail();
            }
        }
        private void справкаToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            FormSpravka1 spravka1 = new FormSpravka1();
            spravka1.Show();
        }
        private void Form1_Load(object sender, EventArgs e) { }
        private void загрузитьДанныеToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (pictureBox5.Visible == true && dataGridView4.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            else if (pictureBox5.Visible == true && dataGridView4.RowCount == 0)
            {
                VibratFail();
            }
            else if (pictureBox7.Visible == true && dataGridView5.RowCount > 0)
            {
                DialogResult result = MessageBox.Show("Открыть новый файл?\nВсе расчеты будут утеряны.", "Сообщение",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (result == DialogResult.Yes)
                {
                    VibratFail();
                }
            }
            else if (pictureBox7.Visible == true && dataGridView5.RowCount == 0)
            {
                VibratFail();
            }
            else
            {
                VibratFail();
            }
        }
        private void дискретныйРядToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            label3.Text = "Сейчас выбран: Дискретный ряд";
            viborRyada = true;
            panel10.Visible = true;
            panel9.Visible = false;

            pictureBox5.Visible = true;
            pictureBox6.Visible = false;

            pictureBox7.Visible = false;
            pictureBox8.Visible = true;
        }

        private void интервальныйРядToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            label3.Text = "Сейчас выбран: Интервальный ряд";
            viborRyada = false;
            panel9.Visible = true;
            panel10.Visible = false;

            pictureBox5.Visible = false;
            pictureBox6.Visible = true;

            pictureBox7.Visible = true;
            pictureBox8.Visible = false;
        }

        private void справкаToolStripMenuItem_Click(object sender, EventArgs e)
        {
            FormSpravka1 spravka1 = new FormSpravka1();
            spravka1.Show();
        }

        private void dataGridView3_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView4_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void pDFToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.pdf)|*.pdf|Документ(*.xlsx)|*.xlsx";
            if (dataGridView4.Rows.Count != 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            //Создаем объект записи пдф-документа в файл
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                            var table1 = new PdfPTable(dataGridView3.ColumnCount);
                            var table2 = new PdfPTable(dataGridView4.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для дискретного ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            //Заголовки
                            for (int j = 0; j < dataGridView3.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView3.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            string frst = "0";
                            string lastt = "0";
                            //Добавляем ячейки исходных данных
                            for (int j = 0; j <= dataGridView3.RowCount - 1; j++)
                            {

                                for (int k = 0; k <= dataGridView3.ColumnCount - 1; k++)
                                {
                                    if (k == 0)
                                    {
                                        frst = dataGridView3.Rows[j].Cells[k].Value.ToString();
                                    }
                                    if (k == dataGridView3.ColumnCount - 1)
                                    {
                                        lastt = dataGridView3.Rows[j].Cells[k].Value.ToString();
                                    }
                                    table1.AddCell(new Phrase(dataGridView3.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);
                            var par05 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par05);
                            //Добавляем ячейки расчетов
                            for (int j = 0; j <= dataGridView4.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView4.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView4.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart1.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Полигон:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            var par07 = new iTextSharp.text.Paragraph("Вывод: общая тенденция на рынке отрицательная, зарплаты снижаются", font);
                            doc.Add(par07);

                            doc.Close();

                        }
                        if (saveFileDialog1.FilterIndex.ToString() == "2")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView4.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView4[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView3.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView3[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = Convert.ToString(dataGridView3[0, i].Value);
                                wsh.Cells[2, h + 12] = dataGridView3[1, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlLine;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void pDFToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.pdf)|*.pdf|Документ(*.xlsx)|*.xlsx";
            if (dataGridView5.Rows.Count != 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            iTextSharp.text.Document doc = new iTextSharp.text.Document();
                            PdfWriter.GetInstance(doc, new FileStream(filename, FileMode.Create));
                            doc.Open();
                            //Определение шрифта
                            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, iTextSharp.text.Font.DEFAULTSIZE, iTextSharp.text.Font.NORMAL);

                            var table1 = new PdfPTable(dataGridView2.ColumnCount);
                            var table2 = new PdfPTable(dataGridView5.ColumnCount);
                            var par0 = new iTextSharp.text.Paragraph("Первичный статистический анализ данных для интервального ряда", font);
                            doc.Add(par0);
                            var par01 = new iTextSharp.text.Paragraph("Исходные данные:", font);
                            doc.Add(par01);
                            var par03 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par03);
                            //Заголовки
                            for (int j = 0; j < dataGridView2.ColumnCount; j++)
                            {
                                var cell = new PdfPCell(new Phrase(new Phrase(dataGridView2.Columns[j].HeaderText.ToString(), font)));
                                {
                                    //Фоновый цвет 
                                    cell.BackgroundColor = iTextSharp.text.BaseColor.LIGHT_GRAY;
                                };
                                table1.AddCell(cell);
                            }
                            //Добавляем ячейки исходных данных
                            for (int j = 0; j <= dataGridView2.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView2.ColumnCount - 1; k++)
                                {
                                    table1.AddCell(new Phrase(dataGridView2.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            doc.Add(table1);
                            var par04 = new iTextSharp.text.Paragraph("Расчеты:", font);
                            doc.Add(par04);
                            var par05 = new iTextSharp.text.Paragraph(" ", font);
                            doc.Add(par05);

                            //Добавляем ячейки расчетов
                            for (int j = 0; j <= dataGridView5.RowCount - 1; j++)
                            {
                                for (int k = 0; k <= dataGridView5.ColumnCount - 1; k++)
                                {
                                    table2.AddCell(new Phrase(dataGridView5.Rows[j].Cells[k].Value.ToString(), font));
                                }
                            }
                            //Добавляем таблицу в документ
                            doc.Add(table2);
                            var chartimage = new MemoryStream();
                            chart2.SaveImage(chartimage, ChartImageFormat.Png);
                            iTextSharp.text.Image Chart_image = iTextSharp.text.Image.GetInstance(chartimage.GetBuffer());
                            Chart_image.ScalePercent(75f);
                            var par06 = new iTextSharp.text.Paragraph("Гистограмма:", font);
                            doc.Add(par06);
                            doc.Add(Chart_image);
                            var par07 = new iTextSharp.text.Paragraph("Вывод: общая тенденция на рынке отрицательная, зарплаты снижаются", font);
                            doc.Add(par07);
                            doc.Close();

                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void excelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.xlsx)|*.xlsx";
            if (dataGridView4.Rows.Count != 0 && pictureBox5.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView4.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView4.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView4[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView3.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView3.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView3[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView3.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = Convert.ToString(dataGridView3[0, i].Value);
                                wsh.Cells[2, h + 12] = dataGridView3[1, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlLine;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            if (dataGridView5.Rows.Count != 0 && pictureBox7.Visible == true)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        if (saveFileDialog1.FilterIndex.ToString() == "1")
                        {
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView5.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView5.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView5[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView2.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView2[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = "(" + Convert.ToString(dataGridView2[0, i].Value) + ") - (" + Convert.ToString(dataGridView2[1, i].Value) + ")";
                                wsh.Cells[2, h + 12] = dataGridView2[2, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        }
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else if (dataGridView4.Rows.Count == 0 && pictureBox5.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            else if (dataGridView5.Rows.Count == 0 && pictureBox7.Visible == true)
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void excelToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Документ(*.xlsx)|*.xlsx";
            if (dataGridView5.Rows.Count != 0)
            {
                if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                         
                            string filename = saveFileDialog1.FileName;
                            object misValue = System.Reflection.Missing.Value;
                            Excel.Application exApp = new Excel.Application();
                            Excel.Workbook workbook = exApp.Workbooks.Add(misValue);
                            Excel.Worksheet wsh = (Excel.Worksheet)exApp.ActiveSheet;
                            for (int i = 0; i <= dataGridView5.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView5.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[i + 1, j + 1] = dataGridView5[j, i].Value.ToString();
                                }
                            }
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                for (int j = 0; j <= dataGridView2.ColumnCount - 1; j++)
                                {
                                    wsh.Cells[1, j + 2] = dataGridView2.Columns[j].HeaderText;
                                    wsh.Cells[i + 2, j + 2] = dataGridView2[j, i].Value;
                                }
                            }
                            int h = 0;
                            for (int i = 0; i <= dataGridView2.RowCount - 1; i++)
                            {
                                wsh.Cells[1, h + 12] = "(" + Convert.ToString(dataGridView2[0, i].Value) + ") - (" + Convert.ToString(dataGridView2[1, i].Value) + ")";
                                wsh.Cells[2, h + 12] = dataGridView2[2, i].Value;
                                h++;
                            }
                            int peremRanga1 = dataGridView3.RowCount;
                            string peremRanga = peremRanga2[peremRanga1];
                            Excel.Range chartRange;
                            Excel.ChartObjects xlCharts = (Excel.ChartObjects)wsh.ChartObjects(Type.Missing);
                            Excel.ChartObject myChart = (Excel.ChartObject)xlCharts.Add(10, 30, 300, 300);
                            Excel.Chart chartPage = myChart.Chart;
                            chartRange = wsh.get_Range("k1", peremRanga + 2);
                            chartPage.SetSourceData(chartRange, misValue);
                            chartPage.ChartType = Excel.XlChartType.xlColumnClustered;
                            exApp.AlertBeforeOverwriting = false;
                            workbook.SaveAs(filename);
                            workbook.Close();
                            exApp.Quit();
                        
                        MessageBox.Show("Документ сохранен", "Сообщение",
                        MessageBoxButtons.OK);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Сообщение",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                }
            }
            else
            {
                MessageBox.Show("Нет данных для сохранения", "Сообщение",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }
    }
}

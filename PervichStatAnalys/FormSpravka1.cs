using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PervichStatAnalys
{
    partial class FormSpravka1 : Form
    {
        public FormSpravka1()
        {
            InitializeComponent();
            this.Text = String.Format("О программе");
            this.labelProductName.Text = "Продукт: Программа для проведения Первичного статистического анализа данных зарплат";
            this.labelVersion.Text = String.Format("Версия 1");
            this.labelCopyright.Text = "Команда проекта:\nРуководитель - Михайлова А.В.\nПрограммист и тестировщик - Штангель В.А \nПрограммист и тестировщик - Лонин И.И. ";
            this.labelCompanyName.Text = "Организация: ЯГТУ";
            this.textBoxDescription.Text += "Приложение предназначено для проведения первичного статистического анализа вариационных рядов, и всех расчетов на основе вводимых данных" +
                "\r\n Программа имеет следующие функции:" +
                "\r\n 1. Загрузка данных формата *.csv;" +
                "\r\n 2. Проверка исходных данных на корректность и промахи;" +
                "\r\n 3. Вычисление аналитических показателей для дискретных и интервальных рядов (частота, частость, среднее значение, мода," +
                " медиана, размах вариации, среднее линейное отклонение, дисперсия, среднее квадратичное отклонение, коэффициент вариации," +
                " нормированный моментный коэффициент асимметрии, оценка существенности асимметрии, эксцесс," +
                " средняя квадратическая ошибка эксцесса);" +
                "\r\n 4. Построение графиков: полигон - для дискретного ряда, гистограмма - для интервального ряда;" +
                "\r\n 5. Выгрузка полученных данных в формате *.xlsx, *.pdf.\r\n";
            this.textBoxDescription.Text += "\r\nМинимальные аппаратные требования:" +
                "\r\n·Процессор Intel, AMD совместимый, тактовая частота не ниже 1000 MHz;" +
                "\r\n·Объем свободной оперативной памяти - не менее 256 Мб;" +
                "\r\n·Не менее 500 МБ свободного дискового пространства;" +
                "\r\n·Мышь;" +
                "\r\n·Монитор с минимальным разрешением - 1280 × 1024 пикселей.";
        }

     

        private void labelVersion_Click(object sender, EventArgs e)
        {

        }

        private void textBoxDescription_TextChanged(object sender, EventArgs e)
        {

        }
    }
}

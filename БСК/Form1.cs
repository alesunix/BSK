using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using Image = System.Drawing.Image;

namespace БСК
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=C:\Program Files (x86)\Alesunix\BSC\components\Database1.mdf;Integrated Security=True");
        public Form1()
        {
            InitializeComponent();

            //--------------инициализировать comboBox------------//
            comboBox4.Items.AddRange(colorsNames);
            comboBox4.DrawItem += ComboBox4_DrawItem;
            comboBox4.DrawMode = DrawMode.OwnerDrawFixed;
            comboBox4.SelectedIndex = 0;
            comboBox4.DropDownStyle = ComboBoxStyle.DropDownList;

            comboBox1.Items.AddRange(vid);
            comboBox2.Items.AddRange(napolnitel);
            comboBox6.Items.AddRange(width);
        }
        string[] vid = File.ReadAllLines(@"C:\\Program Files (x86)\\Alesunix\\BSC\\components\\vid.txt");
        string[] napolnitel = File.ReadAllLines(@"C:\\Program Files (x86)\\Alesunix\\BSC\\components\\napolnitel.txt");
        string[] width = File.ReadAllLines(@"C:\\Program Files (x86)\\Alesunix\\BSC\\components\\width.txt");
        string[] colorsNames = { "RAL 1014", "RAL 1015", "RAL 1018", "RAL 3003", "RAL 3020", "RAL 5002", "RAL 5005", "RAL 5024", "RAL 6002", "RAL 6005", "RAL 7004", "RAL 7035", "RAL 8017", "RAL 9002", "RAL 9003", "RAL 9006", "RAL 9010" };
        Color[] colors = { Color.FromArgb(225, 204, 079), Color.FromArgb(230, 214, 144), Color.FromArgb(248, 243, 053), Color.FromArgb(155, 017, 030), Color.FromArgb(204, 006, 005), Color.FromArgb(032, 033, 079), Color.FromArgb(030, 045, 110), Color.FromArgb(093, 155, 155), Color.FromArgb(045, 087, 044), Color.FromArgb(047, 069, 056), Color.FromArgb(150, 153, 146), Color.FromArgb(215, 215, 215), Color.FromArgb(069, 050, 046), Color.FromArgb(231, 235, 218), Color.FromArgb(244, 244, 244), Color.FromArgb(165, 165, 165), Color.FromArgb(255, 255, 255) };
        void ComboBox4_DrawItem(object sender, DrawItemEventArgs e)//метод добавления цветов
        {
            using (Brush br = new SolidBrush(colors[e.Index]))
            {
                e.Graphics.FillRectangle(br, e.Bounds);
                e.Graphics.DrawString(colorsNames[e.Index], e.Font, Brushes.Black, e.Bounds);
            }
        }
        private void Form1_Load(object sender, EventArgs e)//Загрузка формы
        {
            //-----------------Окраска Гридов-------------------//
            DataGridViewRow row1 = this.dgvzakaz.RowTemplate;
            row1.DefaultCellStyle.BackColor = Color.FromArgb(227, 226, 221);//цвет строк
            row1.DefaultCellStyle.ForeColor = Color.FromArgb(33, 40, 47);//цвет текста
            row1.Height = 40;
            row1.MinimumHeight = 17;
            dgvzakaz.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            //dataGridView1.Columns[0].Width = 5;//Ширина столбца
            dgvzakaz.EnableHeadersVisualStyles = false;
            dgvzakaz.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(179, 96, 61);//цвет заголовка
            dgvzakaz.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке

            DataGridViewRow row2 = this.dgvstena.RowTemplate;
            row2.DefaultCellStyle.BackColor = Color.FromArgb(227, 226, 221);//цвет строк
            row2.DefaultCellStyle.ForeColor = Color.FromArgb(33, 40, 47);//цвет текста
            row2.Height = 40;
            row2.MinimumHeight = 17;
            dgvstena.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dgvstena.EnableHeadersVisualStyles = false;
            dgvstena.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(179, 96, 61);//цвет заголовка
            dgvstena.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке

            DataGridViewRow row3 = this.dgvresult.RowTemplate;
            row3.DefaultCellStyle.BackColor = Color.FromArgb(227, 226, 221);//цвет строк
            row3.DefaultCellStyle.ForeColor = Color.FromArgb(33, 40, 47);//цвет текста
            row3.Height = 40;
            row3.MinimumHeight = 17;
            dgvresult.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dgvresult.EnableHeadersVisualStyles = false;
            dgvresult.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 125, 146);//цвет заголовка
            dgvresult.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке

            DataGridViewRow row4 = this.dgvviev.RowTemplate;
            row4.DefaultCellStyle.BackColor = Color.FromArgb(227, 226, 221);//цвет строк
            row4.DefaultCellStyle.ForeColor = Color.FromArgb(33, 40, 47);//цвет текста
            row4.Height = 40;
            row4.MinimumHeight = 17;
            dgvviev.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dgvviev.EnableHeadersVisualStyles = false;
            dgvviev.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(228, 201, 156);//цвет заголовка
            dgvviev.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке

            DataGridViewRow row5 = this.dgvwindow.RowTemplate;
            row5.DefaultCellStyle.BackColor = Color.FromArgb(227, 226, 221);//цвет строк
            row5.DefaultCellStyle.ForeColor = Color.FromArgb(33, 40, 47);//цвет текста
            row5.Height = 40;
            row5.MinimumHeight = 17;
            dgvwindow.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;//автоподбор ширины столбца по содержимому
            dgvwindow.EnableHeadersVisualStyles = false;
            dgvwindow.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(0, 125, 146);//цвет заголовка
            dgvwindow.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;//Выравнивание текста в заголовке
            //----------------Окраска Гридов--------------------//   

            //-------------Отключить сортировку гридов----------------------//
            foreach (DataGridViewColumn column in dgvzakaz.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            foreach (DataGridViewColumn column in dgvstena.Columns)
            {
                column.SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            //-------------Отключить сортировку гридов----------------------//
            dgvzakaz.RowHeadersVisible = false;//Самая левая колонка
            dgvstena.RowHeadersVisible = false;//Самая левая колонка
            dgvviev.RowHeadersVisible = false;//Самая левая колонка
            dgvwindow.RowHeadersVisible = false;//Самая левая колонка
            dgvresult.RowHeadersVisible = false;//Самая левая колонка

            Select_zakaz();
            Select_stena();
            comboBox6.SelectedIndex = 1;
            comboBox1.SelectedIndex = 0;
            comboBox2.SelectedIndex = 0;
            comboBox3.SelectedIndex = 0;
            comboBox1.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox2.DropDownStyle = ComboBoxStyle.DropDownList;
            comboBox3.DropDownStyle = ComboBoxStyle.DropDownList;

            comboBox5.Items.Add("Окно");
            comboBox5.Items.Add("Дверь");
            comboBox5.Items.Add("Ворота");
            comboBox5.SelectedIndex = 0;
            Disp_data();
            Podschet();
        }
        private void Disp_data()
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT vid AS Вид,napolnitel AS Наполнитель,thickness AS Толщина,width AS Ширина,height AS Высота,color AS Цвет FROM [Table_wall] " +
                "WHERE wall NOT IN (N'0') AND zakaz=@zakaz AND wall=@wall ORDER BY wall DESC", con);
            cmd.Parameters.AddWithValue("@zakaz", label14.Text);
            cmd.Parameters.AddWithValue("@wall", label15.Text);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dgvviev.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение

            con.Open();//Открываем соединение
            SqlCommand cmd2 = new SqlCommand("SELECT MAX(wall) AS Стена,MAX(vid) AS Вид,MAX(thickness) AS Толщина,MAX(napolnitel) AS Наполнитель,MIN(kol_vo) AS 'Кол-во',MAX(length) AS Длина," +
                "MAX(color) AS Цвет FROM [Table_result] " +
                "WHERE kol_vo_window IN (N'0') AND zakaz=@zakaz AND wall = @wall GROUP BY wall ORDER BY wall ", con);
            cmd2.Parameters.AddWithValue("@zakaz", label14.Text);
            cmd2.Parameters.AddWithValue("@wall", label15.Text);
            cmd2.ExecuteNonQuery();
            DataTable dt2 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da2 = new SqlDataAdapter(cmd2);//создаем экземпляр класса SqlDataAdapter
            dt2.Clear();//чистим DataTable, если он был не пуст
            da2.Fill(dt2);//заполняем данными созданный DataTable
            dgvresult.DataSource = dt2;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение

            con.Open();//Открываем соединение
            SqlCommand cmd3 = new SqlCommand("SELECT MIN(wall) AS Стена,SUM(kol_vo_window) AS 'Кол-во',length_window AS Длина, MIN(note) AS Примечание FROM [Table_window] " +
                "WHERE kol_vo_window NOT IN (N'0') AND zakaz=@zakaz AND wall=@wall GROUP BY length_window ORDER BY length_window DESC", con);
            cmd3.Parameters.AddWithValue("@zakaz", label14.Text);
            cmd3.Parameters.AddWithValue("@wall", label15.Text);
            cmd3.ExecuteNonQuery();
            DataTable dt3 = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da3 = new SqlDataAdapter(cmd3);//создаем экземпляр класса SqlDataAdapter
            dt3.Clear();//чистим DataTable, если он был не пуст
            da3.Fill(dt3);//заполняем данными созданный DataTable
            dgvwindow.DataSource = dt3;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
        }
        private void Podschet()
        {
            if(dgvstena.Rows.Count >= 1 & dgvzakaz.Rows.Count >=2)
            {
                if (dgvviev.Rows[0].Cells[0].Value.ToString() == "Стена")
                {
                    comboBox6.SelectedIndex = 0;
                }
                else if (dgvviev.Rows[0].Cells[0].Value.ToString() == "Кровля")
                {
                    comboBox6.SelectedIndex = 1;
                }
                
                    //Количество панелей *Окна
                    double panel2 = 0;
                    foreach (DataGridViewRow row in dgvwindow.Rows)
                    {
                        double incom;
                        double.TryParse((row.Cells[1].Value ?? "0").ToString().Replace(".", ","), out incom);
                        panel2 += incom;
                    }
                    label23.Text = panel2.ToString();
        
                    //Подсчет количества строк (не учитывая пустые строки и колонки)
                        int count = 0;
                        for (int j = 0; j < dgvstena.RowCount; j++)
                        {
                            for (int i = 0; i < dgvstena.ColumnCount; i++)
                            {
                                if (dgvstena[i, j].Value != null)
                                {
                                    label19.Text = Convert.ToString(dgvstena.Rows.Count /*- 1*/) + " стен";// -1 это нижняя пустая строка
                                    count++;
                                    break;
                                }
                            }
                        }
            }
        }
        private void RESULT()
        {
            //RESULT
            double kol_vo = Math.Ceiling(Convert.ToDouble(dgvviev.Rows[0].Cells[3].Value) / Convert.ToDouble(comboBox6.Text));//Штук панелей
            double length = Math.Ceiling(Convert.ToDouble(dgvviev.Rows[0].Cells[4].Value) / 100);//Длина панелей в метрах           

            con.Open();//открыть соединение
            SqlCommand cmd = new SqlCommand("INSERT INTO [Table_result] (zakaz,wall,vid,thickness,napolnitel,kol_vo,length,kol_vo_window,length_window,color) VALUES (@zakaz,@wall,@vid,@thickness,@napolnitel,@kol_vo,@length,@kol_vo_window,@length_window,@color)", con);
            cmd.Parameters.AddWithValue("@zakaz", label14.Text);
            cmd.Parameters.AddWithValue("@wall", label15.Text);
            cmd.Parameters.AddWithValue("@vid", dgvviev.Rows[0].Cells[0].Value.ToString());
            cmd.Parameters.AddWithValue("@napolnitel", dgvviev.Rows[0].Cells[1].Value.ToString());
            cmd.Parameters.AddWithValue("@thickness", dgvviev.Rows[0].Cells[2].Value.ToString());            
            cmd.Parameters.AddWithValue("@kol_vo", kol_vo);
            cmd.Parameters.AddWithValue("@length", length);
            if (textBox5.Text != "")
            {
                cmd.Parameters.AddWithValue("@kol_vo_window", Math.Floor(Convert.ToDouble(textBox5.Text) / Convert.ToDouble(comboBox6.Text)));//Штук панелей где окна
            }
            else if (textBox5.Text == "")
            {
                cmd.Parameters.AddWithValue("@kol_vo_window", 0);
            }
            if (textBox6.Text != "")
            {
                cmd.Parameters.AddWithValue("@length_window", Math.Floor((Convert.ToDouble(dgvviev.Rows[0].Cells[4].Value) - Convert.ToDouble(textBox6.Text)) / 100));//Длина панелей в метрах где окна
            }
            else if (textBox6.Text == "")
            {
                cmd.Parameters.AddWithValue("@length_window", 0);
            }
            cmd.Parameters.AddWithValue("@color", dgvviev.Rows[0].Cells[5].Value.ToString());
            cmd.ExecuteNonQuery();
            con.Close();//закрыть соединение 

            textBox3.Text = "";
            textBox4.Text = "";
            textBox5.Text = "";
            textBox6.Text = "";
        }
        private void ComboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBox1.Text == "Стена")
            {
                comboBox6.SelectedIndex = 0;
            }
            else if (comboBox1.Text == "Кровля")
            {
                comboBox6.SelectedIndex = 1;
            }
        }
        private void ComboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            comboBox3.Items.Clear();
            if (comboBox2.Text == "Пенопласт")
            {
                comboBox3.Items.Add(new ClassComboBox(50, "50"));
                comboBox3.Items.Add(new ClassComboBox(80, "80"));
                comboBox3.Items.Add(new ClassComboBox(100, "100"));
                comboBox3.Items.Add(new ClassComboBox(120, "120"));
                comboBox3.Items.Add(new ClassComboBox(150, "150"));
                comboBox3.Items.Add(new ClassComboBox(170, "170"));
                comboBox3.Items.Add(new ClassComboBox(200, "200"));
            }
            else if (comboBox2.Text == "Базальт")
            {
                comboBox3.Items.Add(new ClassComboBox(50, "50"));
                comboBox3.Items.Add(new ClassComboBox(100, "100"));
                comboBox3.Items.Add(new ClassComboBox(120, "120"));
                comboBox3.Items.Add(new ClassComboBox(150, "150"));
                comboBox3.Items.Add(new ClassComboBox(170, "170"));
                comboBox3.Items.Add(new ClassComboBox(200, "200"));
            }
        }
        private void Select_zakaz()//Заказы
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT MIN(zakaz) AS 'Заказ №' FROM [Table_wall] " +
                "GROUP BY zakaz ORDER BY zakaz DESC", con);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dgvzakaz.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            //dgvzakaz.Columns["Время"].DefaultCellStyle.Format = "HH:mm:ss";
        }
        private void Select_stena()//Стены
        {
            con.Open();//Открываем соединение
            SqlCommand cmd = new SqlCommand("SELECT MIN(wall) AS 'Стена №' FROM [Table_wall] " +
                "WHERE wall NOT IN (N'0') AND zakaz=@zakaz GROUP BY wall ORDER BY wall DESC", con);
            cmd.Parameters.AddWithValue("@zakaz", label14.Text);
            cmd.ExecuteNonQuery();
            DataTable dt = new DataTable();//создаем экземпляр класса DataTable
            SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
            dt.Clear();//чистим DataTable, если он был не пуст
            da.Fill(dt);//заполняем данными созданный DataTable
            dgvstena.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
            con.Close();//Закрываем соединение
            //dgvzakaz.Columns["Время"].DefaultCellStyle.Format = "HH:mm:ss";
        }
        private void dgvzakaz_SelectionChanged(object sender, EventArgs e)//получить данные выделенной строки
        {
            if (dgvzakaz.Rows.Count >= 2)
            {
                label14.Text = dgvzakaz.CurrentRow.Cells[0].Value.ToString();
            }
        }
        private void dgvstena_SelectionChanged(object sender, EventArgs e)
        {
            if (dgvstena.Rows.Count >= 1)
            {
                label15.Text = dgvstena.CurrentRow.Cells[0].Value.ToString();
            }
        }
        private void dgvzakaz_Click(object sender, EventArgs e)//Кликая на грид вызываем метод
        {
            Select_stena();
            Disp_data();
            Podschet();
        }
        private void dgvstena_Click(object sender, EventArgs e)
        {
            Disp_data();
            Podschet();
        }
        private void button1_Click(object sender, EventArgs e)//Добавить заказ
        {
            if (MessageBox.Show("Вы создаете новый заказ, подтвердите действие", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Exclamation) == DialogResult.Yes)
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("INSERT INTO [Table_wall] (zakaz,wall,vid,napolnitel,thickness,width,height,color) VALUES (@zakaz,@wall,@vid,@napolnitel,@thickness,@width,@height,@color)", con);
                cmd.Parameters.AddWithValue("@zakaz", Convert.ToInt32(dgvzakaz.Rows[0].Cells[0].Value) + 1);
                cmd.Parameters.AddWithValue("@wall", 0);
                cmd.Parameters.AddWithValue("@vid", "");
                cmd.Parameters.AddWithValue("@napolnitel", "");
                cmd.Parameters.AddWithValue("@thickness", 0);
                cmd.Parameters.AddWithValue("@width", 0);
                cmd.Parameters.AddWithValue("@height", 0);
                cmd.Parameters.AddWithValue("@color", "");
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение 
                Select_zakaz();
                Disp_data();
                Select_stena();
            }
        }

        private void button3_Click(object sender, EventArgs e)//Добавить стену
        {
            if (comboBox1.Text == "Стена")
            {
                comboBox6.SelectedIndex = 0;
            }
            else if (comboBox1.Text == "Кровля")
            {
                comboBox6.SelectedIndex = 1;
            }
            if (dgvzakaz.Rows.Count >= 2)
            {
                if (comboBox1.Text != "" & comboBox1.Text != "" & comboBox1.Text != "" & comboBox1.Text != "" & textBox3.Text != "" & textBox4.Text != "")
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("INSERT INTO [Table_wall] (zakaz,wall,vid,napolnitel,thickness,width,height,color) VALUES (@zakaz,@wall,@vid,@napolnitel,@thickness,@width,@height,@color)", con);
                    cmd.Parameters.AddWithValue("@zakaz", label14.Text);
                    if (dgvstena.Rows.Count >=1)
                    {
                        cmd.Parameters.AddWithValue("@wall", Convert.ToInt32(dgvstena.Rows[0].Cells[0].Value) + 1);
                    }
                    else cmd.Parameters.AddWithValue("@wall", 1);
                    cmd.Parameters.AddWithValue("@vid", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@napolnitel", comboBox2.Text);
                    cmd.Parameters.AddWithValue("@thickness", comboBox3.Text);
                    cmd.Parameters.AddWithValue("@width", textBox4.Text);
                    cmd.Parameters.AddWithValue("@height", textBox3.Text);
                    cmd.Parameters.AddWithValue("@color", comboBox4.Text);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение 
                    Select_stena();
                    Disp_data();                   
                    RESULT();
                    Disp_data();
                    Podschet();
                }
                else MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Создайте заказ!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }
        private void button6_Click(object sender, EventArgs e)//Удалить стену
        {
                try
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("DELETE FROM [Table_wall] WHERE zakaz = @zakaz AND wall = @wall", con);
                    cmd.Parameters.AddWithValue("@zakaz", label14.Text);
                    cmd.Parameters.AddWithValue("@wall", label15.Text);
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение
                con.Open();//открыть соединение
                SqlCommand cmd1 = new SqlCommand("DELETE FROM [Table_window] WHERE zakaz = @zakaz AND wall = @wall", con);
                cmd1.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd1.Parameters.AddWithValue("@wall", label15.Text);
                cmd1.ExecuteNonQuery();
                con.Close();//закрыть соединение
                con.Open();//открыть соединение
                SqlCommand cmd2 = new SqlCommand("DELETE FROM [Table_result] WHERE zakaz = @zakaz AND wall = @wall", con);
                cmd2.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd2.Parameters.AddWithValue("@wall", label15.Text);
                cmd2.ExecuteNonQuery();
                con.Close();//закрыть соединение
            }
                catch (Exception)
                {
                    MessageBox.Show("Ошибка!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    con.Close();//закрыть соединение
                }
                Select_stena();
                Disp_data();
                Podschet();
        }
        private void button7_Click(object sender, EventArgs e)//Удалить заказ
        {
            try
            {
                con.Open();//открыть соединение
                SqlCommand cmd = new SqlCommand("DELETE FROM [Table_wall] WHERE zakaz = @zakaz", con);
                cmd.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd.ExecuteNonQuery();
                con.Close();//закрыть соединение
                con.Open();//открыть соединение
                SqlCommand cmd1 = new SqlCommand("DELETE FROM [Table_window] WHERE zakaz = @zakaz", con);
                cmd1.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd1.ExecuteNonQuery();
                con.Close();//закрыть соединение
                con.Open();//открыть соединение
                SqlCommand cmd2 = new SqlCommand("DELETE FROM [Table_result] WHERE zakaz = @zakaz", con);
                cmd2.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd2.ExecuteNonQuery();
                con.Close();//закрыть соединение
            }
            catch (Exception)
            {
                MessageBox.Show("Ошибка!", "Готово", MessageBoxButtons.OK, MessageBoxIcon.Information);
                con.Close();//закрыть соединение
            }           
            Select_zakaz();
            Select_stena();
            Disp_data();
            Podschet();
        }
        private void button2_Click(object sender, EventArgs e)//Добавить окно
        {
            if (dgvzakaz.Rows.Count >= 2)
            {
                if (textBox5.Text != "" & textBox6.Text != "" & dgvviev.Rows[0].Cells[0].Value.ToString() != "Кровля")
                {
                    con.Open();//открыть соединение
                    SqlCommand cmd = new SqlCommand("INSERT INTO [Table_window] (zakaz,wall,kol_vo_window,length_window,note) VALUES (@zakaz,@wall,@kol_vo_window,@length_window,@note)", con);
                    cmd.Parameters.AddWithValue("@zakaz", label14.Text);
                    cmd.Parameters.AddWithValue("@wall", label15.Text);
                    cmd.Parameters.AddWithValue("@note", comboBox5.Text);
                    cmd.Parameters.AddWithValue("@kol_vo_window", Math.Floor(Convert.ToDouble(textBox5.Text) / Convert.ToDouble(comboBox6.Text)));//Штук панелей где окна
                    cmd.Parameters.AddWithValue("@length_window", Math.Floor((Convert.ToDouble(dgvviev.Rows[0].Cells[4].Value) - Convert.ToDouble(textBox6.Text)) / 100));//Длина панелей в метрах где окна
                    cmd.ExecuteNonQuery();
                    con.Close();//закрыть соединение 

                    Disp_data();
                    RESULT();
                    Disp_data();
                    Podschet();

                    double kol_vo = Math.Ceiling(Convert.ToDouble(dgvviev.Rows[0].Cells[3].Value) / Convert.ToDouble(comboBox6.Text));//Штук панелей
                    for (int i = 0; i < dgvresult.Rows.Count; i++)
                    {
                        con.Open();//открыть соединение
                        SqlCommand cmd1 = new SqlCommand("UPDATE [Table_result] SET kol_vo = @kol_vo WHERE wall = @wall", con);
                        cmd1.Parameters.AddWithValue("@wall", label15.Text);
                        cmd1.Parameters.AddWithValue("@kol_vo", kol_vo - Convert.ToInt32(label23.Text));
                        cmd1.ExecuteNonQuery();
                        con.Close();//закрыть соединение       
                    }
                    Disp_data();
                    RESULT();
                    Disp_data();
                    Podschet();
                }
                else if (dgvviev.Rows[0].Cells[0].Value.ToString() == "Кровля")
                {
                    MessageBox.Show("Для кровли не возможно добавлять окна!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else MessageBox.Show("Не все поля заполнены!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else MessageBox.Show("Создайте заказ!", "Внимание!", MessageBoxButtons.OK, MessageBoxIcon.Stop);
        }
        private void Form1_FormClosed(object sender, FormClosedEventArgs e)//Закрыть программу
        {
            Application.Exit();
        }
        private void ChekPDF()//Выдача чека
        {
            //Определение шрифта необходимо для сохранения кириллического текста
            //Иначе мы не увидим кириллический текст
            //Если мы работаем только с англоязычными текстами, то шрифт можно не указывать
            BaseFont baseFont = BaseFont.CreateFont("C:\\Windows\\Fonts\\Arial.ttf", BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
            iTextSharp.text.Font font = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.NORMAL);
            iTextSharp.text.Font fontBold = new iTextSharp.text.Font(baseFont, 14, iTextSharp.text.Font.BOLD);
            //Обход по всем таблицам датасета          
            for (int i = 0; i < dgvresult.Rows.Count; i++)
            {
                for (int w = 0; w < dgvwindow.Rows.Count; w++)
                {
                    //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                    PdfPTable table = new PdfPTable(dgvresult.Rows[i].Cells.Count);
                    table.DefaultCell.Padding = 1;
                    table.WidthPercentage = 100;
                    float[] widths = new float[] { 7f, 10f, 10f, 15f, 10f, 7f, 10f };
                    table.SetWidths(widths);
                    table.HorizontalAlignment = Element.ALIGN_LEFT;
                    table.DefaultCell.BorderWidth = 1;

                    //Добавим в таблицу общий заголовок
                    string organization = "Бишкекская сварочная компания";
                    PdfPCell cell = new PdfPCell(new Phrase(organization, fontBold));
                    cell.Colspan = dgvresult.Rows[i].Cells.Count;
                    cell.HorizontalAlignment = 1;
                    //Убираем границу первой ячейки, чтобы была как заголовок
                    cell.Border = 0;
                    table.AddCell(cell);
                    //Сначала добавляем заголовки таблицы
                    for (int j = 0; j < dgvresult.Columns.Count; j++)
                    {
                        cell = new PdfPCell(new Phrase(dgvresult.Columns[j].HeaderText.ToString(), font))
                        {
                            //Фоновый цвет (необязательно, просто сделаем по красивее)
                            //cell.BackgroundColor = BaseColor.LIGHT_GRAY;
                        Border = 1
                        };
                        table.AddCell(cell);
                    }
                    //Добавляем все остальные ячейки
                    for (int x = 0; x < dgvresult.Rows.Count; x++)
                    {
                        for (int k = 0; k < dgvresult.Columns.Count; k++)
                        {
                            table.AddCell(new Phrase(dgvresult.Rows[x].Cells[k].Value.ToString(), font));
                        }
                    }
                    /////////------------------------------------------ Вторая таблица --------------------------------------------------////////////
                    //Создаем объект таблицы и передаем в нее число столбцов таблицы из нашего датасета
                    PdfPTable table2 = new PdfPTable(dgvwindow.Rows[w].Cells.Count);
                    table2.DefaultCell.Padding = 1;
                    table2.WidthPercentage = 100;
                    float[] widths2 = new float[] { 5f, 5f, 5f, 15f };
                    table2.SetWidths(widths2);
                    table2.HorizontalAlignment = Element.ALIGN_LEFT;
                    table2.DefaultCell.BorderWidth = 1;

                    
                    //Сначала добавляем заголовки таблицы
                    for (int j = 0; j < dgvwindow.Columns.Count; j++)
                    {
                        cell = new PdfPCell(new Phrase(dgvwindow.Columns[j].HeaderText.ToString(), font));
                        cell.Border = 1;
                        table2.AddCell(cell);
                    }
                    //Добавляем все остальные ячейки
                    for (int x = 0; x < dgvwindow.Rows.Count; x++)
                    {
                        for (int k = 0; k < dgvwindow.Columns.Count; k++)
                        {
                            table2.AddCell(new Phrase(dgvwindow.Rows[x].Cells[k].Value.ToString(), font));
                        }
                    }
                    //----------------------------------------------------------------------------------------------------------//
                    //Exporting to PDF
                    string folderPath = "C:\\Program Files (x86)\\Alesunix\\BSC\\components";
                    if (!Directory.Exists(folderPath))
                    {
                        Directory.CreateDirectory(folderPath);
                    }
                    using (FileStream stream = new FileStream(folderPath + "Print.pdf", FileMode.Create))
                    {
                        iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(@"C:\Program Files (x86)\Alesunix\BSC\components\image.png");
                        png.ScaleToFit(30f, 30f);
                        Document Doc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
                        //Document Doc = new Document(new iTextSharp.text.Rectangle(Width, Height), 0, 0, 0, 0);
                        //Document Doc = new Document(new iTextSharp.text.Rectangle(120, 1000), 0f, 0f, 0f, 0f);
                        PdfWriter.GetInstance(Doc, stream);
                        Doc.Open();
                        DateTime date = DateTime.Now;
                        
                        Doc.Add(new Paragraph("Расчет " + "Заказ № " + label14.Text, font));
                        Doc.Add(new Paragraph("Дата: " + date, font));
                        Doc.Add(new Paragraph("Адрес: " + " ", font));

                        Doc.Add(png);
                        Doc.Add(table);//таблица
                        Doc.Add(table2);//таблица2
                        Doc.Close();
                        stream.Close();
                    }
                }
            }
            // Печать на устройство, установленное используемым по умолчанию
            Process printJob = new Process();
            printJob.StartInfo.FileName = @"C:\\Program Files (x86)\\Alesunix\\BSC\\components\\Print.pdf";
            printJob.StartInfo.UseShellExecute = true;
            //printJob.StartInfo.Verb = "print";
            printJob.Start();

            printJob.WaitForInputIdle();
            //printJob.Kill();
        }
        private void button5_Click(object sender, EventArgs e)//Печать
        {
            const string filePatch = @"C:\\Program Files (x86)\\Alesunix\\BSC\\components\\Print.pdf";//путь к файлу чека
            FileStream stream = null;
            try
            {
                stream = File.Open(filePatch, FileMode.Open, FileAccess.Read, FileShare.None);//Проверка открыт ли файл если да то отказ
            }
            catch (Exception)
            {
                MessageBox.Show("Закройте файл PDF!", "ОТКАЗАНО", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                if (stream != null)
                    stream.Close();
                //запускаем_закрытый_файл();

                con.Open();//Открываем соединение
                SqlCommand cmd = new SqlCommand("SELECT MAX(wall) AS Стена,MAX(vid) AS Вид,MAX(thickness) AS Толщина,MAX(napolnitel) AS Наполнитель,MIN(kol_vo) AS 'Кол-во',MAX(length) AS Длина," +
                    "MAX(color) AS Цвет FROM [Table_result] " +
                    "WHERE kol_vo_window IN (N'0') AND zakaz=@zakaz GROUP BY wall ORDER BY wall ", con);
                cmd.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd.ExecuteNonQuery();
                DataTable dt = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da = new SqlDataAdapter(cmd);//создаем экземпляр класса SqlDataAdapter
                dt.Clear();//чистим DataTable, если он был не пуст
                da.Fill(dt);//заполняем данными созданный DataTable
                dgvresult.DataSource = dt;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//Закрываем соединение

                con.Open();//Открываем соединение
                SqlCommand cmd1 = new SqlCommand("SELECT MIN(wall) AS Стена,SUM(kol_vo_window) AS 'Кол-во',length_window AS Длина, MIN(note) AS Примечание FROM [Table_window] " +
                    "WHERE kol_vo_window NOT IN (N'0') AND zakaz=@zakaz GROUP BY length_window ORDER BY length_window DESC", con);
                cmd1.Parameters.AddWithValue("@zakaz", label14.Text);
                cmd1.ExecuteNonQuery();
                DataTable dt1 = new DataTable();//создаем экземпляр класса DataTable
                SqlDataAdapter da1 = new SqlDataAdapter(cmd1);//создаем экземпляр класса SqlDataAdapter
                dt1.Clear();//чистим DataTable, если он был не пуст
                da1.Fill(dt1);//заполняем данными созданный DataTable
                dgvwindow.DataSource = dt1;//в качестве источника данных у dataGridView используем DataTable заполненный данными
                con.Close();//Закрываем соединение
                ChekPDF();
            }
        }
    }
}

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;


namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form8 : Form
    {
        
        string nombreUsuario;
        Form formularioAnterior;
        public Form8(string nombre, Form anterior)
        {
            InitializeComponent();
            nombreUsuario = nombre;
            formularioAnterior = anterior;
        }

        void CargarTotales()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                SqlCommand cmd1 = new SqlCommand("SELECT COUNT(*) FROM Difusiones", conn);
                label12.Text = cmd1.ExecuteScalar().ToString();

                SqlCommand cmd2 = new SqlCommand("SELECT COUNT(*) FROM Capacitacion", conn);
                label13.Text = cmd2.ExecuteScalar().ToString();

                SqlCommand cmd3 = new SqlCommand("SELECT COUNT(*) FROM Vinculacion", conn);
                label14.Text = cmd3.ExecuteScalar().ToString();

            }
        }

        void CargarGrafica()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                //DIFUSIONES
                conn.Open();
                SqlCommand cmd1 = new SqlCommand("SELECT COUNT(*) FROM Difusiones", conn);
                int difusiones = (int)cmd1.ExecuteScalar();


                //CAPACITACIONES
                
                SqlCommand cmd2 = new SqlCommand("SELECT COUNT(*) FROM Capacitacion", conn);
                int capacitaciones = (int)cmd1.ExecuteScalar();


                //VINCULACION
                SqlCommand cmd3 = new SqlCommand("SELECT COUNT(*) FROM Vinculacion", conn);
                int vinculacion = (int)cmd1.ExecuteScalar();

                chart1.Series.Clear();

                var serie = chart1.Series.Add("General");
                serie.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;

                serie.Points.AddXY("Difusiones", difusiones);
                serie.Points.AddXY("Capacitaciones", capacitaciones);
                serie.Points.AddXY("Vinculacion", vinculacion);
            }
        }
        
        

        void CargarRecientes()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"SELECT TOP 10 * FROM (
                SELECT 'Difusion' AS area, fecha_publicacion AS fecha, tipo_publicacion AS tipo, descripcion
                FROM Difusiones
                UNION ALL
                SELECT 'Capacitacion' AS area, fecha, tipo, nombre_inmueble 
                FROM Capacitacion
                UNION ALL 
                SELECT 'Vinculacion' AS area, fecha, sector, nombre 
                FROM Vinculacion)
                AS resultado ORDER BY fecha DESC";

                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.DataSource = dt;
            }
        }

        void CargarGraficaMensual()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                chart1.Series.Clear();

                chart1.ChartAreas[0].BackColor = Color.White;
                chart1.BackColor = Color.White;

                var serieDif = chart1.Series.Add("Difusiones");
                serieDif.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                serieDif.Color = Color.DarkRed;
                serieDif.IsValueShownAsLabel = true;
                serieDif["PointWidth"] = "0.4";
                serieDif.IsValueShownAsLabel = false;

                var serieCap = chart1.Series.Add("Capacitaciones");
                serieCap.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                serieCap.Color = Color.SteelBlue;
                serieCap.IsValueShownAsLabel = true;
                serieCap["PointWidth"] = "0.4";
                serieCap.IsValueShownAsLabel = false;

                var serieVinc = chart1.Series.Add("Vinculaciones");
                serieVinc.ChartType = System.Windows.Forms.DataVisualization.Charting.SeriesChartType.Column;
                serieVinc.Color = Color.ForestGreen;
                serieVinc.IsValueShownAsLabel = true;
                serieVinc["PointWidth"] = "0.4";
                serieVinc.IsValueShownAsLabel = false;

                chart1.ChartAreas[0].AxisX.Interval = 1;
                chart1.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                chart1.ChartAreas[0].AxisX.LabelStyle.Angle = -45;
                chart1.ChartAreas[0].AxisY.Interval = 1;

                string[] meses = { "Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic" };

                int[] dif = new int[12];
                int[] cap = new int[12];
                int[] vinc = new int[12];


                SqlCommand cmd1 = new SqlCommand(@"
                SELECT MONTH(fecha_publicacion) mes, COUNT(*) total 
                FROM Difusiones
                GROUP BY MONTH(fecha_publicacion)", conn);
                SqlDataReader dr1 = cmd1.ExecuteReader();
                while (dr1.Read())
                {
                    int mes = Convert.ToInt32(dr1["mes"]) - 1;
                    dif[mes] = Convert.ToInt32(dr1["total"]);
                }
                dr1.Close();


                SqlCommand cmd2 = new SqlCommand(@"
                SELECT MONTH(fecha) mes, COUNT(*) total
                FROM Capacitacion
                GROUP BY MONTH(fecha)", conn);
                SqlDataReader dr2 = cmd2.ExecuteReader();
                while (dr2.Read())
                {
                    int mes = Convert.ToInt32(dr2["mes"]) - 1;
                    cap[mes] = Convert.ToInt32(dr2["total"]);
                }
                dr2.Close();

                SqlCommand cmd3 = new SqlCommand(@"
                SELECT MONTH(fecha) mes, COUNT(*) total
                FROM Vinculacion
                GROUP BY MONTH(fecha)", conn);
                SqlDataReader dr3 = cmd3.ExecuteReader();
                while (dr3.Read())
                {
                    int mes = Convert.ToInt32(dr3["mes"]) - 1;
                    vinc[mes] = Convert.ToInt32(dr3["total"]);
                }
                dr3.Close();

                for (int i = 0; i < 12; i++)
                {
                    
                        serieDif.Points.AddXY(meses[i], dif[i]);
                        serieCap.Points.AddXY(meses[i], cap[i]);
                        serieVinc.Points.AddXY(meses[i], vinc[i]);
                }

            }
        }

        private void Form8_Load(object sender, EventArgs e)
        {
            CargarTotales();
            CargarGraficaMensual();
            CargarRecientes();

            label20.Text = "Usuario: " + nombreUsuario;
            panel1.Cursor = Cursors.Hand;
            panel2.Cursor = Cursors.Hand;
            panel3.Cursor = Cursors.Hand;


            dataGridView1.Columns["area"].HeaderText = "Área";
            dataGridView1.Columns["fecha"].HeaderText = "Fecha";
            dataGridView1.Columns["tipo"].HeaderText = "Tipo";
            dataGridView1.Columns["descripcion"].HeaderText = "Descripción";

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;


            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.Tan;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.LightGray;

            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.SteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.White;
            dataGridView1.BackgroundColor = Color.White;


            chart1.BackColor = Color.SeaShell;
            chart1.ChartAreas[0].BackColor = Color.SeaShell;
            chart1.Legends[0].BackColor = Color.SeaShell;

            //LINEAS HORIZONTALES//
            chart1.ChartAreas[0].AxisY.MajorGrid.LineColor = Color.SteelBlue;
            chart1.ChartAreas[0].AxisY.MajorGrid.LineWidth = 1;

            //LINEAS VERTICALES//
            chart1.ChartAreas[0].AxisX.MajorGrid.LineColor = Color.SteelBlue;
            chart1.ChartAreas[0].AxisX.MajorGrid.LineWidth = 1;

            chart1.BorderlineWidth = 0;
            chart1.ChartAreas[0].AxisX.LineColor = Color.White;
            chart1.ChartAreas[0].AxisY.LineColor = Color.White;


            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;

            panel5.Click += panel5_Click;
            label5.Click += panel5_Click;
            pictureBox5.Click += panel5_Click;


            panel6.Click += panel6_Click;
            label7.Click += panel6_Click;
            pictureBox7.Click += panel6_Click;

            panel7.Click += panel7_Click;
            label10.Click += panel7_Click;
            pictureBox8.Click += panel7_Click;

            panel8.Click += panel8_Click;
            label18.Click += panel8_Click;
            pictureBox10.Click += panel8_Click;

            panel9.Click += panel9_Click;
            label19.Click += panel9_Click;
            pictureBox12.Click += panel9_Click;

            label9.Text = DateTime.Now.ToString("dddd, dd 'de' MMMM 'de' yyyy");

            RedondearPanel(panel2);
            RedondearPanel(panel3);
            RedondearPanel(panel4);


            HoverPanel(panel6);
            HoverPanel(panel7);
            HoverPanel(panel8);
            HoverPanel(panel9);
            HoverPanel(panel5);
        }

        private void panel9_Click(object sender, EventArgs e)
        {
            Form5 dif = new Form5(this);
            this.Hide();
            dif.Show();
        }

        private void panel8_Click(object sender, EventArgs e)
        {
            Form4 vinc = new Form4(this);
            this.Hide();
            vinc.Show();
        }

        private void panel7_Click(object sender, EventArgs e)
        {
            Form3 cap = new Form3(this);
            this.Hide();
            cap.Show();
        }

        void HoverPanel(Panel p)
        {
            p.MouseEnter += (s, e) => p.BackColor = Color.FromArgb(230, 210, 180);
            p.MouseLeave += (s, e) => p.BackColor = Color.FromArgb(245, 225, 200);

            foreach (Control ctrl in p.Controls)
            {
                ctrl.MouseEnter += (s, e) => p.BackColor = Color.FromArgb(230, 210, 180);
                ctrl.MouseLeave += (s, e) => p.BackColor = Color.FromArgb(245, 225, 200);
            }
        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void dataGridView1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "area")
            {
                if (e.Value != null)
                {
                    string valor = e.Value.ToString();

                    if (valor == "Difusion")
                        e.CellStyle.ForeColor = Color.MistyRose;

                    else if (valor == "Capacitacion")
                        e.CellStyle.ForeColor = Color.LightCyan;
                }

            }
        }

        private void panel5_Click(object sender, EventArgs e)
        {
            DialogResult resultado = MessageBox.Show(
               "¿Estás seguro que deseas salir?",
               "Confirmar salida",
               MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                Application.Exit();
            }
        }

        private void panel6_Click(object sender, EventArgs e)
        {
            if (formularioAnterior != null && !formularioAnterior.IsDisposed)
            {
                formularioAnterior.Show();
            }
            else
            {
                Form2 menu = new Form2(Sesion.Usuario);
                menu.Show();
            }
            this.Close();
        }

        private void RedondearPanel(Panel panel)
        {
            int radio = 20; // ideal para 320x165

            GraphicsPath path = new GraphicsPath();
            path.StartFigure();

            path.AddArc(0, 0, radio, radio, 180, 90);
            path.AddArc(panel.Width - radio, 0, radio, radio, 270, 90);
            path.AddArc(panel.Width - radio, panel.Height - radio, radio, radio, 0, 90);
            path.AddArc(0, panel.Height - radio, radio, radio, 90, 90);

            path.CloseFigure();
            panel.Region = new Region(path);
        }

        private void panel2_Resize(object sender, EventArgs e)
        {
            RedondearPanel(panel2);
        }

        private void panel3_Resize(object sender, EventArgs e)
        {
            RedondearPanel(panel3);
        }

        private void panel4_Resize(object sender, EventArgs e)
        {
            RedondearPanel(panel4);
        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void chart1_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {
            
        }

        private void label22_Click(object sender, EventArgs e)
        {

        }
    }
}

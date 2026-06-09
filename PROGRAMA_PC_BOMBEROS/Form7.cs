using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Excel = Microsoft.Office.Interop.Excel;

namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form7 : Form
    {
        
        Form formularioAnterior;
        public Form7(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
        }

        void CargarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();
                string query = "SELECT * FROM Difusiones";
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView1.DataSource = dt;
            }

        }

        private void Form7_Load(object sender, EventArgs e)
        {
            CargarDatos();

            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);



            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.SteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.White;

            dataGridView1.BackgroundColor = Color.White;
            dataGridView1.BorderStyle = BorderStyle.None;

            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AntiqueWhite;

            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.ScrollBars = ScrollBars.Both;


           



            dataGridView1.Columns["id_difusion"].Visible = false;
            dataGridView1.Columns["enlace"].DefaultCellStyle.ForeColor = Color.Blue;

            dataGridView1.Columns["id_difusion"].HeaderText = "ID";
            dataGridView1.Columns["red_social"].HeaderText = "Red Social";
            dataGridView1.Columns["tipo_publicacion"].HeaderText = "Tipo de publicación";
            dataGridView1.Columns["fecha_publicacion"].HeaderText = "Fecha";
            dataGridView1.Columns["enlace"].HeaderText = "Enlace";
            dataGridView1.Columns["responsable"].HeaderText = "Responsable";
            dataGridView1.Columns["descripcion"].HeaderText = "Descripción";
            dataGridView1.Columns["diseño_contenido"].HeaderText = "Diseño del contenido";
            dataGridView1.Columns["usuario_creacion"].HeaderText = "Creado por";
            dataGridView1.Columns["fecha_creacion"].HeaderText = "Fecha de creacion"; 
            dataGridView1.Columns["usuario_modificacion"].HeaderText = "Modificado por";
            dataGridView1.Columns["fecha_modificacion"].HeaderText = "Ultima modificacion";

            dataGridView1.Columns["fecha_creacion"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
            dataGridView1.Columns["fecha_modificacion"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";

            dataGridView1.Columns["usuario_creacion"].DefaultCellStyle.ForeColor = Color.DarkBlue;
            dataGridView1.Columns["usuario_modificacion"].DefaultCellStyle.ForeColor = Color.DarkBlue;

            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;

            
        }

        void FiltrarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();


                string query = @"SELECT * FROM Difusiones
                WHERE (@usarFechas = 0 OR fecha_publicacion BETWEEN @inicio AND @fin)
                AND (@tipo = '' OR tipo_publicacion COLLATE Latin1_General_CI_AI LIKE '%' + @tipo + '%')
                AND (@desc = '' OR descripcion COLLATE Latin1_General_CI_AI LIKE '%' + @desc + '%')";

                SqlCommand cmd = new SqlCommand(query, conn);
                bool usarFechas = checkBox7.Checked;

                cmd.Parameters.AddWithValue("@usarFechas", usarFechas ? 1 : 0);
                cmd.Parameters.AddWithValue("@inicio", dateTimePicker1.Value.Date);
                cmd.Parameters.AddWithValue("@fin", dateTimePicker2.Value.Date);


                string tipo = textBox2.Text.Trim();
                cmd.Parameters.AddWithValue("@tipo", tipo);

                string desc = textBox1.Text.Trim();
                cmd.Parameters.AddWithValue("@desc", desc);

               
                DataTable dt = new DataTable();
                SqlDataAdapter da = new SqlDataAdapter(cmd);
                da.Fill(dt);

                dataGridView1.DataSource = dt;

                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[0].Selected = true;
                }
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            FiltrarDatos();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            FiltrarDatos();
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            timer1.Stop();
            FiltrarDatos();
        }

        private void dateTimePicker1_ValueChanged(object sender, EventArgs e)
        {
            FiltrarDatos();

        }

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            FiltrarDatos();

        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            FiltrarDatos();

        }

       
        private void comboBox2_TextChanged(object sender, EventArgs e)
        {

            FiltrarDatos();


        }

        private void button4_Click(object sender, EventArgs e)
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

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un registro", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int idSeleccionado = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_difusion"].Value);

            Form5 form5 = new Form5(this);
            form5.idEditar = idSeleccionado;

            form5.Show();

            this.Close();
            CargarDatos();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un registro para eliminar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            int idSeleccionado = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_difusion"].Value);
            DialogResult resultado = MessageBox.Show("¿Eliminar registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(Conexion.cadena))
                {
                    try
                    {
                        conn.Open();

                        string query = "DELETE FROM Difusiones WHERE id_difusion = @id";
                        SqlCommand cmd = new SqlCommand(query, conn);
                        cmd.Parameters.AddWithValue("@id", idSeleccionado);

                        cmd.ExecuteNonQuery();

                        MessageBox.Show("Se elimino correctamente", "Registro eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                        CargarDatos();
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }

                }
            }
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            bool activo = checkBox7.Checked;
            dateTimePicker1.Enabled = activo;
            dateTimePicker2.Enabled = activo;

            FiltrarDatos();
        }

        

        private void button3_Click(object sender, EventArgs e)
        {
            textBox2.Clear();
            checkBox7.Checked = false;
            textBox1.Clear();

           

            CargarDatos();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.Columns[e.ColumnIndex].Name == "enlace")
            {
                string url = dataGridView1.CurrentRow.Cells["enlace"].Value.ToString();
                System.Diagnostics.Process.Start(url);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add(true);

            int columna = 0;

            foreach (DataGridViewColumn col in dataGridView1.Columns)
            {
                columna++;
                excel.Cells[1, columna] = col.HeaderText;
            }

            int fila = 0;

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (!row.IsNewRow)
                {
                    fila++;
                    columna = 0;
                    foreach (DataGridViewColumn col in dataGridView1.Columns)
                    {
                        columna++;
                        object valor = row.Cells[col.Index].Value;

                        excel.Cells[fila + 1, columna] = valor != null ? valor.ToString() : "";
                    }
                }
            }
            int totalColumnas = dataGridView1.Columns.Count;
            int totalFilas = fila + 1;

            for (int i = 1; i <= totalColumnas; i++)
            {
                var celda = excel.Cells[1, i];
                celda.Font.Bold = true;
                celda.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.DarkRed);
                celda.Font.Color = System.Drawing.ColorTranslator.ToOle(Color.White);
            }

            var rango = excel.Range[
                excel.Cells[1, 1],
                excel.Cells[totalFilas, totalColumnas]
                ];
            rango.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

            rango.Rows[1].HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;

            excel.Columns.AutoFit();

            excel.Application.ActiveWindow.SplitRow = 1;
            excel.Application.ActiveWindow.FreezePanes = true;

            excel.ActiveSheet.Name = "Reporte";

            excel.Visible = true;
        }
    }
}

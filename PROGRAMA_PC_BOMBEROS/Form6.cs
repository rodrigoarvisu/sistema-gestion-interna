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
    
    public partial class Form6 : Form
    {
       
        Form formularioAnterior;
        public Form6(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
        }

        void CargarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();
                string query = "SELECT * FROM Capacitacion";
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.AutoGenerateColumns = true; 
                dataGridView1.DataSource = null;
                dataGridView1.DataSource = dt;
            }
        }

        private void button5_Click(object  sender, EventArgs e)
        {

        }

        

        private void Form6_Load(object sender, EventArgs e)
        {
            

            CargarDatos();

            dataGridView1.Columns["fecha_modificacion"].Visible = true;


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

            dataGridView1.Columns["id_capacitacion"].Visible = false;


            


            dataGridView1.Columns["id_capacitacion"].HeaderText = "ID";
            dataGridView1.Columns["nombre_inmueble"].HeaderText = "Nombre de Inmueble";
            dataGridView1.Columns["tipo"].HeaderText = "Tipo de inmueble";
            dataGridView1.Columns["direccion"].HeaderText = "Dirección";
            dataGridView1.Columns["fecha"].HeaderText = "Fecha";
            dataGridView1.Columns["introduccion_proteccion_civil"].HeaderText = "Introducción a P.Civil";
            dataGridView1.Columns["primeros_aux"].HeaderText = "Primeros Auxilios";
            dataGridView1.Columns["combate_incendios"].HeaderText = "Combate contra incendios";
            dataGridView1.Columns["plan_emergencia"].HeaderText = "Plan de emergencia";
            dataGridView1.Columns["evacuacion_inmuebles"].HeaderText = "Evacuación de inmuebles";
            dataGridView1.Columns["manejo_sustancias_quimicas"].HeaderText = "Manejo de sustancias quimicas";
            dataGridView1.Columns["busqueda_rescate"].HeaderText = "Busqueda y rescate";
            dataGridView1.Columns["practica_combate_contra_incendios"].HeaderText = "Practica de combate contra incendios";
            dataGridView1.Columns["simulacro"].HeaderText = "Simulacro";
            dataGridView1.Columns["uipc"].HeaderText = "UIPC";
            dataGridView1.Columns["comite"].HeaderText = "Comité vecinal";
            dataGridView1.Columns["hipotesis"].HeaderText = "Hipotesis";
            dataGridView1.Columns["estatus"].HeaderText = "Estatus";
            dataGridView1.Columns["poblacion_fija"].HeaderText = "Población fija";
            dataGridView1.Columns["poblacion_flotante"].HeaderText = "Población flotante";
            dataGridView1.Columns["total"].HeaderText = "Población total";
            dataGridView1.Columns["hora_inicio"].HeaderText = "Hora de inicio";
            dataGridView1.Columns["hora_fin"].HeaderText = "Hora de fin";
            dataGridView1.Columns["tiempo_evacuacion"].HeaderText = "Tiempo de evacuación";
            dataGridView1.Columns["alertamiento"].HeaderText = "Alertamiento";
            dataGridView1.Columns["observaciones"].HeaderText = "Observaciones";
            dataGridView1.Columns["simulacro_gabinete"].HeaderText = "¿Simulacro de gabinete?";

            dataGridView1.Columns["usuario_creacion"].HeaderText = "Creado por";
            dataGridView1.Columns["fecha_creacion"].HeaderText = "Fecha de creacion";
            dataGridView1.Columns["usuario_modificacion"].HeaderText = "Modificado por";
            dataGridView1.Columns["fecha_modificacion"].HeaderText = "Ultima modificacion";

            dataGridView1.Columns["fecha_creacion"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
            dataGridView1.Columns["fecha_modificacion"].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";

            dataGridView1.Columns["usuario_creacion"].DefaultCellStyle.ForeColor = Color.DarkBlue;
            dataGridView1.Columns["usuario_modificacion"].DefaultCellStyle.ForeColor = Color.DarkBlue;

            dateTimePicker3.Enabled = false;
            dateTimePicker2.Enabled = false;

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

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un registro para eliminar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            int idSeleccionado = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_capacitacion"].Value);
            DialogResult resultado = MessageBox.Show("¿Eliminar registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(Conexion.cadena))
                {

                    conn.Open();

                    string query = "DELETE FROM Capacitacion WHERE id_capacitacion = @id";
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@id", idSeleccionado);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Se elimino correctamente", "Registro eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    CargarDatos();
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un registro", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            int idSeleccionado = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_capacitacion"].Value);

            Form3 frm = new Form3(this);
            frm.idEditar = idSeleccionado;

            frm.Show();

            this.Close();

            
        }

        void FiltrarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"SELECT * FROM Capacitacion
                 WHERE 
                  (
                      @texto = '' OR 
                      ISNULL(nombre_inmueble,'') COLLATE Latin1_General_CI_AI LIKE '%' + @texto + '%' OR 
                      ISNULL(tipo,'') COLLATE Latin1_General_CI_AI LIKE '%' + @texto + '%'
                  )
                      AND 
                     (@usarFechas = 0 OR fecha BETWEEN @inicio AND @fin)";

                SqlCommand cmd = new SqlCommand(query, conn);

                
                
                string texto = textBox1.Text.Trim();
                cmd.Parameters.AddWithValue("@texto", texto);


                bool usarFechas = checkBox7.Checked;
                cmd.Parameters.AddWithValue("@usarFechas", usarFechas ? 1 : 0);
                cmd.Parameters.AddWithValue("@inicio", dateTimePicker3.Value.Date);
                cmd.Parameters.AddWithValue("@fin", dateTimePicker2.Value.Date);

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

        private void dateTimePicker2_ValueChanged(object sender, EventArgs e)
        {
            FiltrarDatos();

        }

        private void dateTimePicker3_ValueChanged(object sender, EventArgs e)
        {
            FiltrarDatos();

        }
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            FiltrarDatos();
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            bool activo = checkBox7.Checked;
            dateTimePicker3.Enabled = activo;
            dateTimePicker2.Enabled = activo;

            FiltrarDatos();
        }

        private void dateTimePicker1_Changed(object sender, EventArgs e)
        {
            FiltrarDatos();
        }

       

        private void button5_Click_1(object sender, EventArgs e)
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

        private void button3_Click_1(object sender, EventArgs e)
        {
            textBox1.Clear();
            checkBox7.Checked = false;
            CargarDatos();
        }
    }
}

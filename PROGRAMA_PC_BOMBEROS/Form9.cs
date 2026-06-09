using Microsoft.Office.Core;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;

using Excel = Microsoft.Office.Interop.Excel;

namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form9 : Form
    {
       
        Form formularioAnterior;
        public Form9(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
        }

        private void Form9_Load(object sender, EventArgs e)
        {
            

            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            dataGridView1.MultiSelect = false;
            CargarVinculacion();

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




            dataGridView2.EnableHeadersVisualStyles = false;

            dataGridView2.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView2.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView2.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);



            dataGridView2.DefaultCellStyle.SelectionBackColor = Color.SteelBlue;
            dataGridView2.DefaultCellStyle.SelectionForeColor = Color.White;

            dataGridView2.BackgroundColor = Color.White;
            dataGridView2.BorderStyle = BorderStyle.None;

            dataGridView2.AlternatingRowsDefaultCellStyle.BackColor = Color.AntiqueWhite;
            dataGridView2.RowHeadersVisible = false;
            dataGridView2.AllowUserToAddRows = false;

            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView2.ScrollBars = ScrollBars.Both;

            dataGridView1.Columns["id_vinculacion"].HeaderText = "ID";
            dataGridView1.Columns["sector"].HeaderText = "Sector";
            dataGridView1.Columns["nombre"].HeaderText = "Nombre";
            dataGridView1.Columns["direccion"].HeaderText = "Direccion";
            dataGridView1.Columns["fecha"].HeaderText = "Fecha";
            dataGridView1.Columns["hora_creacion"].HeaderText = "Hora de creación del acta";
            dataGridView1.Columns["hora_termino"].HeaderText = "Hora de termino del acta";
            dataGridView1.Columns["coordinador"].HeaderText = "Coordinador";
            dataGridView1.Columns["jefe_comunicacion"].HeaderText = "Jefe de comunicación";
            dataGridView1.Columns["jefe_evaluacion"].HeaderText = "Jefe de evaluacion";
            dataGridView1.Columns["jefe_primeros_auxilios"].HeaderText = "Jefe de primeros auxilios";
            dataGridView1.Columns["jefe_incendios"].HeaderText = "Jefe de incendios";
            dataGridView1.Columns["observaciones"].HeaderText = "Observaciones";
            dataGridView1.Columns["usuario_creacion"].HeaderText = "Creado por";
            dataGridView1.Columns["fecha_creacion"].HeaderText = "Fecha de creación";
            dataGridView1.Columns["usuario_modificacion"].HeaderText = "Modificado por";
            dataGridView1.Columns["fecha_modificacion"].HeaderText = "Fecha de modificación";


            dataGridView2.Columns["cargo"].HeaderText = "Cargo";
            dataGridView2.Columns["nombre"].HeaderText = "Nombre";
            dataGridView2.Columns["telefono"].HeaderText = "Teléfono";
            dataGridView2.Columns["correo"].HeaderText = "Correo";




            dateTimePicker1.Enabled = false;
            dateTimePicker2.Enabled = false;

            
        }


        void FiltrarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"SELECT * FROM Vinculacion
                  WHERE (@usarFechas = 0 OR fecha BETWEEN @inicio AND @fin)
                  AND (@sector = '' OR sector COLLATE Latin1_General_CI_AI LIKE @sector)
                  AND (@nombre = '' OR nombre COLLATE Latin1_General_CI_AI LIKE @nombre)";

                SqlCommand cmd = new SqlCommand(query, conn);

                // 🔹 Fechas
                bool usarFechas = checkBox7.Checked;
                cmd.Parameters.AddWithValue("@usarFechas", usarFechas ? 1 : 0);
                cmd.Parameters.AddWithValue("@inicio", dateTimePicker1.Value.Date);
                cmd.Parameters.AddWithValue("@fin", dateTimePicker2.Value.Date);

                // 🔹 Sector
                cmd.Parameters.AddWithValue("@sector", "%" + textBox2.Text.Trim() + "%");

                // 🔹 Nombre
                cmd.Parameters.AddWithValue("@nombre", "%" + textBox1.Text.Trim() + "%");

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

        

        void CargarIntegrantes(int id)
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();
                string query = @"
                SELECT
                  cargo, nombre, telefono, correo FROM Integrantes 
                WHERE id_vinculacion = @id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", id);

                SqlDataAdapter da = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView2.DataSource = dt;
            }
        }

        void CargarVinculacion()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"
                SELECT
                id_vinculacion, nombre, sector, direccion, fecha, hora_creacion, hora_termino, coordinador, jefe_comunicacion, jefe_evaluacion,
                jefe_primeros_auxilios, jefe_incendios, observaciones, usuario_creacion, fecha_creacion, usuario_modificacion, fecha_modificacion
                FROM Vinculacion ORDER BY id_vinculacion ASC";

                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.DataSource = dt;

                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.ClearSelection();
                    dataGridView1.Rows[0].Selected = true;

                    int id = Convert.ToInt32(
                        dataGridView1.Rows[0].Cells["id_vinculacion"].Value);
                    CargarIntegrantes(id);
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
                dataGridView1.CurrentRow.Cells["id_vinculacion"].Value);

            Form4 form4 = new Form4(this);
            form4.idEditar = idSeleccionado;

            form4.Show();

            this.Close();
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.Cells["id_vinculacion"].Value != null)
            {
                int id = Convert.ToInt32(dataGridView1.CurrentRow.Cells["id_vinculacion"].Value);

                CargarIntegrantes(id);
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow != null && dataGridView1.CurrentRow.Cells["id_vinculacion"].Value != null)
            {
                int id = Convert.ToInt32(
                    dataGridView1.CurrentRow.Cells["id_vinculacion"].Value);

                CargarIntegrantes(id);
            }
            else
            {
                dataGridView2.DataSource = null; 
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (formularioAnterior != null && !formularioAnterior.IsDisposed)
            {
                formularioAnterior.Show();
            }
            else
            {
                Form4 vinc = new Form4(this);
                vinc.Show();
            }
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un registro para eliminar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            int id = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_vinculacion"].Value);
            DialogResult resultado = MessageBox.Show("¿Eliminar registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(Conexion.cadena))
                {
                    
                        conn.Open();
                        SqlCommand cmd = new SqlCommand("DELETE FROM Vinculacion WHERE id_vinculacion = @id", conn);
                        cmd.Parameters.AddWithValue("@id", id);

                        cmd.ExecuteNonQuery();
                }
                MessageBox.Show("Se elimino correctamente", "Registro eliminado", MessageBoxButtons.OK, MessageBoxIcon.Information);
                CargarVinculacion();
                dataGridView2.DataSource = null;
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
        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            textBox2.Clear();
            checkBox7.Checked = false;

            CargarVinculacion();
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
            bool activo = checkBox7.Checked;
            dateTimePicker1.Enabled = activo;
            dateTimePicker2.Enabled = activo;

            FiltrarDatos();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("No hay datos para exportar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Excel.Application excel = new Excel.Application();
            excel.Workbooks.Add(true);

            // EXPORTAR GRID
            int ExportarGrid(DataGridView grid, int filaInicio)
            {
                int fila = filaInicio;
                int columna;

                //  Encabezados
                columna = 0;
                foreach (DataGridViewColumn col in grid.Columns)
                {
                    if (!col.Visible) continue;

                    columna++;
                    excel.Cells[fila, columna] = col.HeaderText;

                    var celda = excel.Cells[fila, columna];
                    celda.Font.Bold = true;
                    celda.Interior.Color = ColorTranslator.ToOle(Color.DarkRed);
                    celda.Font.Color = ColorTranslator.ToOle(Color.White);
                }

                //  Datos
                foreach (DataGridViewRow row in grid.Rows)
                {
                    if (row.IsNewRow) continue;

                    fila++;
                    columna = 0;

                    foreach (DataGridViewColumn col in grid.Columns)
                    {
                        if (!col.Visible) continue;

                        columna++;
                        object valor = row.Cells[col.Index].Value;

                        excel.Cells[fila, columna] = valor != null ? valor.ToString() : "";
                    }
                }
                // Exportar tabla de integrantes
                

                //  Bordes
                var rango = excel.Range[
                    excel.Cells[filaInicio, 1],
                    excel.Cells[fila, columna]
                ];

                rango.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                return fila;
            }



            int ExportarDataTable(Excel.Application excel2, DataTable dt, int filaInicio2)
            {
                int fila2 = filaInicio2;
                int columna2;

                Dictionary<string, string> nombresColumnas = new Dictionary<string, string>()
                {
                    { "id_vinculacion", "ID Vinculación" },
                    { "vinculacion", "Nombre Vinculación" },
                    { "cargo", "Cargo" },
                    { "nombre", "Nombre" },
                    { "telefono", "Teléfono" },
                    { "correo", "Correo Electrónico" }
                };

                // Encabezados
                columna2 = 0;
                foreach (DataColumn col in dt.Columns)
                {
                    columna2++;
                    string nombreMostrar = nombresColumnas.ContainsKey(col.ColumnName)
                     ? nombresColumnas[col.ColumnName] : col.ColumnName;

                    excel2.Cells[fila2, columna2] = nombreMostrar;

                    var celda2 = excel2.Cells[fila2, columna2];
                    celda2.Font.Bold = true;
                    celda2.Interior.Color = ColorTranslator.ToOle(Color.DarkRed);
                    celda2.Font.Color = ColorTranslator.ToOle(Color.White);
                }

                // Datos
                foreach (DataRow row in dt.Rows)
                {
                    fila2++;
                    columna2 = 0;

                    foreach (DataColumn col in dt.Columns)
                    {
                        columna2++;
                        excel2.Cells[fila2, columna2] = row[col]?.ToString() ?? "";
                    }
                }

                var rango2 = excel2.Range[
                    excel2.Cells[filaInicio2, 1],
                    excel2.Cells[fila2, columna2] 
                ];

                rango2.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                return fila2;
            }
            DataTable dtIntegrantes = new DataTable();

            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"
                    SELECT 
                    v.id_vinculacion,
                    v.nombre AS vinculacion,
                    i.cargo,
                    i.nombre,
                    i.telefono,
                    i.correo
                    FROM Integrantes i
                    INNER JOIN Vinculacion v 
                    ON i.id_vinculacion = v.id_vinculacion
                    ORDER BY v.id_vinculacion";

                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                da.Fill(dtIntegrantes);
            }

            //  Exportar primer grid
            int ultimaFila = ExportarGrid(dataGridView1, 1);

            //  Espacio + título
            excel.Cells[ultimaFila + 2, 1] = "Integrantes";
            excel.Cells[ultimaFila + 2, 1].Font.Bold = true;



            //  Exportar segundo grid
            ExportarDataTable(excel, dtIntegrantes, ultimaFila + 3);
           

            excel.Columns.AutoFit();

            excel.Application.ActiveWindow.SplitRow = 1;
            excel.Application.ActiveWindow.FreezePanes = true;

            excel.ActiveSheet.Name = "Reporte";
            excel.Visible = true;
        }
    }
}

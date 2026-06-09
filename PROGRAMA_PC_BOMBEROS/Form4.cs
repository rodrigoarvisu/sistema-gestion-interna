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

namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form4 : Form
    {
       
        public int idEditar = 0;
        Form formularioAnterior;
        bool actualizando = false;
        public Form4(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
        }

        void ConfigurarDataGrid()
        {
            dataGridView1.Columns.Clear();

            DataGridViewTextBoxColumn colCargo = new DataGridViewTextBoxColumn();
            colCargo.Name = "Cargo";
            colCargo.HeaderText = "Cargo";
            colCargo.ReadOnly = true;

            dataGridView1.Columns.Add(colCargo);

            dataGridView1.Columns.Add("Nombre", "Nombre");
            dataGridView1.Columns.Add("Telefono", "Telefono");
            dataGridView1.Columns.Add("Correo", "Correo");

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.RowHeadersVisible = false;

            AgregarFilasBase();
        }

        void AgregarFilasBase()
        {
            dataGridView1.Rows.Clear();

            string[] cargos =
            {
                "Presidente", "Suplente Presidente", "Secretario", "Suplente Secretario", "Vocal 1", "Vocal 2", "Vocal 3"
            };

            foreach (string cargo in cargos)
            {
                dataGridView1.Rows.Add(cargo, "", "", "");
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.Enter)
            {
                this.SelectNextControl(this.ActiveControl, true, true, true, true);
                return true;
            }
            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void Form4_Load(object sender, EventArgs e)
        {
            comboBox2.Focus();
            ConfigurarDataGrid();

            
            dataGridView1.Columns["Cargo"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            dataGridView1.Columns["Cargo"].DefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.Columns["Cargo"].ReadOnly = true;

            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);

            dataGridView1.BackgroundColor = Color.AntiqueWhite;
            dataGridView1.BorderStyle = BorderStyle.None;

            dataGridView1.ClearSelection();
            dataGridView1.CurrentCell = dataGridView1.Rows[0].Cells[1];
            dataGridView1.Rows[0].Cells[1].Selected = true;

            if (idEditar != 0)
            {
                button1.Text = "Actualizar";
                CargarDatosEditar();
               
            }
            

            panel2.Click += panel2_Click;
            label21.Click += panel2_Click;
            pictureBox6.Click += panel2_Click;
            HoverPanel(panel2);
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

        private void panel2_Click(object sender, EventArgs e)
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
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();
                try
                {
                    TimeSpan? horaInicio = null;
                    TimeSpan? horaFin = null;

                    TimeSpan tempInicio;
                    TimeSpan tempFin;

                    if (!string.IsNullOrWhiteSpace(textBox7.Text))
                    {
                        if  (!TimeSpan.TryParse(textBox7.Text, out tempInicio))
                        {
                            MessageBox.Show("Formato de hora inicio inválido (HH:mm)");
                            return;
                        }
                        horaInicio = tempInicio;
                    }

                    if (!string.IsNullOrWhiteSpace(textBox2.Text))
                    {
                        if (!TimeSpan.TryParse(textBox2.Text, out tempFin))
                        {
                            MessageBox.Show("Formato de hora final inválido (HH:mm)");
                            return;
                        }
                        horaFin = tempFin;
                    }



                    if (comboBox2.SelectedIndex == -1)
                    {
                        MessageBox.Show("Selecciona el sector", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                    }

                    if (textBox1.Text == "")
                    {
                        MessageBox.Show("Ingresa el nombre", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                    }

                    int idVinculacion;

                    if (idEditar != 0)
                    {
                        
                        SqlCommand cmd = new SqlCommand(@"
                UPDATE Vinculacion SET
                sector=@sector,
                nombre=@nombre,
                direccion=@direccion,
                fecha=@fecha,
                hora_creacion=@horacre,
                hora_termino=@horaterm,
                coordinador=@coord,
                jefe_comunicacion=@jfecomunic,
                jefe_evaluacion=@jfeeval,
                jefe_primeros_auxilios=@jfepriaux,
                jefe_incendios=@jfeincend,
                observaciones=@obs,
                usuario_modificacion=@usumodif,
                fecha_modificacion=@fechamodif
                WHERE id_vinculacion=@id", conn);

                        cmd.Parameters.AddWithValue("@id", idEditar);
                        cmd.Parameters.AddWithValue("@sector", comboBox2.Text);
                        cmd.Parameters.AddWithValue("@nombre", textBox1.Text);
                        cmd.Parameters.AddWithValue("@direccion", textBox3.Text);
                        cmd.Parameters.AddWithValue("@fecha", dateTimePicker1.Value);
                        if (horaInicio == null)
                        {
                            cmd.Parameters.AddWithValue(
                            "@horacre",
                            DBNull.Value);
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue(
                            "@horacre",
                            horaInicio);
                        }

                        if (horaFin == null)
                        {
                            cmd.Parameters.AddWithValue(
                            "@horaterm",
                            DBNull.Value);
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue(
                            "@horaterm",
                            horaFin);
                        }
                        cmd.Parameters.AddWithValue("@coord", textBox26.Text);
                        cmd.Parameters.AddWithValue("@jfecomunic", textBox27.Text);
                        cmd.Parameters.AddWithValue("@jfeeval", textBox28.Text);
                        cmd.Parameters.AddWithValue("@jfepriaux", textBox29.Text);
                        cmd.Parameters.AddWithValue("@jfeincend", textBox30.Text);
                        cmd.Parameters.AddWithValue("@obs", textBox4.Text);

                        cmd.Parameters.AddWithValue("@usumodif", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechamodif", DateTime.Now);

                        cmd.ExecuteNonQuery();

                        idVinculacion = idEditar;

                        
                        SqlCommand delete = new SqlCommand(
                            "DELETE FROM Integrantes WHERE id_vinculacion=@id", conn);
                        delete.Parameters.AddWithValue("@id", idVinculacion);
                        delete.ExecuteNonQuery();
                    }
                    else
                    {
                        
                        SqlCommand cmd = new SqlCommand(@"
                           INSERT INTO Vinculacion
                           (sector, nombre, direccion, fecha, hora_creacion, hora_termino, coordinador,
                           jefe_comunicacion, jefe_evaluacion, jefe_primeros_auxilios, jefe_incendios, observaciones, usuario_creacion, fecha_creacion)
 
                           OUTPUT INSERTED.id_vinculacion

                           VALUES
                           (@sector, @nombre, @direccion, @fecha, @horacre, @horaterm, @coord, @jfecomunic, @jfeeval, @jfepriaux, @jfeincend, @obs, @usucre, @fechacre)", conn);

                        cmd.Parameters.AddWithValue("@sector", comboBox2.Text);
                        cmd.Parameters.AddWithValue("@nombre", textBox1.Text);
                        cmd.Parameters.AddWithValue("@direccion", textBox3.Text);
                        cmd.Parameters.AddWithValue("@fecha", dateTimePicker1.Value);
                        if(horaInicio == null)
                        {
                            cmd.Parameters.AddWithValue(
                            "@horacre",
                            DBNull.Value);
                        }
                         else
                        {
                            cmd.Parameters.AddWithValue(
                            "@horacre",
                            horaInicio);
                        }

                        if (horaFin == null)
                        {
                            cmd.Parameters.AddWithValue(
                            "@horaterm",
                            DBNull.Value);
                        }
                        else
                        {
                            cmd.Parameters.AddWithValue(
                            "@horaterm",
                            horaFin);
                        }
                        cmd.Parameters.AddWithValue("@coord", textBox26.Text);
                        cmd.Parameters.AddWithValue("@jfecomunic", textBox27.Text);
                        cmd.Parameters.AddWithValue("@jfeeval", textBox28.Text);
                        cmd.Parameters.AddWithValue("@jfepriaux", textBox29.Text);
                        cmd.Parameters.AddWithValue("@jfeincend", textBox30.Text);
                        cmd.Parameters.AddWithValue("@obs", textBox4.Text);

                        cmd.Parameters.AddWithValue("@usucre", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechacre", DateTime.Now);


                       

                        idVinculacion = (int)cmd.ExecuteScalar();
                    }



                    
                    foreach (DataGridViewRow fila in dataGridView1.Rows)
                    {
                        if (fila.IsNewRow) continue;

                        string cargo = fila.Cells["Cargo"].Value?.ToString();
                        string nombre = fila.Cells["Nombre"].Value?.ToString();
                        string telefono = fila.Cells["Telefono"].Value?.ToString();
                        string correo = fila.Cells["Correo"].Value?.ToString();

                        if (string.IsNullOrWhiteSpace(nombre)) continue;

                        SqlCommand cmd2 = new SqlCommand(@"
                INSERT INTO Integrantes
                (id_vinculacion, cargo, nombre, telefono, correo)
                VALUES (@id, @cargo, @nombre, @telefono, @correo)", conn);

                        cmd2.Parameters.AddWithValue("@id", idVinculacion);
                        cmd2.Parameters.AddWithValue("@cargo", cargo);
                        cmd2.Parameters.AddWithValue("@nombre", nombre);
                        cmd2.Parameters.AddWithValue("@telefono", telefono);
                        cmd2.Parameters.AddWithValue("@correo", correo);

                        cmd2.ExecuteNonQuery();
                    }

                    MessageBox.Show("Guardado correctamente", "Exito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    LimpiarCampos();
                    LimpiarDataGrid();

                    idEditar = 0; 
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button2_Click(object sender, EventArgs e)
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

        private void button3_Click(object sender, EventArgs e)
        {

        }

        void LimpiarCampos()
        {
            comboBox2.SelectedIndex = -1;
            textBox1.Clear();
            textBox3.Clear();
            textBox7.Clear();
            textBox2.Clear();
            textBox4.Clear();
            textBox26.Clear();
            textBox27.Clear();
            textBox28.Clear();
            textBox29.Clear();
            textBox30.Clear();
        }

        void LimpiarDataGrid()
        {
            dataGridView1.Rows.Clear();
            AgregarFilasBase();
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            LimpiarDataGrid();
            LimpiarCampos();
        }

        void CargarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = "SELECT * FROM Vinculacion WHERE id_vinculacion = @id";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", idEditar);

                SqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    comboBox2.Text = dr["sector"].ToString();
                    textBox1.Text = dr["nombre"].ToString();
                    textBox3.Text = dr["direccion"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(dr["fecha"]);
                    textBox7.Text = dr["hora_creacion"].ToString();
                    textBox2.Text = dr["hora_termino"].ToString();

                    textBox26.Text = dr["coordinador"].ToString();
                    textBox27.Text = dr["jefe_comunicacion"].ToString();
                    textBox28.Text = dr["jefe_evaluacion"].ToString();
                    textBox29.Text = dr["jefe_primeros_auxilios"].ToString();
                    textBox30.Text = dr["jefe_incendios"].ToString();

                    textBox4.Text = dr["observaciones"].ToString();
                }
                dr.Close();

                CargarIntegrantes(conn);
            }

            void CargarIntegrantes(SqlConnection conn)
            {
                string query = "SELECT * FROM Integrantes WHERE id_vinculacion = @id";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", idEditar);

                SqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    string cargo = dr["cargo"].ToString();

                    foreach (DataGridViewRow fila in dataGridView1.Rows)
                    {
                        if (fila.Cells["Cargo"].Value.ToString() == cargo)
                        {
                            fila.Cells["Nombre"].Value = dr["nombre"].ToString();
                            fila.Cells["Telefono"].Value = dr["telefono"].ToString();
                            fila.Cells["Correo"].Value = dr["correo"].ToString();
                            break;
                        }
                    }
                }
                dr.Close();
            }
        }

        void CargarDatosEditar()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"SELECT * FROM Vinculacion WHERE id_vinculacion = @id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", idEditar);

                SqlDataReader dr = cmd.ExecuteReader();

                if (dr.Read())
                {
                    comboBox2.Text = dr["sector"].ToString();
                    textBox1.Text = dr["nombre"].ToString();
                    textBox3.Text = dr["direccion"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(dr["fecha"]);
                    textBox7.Text = dr["hora_creacion"].ToString();
                    textBox2.Text = dr["hora_termino"].ToString();
                    textBox26.Text = dr["coordinador"].ToString();
                    textBox27.Text = dr["jefe_comunicacion"].ToString();
                    textBox28.Text = dr["jefe_evaluacion"].ToString();
                    textBox29.Text = dr["jefe_primeros_auxilios"].ToString();
                    textBox30.Text = dr["jefe_incendios"].ToString();
                    textBox4.Text = dr["observaciones"].ToString();
                }

                dr.Close();
            }

            CargarIntegrantesEditar();
        }


        void CargarIntegrantesEditar()
        {
            dataGridView1.Rows.Clear();

            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = @"SELECT cargo, nombre, telefono, correo 
                         FROM Integrantes 
                         WHERE id_vinculacion = @id";

                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", idEditar);

                SqlDataReader dr = cmd.ExecuteReader();

                while (dr.Read())
                {
                    dataGridView1.Rows.Add(
                        dr["cargo"].ToString(),
                        dr["nombre"].ToString(),
                        dr["telefono"].ToString(),
                        dr["correo"].ToString()
                    );
                }

                dr.Close();
            }
        }


       

        private void button4_Click(object sender, EventArgs e)
        {
            Form9 vistavinc = new Form9(this);
            vistavinc.Show();

            this.Hide();
        }

        private void label9_Click(object sender, EventArgs e)
        {

        }

        private void textBox26_TextChanged(object sender, EventArgs e)
        {
            if (actualizando) return;
            if (dataGridView1.CurrentRow.Cells["Cargo"].Value.ToString() == "Presidente")
            {
                actualizando = true;  
                dataGridView1.CurrentCell.Value = textBox26.Text;
                actualizando = false;
            }
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {

            if (dataGridView1.CurrentCell == null) return;
            if (dataGridView1.CurrentCell.OwningColumn.Name != "Nombre") return;
            if (actualizando) return;

                actualizando = true;
                textBox26.Text = dataGridView1.CurrentCell.Value?.ToString();
                actualizando = false;
        }
    }
}
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
    public partial class Form10 : Form
    {
        
        Form formularioAnterior;
        public Form10(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
        }

        private void Form10_Load(object sender, EventArgs e)
        {
            dataGridView1.EnableHeadersVisualStyles = false;

            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.DarkRed;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10, FontStyle.Bold);



            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.SteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.White;

            dataGridView1.BackgroundColor = Color.WhiteSmoke;
            dataGridView1.BorderStyle = BorderStyle.None;

            dataGridView1.AlternatingRowsDefaultCellStyle.BackColor = Color.AntiqueWhite;

            dataGridView1.RowHeadersVisible = false;
            dataGridView1.AllowUserToAddRows = false;

            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            dataGridView1.ScrollBars = ScrollBars.Both;


            panel2.Click += panel2_Click;
            label16.Click += panel2_Click;
            pictureBox6.Click += panel2_Click;

            HoverPanel(panel2);
            CargarUsuarios();

            dataGridView1.Columns["id_usuario"].HeaderText = "ID";
            dataGridView1.Columns["nombre_usuario"].HeaderText = "Usuario";
            dataGridView1.Columns["nombre_completo"].HeaderText = "Nombre";
            dataGridView1.Columns["rol"].HeaderText = "Rol del usuario";
            dataGridView1.Columns["estatus"].HeaderText = "Estatus";

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

        void CargarUsuarios()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();
                string query = "SELECT id_usuario, nombre_usuario, nombre_completo, rol, estatus FROM Usuarios";

                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                dataGridView1.DataSource = dt;

            }

        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (comboBox2.Text == "Inactivo")
            {
                MessageBox.Show("El usuario se encuentra inactivo", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            if (textBox5.Text != textBox1.Text)
            {
                MessageBox.Show("Verifica que las contraseñas sean las mismas", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            if (string.IsNullOrWhiteSpace(textBox2.Text) || string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox5.Text) ||
                    string.IsNullOrWhiteSpace(textBox1.Text) || string.IsNullOrWhiteSpace(comboBox1.Text) || string.IsNullOrWhiteSpace(comboBox2.Text))
            {
                MessageBox.Show("Completa todos los campos", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }


            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                if (idSeleccionado == 0)
                {


                    SqlCommand check = new SqlCommand("SELECT COUNT(*) FROM Usuarios WHERE nombre_usuario=@usuario", conn);
                    check.Parameters.AddWithValue("@usuario", textBox2.Text);

                    int existe = (int)check.ExecuteScalar();

                    if (existe > 0)
                    {
                        MessageBox.Show("El usuario ya existe", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                    }

                    string query = @"INSERT INTO Usuarios
                    (nombre_usuario, contrasena, nombre_completo, rol, estatus) VALUES (@usuario, @contrasena, @nombre, @rol, @estatus)";

                    SqlCommand cmd = new SqlCommand(query, conn);

                    cmd.Parameters.AddWithValue("@usuario", textBox2.Text);
                    cmd.Parameters.AddWithValue("@nombre", textBox4.Text);

                    string hash = Seguridad.Encriptar(textBox5.Text);
                    cmd.Parameters.AddWithValue("@contrasena", hash);
                    cmd.Parameters.AddWithValue("@rol", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@estatus", comboBox2.Text);

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Usuario guardado exitosamente", "¡MUY BIEN!", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    
                }
                else
                {
                    string update = @"UPDATE Usuarios SET 
                     nombre_completo=@nombre,
                     nombre_usuario=@usuario,
                     contrasena=@password,
                     rol=@rol,
                     estatus=@estatus
                     WHERE id_usuario=@id";

                    SqlCommand cmd = new SqlCommand(update, conn);
                    cmd.Parameters.AddWithValue("@id", idSeleccionado);
                    cmd.Parameters.AddWithValue("@nombre", textBox4.Text);
                    cmd.Parameters.AddWithValue("@usuario", textBox2.Text);
                    cmd.Parameters.AddWithValue("@password", textBox5.Text);
                    cmd.Parameters.AddWithValue("@rol", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@estatus", comboBox2.Text);

                    cmd.ExecuteNonQuery();
                    MessageBox.Show("Usuario actualizado exitosamente", "¡MUY BIEN!", MessageBoxButtons.OK, MessageBoxIcon.Information);

                }
            }



            LimpiarCampos();
            CargarUsuarios();


        }


       

        int idSeleccionado = 0;
       

        

        private void button4_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("Selecciona un usuario para eliminar", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }
            idSeleccionado = Convert.ToInt32(
                 dataGridView1.CurrentRow.Cells["id_usuario"].Value);
            DialogResult resultado = MessageBox.Show("¿Eliminar registro?", "Confirmar", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (resultado == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(Conexion.cadena))
                {

                    conn.Open();
                    SqlCommand cmd = new SqlCommand("UPDATE Usuarios SET estatus = 'Inactivo' WHERE id_usuario = @id", conn);
                    cmd.Parameters.AddWithValue("@id", idSeleccionado);

                    cmd.ExecuteNonQuery();
                }
                MessageBox.Show("Usuario eliminado correctamente", "¡ELIMINADO!", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
            CargarUsuarios();
            LimpiarCampos();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            LimpiarCampos();
        }

        void LimpiarCampos()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox2.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox1.Clear();
            idSeleccionado = 0;
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

        private void dataGridView1_SelectionChanged(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridView1.CurrentRow == null) return;

            if (dataGridView1.CurrentRow.Cells["id_usuario"].Value == null) return;

            idSeleccionado = Convert.ToInt32(
                dataGridView1.CurrentRow.Cells["id_usuario"].Value);

            textBox4.Text = dataGridView1.CurrentRow.Cells["nombre_completo"].Value.ToString();
            textBox2.Text = dataGridView1.CurrentRow.Cells["nombre_usuario"].Value.ToString();
            comboBox1.Text = dataGridView1.CurrentRow.Cells["rol"].Value.ToString();
            comboBox2.Text = dataGridView1.CurrentRow.Cells["estatus"].Value.ToString();

            button1.Text = "Actualizar";
        }
    }



}

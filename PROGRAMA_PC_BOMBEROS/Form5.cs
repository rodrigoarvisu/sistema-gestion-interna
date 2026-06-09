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
    public partial class Form5 : Form
    {
        
        
        public int idEditar = 0;
        Form formularioAnterior;
        public Form5(Form anterior)
        {
            InitializeComponent();
            formularioAnterior = anterior;
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
            LimpiarCampos();
        }

        void LimpiarCampos()
        {
            comboBox1.SelectedIndex = -1;
            comboBox2.SelectedIndex = -1;
            textBox3.Clear();
            textBox1.Clear();
            textBox8.Clear();
            comboBox3.SelectedIndex = -1;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox3.Text))
            {
               

                Uri uriResultado;
                bool valido = Uri.TryCreate(textBox3.Text, UriKind.Absolute, out uriResultado) 
                    && (uriResultado.Scheme == Uri.UriSchemeHttp || uriResultado.Scheme == Uri.UriSchemeHttps);

                if (!valido)
                {
                    MessageBox.Show("Ingresa un enlace valido (http/https)"); return;
                }
            }

            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Selecciona la red social", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("Selecciona el tipo de publicación", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            if (string.IsNullOrWhiteSpace(textBox8.Text))
            {
                MessageBox.Show("Ingresa la descripción de la publicación", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            button1.Enabled = false;

            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                try
                {
                    conn.Open();

                    string query; if (idEditar == 0) 
                    {

                        string validar = @"SELECT COUNT (*) FROM Difusiones WHERE enlace = @enlace AND fecha_publicacion = @fecha";
                        SqlCommand cmdVal = new SqlCommand(validar, conn);
                        cmdVal.Parameters.AddWithValue("@enlace", textBox3.Text);
                        cmdVal.Parameters.AddWithValue("@fecha", dateTimePicker1.Value.Date);

                        int existe = (int)cmdVal.ExecuteScalar();

                        if (existe > 0)
                        {
                            MessageBox.Show("Ya existe una publicacion con este enlace y fecha", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
                        }

                        // INSERT
                        query = @"INSERT INTO Difusiones
                        (red_social, tipo_publicacion, fecha_publicacion, enlace, responsable, descripcion, diseño_contenido,
                        usuario_creacion, fecha_creacion)
                        VALUES
                        (@red, @tipo, @fecha, @enlace, @responsable, @desc, @diseno, @usucreacion, @fechacreacion)";
                        }
                        else
                        {
                        // UPDATE
                        query = @"UPDATE Difusiones SET
                        red_social=@red,
                        tipo_publicacion=@tipo,
                        fecha_publicacion=@fecha,
                        enlace=@enlace,
                        responsable=@responsable,
                        descripcion=@desc,
                        diseño_contenido=@diseno,
                        usuario_modificacion=@usumodif,
                        fecha_modificacion=@fechamodif
                        WHERE id_difusion=@id";
                     }

                    SqlCommand cmd = new SqlCommand(query, conn);

                    cmd.Parameters.AddWithValue("@red", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@tipo", comboBox2.Text);
                    cmd.Parameters.AddWithValue("@fecha", dateTimePicker1.Value.Date);
                    cmd.Parameters.AddWithValue("@enlace", textBox3.Text);
                    cmd.Parameters.AddWithValue("@responsable", textBox1.Text);
                    cmd.Parameters.AddWithValue("@desc", textBox8.Text);
                    cmd.Parameters.AddWithValue("@diseno", comboBox3.Text);

                    if (idEditar == 0)
                    {
                        cmd.Parameters.AddWithValue("@usucreacion", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechacreacion", DateTime.Now);
                    }
                    else
                    {
                        cmd.Parameters.AddWithValue("@id", idEditar);
                        cmd.Parameters.AddWithValue("@usumodif", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechamodif", DateTime.Now);
                        
                    }
                    

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Guardado correctamente", "Exito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    LimpiarCampos();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
                finally
                {
                    button1.Enabled = true;
                }
            }

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form7 vista2 = new Form7(this);
            vista2.Show();
            this.Hide();

            
        }

        void CargarDatos()
        {
            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                conn.Open();

                string query = "SELECT * FROM Difusiones WHERE id_difusion = @id";
                SqlCommand cmd = new SqlCommand(query, conn);
                cmd.Parameters.AddWithValue("@id", idEditar);

                SqlDataReader reader = cmd.ExecuteReader();

                if (reader.Read())
                {
                    comboBox1.Text = reader["red_social"].ToString();
                    comboBox2.Text = reader["tipo_publicacion"].ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(reader["fecha_publicacion"]);
                    textBox3.Text = reader["enlace"].ToString();
                    textBox1.Text = reader["responsable"].ToString();
                    textBox8.Text = reader["descripcion"].ToString();
                    comboBox3.Text = reader["diseño_contenido"].ToString();
                }
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
        private void Form5_Load(object sender, EventArgs e)
        {
            panel2.Click += panel2_Click;
            label16.Click += panel2_Click;
            pictureBox6.Click += panel2_Click;
            HoverPanel(panel2);

            if (idEditar != 0)
            {
                button1.Text = "Actualizar";
                CargarDatos();
            }

           
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

        private void pictureBox6_Click(object sender, EventArgs e)
        {

        }

        private void label16_Click(object sender, EventArgs e)
        {

        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

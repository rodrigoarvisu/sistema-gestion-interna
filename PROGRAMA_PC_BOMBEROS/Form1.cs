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
    
    public partial class Form1 : Form
    {
        
        bool mostrar = false;
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {

            pictureBox2.Cursor = Cursors.Hand;
            this.AcceptButton = button1;
        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {
            

            mostrar = !mostrar;

            if (mostrar)
            {
                textBox2.UseSystemPasswordChar = false;
            }
            else
            {
                textBox2.UseSystemPasswordChar = true;
            }
        }

        

        private void button1_Click(object sender, EventArgs e)
        {

            string usuario = textBox1.Text.Trim();
            string contrasena = textBox2.Text.Trim();

            if (string.IsNullOrWhiteSpace(usuario) || string.IsNullOrWhiteSpace(contrasena))
            {
                MessageBox.Show("Ingresa usuario y contraseña", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                try
                {
                    conn.Open();
                    string query = @"SELECT id_usuario, nombre_usuario, nombre_completo, rol 
                                     FROM Usuarios 
                                     WHERE nombre_usuario = @usuario
                                     AND contrasena = @contrasena
                                     AND  estatus = 'Activo'";
                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@usuario", usuario);

                    string contraHash = Seguridad.Encriptar(contrasena);
                    cmd.Parameters.Add("@contrasena", SqlDbType.VarChar).Value = contraHash;

                    SqlDataReader reader = cmd.ExecuteReader();
                    
                    if (reader.Read())
                    {
                        Sesion.Usuario = reader["nombre_usuario"].ToString();
                        Sesion.Rol = reader["rol"].ToString().ToLower(); 
                        Sesion.NombreCompleto = reader["nombre_completo"].ToString();
                       
                        MessageBox.Show("Bienvenido " + Sesion.NombreCompleto, "Acceso Correcto", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        contrasena = "";
                        string nombre = reader["nombre_completo"].ToString();

                        Form2 form2 = new Form2(Sesion.NombreCompleto);
                        this.Hide();

                        form2.FormClosed += (s, args) => this.Close();
                        form2.Show();
                    }
                    else
                    {
                        MessageBox.Show("Usuario o contraseña incorrectos", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        
                        textBox1.Clear();
                        textBox1.Focus();
                        textBox2.Clear();
                    }
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

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            textBox2.UseSystemPasswordChar = true;
        }
    }
}

using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form3 : Form
    {
       
        public int idEditar = 0;
        Form formularioAnterior;
        public Form3(Form anterior)
        {
            InitializeComponent();
            
            formularioAnterior = anterior;
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

        private void Form3_Load(object sender, EventArgs e)
        {
            
        


        panel2.Click += panel2_Click;
            label16.Click += panel2_Click;
            pictureBox6.Click += panel2_Click;
            HoverPanel(panel2);

            groupBox4.Visible = false;
            textBox2.TextChanged += (s, e2) => CalcularTotal();
            textBox4.TextChanged += (s, e2) => CalcularTotal();

            if (idEditar != 0)
            {
                button1.Text = "Actualizar";
                

                using (SqlConnection conn = new SqlConnection(Conexion.cadena))
                {
                    conn.Open();

                    string query = "SELECT * FROM Capacitacion WHERE id_capacitacion = @id";

                    SqlCommand cmd = new SqlCommand(query, conn);
                    cmd.Parameters.AddWithValue("@id", idEditar);

                    SqlDataReader reader = cmd.ExecuteReader();

                    if (reader.Read())
                    {
                        textBox1.Text = reader["nombre_inmueble"].ToString();
                        comboBox1.Text = reader["tipo"].ToString();
                        textBox3.Text = reader["direccion"].ToString();

                        if (reader["fecha"] != DBNull.Value)
                            dateTimePicker1.Value = Convert.ToDateTime(reader["fecha"]);

                        checkBox13.Checked = reader["alertamiento"] != DBNull.Value &&
                                             Convert.ToBoolean(reader["alertamiento"]);

                        checkBox11.Checked = Convert.ToBoolean(reader["simulacro"]);

                        // 🔹 CAPACITACIONES
                        checkBox1.Checked = Convert.ToBoolean(reader["introduccion_proteccion_civil"]);
                        checkBox2.Checked = Convert.ToBoolean(reader["primeros_aux"]);
                        checkBox3.Checked = Convert.ToBoolean(reader["combate_incendios"]);
                        checkBox4.Checked = Convert.ToBoolean(reader["plan_emergencia"]);
                        checkBox5.Checked = Convert.ToBoolean(reader["evacuacion_inmuebles"]);
                        checkBox8.Checked = Convert.ToBoolean(reader["manejo_sustancias_quimicas"]);
                        checkBox9.Checked = Convert.ToBoolean(reader["busqueda_rescate"]);
                        checkBox10.Checked = Convert.ToBoolean(reader["practica_combate_contra_incendios"]);

                        checkBox7.Checked = Convert.ToBoolean(reader["uipc"]);
                        checkBox6.Checked = Convert.ToBoolean(reader["comite"]);

                        // 🔹 SIMULACRO
                        comboBox2.Text = reader["hipotesis"].ToString();
                        comboBox3.Text = reader["estatus"].ToString();

                        textBox2.Text = reader["poblacion_fija"].ToString();
                        textBox4.Text = reader["poblacion_flotante"].ToString();
                        textBox5.Text = reader["total"].ToString();

                        textBox7.Text = reader["hora_inicio"].ToString();
                        textBox9.Text = reader["hora_fin"].ToString();

                        textBox6.Text = reader["tiempo_evacuacion"].ToString();

                       

                        textBox8.Text = reader["observaciones"].ToString();

                        comboBox4.Text = reader["simulacro_gabinete"].ToString();

                        // 🔹 MOSTRAR SIMULACRO
                        groupBox4.Visible = checkBox11.Checked;

                        
                    }
                }
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

        private void panel2_Click(object sender,  EventArgs e)
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
            DialogResult resultado = MessageBox.Show(
                "¿Estás seguro que deseas salir?",
                "Confirmar salida",
                MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (resultado == DialogResult.Yes)
            {
                System.Windows.Forms.Application.Exit();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (textBox1.Text == "" || comboBox1.Text == "")
            {
                MessageBox.Show("Completa los campos obligatorios", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (checkBox11.Checked && comboBox2.Text == "")
            {
                MessageBox.Show("Selecciona una hipotesis", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            if (checkBox11.Checked && comboBox3.Text == "")
            {
                MessageBox.Show("Selecciona el estatus del simulacro", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            using (SqlConnection conn = new SqlConnection(Conexion.cadena))
            {
                try
                {
                    conn.Open();
                    string query;
                    if(idEditar == 0)
                    {
                        query = @"INSERT INTO Capacitacion
                    (nombre_inmueble, tipo, direccion, fecha, introduccion_proteccion_civil, primeros_aux, combate_incendios, plan_emergencia, 
                    evacuacion_inmuebles, manejo_sustancias_quimicas, busqueda_rescate, practica_combate_contra_incendios, simulacro,
                    uipc, comite, hipotesis, estatus, poblacion_fija, poblacion_flotante, total, hora_inicio, hora_fin, tiempo_evacuacion, alertamiento,
                    observaciones, simulacro_gabinete, usuario_creacion, fecha_creacion, usuario_modificacion, fecha_modificacion) VALUES
                    (@nombre, @tipo, @direccion, @fecha, @intro, @primaux, @combateincendios, @planemerg, @evacuacion, @sustquimicas, @busquedarescate,
                    @practicacombcontraincendios, @simulacro, @uipc, @comite, @hipotesis, @estatus, @fija, @flotante, @total, @horainicio, @horafin,
                    @tiempo, @alertamiento, @observaciones, @gabinete, @usucreacion, @fechacreacion, @usumodif, @fechamodif)";
                    }
                    else
                    {
                        query = @"UPDATE Capacitacion SET
                        nombre_inmueble=@nombre,
                        tipo=@tipo,
                        direccion=@direccion,
                        fecha=@fecha,
                        introduccion_proteccion_civil=@intro,
                        primeros_aux=@primaux,
                        combate_incendios=@combateincendios,
                        plan_emergencia=@planemerg,
                        evacuacion_inmuebles=@evacuacion,
                        manejo_sustancias_quimicas=@sustquimicas,
                        busqueda_rescate=@busquedarescate,
                        practica_combate_contra_incendios=@practicacombcontraincendios,
                        simulacro=@simulacro,
                        uipc=@uipc,
                        comite=@comite,
                        hipotesis=@hipotesis,
                        estatus=@estatus,
                        poblacion_fija=@fija,
                        poblacion_flotante=@flotante,
                        total=@total,
                        hora_inicio=@horainicio,
                        hora_fin=@horafin,
                        tiempo_evacuacion=@tiempo,
                        alertamiento=@alertamiento,
                        observaciones=@observaciones,
                        simulacro_gabinete=@gabinete,
                        usuario_modificacion=@usumodif,
                        fecha_modificacion=@fechamodif
                        WHERE id_capacitacion=@id";
                    }

                    SqlCommand cmd = new SqlCommand(query, conn);

                    cmd.Parameters.AddWithValue("@nombre", textBox1.Text);
                    cmd.Parameters.AddWithValue("@tipo", comboBox1.Text);
                    cmd.Parameters.AddWithValue("@direccion", textBox3.Text);
                    cmd.Parameters.AddWithValue("@fecha", dateTimePicker1.Value);
                    cmd.Parameters.AddWithValue("@intro", checkBox1.Checked);
                    cmd.Parameters.AddWithValue("@primaux", checkBox2.Checked);
                    cmd.Parameters.AddWithValue("@combateincendios", checkBox3.Checked);
                    cmd.Parameters.AddWithValue("@planemerg", checkBox4.Checked);
                    cmd.Parameters.AddWithValue("@evacuacion", checkBox5.Checked);
                    cmd.Parameters.AddWithValue("@sustquimicas", checkBox8.Checked);
                    cmd.Parameters.AddWithValue("@busquedarescate", checkBox9.Checked);
                    cmd.Parameters.AddWithValue("@practicacombcontraincendios", checkBox10.Checked);
                    cmd.Parameters.AddWithValue("@simulacro", checkBox11.Checked);
                    cmd.Parameters.AddWithValue("@uipc", checkBox7.Checked);
                    cmd.Parameters.AddWithValue("@comite", checkBox6.Checked);
                    cmd.Parameters.AddWithValue("@hipotesis", comboBox2.Text);
                    cmd.Parameters.AddWithValue("@estatus", comboBox3.Text);
                    int fija = 0;
                    int flotante = 0;
                    int total = 0;

                    int.TryParse(textBox2.Text, out fija);
                    int.TryParse(textBox4.Text, out flotante);
                    int.TryParse(textBox5.Text, out total);

                    cmd.Parameters.AddWithValue("@fija", fija);
                    cmd.Parameters.AddWithValue("@flotante", flotante);
                    cmd.Parameters.AddWithValue("@total", total);
                    cmd.Parameters.AddWithValue("@horainicio", textBox7.Text);
                    cmd.Parameters.AddWithValue("@horafin", textBox9.Text);
                    cmd.Parameters.AddWithValue("@tiempo", textBox6.Text);
                    cmd.Parameters.AddWithValue("@alertamiento", checkBox13.Checked);
                    cmd.Parameters.AddWithValue("@observaciones", textBox8.Text);
                    cmd.Parameters.AddWithValue("@gabinete", comboBox4.Text);
                    

                    if (idEditar == 0)
                    {
                        // INSERT
                        cmd.Parameters.AddWithValue("@usucreacion", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechacreacion", DateTime.Now);
                        cmd.Parameters.AddWithValue("@usumodif", DBNull.Value);
                        cmd.Parameters.AddWithValue("@fechamodif", DBNull.Value);
                    }
                    else
                    {
                        // UPDATE
                        cmd.Parameters.AddWithValue("@id", idEditar);
                        cmd.Parameters.AddWithValue("@usumodif", Sesion.Usuario);
                        cmd.Parameters.AddWithValue("@fechamodif", DateTime.Now);
                    }

                    

                    cmd.ExecuteNonQuery();

                    MessageBox.Show("Guardado correctamente", "Exito", MessageBoxButtons.OK, MessageBoxIcon.Information);

                    if (idEditar != 0)
                    {
                        this.Close();
                       
                    }
                    else
                    {
                        LimpiarCampos();
                        
                    }
                    
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error: " + ex.Message);
                }
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            textBox1.Clear();
            comboBox1.SelectedIndex = -1;
            textBox3.Clear();
            textBox2.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;

            checkBox7.Checked = false;
            checkBox6.Checked = false;

            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;
        }

        void LimpiarCampos()
        {
            textBox1.Clear();
            comboBox1.SelectedIndex = -1;
            textBox3.Clear();
            textBox2.Clear();
            textBox4.Clear();
            textBox5.Clear();
            textBox6.Clear();
            textBox7.Clear();
            textBox8.Clear();
            textBox9.Clear();

            checkBox1.Checked = false;
            checkBox2.Checked = false;
            checkBox3.Checked = false;
            checkBox4.Checked = false;
            checkBox5.Checked = false;
            checkBox8.Checked = false;
            checkBox9.Checked = false;
            checkBox10.Checked = false;
            checkBox11.Checked = false;
            checkBox12.Checked = false;
            checkBox13.Checked = false;

            checkBox7.Checked = false;
            checkBox6.Checked = false;

            comboBox2.SelectedIndex = -1;
            comboBox3.SelectedIndex = -1;
            comboBox4.SelectedIndex = -1;

            
        }

        private void CalcularTotal()
        {
            int fija = 0;
            int flotante = 0;

            int.TryParse(textBox2.Text, out fija);
            int.TryParse(textBox4.Text, out flotante);
            textBox5.Text = (fija + flotante).ToString();
        }










        private void groupBox4_Enter(object sender, EventArgs e)
        {

        }

        private void checkBox11_CheckedChanged(object sender, EventArgs e)
        {
            groupBox4.Visible = checkBox11.Checked;
        }

        private void label8_Click(object sender, EventArgs e)
        {

        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {

        }

        private void label10_Click(object sender, EventArgs e)
        {

        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {

        }

        private void label13_Click(object sender, EventArgs e)
        {

        }

        private void button4_Click(object sender, EventArgs e)
        {
            Form6 vista = new Form6(this);
            this.Hide();
            vista.Show();

        }

        private void checkBox12_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox7_CheckedChanged(object sender, EventArgs e)
        {
           
        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            
        }

        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}

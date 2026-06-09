using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PROGRAMA_PC_BOMBEROS
{
    public partial class Form2 : Form
    {
        string nombreUsuario;
        public Form2(string nombre)
        {
            InitializeComponent();
            nombreUsuario = nombre;
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            
        }

        Dictionary<string, string> formatos = new Dictionary<string, string>()
{
             { "ACTA CONSTITUTIVA", "ACTA CONSTITUTIVA DE LA INTEGRACIÓN DEL COMITÉ CIUDADANO DE PREVENCIÓN DE PROTECCIÓN CIVIL 2026 DEFINITIVA.pdf" },
             { "CONSENTIMIENTO MENORES", "CONSENTIMIENTO MENORES.pdf" },
             { "LISTA DE ASISTENCIA COORDINACIÓN DE CAPACITACIÓN, DIFUSIÓN Y VINCULACIÓN", "LISTA DE AISITENCIA COORDINACIÓN DE CAPACITACIÓN, DIFUSIÓN Y VINCULACIÓN.pdf" },
             { "LISTAS DE ASISTENCIA DE COMITE VECINAL", "LISTAS DE ASISTENCIA DE COMITE VECINAL.pdf" },
             { "HOJA MEMBRETADA", "HOJA MEMBRETADA.docx" }
             };

        private void Form2_Load(object sender, EventArgs e)
        {
            

            label1.Text = "Bienvenido " + nombreUsuario;

            panel2.Click += panel2_Click;
            label2.Click += panel2_Click;
            pictureBox2.Click += panel2_Click;
            pictureBox6.Click += panel2_Click;

            panel3.Click += panel3_Click;
            label3.Click += panel3_Click;
            pictureBox3.Click += panel3_Click;
            pictureBox7.Click += panel3_Click;

            panel4.Click += panel4_Click;
            label4.Click += panel4_Click;
            pictureBox4.Click += panel4_Click;
            pictureBox8.Click += panel4_Click;

            panel5.Click += panel5_Click;
            label5.Click += panel5_Click;
            pictureBox5.Click += panel5_Click;

            panel6.Click += panel6_Click;
            label11.Click += panel6_Click;
            pictureBox10.Click += panel6_Click;

            panel7.Click += panel7_Click;
            label13.Click += panel7_Click;
            pictureBox11.Click += panel7_Click;
            pictureBox9.Click += panel7_Click;

            panel9.Click += panel9_Click;
            label15.Click += panel9_Click;
            pictureBox14.Click += panel9_Click;
            pictureBox15.Click += panel9_Click;

            comboBox1.Items.AddRange(formatos.Keys.ToArray());


            label16.Text = DateTime.Now.ToString("dddd, dd 'de' MMMM 'de' yyyy");

            HoverPanel(panel2);
            HoverPanel(panel3);
            HoverPanel(panel4);
            HoverPanel(panel5);
            HoverPanel(panel6);
            HoverPanel(panel7);
            HoverPanel(panel9);
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
        private void panel9_Click(object sender, EventArgs e)
        {
            Form8 dash = new Form8(Sesion.Usuario, this);
            this.Hide();
            dash.Show();
            
        }

        private void panel2_Click(object sender, EventArgs e)
        {
            Form3 cap = new Form3(this);
            this.Hide();
            cap.Show();
        }

        

        private void panel6_Click(object sender, EventArgs e)
        {
            if (comboBox1.SelectedIndex == -1)
            {
                MessageBox.Show("Selecciona un formato", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Warning); return;
            }

            string seleccion = comboBox1.SelectedItem.ToString();

            if (string.IsNullOrEmpty(seleccion))
            {
                MessageBox.Show("Seleccion invalida"); return;
            }

            if (!formatos.ContainsKey(seleccion))
            {
                MessageBox.Show("Formato no válido");
                return;
            }

            string rutaArchivo = @"Formatos\" + formatos[seleccion];

            
            string rutaCompleta = Path.Combine(Application.StartupPath, rutaArchivo);

            if (!File.Exists(rutaCompleta))
            {
                MessageBox.Show("El archivo no existe en la carpeta Formatos"); return;   
            }

            SaveFileDialog save = new SaveFileDialog();
            save.FileName = Path.GetFileName(rutaCompleta);

            if (save.ShowDialog() == DialogResult.OK)
            {
                File.Copy(rutaCompleta, save.FileName, true);
                MessageBox.Show("Archivo descargado correctamente", "¡EXCLENTE!", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }


        }

        private void panel3_Click(object sender, EventArgs e)
        {
            Form4 vinc = new Form4(this);
            this.Hide();
            vinc.Show();
        }

        private void panel4_Click(object sender, EventArgs e)
        {
            Form5 dif = new Form5(this);
            this.Hide(); 
            dif.Show();
        }

        private void Dif_FormClosed(object sender, FormClosedEventArgs e)
        {
            throw new NotImplementedException();
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

        private void panel7_Click(object sender, EventArgs e)
        {
            if (Sesion.Rol != "admin")
            {
                MessageBox.Show("No tienes permisos para acceder a este módulo, contacta al administrador", "¡AVISO!", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
            {


                Form10 us = new Form10(this);
                this.Hide();
                us.Show();
            }
        }
        private void panel2_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox2_Click(object sender, EventArgs e)
        {

        }

        private void panel6_Paint(object sender, PaintEventArgs e)
        {

        }

        private void label2_Click(object sender, EventArgs e)
        {

        }

        private void panel3_Paint(object sender, PaintEventArgs e)
        {

        }

        private void pictureBox3_Click(object sender, EventArgs e)
        {

        }
    }
}

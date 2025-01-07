using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pantalla_De_Control
{
    public partial class MessageBoxCustom : Form
    {
        private bool isDragging = false;
        private System.Drawing.Point dragStart;
        public delegate void EnviarVariableDelegate4();
        public event EnviarVariableDelegate4 EnviarVariableEvent4;
        public MessageBoxCustom()
        {
            InitializeComponent();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void flowLayoutPanel1_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                // Obtén la posición actual del cursor en pantalla
                Point currentScreenPos = Cursor.Position;

                // Calcula la nueva posición del formulario
                this.Left += currentScreenPos.X - dragStart.X;
                this.Top += currentScreenPos.Y - dragStart.Y;

                // Actualiza la posición inicial para el próximo movimiento
                dragStart = currentScreenPos;
            }
        }

        private void flowLayoutPanel1_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;

                // Guarda la posición actual del mouse en pantalla
                dragStart = Cursor.Position;
            }
        }

        private void flowLayoutPanel1_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false; // Detén el arrastre
            }
        }

        private void Ajuste_Click(object sender, EventArgs e)
        {
            ControlAcceso control = new ControlAcceso();
            control.EnviarVariableEvent3 += new ControlAcceso.EnviarVariableDelegate3(ejecutar);
            control.ShowDialog();  
        }
        public void ejecutar()
        {
            if (GlobalSettings.Instance.aceptado == true)
            {
                EnviarVariableEvent4();
                this.Close();
            }
        }

        private void GridExistencia_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void MessageBoxCustom_Load(object sender, EventArgs e)
        {

        }
    }
}

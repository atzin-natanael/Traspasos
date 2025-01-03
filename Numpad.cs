using System;
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
    public partial class Numpad : Form
    {
        public delegate void EnviarVariableDelegate2(decimal cantidad);
        public event EnviarVariableDelegate2 EnviarVariableEvent2;
        public Numpad()
        {
            InitializeComponent();
        }
        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void N7_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "7";
        }

        private void N8_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "8";
        }

        private void N9_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "9";
        }

        private void N4_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "4";
        }

        private void N5_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "5";
        }

        private void N6_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "6";
        }

        private void N1_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "1";
        }

        private void N2_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "2";
        }

        private void N3_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "3";
        }

        private void OK_Click(object sender, EventArgs e)
        {
            if(Cantidad.Text != string.Empty && Cantidad.Text != "0"){
                EnviarVariableEvent2(decimal.Parse(Cantidad.Text));
                this.Close();
            
            }
            else{
                MessageBox.Show("Ingrese una cantidad valida");
                Cantidad.Focus();
                Cantidad.Select(0, Cantidad.TextLength);
            }
        }

        private void N0_Click(object sender, EventArgs e)
        {
            Cantidad.Text += "0";
        }

        private void Cantidad_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                OK.Focus();
            }
        }

        private void Numpad_Load(object sender, EventArgs e)
        {
            Cantidad.Focus();
            Cantidad.Select(0, Cantidad.TextLength);
        }
    }
}

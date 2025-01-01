using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection.Emit;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pantalla_De_Control
{
    public partial class Mensajes : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);
        public Mensajes()
        {
            InitializeComponent();
        }

        private void Aceptar_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Close_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Mensajes_Load(object sender, EventArgs e)
        {
            int formWidth = this.Width;

            // Obtén el tamaño del label
            int labelWidth = Texto.Width;

            // Calcula la nueva posición para centrar el label horizontalmente
            int x = (formWidth - labelWidth) / 2;

            // Establece la nueva posición del label
            Texto.Location = new Point(x, Texto.Location.Y);
            Aceptar.Focus();
        }
    }
}

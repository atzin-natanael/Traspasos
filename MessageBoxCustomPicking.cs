using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Pantalla_De_Control
{
    public partial class MessageBoxCustomPicking : Form
    {
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);
        public delegate void EnviarVariableDelegate6();
        public event EnviarVariableDelegate6 EnviarVariableEvent6;
        public MessageBoxCustomPicking()
        {
            InitializeComponent();
            GridPicking.ClearSelection();
            GridPicking.CurrentCell = null;
            Permiso.Focus();
            Permiso.Select();
        }

        private void OK_Click(object sender, EventArgs e)
        {
            this.Close();
            GlobalSettings.Instance.aceptadoP = false;
        }

        private void Permiso_Click(object sender, EventArgs e)
        {
            GlobalSettings.Instance.BanderaPicking = true;
            ControlAcceso control = new ControlAcceso();
            control.EnviarVariableEvent5 += new ControlAcceso.EnviarVariableDelegate5(aceptar);
            control.ShowDialog();
        }
        public void aceptar()
        {
            if (GlobalSettings.Instance.aceptadoP == true)
            {
                EnviarVariableEvent6();
                this.Close();
            }
        }

        private void panel1_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void MessageBoxCustomPicking_Load(object sender, EventArgs e)
        {
            GridPicking.ClearSelection();
            GridPicking.CurrentCell = null;
            Permiso.Focus();
            Permiso.Select();
        }
    }
}

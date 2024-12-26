using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Pantalla_De_Control
{
    public partial class SinCodigo : Form
    {
        public delegate void EnviarArticulo(int articulo);
        public event EnviarArticulo EnviarArticuloEvent;
        public SinCodigo()
        {
            InitializeComponent();
            ATienda();
        }
        [DllImport("user32.DLL", EntryPoint = "ReleaseCapture")]
        private extern static void ReleaseCapture();
        [DllImport("user32.DLL", EntryPoint = "SendMessage")]
        private extern static void SendMessage(System.IntPtr hwnd, int wmsg, int wparam, int lparam);
        public void ATienda()
        {
            LbTitulo.Text = "Articulos de Tienda";
            P6.Visible = false;
            PT6.Visible = false;
            Pb6.Visible = false;
            Cb1.Items.Clear();
            Cb2.Items.Clear();
            Cb3.Items.Clear();
            Cb4.Items.Clear();
            Cb5.Items.Clear();
            Cb1.Visible = false;
            Pb1.Image = Image.FromFile("C:\\Img\\imprenta.png");
            Lb1.Text = "Papel Imprenta";
            Lb2.Text = "Papel Manila";
            Pb2.Image = Image.FromFile("C:\\Img\\manila.jpg");
            Cb2.Items.Add("Elegir modelo:");
            Cb2.Items.Add("AMARILLO");
            Cb2.Items.Add("ROJO");
            Cb2.Items.Add("AZUL");
            Cb2.Items.Add("VERDE");
            Lb3.Text = "Foamy Guine";
            Pb3.Image = Image.FromFile("C:\\Img\\foamy.jpg");
            Cb3.Items.Add("Elegir modelo:");
            Cb3.Items.Add("CARTA");
            Cb3.Items.Add("PLIEGO CORTO");
            Lb4.Text = "Papel Picado";
            Pb4.Image = Image.FromFile("C:\\Img\\picado.jpg");
            Cb4.Items.Add("Elegir modelo:");
            Cb4.Items.Add("NORMAL");
            Cb4.Items.Add("COMPRIMIDO");
            Lb5.Text = "Cartoncillo";
            Pb5.Image = Image.FromFile("C:\\Img\\cartoncillo.jpg");
            Cb5.Items.Add("Elegir modelo:");
            Cb5.Items.Add("BLANCO");
            Cb5.Items.Add("NEGRO");
            Cb5.Items.Add("CANARIO");
            Cb5.Items.Add("NARANJA");
            Cb5.Items.Add("ROJO");
            Cb5.Items.Add("ROSA");
            Cb5.Items.Add("VERDE BANDERA");
            Cb2.SelectedItem = 0;
            Cb3.SelectedIndex = 0;
            Cb4.SelectedIndex = 0;
            Cb5.SelectedIndex = 0;
        }
        private void Tienda_Click(object sender, EventArgs e)
        {
            ATienda();
        }

        private void Regalos_Click(object sender, EventArgs e)
        {
            Cb1.Visible = true;
            P6.Visible = true;
            PT6.Visible = true;
            Pb6.Visible = true;
            Cb1.Items.Clear();
            Cb2.Items.Clear();
            Cb3.Items.Clear();
            Cb4.Items.Clear();
            Cb5.Items.Clear();
            Cb6.Items.Clear();
            Cb5.Visible = false;
            LbTitulo.Text = "Articulos de Regalos";
            Pb1.Image = Image.FromFile("C:\\Img\\raf mono.jpg");
            Lb1.Text = "Monografias";
            Cb1.Items.Add("Elegir modelo:");
            Cb1.Items.Add("RAF");
            Cb1.Items.Add("SUN RISE");
            Lb2.Text = "Esferas Unicel";
            Pb2.Image = Image.FromFile("C:\\Img\\esferas.jpg");
            Cb2.Items.Add("Elegir modelo:");
            Cb2.Items.Add("0");
            Cb2.Items.Add("00");
            Cb2.Items.Add("1");
            Cb2.Items.Add("1A");
            Cb2.Items.Add("2");
            Cb2.Items.Add("3");
            Cb2.Items.Add("4");
            Cb2.Items.Add("5");
            Cb2.Items.Add("6");
            Cb2.Items.Add("7");
            Cb2.Items.Add("8");
            Cb2.Items.Add("9");
            Cb2.Items.Add("10");
            Cb2.Items.Add("11");
            Cb2.Items.Add("12");
            Cb2.Items.Add("13");
            Cb2.SelectedIndex = 0;
            Lb3.Text = "Placas";
            Pb3.Image = Image.FromFile("C:\\Img\\placas.jpg");
            Cb3.Items.Add("Elegir modelo:");
            Cb3.Items.Add("50 X 25");
            Cb3.Items.Add("50 X 50");
            Cb3.Items.Add("100 X 50");
            Cb3.Items.Add("100 X 100");
            Lb4.Text = "Pelotas";
            Pb4.Image = Image.FromFile("C:\\Img\\pelotas.png");
            Cb4.Items.Add("Elegir modelo:");
            Cb4.Items.Add("DECORADA");
            Cb4.Items.Add("HEAVY METAL");
            Cb4.Items.Add("LISA");
            Lb5.Text = "Biografias";
            Pb5.Image = Image.FromFile("C:\\Img\\biografias.jpg");
            Lb6.Text = "Cromos Raf";
            Pb6.Image = Image.FromFile("C:\\Img\\cromos.jpg");
            Cb6.Items.Add("Elegir modelo:");
            Cb6.Items.Add("CROMO NO. 6");
            Cb6.Items.Add("CROMO NO. 10");
            Cb1.SelectedIndex = 0;
            Cb2.SelectedItem = 0;
            Cb3.SelectedIndex = 0;
            Cb4.SelectedIndex = 0;
            Cb6.SelectedIndex = 0;
        }

        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel2_MouseDown(object sender, MouseEventArgs e)
        {
            ReleaseCapture();
            SendMessage(this.Handle, 0x112, 0xf012, 0);
        }

        private void Pb1_Click(object sender, EventArgs e)
        {
            if(LbTitulo.Text == "Articulos de Tienda")
            {
                this.Close();
                EnviarArticuloEvent(248450);
            }
            else if(LbTitulo.Text == "Articulos de Regalos")
            {
                switch (Cb1.Text)
                {
                    case "RAF":
                        this.Close();
                        EnviarArticuloEvent(184012);
                        break;
                    case "SUN RISE":
                        this.Close();
                        EnviarArticuloEvent(186003);
                        break;
                    default:
                        MessageBox.Show("Operación no válida");
                        break;
                }
            }
        }

        private void Pb2_Click(object sender, EventArgs e)
        {
            if (LbTitulo.Text == "Articulos de Tienda")
            {
                if(Cb2.Text != string.Empty && Cb2.Text != "Elegir color:")
                {
                    switch (Cb2.Text)
                    {
                        case "AMARILLO":
                            this.Close();
                            EnviarArticuloEvent(248118);
                            break;
                        case "AZUL":
                            this.Close();
                            EnviarArticuloEvent(248120);
                            break;
                        case "ROJO":
                            this.Close();
                            EnviarArticuloEvent(248122);
                            break;
                        case "VERDE":
                            this.Close();
                            EnviarArticuloEvent(248116);
                            break;
                        default:
                            MessageBox.Show("Operación no válida");
                            break;
                    }
                }
            }
            else if (LbTitulo.Text == "Articulos de Regalos")
            {
                if (Cb2.Text != string.Empty && Cb2.Text != "Elegir color:")
                {
                    switch (Cb2.Text)
                    {
                        case "0":
                            this.Close();
                            EnviarArticuloEvent(306002);
                            break;
                        case "00":
                            this.Close();
                            EnviarArticuloEvent(306004);
                            break;
                        case "1":
                            this.Close();
                            EnviarArticuloEvent(306006);
                            break;
                        case "1A":
                            this.Close();
                            EnviarArticuloEvent(306008);
                            break;
                        case "2":
                            this.Close();
                            EnviarArticuloEvent(306010);
                            break;
                        case "3":
                            this.Close();
                            EnviarArticuloEvent(306012);
                            break;
                        case "4":
                            this.Close();
                            EnviarArticuloEvent(306014);
                            break;
                        case "5":
                            this.Close();
                            EnviarArticuloEvent(306016);
                            break;
                        case "6":
                            this.Close();
                            EnviarArticuloEvent(306020);
                            break;
                        case "7":
                            this.Close();
                            EnviarArticuloEvent(306022);
                            break;
                        case "8":
                            this.Close();
                            EnviarArticuloEvent(306024);
                            break;
                        case "9":
                            this.Close();
                            EnviarArticuloEvent(306026);
                            break;
                        case "10":
                            this.Close();
                            EnviarArticuloEvent(306028);
                            break;
                        case "11":
                            this.Close();
                            EnviarArticuloEvent(306030);
                            break;
                        case "12":
                            this.Close();
                            EnviarArticuloEvent(306032);
                            break;
                        case "13":
                            this.Close();
                            EnviarArticuloEvent(306034);
                            break;
                        default:
                            MessageBox.Show("Operación no válida");
                            break;
                    }
                }
            }
        }

        private void Pb3_Click(object sender, EventArgs e)
        {
            if (LbTitulo.Text == "Articulos de Tienda")
            {
                switch (Cb3.Text)
                {
                    case "CARTA":
                        this.Close();
                        EnviarArticuloEvent(292006);
                        break;
                    case "PLIEGO CORTO":
                        this.Close();
                        EnviarArticuloEvent(292012);
                        break;
                    default:
                        MessageBox.Show("Operación no válida");
                        break;
                }
            }
            else if (LbTitulo.Text == "Articulos de Regalos")
            {
                if (Cb3.Text != string.Empty && Cb3.Text != "Elegir color:")
                {
                    switch (Cb3.Text)
                    {
                        case "50 X 25":
                            this.Close();
                            EnviarArticuloEvent(306046);
                            break;
                        case "50 X 50":
                            this.Close();
                            EnviarArticuloEvent(306044);
                            break;
                        case "100 X 50":
                            this.Close();
                            EnviarArticuloEvent(306042);
                            break;
                        case "100 X 100":
                            this.Close();
                            EnviarArticuloEvent(306040);
                            break;
                        default:
                            MessageBox.Show("Operación no válida");
                            break;
                    }
                }
            }
        }

        private void Pb4_Click(object sender, EventArgs e)
        {
            if (LbTitulo.Text == "Articulos de Tienda")
            {
                switch (Cb4.Text)
                {
                    case "NORMAL":
                        this.Close();
                        EnviarArticuloEvent(640008);
                        break;
                    case "COMPRIMIDO":
                        this.Close();
                        EnviarArticuloEvent(640002);
                        break;
                    default:
                        MessageBox.Show("Operación no válida");
                        break;
                }
            }
            else if (LbTitulo.Text == "Articulos de Regalos")
            {
                if (Cb4.Text != string.Empty && Cb4.Text != "Elegir color:")
                {
                    switch (Cb4.Text)
                    {
                        case "DECORADA":
                            this.Close();
                            EnviarArticuloEvent(195320);
                            break;
                        case "HEAVY METAL":
                            this.Close();
                            EnviarArticuloEvent(195316);
                            break;
                        case "LISA":
                            this.Close();
                            EnviarArticuloEvent(195318);
                            break;
                        default:
                            MessageBox.Show("Operación no válida");
                            break;
                    }
                }
            }
        }

        private void Pb5_Click(object sender, EventArgs e)
        {
            if (LbTitulo.Text == "Articulos de Tienda")
            {
                switch (Cb5.Text)
                {
                    case "BLANCO":
                        this.Close();
                        EnviarArticuloEvent(248100);
                        break;
                    case "NEGRO":
                        this.Close();
                        EnviarArticuloEvent(248114);
                        break;
                    case "CANARIO":
                        this.Close();
                        EnviarArticuloEvent(248104);
                        break;
                    case "NARANJA":
                        this.Close();
                        EnviarArticuloEvent(248108);
                        break;
                    case "ROJO":
                        this.Close();
                        EnviarArticuloEvent(248112);
                        break;
                    case "ROSA":
                        this.Close();
                        EnviarArticuloEvent(248110);
                        break;
                    case "VERDE BANDERA":
                        this.Close();
                        EnviarArticuloEvent(248102);
                        break;
                    case "AZUL":
                        this.Close();
                        EnviarArticuloEvent(248106);
                        break;
                    default:
                        MessageBox.Show("Operación no válida");
                        break;
                }
            }
            else if (LbTitulo.Text == "Articulos de Regalos")
            {
                    this.Close();
                    EnviarArticuloEvent(184002);
            }
        }

        private void Pb6_Click(object sender, EventArgs e)
        {
            if (LbTitulo.Text == "Articulos de Regalos")
            {
                switch (Cb6.Text)
                {
                    case "CROMO NO. 6":
                        this.Close();
                        EnviarArticuloEvent(184018);
                        break;
                    case "CROMO NO. 10":
                        this.Close();
                        EnviarArticuloEvent(184020);
                        break;
                    default:
                        MessageBox.Show("Operación no válida");
                        break;
                }
            }
        }
    }
}

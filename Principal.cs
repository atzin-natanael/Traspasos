using FirebirdSql.Data.FirebirdClient;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.Common;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ApiBas = ApisMicrosip.ApiMspBasicaExt;
using ApiInv = ApisMicrosip.ApiMspInventExt;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using System.CodeDom.Compiler;
using System.Reflection;
using System.Security.Policy;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using DocumentFormat.OpenXml.Drawing.Charts;
using System.Drawing.Printing;
using static System.Net.Mime.MediaTypeNames;
using DocumentFormat.OpenXml.Drawing;

namespace Pantalla_De_Control
{
    public partial class Principal : Form
    {

        private List<string> nombresValor = new List<string>();

        private List<string> nombresArray = new List<string>();

        MessageBoxCustom mensaje = new MessageBoxCustom();
        MessageBoxCustomPicking mensajePicking = new MessageBoxCustomPicking();


        private bool isDragging = false; // Para controlar si el mouse está presionado
        private System.Drawing.Point startPoint; // Para guardar la posición inicial del clic del mouse
        public Principal()
        {
            InitializeComponent();
            Config();
            Leer_Datos();
            Cb_Surtidor.SelectedIndex = -1;
            Cb_Surtidor.DropDownHeight = 250;
            Pedido.Focus();
            Pedido.Select();
            CargarExcel();
            GlobalSettings.Instance.ContadorIncompletos = 0;
            GlobalSettings.Instance.ContadorExcedentes = 0;


        }
        public void CargarExcel()
        {
            string filePath = "\\\\SRVPRINCIPAL\\clavesSurtido\\Picking Almacen.xlsx";
           //string filePath = "C:\\clavesSurtido\\Claves.xlsx";
            using (var workbook = new XLWorkbook(filePath))
            {
                // Seleccionar la primera hoja del archivo
                var worksheet = workbook.Worksheet(1);
                int filas = worksheet.LastRowUsed().RowNumber();
                //List<List<string>> Picking = new List<List<string>>();
                for (int i = 2; i < filas + 1; ++i)
                {
                    List<string> temp = new List<string>();
                    temp.Add(worksheet.Row(i).Cell(1).Value.ToString());
                    temp.Add(worksheet.Row(i).Cell(4).Value.ToString());
                    temp.Add(worksheet.Row(i).Cell(6).Value.ToString());
                    GlobalSettings.Instance.Picking.Add(temp);
                }
            }


        }
            //Cb_Empacador.Text = "";
        public void Leer_Datos()
        {
            nombresArray.Clear();
            nombresValor.Clear();
            //string filePath = "\\\\192.168.0.2\\C$\\clavesSurtido\\Claves.xlsx";

            string filePath = "\\\\SRVPRINCIPAL\\clavesSurtido\\Claves.xlsx";
            // string filePath = "C:\\clavesSurtido\\Claves.xlsx";
            using (var workbook = new XLWorkbook(filePath))
            {
                // Seleccionar la primera hoja del archivo
                var worksheet = workbook.Worksheet(1);
                int filas = worksheet.LastRowUsed().RowNumber();
                for (int i = 2; i < filas + 1; ++i)
                {
                    string temp_name = worksheet.Row(i).Cell(1).Value.ToString();
                    string temp_value = worksheet.Row(i).Cell(2).Value.ToString();
                    string temp_status = worksheet.Row(i).Cell(3).Value.ToString();

                    string name = temp_name + " " + temp_value;

                    nombresArray.Add(name);
                    nombresValor.Add(temp_status);
                }
            }
            Cb_Surtidor.DataSource = nombresArray;
            Cb_Surtidor.AutoCompleteMode = AutoCompleteMode.Append;
            Cb_Surtidor.AutoCompleteSource = AutoCompleteSource.CustomSource;
            Cb_Surtidor.AutoCompleteCustomSource.AddRange(nombresArray.ToArray());
            //Cb_Empacador.Text = "";
        }
        private int ConectaBD()
        {

            ApiBas.SetErrorHandling(0, 0);
            if (GlobalSettings.Instance.Bd == 0)
                GlobalSettings.Instance.Bd = ApiBas.NewDB();
            //Objeto transaccion
            //TrnHandle
            GlobalSettings.Instance.Trn = ApiBas.NewTrn(GlobalSettings.Instance.Bd, 3);
            string path = "192.168.0.11:D:\\Microsip datos\\PAPELERIA CORIBA CORNEJO.fdb";
            //string path = "C:\\Microsip datos\\PAPELERIA CORIBA CORNEJO.fdb";

            int conecta = ApiBas.DBConnect(GlobalSettings.Instance.Bd, path, "SYSDBA", "C0r1b423");
            //int conecta = ApiBas.DBConnect(GlobalSettings.Instance.Bd, path, "SYSDBA", "masterkey");
            StringBuilder obtieneError = new StringBuilder(1000);
            int codigoError = ApiBas.GetLastErrorMessage(obtieneError);
            String mensajeError = codigoError.ToString();
            if (codigoError > 0)
            {
                MessageBox.Show(obtieneError.ToString());
                return 0;
            }
            else
            {
                //btnFacturar.Enabled = true;
                return 1;
            }

        }
        public void Config()
        {
            string filePath = "C:\\ConfigDB\\DB.txt"; // Ruta de tu archivo de texto
            List<string> lineas = new List<string>();

            // Verificar si el archivo existe
            if (File.Exists(filePath))
            {
                // Leer todas las líneas del archivo
                using (StreamReader sr = new StreamReader(filePath))
                {
                    string linea;
                    while ((linea = sr.ReadLine()) != null)
                    {
                        GlobalSettings.Instance.Config.Add(linea);
                    }

                }
                GlobalSettings.Instance.Ip = GlobalSettings.Instance.Config[0];
                GlobalSettings.Instance.Puerto = GlobalSettings.Instance.Config[1];
                GlobalSettings.Instance.Direccion = GlobalSettings.Instance.Config[2];
                GlobalSettings.Instance.User = GlobalSettings.Instance.Config[3];
                GlobalSettings.Instance.Pw = GlobalSettings.Instance.Config[4];
            }

        }
        public string Descripcion(string articulo_id)
        {
            FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
            try
            {
                con.Open();
                string query = "SELECT * FROM ARTICULOS WHERE ARTICULO_ID = '" + articulo_id + "';";
                FbCommand command = new FbCommand(query, con);
                FbDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    string descripcion = reader.GetString(1);
                    return descripcion;
                }
                else
                {
                    return "articulo no encontrado";
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(ex.ToString());
                return "articulo no encontrado";
            }
            finally
            {
                con.Close();
            }
        }
        public int Clave_Principal(string articulo_id)
        {
            FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
            try
            {
                con.Open();
                string query = "SELECT * FROM CLAVES_ARTICULOS WHERE ARTICULO_ID = '" + articulo_id + "' AND ROL_CLAVE_ART_ID = '17';";
                FbCommand command = new FbCommand(query, con);
                FbDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    int clave_principal = reader.GetInt32(1);
                    return clave_principal;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(ex.ToString());
                return 0;
            }
            finally
            {
                con.Close();
            }

        }
        public int Art_Id(string clave_principal)
        {
            FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
            try
            {
                con.Open();
                string query = "SELECT ARTICULO_ID FROM CLAVES_ARTICULOS WHERE CLAVE_ARTICULO = '" + clave_principal+ "';";
                FbCommand command = new FbCommand(query, con);
                FbDataReader reader = command.ExecuteReader();
                if (reader.Read())
                {
                    int art_id = reader.GetInt32(0);
                    return art_id;
                }
                else
                {
                    return 0;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(ex.ToString());
                return 0;
            }
            finally
            {
                con.Close();
            }

        }
        private void Ingresar_Click(object sender, EventArgs e)
        {
            if(Codigo.Text != string.Empty && Pedido.Text != string.Empty && Cb_Surtidor.Text != string.Empty)
            {
                FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
                try
                {
                    con.Open();
                    string query = "SELECT * FROM CLAVES_ARTICULOS WHERE CLAVE_ARTICULO = '" + Codigo.Text + "';";
                    FbCommand command = new FbCommand(query, con);
                    bool Find = false;
                    FbDataReader reader = command.ExecuteReader();
                    string descripcion = "";
                    int cantidad = 0;
                    int clave_principal = 0;
                    string articulo_id = "";
                    bool paquetes = false ,No_incluido = false, repetido = false, excedente = false , completo = false, inicio = false, incompleto = false;
                    if (reader.Read())
                    {
                        articulo_id = reader.GetString(2);
                        cantidad = reader.GetInt32(4);
                        clave_principal = Clave_Principal(articulo_id);
                        descripcion = Descripcion(articulo_id);
                        Find = true;
                    }
                    else
                    {
                        Find = false;
                        MessageBox.Show("Código no encontrado","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Codigo.Select(0, Codigo.TextLength);
                        Codigo.Focus();
                        return;
                    }
                    for(int i =0; i<GlobalSettings.Instance.Renglones.Count; i++)
                    {
                        if (GlobalSettings.Instance.Renglones[i][0] == articulo_id)
                        {
                            No_incluido = false;
                            break;
                        }
                        else
                            No_incluido = true;
                    }
                    reader.Close();
                    decimal piezas_pedido = 0;
                    for (int i = 0; i < Grid.Rows.Count; i++)
                    {
                        if (Grid.Rows[i].Cells[1].Value.ToString() == clave_principal.ToString())
                        {
                            repetido = true;
                            GlobalSettings.Instance.Index = i;
                            break;
                        }
                    }
                    if (!repetido)
                    {
                        GlobalSettings.Instance.posicion++;
                        Grid.Rows.Add(GlobalSettings.Instance.posicion, clave_principal, descripcion ,cantidad);
                        inicio = true;
                    }
                    for(int i = 0; i<Grid.Rows.Count; i++)
                    {
                        if(Grid.Rows[i].Cells[1].Value.ToString() == clave_principal.ToString())
                        {
                            decimal ant_cant;
                            if (inicio == false)
                                ant_cant = decimal.Parse(Grid.Rows[i].Cells[3].Value.ToString());
                            else
                                ant_cant = 0;
                            if(!No_incluido)
                            {
                                for (int j = 0; j < GlobalSettings.Instance.Renglones.Count; j++)
                                {
                                    if (articulo_id == GlobalSettings.Instance.Renglones[j][0].ToString())
                                    {
                                        piezas_pedido = decimal.Parse(GlobalSettings.Instance.Renglones[j][1]);
                                    }
                                }
                            decimal suma = cantidad + ant_cant;
                            if(suma > piezas_pedido) {
                                    excedente = true;
                            }
                            else if(suma == piezas_pedido) {
                                    completo = true;
                            }
                            else
                                incompleto = true;
                            }
                            Grid.Rows[i].Cells[3].Value = cantidad + ant_cant;
                        }
                    }
                    paquetes = GlobalSettings.Instance.Picking.Any(subLista => subLista.Count > 0 && subLista[0] == clave_principal.ToString());
                    int index = 0;
                    for(int i = 0; i< Grid.Rows.Count; i++)
                    {
                        if (repetido)
                        {
                            index = GlobalSettings.Instance.Index;
                            Grid.FirstDisplayedScrollingRowIndex = index;
                            Grid.ClearSelection();
                            Grid.Rows[index].Cells[0].Selected = true;
                            Grid.Rows[index].Cells[1].Selected = true;
                            break;
                        }
                        else if (Grid.Rows[i].Cells[0].Value.ToString() == GlobalSettings.Instance.posicion.ToString())
                        {
                            index = i;
                            Grid.FirstDisplayedScrollingRowIndex = index;
                            Grid.ClearSelection();
                            Grid.Rows[index].Cells[0].Selected = true;
                            Grid.Rows[index].Cells[1].Selected = true;
                            Grid.Rows[index].Cells[3].Style.Font = new Font("Arial", 18, FontStyle.Bold);
                            break;
                        }
                    }
                    if (incompleto)
                    {
                        if (Grid.Rows[index].DefaultCellStyle.BackColor != System.Drawing.Color.LightBlue)
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            GlobalSettings.Instance.ContadorIncompletos += 1;
                        }
                    }
                    if (No_incluido)
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Tomato;
                    if (completo)
                    {
                        if (Grid.Rows[index].DefaultCellStyle.BackColor == System.Drawing.Color.LightBlue)
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                            GlobalSettings.Instance.ContadorIncompletos -= 1;
                        }
                        else
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;

                    }
                    if (excedente)
                    {
                        if (Grid.Rows[index].DefaultCellStyle.BackColor == System.Drawing.Color.LightBlue)
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                            GlobalSettings.Instance.ContadorIncompletos -= 1;
                            GlobalSettings.Instance.ContadorExcedentes += 1;
                        }
                        else if (Grid.Rows[index].DefaultCellStyle.BackColor == System.Drawing.Color.LightGreen)
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                            GlobalSettings.Instance.ContadorExcedentes += 1;
                        }
                        else if(Grid.Rows[index].DefaultCellStyle.BackColor == System.Drawing.Color.Orange)
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                        }
                        else
                        {
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                            GlobalSettings.Instance.ContadorExcedentes += 1;
                        }
                    }
                    if (paquetes)
                    {

                        Grid.Rows[index].DefaultCellStyle.ForeColor = Color.Crimson;
                        Grid.Rows[index].Cells[2].Style.Font = new System.Drawing.Font("Arial",13, FontStyle.Underline);

                    }
                    ContadorIncompletos.Text = GlobalSettings.Instance.ContadorIncompletos.ToString();
                    ContadorExcedentes.Text = GlobalSettings.Instance.ContadorExcedentes.ToString();
                    Grid.Rows[Grid.Rows.Count - 1].Height = 50;
                    Codigo.Select(0, Codigo.TextLength);
                    Codigo.Focus();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(ex.ToString());
                    return;
                }
                finally
                {
                    con.Close();
                }
            }
            else
            {
                MessageBox.Show("Completa todos los campos", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                if(Pedido.Text == string.Empty)
                    Pedido.Focus();
                else if (Cb_Surtidor.Text == string.Empty)
                    Cb_Surtidor.Focus();
                else
                    Codigo.Focus();
            }
        }

        public decimal Existencia(string articulo_id)
        {
            FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
            //command.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108404; //peri
            try
            {
                con.Open();
                FbCommand command = new FbCommand("EXIVAL_ART", con);
                command.CommandType = CommandType.StoredProcedure;
                command.Parameters.Add("V_ARTICULO_ID", FbDbType.Integer).Value = articulo_id;
                command.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108401; //ALMACEN
                //command.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108405; culiacan
                command.Parameters.Add("V_FECHA", FbDbType.Date).Value = DateTime.Today;
                command.Parameters.Add("V_ES_ULTIMO_COSTO", FbDbType.Char).Value = 'S';
                command.Parameters.Add("V_SUCURSAL_ID", FbDbType.Integer).Value = 0;

                // Parámetro de salida
                FbParameter paramARTICULO = new FbParameter("ARTICULO_ID", FbDbType.Numeric);
                paramARTICULO.Direction = ParameterDirection.Output;
                command.Parameters.Add(paramARTICULO);
                FbParameter paramEXISTENCIA = new FbParameter("EXISTENCIAS", FbDbType.Numeric);
                paramEXISTENCIA.Direction = ParameterDirection.Output;
                command.Parameters.Add(paramEXISTENCIA);
                // Ejecutar el procedimiento almacenado
                command.ExecuteNonQuery();
                GlobalSettings.Instance.ExistenciaAl = Convert.ToInt32(command.Parameters[6].Value);

                FbCommand command2 = new FbCommand("EXIVAL_ART", con);
                command2.CommandType = CommandType.StoredProcedure;
                // Parámetros de entrada
                command2.Parameters.Add("V_ARTICULO_ID", FbDbType.Integer).Value = articulo_id;
                command2.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108403; //PERI
                //command2.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108405; culiacan
                command2.Parameters.Add("V_FECHA", FbDbType.Date).Value = DateTime.Today;
                command2.Parameters.Add("V_ES_ULTIMO_COSTO", FbDbType.Char).Value = 'S';
                command2.Parameters.Add("V_SUCURSAL_ID", FbDbType.Integer).Value = 0;

                // Parámetro de salida
                FbParameter paramARTICULO2 = new FbParameter("ARTICULO_ID", FbDbType.Numeric);
                paramARTICULO2.Direction = ParameterDirection.Output;
                command2.Parameters.Add(paramARTICULO2);
                FbParameter paramEXISTENCIA2 = new FbParameter("EXISTENCIAS", FbDbType.Numeric);
                paramEXISTENCIA2.Direction = ParameterDirection.Output;
                command2.Parameters.Add(paramEXISTENCIA2);
                // Ejecutar el procedimiento almacenadoienda
                command2.ExecuteNonQuery();
                decimal ExistenciaTienda = Convert.ToInt32(command2.Parameters[6].Value);
                return ExistenciaTienda;
                //MessageBox.Show("ALMACÉN: "+ Existencia.ToString() +"\n TIENDA: "+ ExistenciaTienda.ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                MessageBox.Show(ex.ToString());
                return 0;
            }
            finally
            {
                con.Close();
            }
        }
        public void Traspaso()
        {
            GlobalSettings.Instance.ListaArt.Clear();
            DateTime fecha = DateTime.Now;
            string des = Pedido.Text + " Realizado por: " + Cb_Surtidor.Text;
            //int ErrorFolio = ApiVe.NuevoPedido(Fecha, "IP", int.Parse(Cliente_Id), Dir_Consig_Id, Almacen_Id,"", Tipo_Desc,Descuento,"",Descripcion,Vendedor_Id, 0, 0, Moneda_Id);
            if (GlobalSettings.Instance.aceptadoP)
            {
                des = Pedido.Text + " Realizado por: " + Cb_Surtidor.Text+ "\n Autorizado por: " + GlobalSettings.Instance.Usuario;
                GlobalSettings.Instance.aceptadoP = false;
            }
            int ErrorFolio = ApiInv.NuevaSalida(36, 108403, 108401, fecha.ToString(), "", des, 0);
            for (int i = 0; i < Grid.Rows.Count; ++i)
            {
                List<string> Articulos = new List<string>();
                int articulo_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                int Renglon = ApiInv.RenglonSalida(articulo_id, double.Parse(Grid.Rows[i].Cells[3].Value.ToString()), 0, 0);
                Articulos.Add(Grid.Rows[i].Cells[1].Value.ToString());
                Articulos.Add(Grid.Rows[i].Cells[2].Value.ToString());
                Articulos.Add(Grid.Rows[i].Cells[3].Value.ToString());
                GlobalSettings.Instance.ListaArt.Add(Articulos);
            }
            int final = ApiInv.AplicaSalida();
            if (final == 3)
            {         
                mensaje.Titulo.Text = "Hay algunos articulos con existencia insuficiente";
                mensaje.GridExistencia.Rows.Clear();
                mensaje.Height = 150;
                for (int i = 0; i < Grid.Rows.Count; ++i)
                {
                    int articulo_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                    decimal existencia = Existencia(articulo_id.ToString());
                    if (existencia < decimal.Parse(Grid.Rows[i].Cells[3].Value.ToString()))
                    {
                        mensaje.Height += 25;
                        mensaje.GridExistencia.Rows.Add(Grid.Rows[i].Cells[1].Value.ToString(), existencia, Grid.Rows[i].Cells[3].Value.ToString(), decimal.Parse(Grid.Rows[i].Cells[3].Value.ToString()) - existencia );
                        mensaje.GridExistencia.Height += 25;

                    }
                }
                ApiBas.DBDisconnect(GlobalSettings.Instance.Bd);
                mensaje.EnviarVariableEvent4 += new MessageBoxCustom.EnviarVariableDelegate4(ajuste);
                mensaje.ShowDialog();
            }
            else if( final == 0)
            {
                GlobalSettings.Instance.ContadorIncompletos = 0;
                GlobalSettings.Instance.ContadorExcedentes = 0;
                ContadorIncompletos.Text = "";
                ContadorExcedentes.Text = "";
                Mensajes mensajecu = new Mensajes();
                mensajecu.Texto.Text = "Salida con éxito";
                mensajecu.ShowDialog();
                ApiBas.DBDisconnect(GlobalSettings.Instance.Bd);
                printDocument1.Print();
                //if (printDialog1.ShowDialog() == DialogResult.OK)
                //{
                //    printDocument1.Print();
                //}

                Grid.Rows.Clear();
                GlobalSettings.Instance.posicion = 0;
                GlobalSettings.Instance.Id = 0;
                Codigo.Text = string.Empty;
                Pedido.Enabled = true;
                Agregar_Pedido.BackColor = System.Drawing.Color.Black;
                Agregar_Pedido.Enabled = true;
                Agregar_Pedido.Text = "Agregar";
                GlobalSettings.Instance.ListaArt.Clear();
                GlobalSettings.Instance.aceptadoP = false;
                GlobalSettings.Instance.aceptado = false;
                GlobalSettings.Instance.Renglones.Clear();
                GlobalSettings.Instance.Usuario = string.Empty;
                GlobalSettings.Instance.SinEfecto = false;
                Pedido.Text = string.Empty;
                Pedido.Focus();
                //RestoreOriginalSize();
            }
        }
        public void ajuste()
        {
            if(GlobalSettings.Instance.aceptado == true)
            {

                int conecta = ConectaBD();
                if (conecta == 1)
                {
                    ApiBas.DBConnected(GlobalSettings.Instance.Bd);
                    ApiInv.SetDBInventarios(GlobalSettings.Instance.Bd);

                }
                DateTime fecha = DateTime.Now;
                //int ErrorFolio = ApiVe.NuevoPedido(Fecha, "IP", int.Parse(Cliente_Id), Dir_Consig_Id, Almacen_Id,"", Tipo_Desc,Descuento,"",Descripcion,Vendedor_Id, 0, 0, Moneda_Id);
                int ErrorFolio2 =  ApiInv.NuevaEntrada(1490318, 108403, fecha.ToString(), "", Pedido.Text + " Realizado por: " + Cb_Surtidor.Text + "\nAutorizado por: "+ GlobalSettings.Instance.Usuario,0);
                for (int i = 0; i < mensaje.GridExistencia.Rows.Count; ++i)
                {
                    int articulo_id = Art_Id(mensaje.GridExistencia.Rows[i].Cells[0].Value.ToString());
                    int Renglon = ApiInv.RenglonEntrada(articulo_id, double.Parse(mensaje.GridExistencia.Rows[i].Cells[3].Value.ToString()), 0,0);
                }
                int final = ApiInv.AplicaEntrada();
                Mensajes mensajesc = new Mensajes();
                mensajesc.Texto.Text = "Entrada con éxito"+ "\n Autorizada por: "+GlobalSettings.Instance.Usuario;
                mensajesc.ShowDialog();
                GlobalSettings.Instance.aceptado = false;
                GlobalSettings.Instance.ListaArt.Clear();
                Traspaso();
            }
        }
        public void ValidaExcel()
        {
            List<List<string>> TablaPicking = new List<List<string>>();
            bool encontrado = false;
            for (int i = 0; i < Grid.Rows.Count; i++)
            {
                int art_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                decimal existencia = Existencia(art_id.ToString()); 
                foreach (var sublista in GlobalSettings.Instance.Picking)
                {
                    // Verificar si la sublista contiene el valor que buscas
                    if (sublista.Contains(Grid.Rows[i].Cells[1].Value.ToString()) && GlobalSettings.Instance.ExistenciaAl > decimal.Parse(sublista[2]))
                    {
                        List<string> temp = new List<string>();
                        temp.Add(Grid.Rows[i].Cells[1].Value.ToString());
                        temp.Add(existencia.ToString());
                        temp.Add(GlobalSettings.Instance.ExistenciaAl.ToString());
                        temp.Add(Grid.Rows[i].Cells[3].Value.ToString());
                        temp.Add(sublista[1]);
                        TablaPicking.Add(temp);
                        encontrado = true;
                    }
                }
            }
            if (encontrado)
            {  
                mensajePicking.GridPicking.Rows.Clear();
                mensajePicking.Height = 150;
                mensajePicking.Titulo.Text = "Articulos en Picking Almacén";
                mensajePicking.GridPicking.Height = 50;
                for (int i = 0; i < TablaPicking.Count; i++)
                {
                    mensajePicking.Height += 25;
                    mensajePicking.GridPicking.Rows.Add(TablaPicking[i][0], TablaPicking[i][1], TablaPicking[i][2], TablaPicking[i][3], TablaPicking[i][4]);
                    mensajePicking.GridPicking.Height += 25;
                }
                mensajePicking.EnviarVariableEvent6 += new MessageBoxCustomPicking.EnviarVariableDelegate6(aceptado);
                mensajePicking.GridPicking.ClearSelection();
                mensajePicking.ShowDialog();
                TablaPicking.Clear();
            }
            if(encontrado == false)
            {
                GlobalSettings.Instance.SinEfecto = true;
            }
        }
        public void aceptado()
        {
            mensajePicking.Close();
        }
        private void Generar_Click(object sender, EventArgs e)
        {
            if (Grid.Rows.Count == 0)
            {
                MessageBox.Show("Primero ingresa renglones", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (Cb_Surtidor.Text == string.Empty)
            {
                MessageBox.Show("Te falta asignar un surtidor", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            DialogResult result = MessageBox.Show("¿Estás seguro que deseas realizar el traspaso?\n Revisa bien las cantidades \n ¿Deseas continuar?", "Advertencia", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            if (result == DialogResult.Cancel)
            {
                return;
            }
            ValidaExcel();
            if(GlobalSettings.Instance.aceptadoP == false && GlobalSettings.Instance.SinEfecto == false)
            {
                return;
            }
            int conecta = ConectaBD();
            if (conecta == 1)
            {
                ApiBas.DBConnected(GlobalSettings.Instance.Bd);
                ApiInv.SetDBInventarios(GlobalSettings.Instance.Bd);
                //int contadorFolio = Contador();
                //MessageBox.Show(contadorFolio.ToString());
                //int Contador = ContarFolios();
                Traspaso();
                
            }
        }
        private void RestoreOriginalSize()
        {
            //this.Width = GlobalSettings.Instance.originalWidth;
            //this.Height = GlobalSettings.Instance.originalHeight;
            //this.Size = GlobalSettings.Instance.originalSize;
        }
        private void Principal_MouseMove(object sender, MouseEventArgs e)
        {
            if (isDragging)
            {
                // Calcula el desplazamiento
                int deltaX = (int)((startPoint.X - e.Location.X) * 0.05f);
                int deltaY = (int)((startPoint.Y - e.Location.Y) * 0.05f);

                // Ajusta el scroll del formulario
                this.AutoScrollPosition = new System.Drawing.Point(
                    deltaX - this.AutoScrollPosition.X,
                    deltaY - this.AutoScrollPosition.Y
                );
            }
        }

        private void Principal_MouseUp(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = false; // Detiene el "arrastre"
            }
        }

        private void Principal_MouseDown(object sender, MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Left)
            {
                isDragging = true;
                startPoint = e.Location; // Guarda la posición inicial del clic
            }
        }

        private void Codigo_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Ingresar.Focus();
            }
        }

        private void Cb_Surtidor_Leave(object sender, EventArgs e)
        {
            if (!nombresArray.Contains(Cb_Surtidor.Text) && Cb_Surtidor.Text != "")
            {
                MessageBox.Show("Este usuario no está registrado", "¡Espera!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Cb_Surtidor.Text = "";
                Cb_Surtidor.Focus();
            }
        }

        private void Cb_Surtidor_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (Char.IsLower(e.KeyChar))
            {
                e.KeyChar = Char.ToUpper(e.KeyChar);
            }
        }

        private void Cb_Surtidor_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                e.SuppressKeyPress = true;
                Codigo.Focus();
            }
        }

        private void Grid_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if(Grid.CurrentCell != null)
            {
                if (Grid.CurrentCell.RowIndex >= 0 && Grid.CurrentCell.ColumnIndex == 4)
                {
                    DialogResult result = MessageBox.Show("¿Estás seguro que deseas eliminar este artículo?\n ¿Deseas continuar?", "Advertencia", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
                    if (result == DialogResult.Cancel)
                    {
                        Codigo.Focus();
                        Codigo.Select(0, Codigo.Text.Length);
                        return;
                    }
                    else
                    {
                        if (Grid.Rows[Grid.CurrentCell.RowIndex].DefaultCellStyle.BackColor == System.Drawing.Color.LightBlue)
                        {
                            GlobalSettings.Instance.ContadorIncompletos -= 1;
                            ContadorIncompletos.Text = GlobalSettings.Instance.ContadorIncompletos.ToString();  
                        }
                        if (Grid.Rows[Grid.CurrentCell.RowIndex].DefaultCellStyle.BackColor == System.Drawing.Color.Orange)
                        {
                            GlobalSettings.Instance.ContadorExcedentes -= 1;
                            ContadorExcedentes.Text = GlobalSettings.Instance.ContadorExcedentes.ToString();
                        }
                        Grid.Rows.RemoveAt(Grid.CurrentCell.RowIndex);
                        //Grid.CurrentCell = null;
                        Grid.ClearSelection();
                        //Grid.Height -= 50;
                        Codigo.Select(0, Codigo.TextLength);
                        Codigo.Focus();
                    }
                }
                else if (Grid.CurrentCell.RowIndex >= 0 && Grid.CurrentCell.ColumnIndex == 3)
                {
                    Numpad numpad = new Numpad();
                    GlobalSettings.Instance.identificador = int.Parse(Grid.Rows[Grid.CurrentCell.RowIndex].Cells[0].Value.ToString());
                    numpad.EnviarVariableEvent2 += new Numpad.EnviarVariableDelegate2(ejecutar);
                    numpad.ShowDialog();
                    numpad.Cantidad.Focus();
                    numpad.Cantidad.Select(0, numpad.Cantidad.TextLength);
                }
                else if (Grid.CurrentCell.RowIndex >= 0 && Grid.CurrentCell.ColumnIndex == 1)
                {
                    int codigo = int.Parse(Grid.CurrentRow.Cells[1].Value.ToString());
                    int Articulo = Art_Id(codigo.ToString());
                    decimal Ex = Existencia(Articulo.ToString());
                    Mensajes mensajes = new Mensajes();
                    mensajes.Texto.Text = "Existencia Almacén: " + GlobalSettings.Instance.ExistenciaAl.ToString() + "\n Existencia Tienda: " + Ex.ToString();
                    mensajes.ShowDialog();
                }
                Codigo.Focus();
                Codigo.Select(0, Codigo.Text.Length);
                Grid.ClearSelection();
                //Grid.CurrentCell = null;
            }
            
        }
        public void ejecutar(decimal unidad)
        {
            bool encontrado = false;
            decimal piezas_pedido = 0;
            for(int i= 0; i< Grid.Rows.Count; i++)
            {
                if (Grid.Rows[i].Cells[0].Value.ToString() == GlobalSettings.Instance.identificador.ToString())
                {
                    bool azul = Grid.Rows[i].DefaultCellStyle.BackColor == System.Drawing.Color.LightBlue;
                    bool verde = Grid.Rows[i].DefaultCellStyle.BackColor == System.Drawing.Color.LightGreen;
                    bool Naranja = Grid.Rows[i].DefaultCellStyle.BackColor == System.Drawing.Color.Orange;
                    Grid.Rows[i].Cells[3].Value = unidad;
                    int articulo_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                    for (int j = 0; j < GlobalSettings.Instance.Renglones.Count; j++)
                    {
                          if (articulo_id.ToString() == GlobalSettings.Instance.Renglones[j][0].ToString())
                          {
                                  piezas_pedido += decimal.Parse(GlobalSettings.Instance.Renglones[j][1]);
                                  encontrado = true;
                          }
                        
                    }
                    if(encontrado == false)
                    {
                        break;
                    }
                    if (unidad == piezas_pedido)
                    {
                        if (azul)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                            GlobalSettings.Instance.ContadorIncompletos -= 1;
                        }
                        if(Naranja)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                            GlobalSettings.Instance.ContadorExcedentes -= 1;
                        }
                        else
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                    }
                    else if (unidad > piezas_pedido)
                    {
                        if(azul)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                            GlobalSettings.Instance.ContadorIncompletos -= 1;
                            GlobalSettings.Instance.ContadorExcedentes += 1;
                        }
                        if (verde)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                            GlobalSettings.Instance.ContadorExcedentes += 1;
                        }
                        else
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                    }
                    else if (unidad < piezas_pedido)
                    {
                        if (verde)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            GlobalSettings.Instance.ContadorIncompletos += 1;
                        }
                        if (Naranja)
                        {
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                            GlobalSettings.Instance.ContadorIncompletos += 1;
                            GlobalSettings.Instance.ContadorExcedentes -= 1;
                        }
                        else
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightBlue;
                    }
                    bool paquete = GlobalSettings.Instance.Picking.Any(subLista => subLista.Count > 0 && subLista[0] == Grid.Rows[i].Cells[1].Value.ToString());
                    if (paquete)
                    {
                        Grid.Rows[i].DefaultCellStyle.ForeColor = Color.Crimson;
                    }

                }
                    ContadorIncompletos.Text = GlobalSettings.Instance.ContadorIncompletos.ToString();
                ContadorExcedentes.Text = GlobalSettings.Instance.ContadorExcedentes.ToString();
            }
            
            Codigo.Focus();
            Codigo.Select(0, Codigo.Text.Length);
            //Grid.CurrentCell = null;
            Grid.ClearSelection();
        }

        private void Agregar_Pedido_Click(object sender, EventArgs e)
        {
            if (Pedido.Text != string.Empty)
            {
                string Folio_Mod = Pedido.Text;
                if (Folio_Mod[1] == 'O' || Folio_Mod[1] == 'E' || Folio_Mod[1] == 'P' || Folio_Mod[1] == 'M' || Folio_Mod[1] == 'A')
                {
                    int cont = 9 - Folio_Mod.Length;
                    string prefix = Folio_Mod.Substring(0, 2);
                    string suffix = Folio_Mod.Substring(2);
                    string patch = "";
                    for (int i = 0; i < cont; i++)
                    {
                        patch = patch + "0";
                    }
                    Folio_Mod = prefix + patch + suffix;
                }
                else if (Folio_Mod[0] == 'P')
                {
                    int cont = 9 - Folio_Mod.Length;
                    string prefix = Folio_Mod.Substring(0, 1);
                    string suffix = Folio_Mod.Substring(1);
                    string patch = "";
                    for (int i = 0; i < cont; i++)
                    {
                        patch = patch + "0";
                    }
                    Folio_Mod = prefix + patch + suffix;
                }
                FbConnection con = new FbConnection("User=" + GlobalSettings.Instance.User + ";" + "Password=" + GlobalSettings.Instance.Pw + ";" + "Database=" + GlobalSettings.Instance.Direccion + ";" + "DataSource=" + GlobalSettings.Instance.Ip + ";" + "Port=" + GlobalSettings.Instance.Puerto + ";" + "Dialect=3;" + "Charset=UTF8;");
                try
                {
                    con.Open();
                    string Docto_Id = "";
                    string query = "SELECT * FROM DOCTOS_VE WHERE FOLIO = '" + Folio_Mod + "'";
                    FbCommand command = new FbCommand(query, con);
                    FbDataReader reader = command.ExecuteReader();
                    if (reader.Read())
                    {
                        Docto_Id = reader.GetString(0);
                    }
                    else
                    {
                        MessageBox.Show("Pedido no encontrado", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        Pedido.Select(0, Pedido.TextLength);
                        Pedido.Focus();
                        return;
                    }
                    reader.Close();
                    string query2 = "SELECT * FROM DOCTOS_VE_DET WHERE DOCTO_VE_ID = '" + Docto_Id + "'";
                    FbCommand command2 = new FbCommand(query2, con);
                    FbDataReader reader2 = command2.ExecuteReader();
                    while (reader2.Read())
                    {
                        string Articulo_Id = reader2.GetString(3);
                        decimal Unidades = reader2.GetDecimal(4);
                        bool found = false; // Esta variable indica si se encontró el artículo en los renglones.

                        foreach (var list in GlobalSettings.Instance.Renglones)
                        {
                            if (list.Contains(Articulo_Id))
                            {
                                // Convierte list[1] a decimal y suma el valor de Unidades
                                decimal currentValue = decimal.Parse(list[1]);
                                currentValue += Unidades;

                                // Vuelve a asignar el valor actualizado a list[1]
                                list[1] = currentValue.ToString();

                                found = true;
                                break; // Si se encuentra el artículo, ya no es necesario seguir buscando.
                            }
                        }

                        // Si no se encontró el artículo en los renglones, se agrega uno nuevo
                        if (!found)
                        {
                            List<string> renglon = new List<string>();
                            renglon.Add(Articulo_Id);
                            renglon.Add(Unidades.ToString());
                            GlobalSettings.Instance.Renglones.Add(renglon);
                        }

                    }
                    reader2.Close();
                    Agregar_Pedido.BackColor = System.Drawing.Color.SteelBlue;
                    Pedido.Enabled = false;
                    Agregar_Pedido.Text = "Listo";
                    Agregar_Pedido.Enabled = false;
                    Cb_Surtidor.Focus();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Se perdió la conexión :( , contacta a 06 o intenta de nuevo", "¡Error!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    MessageBox.Show(ex.ToString());
                }
                finally
                {
                    con.Close();
                }
            }
        }

        private void Pedido_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Agregar_Pedido.Focus();

            }
        }

        private void Agregar_Pedido_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                Cb_Surtidor.Focus();

            }
        }

        private void SinCodigo_Click(object sender, EventArgs e)
        {
            if(Pedido.Text != string.Empty && Cb_Surtidor.Text != string.Empty)
            {
                SinCodigo Ventana = new SinCodigo();
                Ventana.EnviarArticuloEvent += new SinCodigo.EnviarArticulo(Agregar);
                Ventana.ShowDialog();
                Codigo.Select(0, Codigo.TextLength);
                Codigo.Focus();
            }
            else
            {
                MessageBox.Show("Completa los campos primero", "¡Espera!", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Pedido.Focus();
            }
        }
        public void Agregar(int articulo)
        {
            Codigo.Text = articulo.ToString();
            Ingresar_Click(null, EventArgs.Empty);
        }

        private void Principal_Load(object sender, EventArgs e)
        {
            //GlobalSettings.Instance.originalWidth = this.Width;
            //GlobalSettings.Instance.originalHeight = this.Height;
            //GlobalSettings.Instance.originalSize = this.Size;
            //this.WindowState = FormWindowState.Maximized;
        }

        private void Grid_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.F9)
            {
                if (Grid.CurrentRow != null)
                {
                    int codigo = int.Parse(Grid.CurrentRow.Cells[1].Value.ToString());
                    int Articulo = Art_Id(codigo.ToString());   
                    decimal Ex = Existencia(Articulo.ToString());
                    Mensajes mensajes = new Mensajes();
                    mensajes.Texto.Text = "Existencia Almacén: " + GlobalSettings.Instance.ExistenciaAl.ToString() + "\n Existencia Tienda: " + Ex.ToString();
                    mensajes.ShowDialog();
                }
            }
        }
        private void SaveDocumentAsImage()
        {
            // Definir el tamaño de la imagen
            int width = 600;
            int height = 800;

            // Crear una imagen en blanco con el tamaño definido
            Bitmap bitmap = new Bitmap(width, height);

            // Crear un objeto Graphics para dibujar sobre la imagen
            using (Graphics graphics = Graphics.FromImage(bitmap))
            {
                graphics.Clear(Color.White);  // Limpiar la imagen con color blanco

                // Aquí se renderiza el contenido que querías imprimir
                using (Font font = new Font("Arial", 12, FontStyle.Bold))
                {
                    graphics.DrawString("TRASPASO A ALMACÉN", font, Brushes.Black, new PointF(40, 0));
                }

                using (Font font = new Font("Arial", 20, FontStyle.Bold))
                {
                    graphics.DrawString(Pedido.Text, font, Brushes.Black, new PointF(90, 50));
                }

                using (Font font = new Font("Arial", 10, FontStyle.Bold))
                {
                    graphics.DrawString(DateTime.Now.ToString(" yyyy-MM-dd HH:mm"), font, Brushes.Black, new PointF(160, 20));
                }

                System.Drawing.Image originalImage = System.Drawing.Image.FromFile("C:\\Img\\logo.jpg");
                System.Drawing.Image resizedImage = new Bitmap(originalImage, new System.Drawing.Size(50, 30));
                graphics.DrawImage(resizedImage, new PointF(0, 40));

                int j = 90;
                for (int i = 0; i < GlobalSettings.Instance.ListaArt.Count; ++i)
                {
                    using (Font font = new Font("Arial", 10, FontStyle.Bold))
                    {
                        if (i > 0)
                            graphics.DrawString("_____________________________________________________________", font, Brushes.Black, new PointF(0, j - 25));
                        graphics.DrawString(GlobalSettings.Instance.ListaArt[i][0], font, Brushes.Black, new PointF(0, j));
                        if (GlobalSettings.Instance.ListaArt[i][2].ToString().Length > 3)
                            graphics.DrawString(GlobalSettings.Instance.ListaArt[i][2].ToString(), font, Brushes.Black, new PointF(252, j));
                        else
                            graphics.DrawString(GlobalSettings.Instance.ListaArt[i][2].ToString(), font, Brushes.Black, new PointF(260, j));
                    }

                    System.Drawing.Rectangle DestinationRectangle = new System.Drawing.Rectangle(51, j, 220, 50);
                    using (StringFormat sf = new StringFormat())
                    {
                        graphics.DrawString(GlobalSettings.Instance.ListaArt[i][1], new Font("Arial", 8, FontStyle.Regular), Brushes.Black, DestinationRectangle, sf);
                    }

                    j += 50;
                }

                using (Font font = new Font("Arial", 10, FontStyle.Bold))
                {
                    graphics.DrawString("                      ___________  ", font, Brushes.Black, new PointF(0, j + 50));
                    graphics.DrawString("                        Recibió    ", font, Brushes.Black, new PointF(0, j + 70));
                    graphics.DrawString("Hecho por: " + Cb_Surtidor.Text, font, Brushes.Black, new PointF(0, j + 90));
                }
            }

            // Guardar la imagen como un archivo (por ejemplo, en formato PNG)
            bitmap.Save("C:\\Traspasos\\"+Pedido.Text+".png", System.Drawing.Imaging.ImageFormat.Png);
        }

        private void printDocument1_PrintPage(object sender, System.Drawing.Printing.PrintPageEventArgs e)
        {
            SaveDocumentAsImage();
            using (Font font = new Font("Arial", 12, FontStyle.Bold))
            {
                e.Graphics.DrawString("TRASPASO A ALMACÉN", font, Brushes.Black, new PointF(40,0));

            }
            using (Font font = new Font("Arial", 20, FontStyle.Bold))
            {
                e.Graphics.DrawString(Pedido.Text, font, Brushes.Black, new PointF(90, 50));

            }
            using (Font font = new Font("Arial", 10, FontStyle.Bold))
            {
                e.Graphics.DrawString(DateTime.Now.ToString(" yyyy-MM-dd HH:mm"), font, Brushes.Black, new PointF(160, 20));

            }
            System.Drawing.Image originalImage = System.Drawing.Image.FromFile("C:\\Img\\logo.jpg");
            System.Drawing.Image resizedImage = new Bitmap(originalImage, new System.Drawing.Size(50, 30));
            e.Graphics.DrawImage(resizedImage, new PointF(0, 40));
            int j = 90;
            for (int i = 0; i < GlobalSettings.Instance.ListaArt.Count; ++i)
            {
                using (Font font = new Font("Arial", 10, FontStyle.Bold))
                {
                    if (i > 0)
                        e.Graphics.DrawString("_____________________________________________________________", font, Brushes.Black, new PointF(0, j - 25));
                    e.Graphics.DrawString(GlobalSettings.Instance.ListaArt[i][0], font, Brushes.Black, new PointF(0, j));
                    if (GlobalSettings.Instance.ListaArt[i][2].ToString().Length > 3)
                        e.Graphics.DrawString(GlobalSettings.Instance.ListaArt[i][2].ToString(), font, Brushes.Black, new PointF(252, j));
                    else
                        e.Graphics.DrawString(GlobalSettings.Instance.ListaArt[i][2].ToString(), font, Brushes.Black, new PointF(260, j));

                }

                System.Drawing.Rectangle DestinationRectangle = new System.Drawing.Rectangle(51, j, 220, 50);
                using (StringFormat sf = new StringFormat())
                {
                    e.Graphics.DrawString(GlobalSettings.Instance.ListaArt[i][1], new Font("Arial", 8, FontStyle.Regular), Brushes.Black, DestinationRectangle, sf);
                }

                j += 50;
            }
            using (Font font = new Font("Arial", 10, FontStyle.Bold))
            {
                e.Graphics.DrawString("                      ___________  ", font, Brushes.Black, new PointF(0, j + 50));
                e.Graphics.DrawString("                        Recibió    ", font, Brushes.Black, new PointF(0, j + 70));
                e.Graphics.DrawString("Hecho por: "+Cb_Surtidor.Text, font, Brushes.Black, new PointF(0, j + 90));

            }
        }



        
        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if(Pedido.Text != string.Empty)
            {
                string file = "C:\\Traspasos\\" + Pedido.Text + ".png";
                if (File.Exists(file))
                {
                    GlobalSettings.Instance.ticket = System.Drawing.Image.FromFile(file);
                    printDocument2.Print();
                }
                else
                {
                    MessageBox.Show("El pedido no existe.","Error",MessageBoxButtons.OK,MessageBoxIcon.Error);
                    return;
                }
            }
            else
            {
                MessageBox.Show("Primero ingresa un pedido","Error",MessageBoxButtons.OK,MessageBoxIcon.Stop);
                return;
            }
        }

        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)
        {
            e.Graphics.DrawImage(GlobalSettings.Instance.ticket, 0, 0);
        }
    }
}

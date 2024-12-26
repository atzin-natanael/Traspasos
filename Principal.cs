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
using SpreadsheetLight;
using DocumentFormat.OpenXml.Bibliography;
using System.CodeDom.Compiler;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Diagrams;
using System.Reflection;

namespace Pantalla_De_Control
{
    public partial class Principal : Form
    {

        private List<string> nombresValor = new List<string>();

        private List<string> nombresArray = new List<string>();

        MessageBoxCustom mensaje = new MessageBoxCustom();


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
        
        }
        public void Leer_Datos()
        {
            nombresArray.Clear();
            nombresValor.Clear();
            //string filePath = "\\\\192.168.0.2\\C$\\clavesSurtido\\Claves.xlsx";

            string filePath = "\\\\SRVPRINCIPAL\\clavesSurtido\\Claves.xlsx";
            // string filePath = "C:\\clavesSurtido\\Claves.xlsx";
            using (SLDocument documento = new SLDocument(filePath))
            {
                int filas = documento.GetWorksheetStatistics().NumberOfRows;
                for (int i = 2; i < filas + 1; ++i)
                {
                    string temp_name = documento.GetCellValueAsString("A" + i);
                    string temp_value = documento.GetCellValueAsString("B" + i);
                    string temp_status = documento.GetCellValueAsString("C" + i);
                    string name = temp_name + " " + temp_value;
                    nombresArray.Add(name);
                    nombresValor.Add(temp_status);
                }
                documento.CloseWithoutSaving();
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
            int conecta = ApiBas.DBConnect(GlobalSettings.Instance.Bd, path, "SYSDBA", "C0r1b423");
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
                    bool paquetes = false ,No_incluido = false, repetido = false, excedente = false , completo = false, inicio = false;
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
                            }
                            Grid.Rows[i].Cells[3].Value = cantidad + ant_cant;
                        }
                    }
                    paquetes = descripcion.Contains("C/");
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
                            break;
                        }
                    }
                    if (paquetes)
                        Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orchid;
                    if(No_incluido)
                            Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Tomato;
                    if (completo)
                        Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                    if (excedente)
                        Grid.Rows[index].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
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

                // Parámetros de entrada
                command.Parameters.Add("V_ARTICULO_ID", FbDbType.Integer).Value = articulo_id;
                command.Parameters.Add("V_ALMACEN_ID", FbDbType.Integer).Value = 108403;
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
                int Existencia = Convert.ToInt32(command.Parameters[6].Value);

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
            DateTime fecha = DateTime.Now;
            //int ErrorFolio = ApiVe.NuevoPedido(Fecha, "IP", int.Parse(Cliente_Id), Dir_Consig_Id, Almacen_Id,"", Tipo_Desc,Descuento,"",Descripcion,Vendedor_Id, 0, 0, Moneda_Id);
            int ErrorFolio = ApiInv.NuevaSalida(36, 108403, 108401, fecha.ToString(), "", Pedido.Text + " Realizado por: " + Cb_Surtidor.Text, 0);
            for (int i = 0; i < Grid.Rows.Count; ++i)
            {
                int articulo_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                int Renglon = ApiInv.RenglonSalida(articulo_id, double.Parse(Grid.Rows[i].Cells[3].Value.ToString()), 0, 0);
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
                MessageBox.Show("Salida con éxito");
                ApiBas.DBDisconnect(GlobalSettings.Instance.Bd);
                Grid.Rows.Clear();
                GlobalSettings.Instance.posicion = 0;
                GlobalSettings.Instance.Id = 0;
                Codigo.Text = string.Empty;
                Pedido.Enabled = true;
                Agregar_Pedido.Enabled = true;
                Pedido.Text = string.Empty;
                Pedido.Focus();
                RestoreOriginalSize();
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
                int ErrorFolio2 =  ApiInv.NuevaEntrada(1490318, 108403, fecha.ToString(), "", Pedido.Text + " Realizado por: " + Cb_Surtidor.Text,0);
                for (int i = 0; i < mensaje.GridExistencia.Rows.Count; ++i)
                {
                    int articulo_id = Art_Id(mensaje.GridExistencia.Rows[i].Cells[0].Value.ToString());
                    int Renglon = ApiInv.RenglonEntrada(articulo_id, double.Parse(Grid.Rows[i].Cells[3].Value.ToString()), 0,0);
                }
                int final = ApiInv.AplicaEntrada();
                MessageBox.Show("Entrada éxitosa");
                GlobalSettings.Instance.aceptado = false;
                Traspaso();
            }
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
            this.Width = GlobalSettings.Instance.originalWidth;
            this.Height = GlobalSettings.Instance.originalHeight;
            this.Size = GlobalSettings.Instance.originalSize;
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
            if(Grid.CurrentCell.RowIndex >= 0 && Grid.CurrentCell.ColumnIndex == 4)
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
                    Grid.Rows.RemoveAt(Grid.CurrentCell.RowIndex);
                    Grid.CurrentCell = null;
                    Grid.Height -= 50;
                    Codigo.Select(0, Codigo.TextLength);
                    Codigo.Focus();
                }
            }
            else if( Grid.CurrentCell.RowIndex >= 0 && Grid.CurrentCell.ColumnIndex == 3)
            {
                Numpad numpad = new Numpad();
                GlobalSettings.Instance.identificador =  int.Parse(Grid.Rows[Grid.CurrentCell.RowIndex].Cells[0].Value.ToString());
                numpad.EnviarVariableEvent2 += new Numpad.EnviarVariableDelegate2(ejecutar);
                numpad.Show();
                numpad.Cantidad.Select(0, numpad.Cantidad.TextLength);
                numpad.Cantidad.Focus();
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
                    Grid.Rows[i].Cells[3].Value = unidad;
                    int articulo_id = Art_Id(Grid.Rows[i].Cells[1].Value.ToString());
                    for (int j = 0; j < GlobalSettings.Instance.Renglones.Count; j++)
                    {
                          if (articulo_id.ToString() == GlobalSettings.Instance.Renglones[j][0].ToString())
                          {
                                  piezas_pedido = decimal.Parse(GlobalSettings.Instance.Renglones[j][1]);
                                  encontrado = true;
                                  break;
                          }
                        
                    }
                    if(encontrado == false)
                    {
                        break;
                    }
                    if (unidad == piezas_pedido)
                    {
                        Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                    }
                    else if (unidad > piezas_pedido)
                    {
                        Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Orange;
                    }
                    else
                    {
                        bool paquete = Grid.Rows[i].Cells[2].Value.ToString().Contains("C/");
                        if (paquete)
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.Orchid;
                        else
                            Grid.Rows[i].DefaultCellStyle.BackColor = System.Drawing.Color.White;
                    }
                    break;

                }
            }
            Codigo.Focus();
            Codigo.Select(0, Codigo.Text.Length);
            Grid.CurrentCell = null;
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
                        List<string> renglon = new List<string>();
                        renglon.Add(Articulo_Id);
                        renglon.Add(Unidades.ToString());
                        GlobalSettings.Instance.Renglones.Add(renglon);
                    }
                    reader2.Close();
                    Pedido.Enabled = false;
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
            GlobalSettings.Instance.originalWidth = this.Width;
            GlobalSettings.Instance.originalHeight = this.Height;
            GlobalSettings.Instance.originalSize = this.Size;
            this.WindowState = FormWindowState.Maximized;
        }
    }
}

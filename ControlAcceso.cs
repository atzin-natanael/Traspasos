using Dapper;
using FirebirdSql.Data.Services;
using Microsoft.Data.Sqlite;
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
    public partial class ControlAcceso : Form
    {
        public delegate void EnviarVariableDelegate3();
        public event EnviarVariableDelegate3 EnviarVariableEvent3;
        public ControlAcceso()
        {
            InitializeComponent();
            TxtContrasenia.Focus();
            TxtContrasenia.Select();
        }

        private void Aceptar_Click(object sender, EventArgs e)
        {
            if (TxtContrasenia.Text == "123")
            {
                this.Close();
                GlobalSettings.Instance.aceptado = true;
                EnviarVariableEvent3();
            }
            else
            {
                MessageBox.Show("Contraseña incorrecta","Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                TxtContrasenia.Focus();
                TxtContrasenia.Select(0, TxtContrasenia.TextLength);
            }
        }
      
        private void Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void TxtContrasenia_KeyDown(object sender, KeyEventArgs e)
        {
            if(e.KeyCode == Keys.Enter)
            {
                Aceptar.Focus();
            }
        }
    }
}

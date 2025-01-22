using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pantalla_De_Control
{
    public class GlobalSettings
    {
        private static GlobalSettings instance;
        public string Ip { get; set; }
        public string Puerto { get; set; }
        public string Direccion { get; set; }
        public int posicion { get; set; }
        public string User { get; set; }
        public string Pw { get; set; }
        public string Ruta { get; set; }
        public List<string> Config { get; set; }
        public List<List<string>> Renglones { get; set; }
        public List<List<string>> Picking { get; set; }
        public List<List<string>> ListaArt { get; set; }
        public int Bd { get; set; }
        public int Id { get; set; }
        public decimal ExistenciaAl { get; set; }
        public decimal ContadorIncompletos { get; set; }
        public decimal ContadorExcedentes { get; set; }
        public Image ticket { get; set; }
        public int Index { get; set; }
        public int identificador { get; set; }
        public int Trn { get; set; }
        public bool aceptado { get; set; }
        public bool aceptadoP { get; set; }
        public bool SinEfecto { get; set; }
        public bool BanderaPicking { get; set; }

        public int originalWidth {  get; set; }
        public int originalHeight { get; set; }
        public Size originalSize { get; set; }
        public string Usuario { get; set; }
        private GlobalSettings()
        {
            Config = new List<string>();

            Renglones = new List<List<string>>();
            Picking = new List<List<string>>();

            ListaArt = new List<List<string>>();

        }
        public static GlobalSettings Instance
        {
            get
            {
                if (instance == null)
                {
                    instance = new GlobalSettings();
                }
                return instance;
            }
        }
    }
}

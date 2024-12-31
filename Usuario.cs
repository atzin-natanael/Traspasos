using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Pantalla_De_Control
{
    public class Usuario
    {
        public int Id { get; set; }  // LiteDB genera el ID automáticamente si no se establece
        public string UsuarioName { get; set; }
        public string Password { get; set; }
    }
}

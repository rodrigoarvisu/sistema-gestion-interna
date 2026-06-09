using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.Configuration;

namespace PROGRAMA_PC_BOMBEROS
{
    public class Conexion
    {
        public static string cadena = ConfigurationManager.ConnectionStrings["MiConexion"].ConnectionString;
    }
}

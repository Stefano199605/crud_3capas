using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;

namespace CapaDatos
{
    public class CD_Productos
    {
        private CD_Conexion Conexion = new CD_Conexion();

        SqlDataReader leer;
        DataTable tabla = new DataTable();
        SqlCommand comando = new SqlCommand();

        public DataTable Mostrar()
        {
            comando.Connection = Conexion.AbrirConexion();
            comando.CommandText = "MostrarProductos";
            comando.CommandType = CommandType.StoredProcedure;
            leer = comando.ExecuteReader();
            tabla.Load(leer);
            Conexion.CerrarConexion();
            return tabla;
        }

        public DataTable BuscarxNombre(string nombre)
        {
            comando.Connection = Conexion.AbrirConexion();
            comando.CommandText = "filtro_busqueda";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@filtro", nombre);
            SqlDataAdapter DA = new SqlDataAdapter(comando);
            DA.Fill(tabla);
            Conexion.CerrarConexion();
            return tabla;
        }
      
        public void Insertar(string nombre,string desc,string marca,double precio,int stock)
        {
            comando.Connection = Conexion.AbrirConexion();
            comando.CommandText = "InsetarProductos";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@nombre", nombre);
            comando.Parameters.AddWithValue("@descrip", desc);
            comando.Parameters.AddWithValue("@Marca", marca);
            comando.Parameters.AddWithValue("@precio", precio);
            comando.Parameters.AddWithValue("@stock", stock);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            Conexion.CerrarConexion();

        }
        public void Eliminar(int id)
        {
            comando.Connection = Conexion.AbrirConexion();
            comando.CommandText = "EliminarProducto";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@idpro", id);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            Conexion.CerrarConexion();
        }
        public void Editar(string nombre, string desc, string marca, double precio, int stock, int id)
        {
            comando.Connection = Conexion.AbrirConexion();
            comando.CommandText = "EditarProductos";
            comando.CommandType = CommandType.StoredProcedure;
            comando.Parameters.AddWithValue("@nombre", nombre);
            comando.Parameters.AddWithValue("@descrip", desc);
            comando.Parameters.AddWithValue("@Marca", marca);
            comando.Parameters.AddWithValue("@precio", precio);
            comando.Parameters.AddWithValue("@stock", precio);
            comando.Parameters.AddWithValue("@id", id);
            comando.ExecuteNonQuery();
            comando.Parameters.Clear();
            Conexion.CerrarConexion();
        }
    }
}

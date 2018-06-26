using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.SqlClient;
using Capa_Entidades;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;
using Microsoft.SqlServer.Server;

namespace Capa_Datos
{
    public class ClaseUpdateServicioWMSSAP
    {

        public static  void Update_soporte(Capa_Entidades.errorEntity  oReg)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = @"Data Source=SRVWMS01\V2005;Initial Catalog=DimakDb_QA;User ID=eric_alvarez;Password=dimak2018";

            try
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                SqlCommand cmd_PS = new SqlCommand();
                cmd_PS.Connection = cn;
                cmd_PS.CommandType = CommandType.StoredProcedure;
                cmd_PS.CommandText = "Update_LogServicioIdoc";
                cmd_PS.Parameters.AddWithValue("@LogI_id", oReg.LogI_id);
                cmd_PS.Parameters.AddWithValue("@LogI_ObservSoporte", oReg.LogI_ObservSoporte);
                cmd_PS.Parameters.AddWithValue("@LogI_StatusSoporte", oReg.LogI_StatusSoporte);
                cmd_PS.ExecuteNonQuery();
                
                MessageBox.Show("Observación guardada");
            }
            catch (Exception Ex)

            {

                throw new Exception("Error al editar fila : " + Ex.Message, Ex);

            }

            finally

            {

                cn.Close();

                System.Data.SqlClient.SqlConnection.ClearPool(cn);

                cn.Dispose();

            }
        }
    }
 
}

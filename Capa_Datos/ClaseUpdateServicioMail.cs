using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Capa_Datos
{
   public class ClaseUpdateServicioMail
    {
        public static void Update_soporteMail(Capa_Entidades.errorEntityMail oReg)
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
                cmd_PS.CommandText = "Update_LogServicioMail";
                cmd_PS.Parameters.AddWithValue("@LogI_id", oReg.LogI_id);
                cmd_PS.Parameters.AddWithValue("@LogS_ObservSoporte", oReg.LogS_ObservSoporte);
                cmd_PS.Parameters.AddWithValue("@LogS_StatusSoporte", oReg.LogS_StatusSoporte);
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

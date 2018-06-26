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
    public class ClaseUpdateServicioWMNSOleada
    {
        public static void Update_soporte(Capa_Entidades.errorEntityWMSOleada oReg)
        {
            SqlConnection cn = new SqlConnection();
            cn.ConnectionString = @"Data Source=101.1.1.14\V2012;Initial Catalog=WMS_Picking;User ID=eric_alvarez;Password=dimak2018";

            try
            {
                cn.Open();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = cn;
                SqlCommand cmd_PS = new SqlCommand();
                cmd_PS.Connection = cn;
                cmd_PS.CommandType = CommandType.StoredProcedure;
                cmd_PS.CommandText = "Update_Log_WMNS_Oleada";
                cmd_PS.Parameters.AddWithValue("@LogS_id", oReg.LogS_Id);
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

using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Capa_Datos
{
    public class ClaseLogServicioMail
    {
        public static DataTable CargaLogError_ServicoMail()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=SRVWMS01\V2005;Initial Catalog=DimakDb_QA;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "WNS_LogS_Error_ServicioMail";

                dtt.Load(cmd_PS.ExecuteReader(), LoadOption.OverwriteChanges);

                return dtt;


            }

            catch (Exception Ex)

            {

                throw new Exception("Error al listar comunas : " + Ex.Message, Ex);

            }

            finally

            {

                cn.Close();

                System.Data.SqlClient.SqlConnection.ClearPool(cn);

                cn.Dispose();

            }
        }
        public static DataTable Carga24hrs_ServicioMail()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=SRVWMS01\V2005;Initial Catalog=DimakDb_QA;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "WNS_LogS_Error_24hrs_ServicioMail";

                dtt.Load(cmd_PS.ExecuteReader(), LoadOption.OverwriteChanges);

                return dtt;


            }

            catch (Exception Ex)

            {

                throw new Exception("Error al listar comunas : " + Ex.Message, Ex);

            }

            finally

            {

                cn.Close();

                System.Data.SqlClient.SqlConnection.ClearPool(cn);

                cn.Dispose();

            }
        }

        public static DataTable Carga1000_ServicioMail()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=SRVWMS01\V2005;Initial Catalog=DimakDb_QA;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "WNS_LogS_Error_Ultimos1000_ServicioMail";

                dtt.Load(cmd_PS.ExecuteReader(), LoadOption.OverwriteChanges);

                return dtt;


            }

            catch (Exception Ex)

            {

                throw new Exception("Error al listar comunas : " + Ex.Message, Ex);

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

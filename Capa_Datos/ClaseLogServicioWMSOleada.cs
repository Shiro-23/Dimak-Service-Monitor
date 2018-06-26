using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Capa_Datos
{
    public class ClaseLogServicioWMSOleada
    {
        public static DataTable CargaLogError_ServicioWMSOleada()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=101.1.1.14\V2012;Initial Catalog=WMS_Picking;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "PICKING_WMS_Oleada_Error";

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
        public static DataTable carga24hrs_ServicioWMNSOleada()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=101.1.1.14\V2012;Initial Catalog=WMS_Picking;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "PICKING_WMNS_Oleada_24hrs";

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
        public static DataTable Carga1000_ServicioWMSOleada()

        {

            SqlConnection cn = new SqlConnection();

            cn.ConnectionString = @"Data Source=101.1.1.14\V2012;Initial Catalog=WMS_Picking;User ID=eric_alvarez;Password=dimak2018";

            try

            {

                cn.Open();

                SqlCommand cmd = new SqlCommand();

                cmd.Connection = cn;

                DataTable dtt = new DataTable();

                SqlCommand cmd_PS = new SqlCommand();

                cmd_PS.Connection = cn;

                cmd_PS.CommandType = CommandType.StoredProcedure;

                cmd_PS.CommandText = "PICKING_WMS_Oleada_Error_ultimos1000";

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

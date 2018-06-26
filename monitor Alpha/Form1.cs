using System;
using System.Windows.Forms;
using System.ServiceProcess;
using Capa_Datos;
using System.Data.SqlClient;
using System.Data;
using Capa_Entidades;

namespace monitor_Alpha
{
    public partial class frm_Dimak_Service_Monitor : Form
    {
        public frm_Dimak_Service_Monitor()
        {
            InitializeComponent();
            servEmails();
            servRepDC();
            servWMSSAP();
            servInfo();
            servWMSOlea();
            servSmartphone();
            grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.CargaLogError_ServicoMail();
            grdv_LogServicioWMSSAP.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioWMSSAP;
            grdv_LogServicioEmail.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioEmail;
            grdv_LogServicioReplicaCD.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioReplicaCD;
            grdv_LogServicioWMSOleada.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioWMSOleada;
            grdv_LogServicioServInfo.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioServInfo;
            grdv_LogServicioSmartphone.SelectionMode = DataGridViewSelectionMode.FullRowSelect;
            this.ActiveControl = grdv_LogServicioSmartphone;
        }

        // variables
        //ID es donde se almacena la id de la fila selecionada
        //ObservacionSoporte almacena los comentarios entregados por el usuario.
        //Estado soporte da un valor de true al campo tipo bit.
        string ID;
        string ObservacionSoporte;
        bool EstadoSoporte = true;
        // servEmails() controla el estado de Servicio_Emails, ademas activarlo en caso de no estarlo  
        public void servEmails()
        {
            try
            {
                ServiceController se = new ServiceController("Servicio_Emails", "101.1.1.14");

                if (se != null && se.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServicioEmail.Image = monitor_Alpha.Properties.Resources.checkv;

                }
                else
                {
                    pic_EstadoServicioEmail.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // servRepDC() controla el estado de Servicio_Replica_CD, ademas activarlo en caso de no estarlo  
        private void servRepDC()
        {
            try
            {
                ServiceController sr = new ServiceController("Servicio_Replica_CD", "101.1.1.14");

                if (sr != null && sr.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServicioReplica.Image = monitor_Alpha.Properties.Resources.checkv;
                }
                else
                {
                    pic_EstadoServicioReplica.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // servWMSSAP() controla el estado de Servicio WMS SAP, ademas activarlo en caso de no estarlo  
        private void servWMSSAP()
        {
            try
            {
                // Entre comillas se especifica el nombre del servicio que se quiere testear seguido de la IP del equipo
                // donde se encuentra alojado, en caso estar localmente no es necesario suministrar la IP
                ServiceController sws = new ServiceController("Servicio_WMS_SAP", "101.1.1.14");

                if (sws != null && sws.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServicioWMSSAP.Image = monitor_Alpha.Properties.Resources.checkv;
                }
                else
                {
                    pic_EstadoServicioWMSSAP.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // servInfo() controla el estado de Servicio ServInfo, ademas activarlo en caso de no estarlo  
        private void servInfo()
        {
            try
            {
                ServiceController si = new ServiceController("Servicio ServInfo", "101.1.1.14");

                if (si != null && si.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServicioServInfo.Image = monitor_Alpha.Properties.Resources.checkv;
                }
                else
                {
                    pic_EstadoServicioServInfo.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
        // servWMSOlea() controla el estado de WMS_Oleada, ademas activarlo en caso de no estarlo  
        private void servWMSOlea()
        {
            try
            {
                ServiceController swo = new ServiceController("WMS_Oleadas", "101.1.1.14");

                if (swo != null && swo.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServicioWMSOleada.Image = monitor_Alpha.Properties.Resources.checkv;
                }
                else
                {
                    pic_EstadoServicioWMSOleada.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void servSmartphone()
        {
            try
            {
                ServiceController ssm = new ServiceController("Servicio ServPPCCapGenBD", "101.1.1.14");

                if (ssm != null && ssm.Status == ServiceControllerStatus.Running)
                {
                    pic_EstadoServSmartphone.Image = monitor_Alpha.Properties.Resources.checkv;
                }
                else
                {
                    pic_EstadoServSmartphone.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // timer se encarga de ir ejecutando los metodos que mantienen los servicios cada 5 seg
        // para mantener actualizado su estado
        private void timer1_Tick(object sender, EventArgs e)
        {
            servEmails();
            servRepDC();
            servWMSSAP();
            servInfo();
            servWMSOlea();
            tim_servicios.Interval = Convert.ToInt32(txt_Intervalo_Timer_Servicios.Text) * 1000;
        }
        // boton buscar para ejecutar el filtro deseado u seleccionado
        private void button1_Click(object sender, EventArgs e)
        {
            buscar();
        }
        // no hace nada
        public void toolStripLabel1_Click(object sender, EventArgs e)
        {
            ClaseExportacionExcel exp = new ClaseExportacionExcel();
            if (tab_TablasLog.SelectedTab == tabp_LogServicoEmail)
            {
                exp.ExportarDatosExcel(grdv_LogServicioEmail);
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSSAP)
            {
                exp.ExportarDatosExcel(grdv_LogServicioWMSSAP);
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSOleada)
            {
                exp.ExportarDatosExcel(grdv_LogServicioWMSOleada);
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioSmartphone)
            {
                exp.ExportarDatosExcel(grdv_LogServicioSmartphone);
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioServInfo)
            {
                exp.ExportarDatosExcel(grdv_LogServicioServInfo);
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioReplicaCD)
            {
                exp.ExportarDatosExcel(grdv_LogServicioReplicaCD);
            }
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            tim_servicios.Enabled = false;
            chk_RefrescarServicosAutomatico.Checked = false;
            tim_servicios.Interval = Convert.ToInt32(txt_Intervalo_Timer_Servicios.Text) * 1000;
            tim_logs.Enabled = false;
            chk_RefrescarLogsAutomatico.Checked = false;
            tim_logs.Interval = Convert.ToInt32(txt_Intervalo_Timer_Logs.Text) * 1000;
            rdo_Default.Checked = true;
        }

        /* 
         * grd con una propiedad CellClick, al selecionar una fila y escribir directamente en la celda Obsevacion
         * de soporte, se almacenan los datos de ID Y Observacion de soporte en las respectivas cariables.
        */
        private void grdv_LogServicioWMSSAP_CellClick(object sender, DataGridViewCellEventArgs e)
        {

            ID = grdv_LogServicioWMSSAP.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioWMSSAP.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();

        }

        /*button1_Click_1 la funcion de este boton es enviar todos los datos a la capa entidad de la clase
         * TablaEntidad en el metodo errorEntity e inyectarles los valores correspondientes a los campos, los cuales
         * son consumidos por la clase ClaseUpdateServicioWMSSAP Para ejecutar el procedimiento de almacenado que 
         * contiene.
         * porterior a eso siendo guardados con exitos los datos procede a refrescar los datos de su tabla Log.
        */
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (tab_TablasLog.SelectedTab == tabp_LogServicoEmail)
            {
                errorEntityMail oLog = new errorEntityMail();

                oLog.LogI_id = Convert.ToInt32(ID);
                oLog.LogS_ObservSoporte = ObservacionSoporte;
                oLog.LogS_StatusSoporte = EstadoSoporte;

                try
                {
                    Capa_Datos.ClaseUpdateServicioMail.Update_soporteMail(oLog);
                    grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.CargaLogError_ServicoMail();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }


            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioReplicaCD)
            {
                errorEntityReplicaCD oLog = new errorEntityReplicaCD();

                oLog.LogS_id = Convert.ToInt32(ID);
                oLog.LogS_ObservSoporte = ObservacionSoporte;
                oLog.LogS_StatusSoporte = EstadoSoporte;

                try
                {
                    Capa_Datos.ClaseUpdateServicioReplicaCD.Update_soporteReplicaCD(oLog);
                    grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.CargaLogError_ServicoReplicaCD();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSSAP)
            {
                errorEntity oLog = new errorEntity();

                oLog.LogI_id = Convert.ToInt32(ID);
                oLog.LogI_ObservSoporte = ObservacionSoporte;
                oLog.LogI_StatusSoporte = EstadoSoporte;

                try
                {
                    Capa_Datos.ClaseUpdateServicioWMSSAP.Update_soporte(oLog);
                    grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.CargaLogError();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSOleada)
            {
                errorEntityWMSOleada oLog = new errorEntityWMSOleada();

                oLog.LogS_Id = Convert.ToInt32(ID);
                oLog.LogS_ObservSoporte = ObservacionSoporte;
                oLog.LogS_StatusSoporte = EstadoSoporte;

                try
                {
                    Capa_Datos.ClaseUpdateServicioWMNSOleada.Update_soporte(oLog);
                    grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.CargaLogError_ServicioWMSOleada();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
        }

        // button2_Click Su unica funcion es refrescar los datos de todas las tablas Log.
        private void button2_Click(object sender, EventArgs e)
        {
            if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSSAP)
            {
                grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.CargaLogError();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicoEmail)
            {
                grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.CargaLogError_ServicoMail();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioReplicaCD)
            {
                grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.CargaLogError_ServicoReplicaCD();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSOleada)
            {
                grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.CargaLogError_ServicioWMSOleada();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioServInfo)
            {
                grdv_LogServicioServInfo.DataSource = ClaseLogServicioServInfo.CargaLogError_ServicioServInfo();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioSmartphone)
            {
                grdv_LogServicioSmartphone.DataSource = ClaseLogServicioSmartphone.CargaLogError_ServicioSmartphone();
            }
        }

        // toolStripLabel2_Click su funcion es preguntar si desea cerrar el programa al precionar la opcion cerrar.
        private void toolStripLabel2_Click(object sender, EventArgs e)
        {
            DialogResult res = MessageBox.Show("¿Desea cerrar el programa?",
               "Cerrar el programa", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (res == DialogResult.Yes)
            {
                Close();
            }

        }

        private void chk_RefrescarAutomatico_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_RefrescarServicosAutomatico.Checked == true)
            {
                tim_servicios.Enabled = false;
                tim_servicios.Interval = Convert.ToInt32(txt_Intervalo_Timer_Servicios.Text) * 1000;
                tim_servicios.Enabled = true;
            }
            else
            {
                tim_servicios.Enabled = false;
            }
        }

        private void pic_EstadoServicioEmail_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController se = new ServiceController("Servicio_Emails", "101.1.1.14");

                if (se != null && se.Status == ServiceControllerStatus.Running)
                {
                    se.Stop();
                    pic_EstadoServicioEmail.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    se.Start();
                    pic_EstadoServicioEmail.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_EstadoServicioReplica_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController sr = new ServiceController("Servicio_Replica_CD", "101.1.1.14");

                if (sr != null && sr.Status == ServiceControllerStatus.Running)
                {
                    sr.Stop();
                    pic_EstadoServicioReplica.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    sr.Start();
                    pic_EstadoServicioReplica.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_EstadoServicioWMSSAP_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController sws = new ServiceController("Servicio_WMS_SAP", "101.1.1.14");

                if (sws != null && sws.Status == ServiceControllerStatus.Running)
                {
                    sws.Stop();
                    pic_EstadoServicioWMSSAP.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    sws.Start();
                    pic_EstadoServicioWMSSAP.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_EstadoServicioServInfo_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController si = new ServiceController("Servicio ServInfo", "101.1.1.14");

                if (si != null && si.Status == ServiceControllerStatus.Running)
                {
                    si.Stop();
                    pic_EstadoServicioServInfo.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    si.Start();
                    pic_EstadoServicioServInfo.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_EstadoServicioWMSOleada_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController swo = new ServiceController("WMS_Oleadas", "101.1.1.14");

                if (swo != null && swo.Status == ServiceControllerStatus.Running)
                {
                    swo.Stop();
                    pic_EstadoServicioWMSOleada.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    swo.Start();
                    pic_EstadoServicioWMSOleada.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void pic_EstadoServSmartphone_Click(object sender, EventArgs e)
        {
            try
            {
                ServiceController ssm = new ServiceController("Servicio ServPPCCapGenBD", "101.1.1.14");

                if (ssm != null && ssm.Status == ServiceControllerStatus.Running)
                {
                    ssm.Stop();
                    pic_EstadoServSmartphone.Image = monitor_Alpha.Properties.Resources.Cruz_roja;
                    MessageBox.Show("Servicio Detenido", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    ssm.Start();
                    pic_EstadoServSmartphone.Image = monitor_Alpha.Properties.Resources.checkv;
                    MessageBox.Show("Servicio Iniciado", "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Information);                                       
                }
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.Message, "Monitor servicios", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void grdv_LogServicioEmail_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ID = grdv_LogServicioEmail.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioEmail.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();
        }

        private void grdv_LogServicioReplicaCD_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ID = grdv_LogServicioReplicaCD.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioReplicaCD.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();
        }

        private void grdv_LogServicioWMSOleada_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ID = grdv_LogServicioWMSOleada.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioWMSOleada.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();
        }

        private void grdv_LogServicioServInfo_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ID = grdv_LogServicioServInfo.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioServInfo.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();
        }

        private void grdv_LogServicioSmartphone_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            ID = grdv_LogServicioSmartphone.SelectedRows[0].Cells["ID"].Value.ToString();
            ObservacionSoporte = grdv_LogServicioSmartphone.SelectedRows[0].Cells["Observacion de soporte"].Value.ToString();
        }

        private void tab_TablasLog_Click(object sender, EventArgs e)
        {
            buscar();
        }

        private void tim_logs_Tick(object sender, EventArgs e)
        {
            buscar();
            tim_logs.Interval = Convert.ToInt32(txt_Intervalo_Timer_Servicios.Text) * 1000;
        }

        private void Cargar_Tablas_Logs()
        {
            if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSSAP)
            {
                grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.CargaLogError();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicoEmail)
            {
                grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.CargaLogError_ServicoMail();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioReplicaCD)
            {
                grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.CargaLogError_ServicoReplicaCD();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSOleada)
            {
                grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.CargaLogError_ServicioWMSOleada();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioServInfo)
            {
                grdv_LogServicioServInfo.DataSource = ClaseLogServicioServInfo.CargaLogError_ServicioServInfo();
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioSmartphone)
            {
                grdv_LogServicioSmartphone.DataSource = ClaseLogServicioSmartphone.CargaLogError_ServicioSmartphone();
            }
        }

        private void chk_RefrescarLogsAutomatico_CheckedChanged(object sender, EventArgs e)
        {
            if (chk_RefrescarLogsAutomatico.Checked == true)
            {
                tim_logs.Enabled = false;
                tim_logs.Interval = Convert.ToInt32(txt_Intervalo_Timer_Logs.Text) * 1000;
                tim_logs.Enabled = true;
            }
            else
            {
                tim_servicios.Enabled = false;
            }
        }
        private void buscar()
        {
            //Consulta cual tabPage esta selecionada para aplicarle el filtro seleccionado
            if (tab_TablasLog.SelectedTab == tabp_LogServicoEmail)
            {
                //verifica cual de las 3 opciones de friltro esta selecionada y al precionar el boton buscar carga el stored procedure correspondiente
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.Carga24hrs_ServicioMail();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.Carga1000_ServicioMail();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioEmail.DataSource = ClaseLogServicioMail.CargaLogError_ServicoMail();
                }

            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioReplicaCD)
            {
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.Carga24hrs_ServicioReplicaCD();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.Carga1000_ServicioReplicaCD();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioReplicaCD.DataSource = ClaseLogServicioReplicaCD.CargaLogError_ServicoReplicaCD();
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSSAP)
            {
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.Carga24hrs();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.Carga1000();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioWMSSAP.DataSource = ClaseLogServicioWMSSAP.CargaLogError();
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioWMSOleada)
            {
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.carga24hrs_ServicioWMNSOleada();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.Carga1000_ServicioWMSOleada();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioWMSOleada.DataSource = ClaseLogServicioWMSOleada.CargaLogError_ServicioWMSOleada();
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioServInfo)
            {
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioServInfo.DataSource = ClaseLogServicioServInfo.Carga24hrs_ServicioServInfo();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioServInfo.DataSource = ClaseLogServicioServInfo.Carga1000_ServicioServInfo();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioServInfo.DataSource = ClaseLogServicioServInfo.CargaLogError_ServicioServInfo();
                }
            }
            else if (tab_TablasLog.SelectedTab == tabp_LogServicioSmartphone)
            {
                if (rdo_ErroresUltimas24Hrs.Checked == true)
                {
                    grdv_LogServicioSmartphone.DataSource = ClaseLogServicioSmartphone.Carga24hrs_ServicioSmartphone();
                }
                else if (rdo_Ultimos1000Errores.Checked == true)
                {
                    grdv_LogServicioSmartphone.DataSource = ClaseLogServicioSmartphone.Carga1000_ServicioSmartphone();
                }
                else if (rdo_Default.Checked == true)
                {
                    grdv_LogServicioSmartphone.DataSource = ClaseLogServicioSmartphone.CargaLogError_ServicioSmartphone();
                }
            }
        }

        private void txt_Intervalo_Timer_Logs_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }

        private void txt_Intervalo_Timer_Servicios_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (!(char.IsNumber(e.KeyChar)) && (e.KeyChar != (char)Keys.Back))
            {
                MessageBox.Show("Solo se permiten numeros", "Advertencia", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                e.Handled = true;
                return;
            }
        }
    }
}

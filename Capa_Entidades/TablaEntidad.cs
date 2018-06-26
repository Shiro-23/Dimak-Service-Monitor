using Microsoft.SqlServer.Server;
using System;
using System.Collections.Generic;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.DataFormats;


namespace Capa_Entidades
{
    public  class errorEntity
    {        
            public  string LogI_ObservSoporte { get ; set ;}
            public  bool LogI_StatusSoporte { get ; set ;} 
            public   int LogI_id { get; set; }
            public  int LogI_Fecha { get; set; }
            public  string LogI_TipoDoc { get; set; }
            public  string LogI_Mensaje { get; set; }
            public  int LogI_IndError { get; set; }

        
    }

    public class errorEntityMail
    {
        public int LogI_id { get; set; }

        public DateTime LogI_Fecha { get; set; }

        public String LogI_Mensaje { get; set; }

        public int LogS_IndError { get; set; }

        public string LogS_ObservSoporte { get; set; }

        public bool LogS_StatusSoporte { get; set; }
    }

    public class errorEntityReplicaCD
    {
        public int LogS_id { get; set; }

        public DateTime LogS_Fecha { get; set; }

        public string LogS_Mensaje { get; set; }

        public int LogS_IndError { get; set; }

        public string LogS_ObservSoporte { get; set; }

        public bool LogS_StatusSoporte { get; set; }
    }

    public class errorEntityWMSOleada
    {
        public int LogS_Id { get; set; }

        public DateTime LogS_Fecha { get; set; }

        public string LogS_Mensaje { get; set; }

        public string LogS_ObservSoporte { get; set; }

        public bool LogS_StatusSoporte { get; set; }
    }
}

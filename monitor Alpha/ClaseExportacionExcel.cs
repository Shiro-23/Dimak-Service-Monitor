using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Core;


namespace monitor_Alpha
{
    class ClaseExportacionExcel
    {
        public void ExportarDatosExcel(DataGridView grd)
        {
            try
            {
                SaveFileDialog fichero = new SaveFileDialog();
                fichero.Filter = "Excel (*.xls)|*.xls";
                fichero.FileName = "Archivo Exportado";
                if (fichero.ShowDialog() == DialogResult.OK)
                {
                    Microsoft.Office.Interop.Excel.Application aplicacion;
                    Microsoft.Office.Interop.Excel.Workbook libros_trabajos;
                    Microsoft.Office.Interop.Excel.Worksheet hoja_trabajos;

                    aplicacion = new Microsoft.Office.Interop.Excel.Application();
                    libros_trabajos = aplicacion.Workbooks.Add();
                    hoja_trabajos = (Microsoft.Office.Interop.Excel.Worksheet)libros_trabajos.Worksheets.get_Item(1);
                    for (int i = 1; i <= grd.Columns.Count; i++)
                    {
                        hoja_trabajos.Cells[1, i] = grd.Columns[i - 1].HeaderText;
                    }
                    for (int i = 0; i < grd.Rows.Count - 1; i++)
                    {
                        for (int j = 0; j < grd.Columns.Count; j++)
                        {
                            hoja_trabajos.Cells[i + 2, j + 1]= grd.Rows[i].Cells[j].Value.ToString();
                        }
                    }
                    libros_trabajos.SaveAs(fichero.FileName, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal);
                    libros_trabajos.Close(true);
                    aplicacion.Quit();
                    MessageBox.Show("Archivo Extraido","Monitor Servicio", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error al intentar exporta debido a: " + ex.ToString());
            }
        }
    }
}

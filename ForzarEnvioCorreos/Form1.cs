using MySqlConnector;
using SpreadsheetLight;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Net;
using System.Net.Mail;

namespace ForzarEnvioCorreos
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void Form1_Load(object sender, EventArgs e)
        {
        }

        private void leerexcel()
        {
            SLDocument sl = new SLDocument(@"D:\excell\Libro1.xlsx");
            SLWorksheetStatistics propiedades = sl.GetWorksheetStatistics();

            int ultimaFila = propiedades.EndRowIndex;

            for (int x = 2; x <= ultimaFila; x++)
            {
                String idreclamo = sl.GetCellValueAsString("A" + x);
                String rec_fecha = sl.GetCellValueAsString("B" + x);
                String sucu = sl.GetCellValueAsString("C" + x);
                String rec_nombre = sl.GetCellValueAsString("D" + x);
                String rec_telefono = sl.GetCellValueAsString("E" + x);
                String rec_email = sl.GetCellValueAsString("F" + x);
                String rec_detalle = sl.GetCellValueAsString("G" + x);

                try
                {
                    using (MailMessage mailMessage = new MailMessage())
                    {
                        //mailMessage.To.Add("reclamos@expresomarvisur.com");
                        mailMessage.To.Add("desarrollomarvisur.02@gmail.com");
                        mailMessage.Subject = "RECLAMO - " + idreclamo;
                        mailMessage.Body =
                            "DATOS DEL RECLAMO :\n"
                            + "*****************************************************\n"
                            + "Fecha        : " + rec_fecha + "\n"
                            + "Sucursal    : " + sucu + "\n"
                            + "Nombre     : " + rec_nombre + "\n"
                            + "Telefono    : " + rec_telefono + "\n"
                            + "Email         : " + rec_email + "\n"
                            + "Detalle       : " + rec_detalle + "\n"
                            + "Detalles     : https://www.expresomarvisur.com/reclamaciones/print.php?numero=" + idreclamo + "\n"
                            + "*****************************************************";

                        mailMessage.IsBodyHtml = false;

                        mailMessage.From = new MailAddress("desarrollomarvisur.02@gmail.com", "RECLAMOS");

                        using (SmtpClient cliente = new SmtpClient())
                        {
                            cliente.UseDefaultCredentials = false;
                            cliente.Credentials = new NetworkCredential("desarrollomarvisur.02@gmail.com", "ktzzozirepntebcl");
                            cliente.Port = 587;
                            cliente.EnableSsl = true;
                            cliente.Host = "smtp.gmail.com";
                            cliente.Send(mailMessage);
                        }
                    }
                }
                catch (Exception)
                {
                    throw;
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            leerexcel();
        }
    }
}

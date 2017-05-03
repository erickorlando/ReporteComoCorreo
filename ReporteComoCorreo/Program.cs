/* 
 * Este proyecto pretende mostrar como ejecutar un reporte RDLC (Reporting Services)
 * como un archivo adjunto PDF en un correo electrónico en tiempo de ejecución
 * sin necesidad de hacer un preview del reporte.
 * Autor: Erick Orlando (http://t.me/erickorlando)
 */

using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Mail;
using ErickOrlando.Utilidades.Data;
using Microsoft.Reporting.WebForms;
using ReporteComoCorreo.Data;

namespace ReporteComoCorreo
{
    class Program
    {
        static void Main()
        {
            try
            {
                Console.WriteLine("Demo de Reporte RDLC como Correo Electrónico Adjunto");

                // Cargamos con data la clase Cliente.
                var listaClientes = new List<Cliente>();
                var generadorData = new DataGenerator();
                for (int i = 1; i < 51; i++)
                {
                    listaClientes.Add(new Cliente
                    {
                        Id = i,
                        Nombre = generadorData.GetFirstName(),
                        Apellidos = generadorData.GetLastName(),
                        Correo = generadorData.GetEmail(),
                        Empresa = generadorData.GetCompanyName()
                    });
                }

                var aliasfrom = GetCampo("Nombre Remitente","Pepito");
                var emailFrom = GetCampo("Email Remitente","pepito@mail.com");

                var aliasTo = GetCampo("Nombre Destinatario","Juanito");
                var emailTo = GetCampo("Email Destinatario", "juanito@mail.com");

                var host = GetCampo("Servidor SMTP","smtp.gmail.com");
                var puertoSeguro = GetCampo("El Servidor SMTP usa un puerto Seguro? [S (SÍ)/N (NO)]").ToLower() == "s";

                using (var viewer = new LocalReport())
                {
                    Warning[] warnings;
                    string[] streamIds;
                    string mimeType;
                    string encoding;
                    string filenameExtension;

                    Console.WriteLine("Renderizando reporte....");
                    viewer.DataSources.Add(new ReportDataSource("data", listaClientes));
                    viewer.Refresh();
                    // Para que esta línea funcione se debe escoger el archivo DemoReporte.rdlc y en las propiedades de archivo
                    // ajustarlo a "Copiar siempre" en 'Acción de Compilación'.
                    viewer.ReportPath = "./Reports/DemoReporte.rdlc";
                    var bytes = viewer.Render("PDF", null, out mimeType, out encoding, out filenameExtension,
                        out streamIds, out warnings);

                    var correo = new MailMessage { From = new MailAddress(emailFrom, aliasfrom) };

                    correo.To.Add(new MailAddress(emailTo, aliasTo));
                    correo.Subject = "Reporte como Correo";
                    correo.Attachments.Add(new Attachment(new MemoryStream(bytes), "Reporte.pdf"));

                    correo.Body = "Estimado usuario, se le adjunta el reporte.";

                    using (var smtpClient = new SmtpClient(host))
                    {
                        if (puertoSeguro)
                        {
                            smtpClient.EnableSsl = true;
                            smtpClient.Port = 587;
                        }
                        smtpClient.Credentials =
                            new System.Net.NetworkCredential(emailFrom, GetClave());
                        Console.WriteLine("Espere unos segundos....");
                        smtpClient.Send(correo);
                    }

                    Console.WriteLine("Correo enviado");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                Console.ReadLine();
            }
        }

        private static string GetCampo(string nombreCampo, string sugerencia = "")
        {
            Console.WriteLine("Ingrese el valor para {0}{1}:", nombreCampo, string.IsNullOrEmpty(sugerencia) ? sugerencia : $" Ejm:({sugerencia})" );
            return Console.ReadLine();
        }

        private static string GetClave()
        {
            Console.WriteLine("Ingrese contraseña de servidor SMTP y luego presione Enter");
            ConsoleKeyInfo key;
            var clave = string.Empty;
            do
            {
                key = Console.ReadKey(true);
                // Ignorar tecla Backspace y Enter.
                if (key.Key != ConsoleKey.Backspace && key.Key != ConsoleKey.Enter)
                {
                    clave += key.KeyChar;
                    Console.Write("*");
                }
                else
                {
                    // En caso se desee borrar una letra.
                    if (key.Key != ConsoleKey.Backspace || clave.Length <= 0) continue;
                    clave = clave.Substring(0, clave.Length - 1);
                    Console.Write("\b \b");
                }
            }
            // Detener el bucle cuando se recibe Enter.
            while (key.Key != ConsoleKey.Enter);
            return clave;
        }
    }
}

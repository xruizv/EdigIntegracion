using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using System.Data.OleDb;
using System.Configuration;
using System.IO;
using System.Xml.Linq;
using System.Reflection;
using ClosedXML.Excel;
using System.Globalization;
using System.Threading;

namespace Integracion
{
    class Program
    {
        static void Main(string[] args)
        {
            //Configurración de zona
            Thread.CurrentThread.CurrentCulture = new CultureInfo("es-cl");
            // Inicialización de variables
            var appSettings = ConfigurationManager.AppSettings;
            string rutaContribucionesF29 = appSettings.GetValues("contribucionesF29").First().ToString();
            string rutaRemuneraciones = appSettings.GetValues("remuneraciones").First().ToString();
            string rutaSalida = appSettings.GetValues("salida").First().ToString();
            // Vemos que mes vamos a procesar, si no está en el app.config cargamos el mes actual
            string anio = String.IsNullOrEmpty(appSettings.GetValues("anio").First().ToString()) ? DateTime.Now.Year.ToString("D2") : appSettings.GetValues("anio").First().ToString();
            string mes = String.IsNullOrEmpty(appSettings.GetValues("mes").First().ToString()) ? DateTime.Now.Month.ToString("D4") : appSettings.GetValues("mes").First().ToString();
            IList<Cobro> cobros = new List<Cobro>();
            //Form29
            string[] subdirectoryEntriesF29 = Directory.GetDirectories(rutaContribucionesF29);
            string pathEmpresas = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ConfiguracionEmpresas.xml");
            XElement configuracionEmpresas = XElement.Load(pathEmpresas);
            IEnumerable<XElement> empresas = configuracionEmpresas.Descendants("empresa");
            foreach (var item in subdirectoryEntriesF29)
            {
                string[] filesEntries = Directory.GetFiles(item);
                string nombreArchivoMDBF29 = filesEntries.SingleOrDefault(x => x.Contains(String.Format("F29LGH{0}.MDB", anio.Substring(2, 2))));
                string connectionString = String.Format("Provider=Microsoft.Jet.OLEDB.4.0;Data Source={0}", nombreArchivoMDBF29);//User Id=admin;Password=;
                string queryString = String.Format("SELECT d545 + d546 + d708 FROM Form29 WHERE Mes = {0}", int.Parse(mes).ToString());

                using (OleDbConnection connection = new OleDbConnection(connectionString))
                {
                    double suma = 0;
                    OleDbCommand command = new OleDbCommand(queryString, connection);
                    try
                    {
                        connection.Open();
                        OleDbDataReader reader = command.ExecuteReader();
                        while (reader.Read())
                        {
                            double numero = 0;
                            double.TryParse(reader.GetValue(0).ToString(), out numero);
                            suma += numero;
                        }
                        reader.Close();
                        int rut;
                        int.TryParse(Path.GetFileName(item), out rut);
                        var empresa = empresas.SingleOrDefault(x => x.Element("rut").Value == rut.ToString());
                        Rut bufferRut = new Rut(empresa.Element("rut").Value);
                        Cobro cobro = new Cobro()
                        {
                            Empresa = empresa.Element("nombre").Value,
                            Fecha = File.GetCreationTime(item),
                            Moneda = "Pesos",
                            Monto = suma,
                            Producto = "Formulario 29",
                            Origen = "Interno",
                            RUT = bufferRut.Numero.ToString(),
                            DV = bufferRut.DV.ToString(),
                            NumeroDocumento = String.Empty
                        };
                        cobros.Add(cobro);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine(ex.Message);
                    }
                }
            }
            // Remuneraciones formato Previred
            string linea = "";
            string[] subdirectoryEntriesRemuneraciones = Directory.GetDirectories(rutaRemuneraciones);
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), "ConfiguracionPrevired.xml");
            XElement configuracionPrevired = XElement.Load(path);
            IEnumerable<XElement> raiz = configuracionPrevired.Descendants("previred");
            foreach (var item in subdirectoryEntriesRemuneraciones)
            {
                double acumulador = 0;
                string[] filesEntries = Directory.GetFiles(item);
                string nombreArchivoPrevired = filesEntries.SingleOrDefault(x => x.Contains(String.Format("PREVIRED{0}{1}", mes, anio)));
                using (StreamReader sr = new StreamReader(nombreArchivoPrevired))
                {
                    while ((linea = sr.ReadLine()) != null)
                    {
                        foreach (var subitem in raiz)
                        {
                            int largo = 0;
                            int.TryParse(subitem.Element("largo").Value, out largo);
                            int inicio = 0;
                            int.TryParse(subitem.Element("inicio").Value, out inicio);
                            int monto = 0;
                            if (int.TryParse(linea.Substring(inicio, largo), out monto))
                            {
                                acumulador += monto;
                            }
                        }
                    }
                }
                int identificador;
                int.TryParse(Path.GetFileName(item).Replace("EMP",""), out identificador);
                var empresa = empresas.SingleOrDefault(x => x.Element("identificador").Value == identificador.ToString());
                Rut bufferRut = new Rut(empresa.Element("rut").Value);
                Cobro cobro = new Cobro() {
                    Empresa = empresa.Element("nombre").Value,
                    Fecha = File.GetCreationTime(item),
                    Moneda = "Pesos",
                    Monto = acumulador,
                    Producto = "Imposiciones",
                    Origen = "Interno",
                    RUT = bufferRut.Numero.ToString(),
                    DV = bufferRut.DV.ToString(),
                    NumeroDocumento = String.Empty
                };
                cobros.Add(cobro);
            }
            var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Cobros");
            IXLRow cabecera = worksheet.Row(1);
            cabecera.Cell(1).SetValue("Empresa - Nombre de la empresa");
            cabecera.Cell(2).SetValue("Rut ");
            cabecera.Cell(3).SetValue("DV");
            cabecera.Cell(4).SetValue("Origen");
            cabecera.Cell(5).SetValue("Producto");
            cabecera.Cell(6).SetValue("Fecha");
            cabecera.Cell(7).SetValue("Monto");
            cabecera.Cell(8).SetValue("Moneda - CLP");
            cabecera.Cell(9).SetValue("Numero Documento");
            cabecera.Cell(10).SetValue("Correlativo");
            int i = 2;
            foreach(var item in cobros)
            {
                IXLRow fila = worksheet.Row(i);
                fila.Cell(1).SetValue(item.Empresa);
                fila.Cell(2).SetValue(item.RUT);
                fila.Cell(3).SetValue(item.DV);
                fila.Cell(4).SetValue(item.Origen);
                fila.Cell(5).SetValue(item.Producto);
                fila.Cell(6).SetValue(item.Fecha.ToShortDateString());
                fila.Cell(7).SetValue(item.Monto.ToString("C0"));
                fila.Cell(8).SetValue(item.Moneda);
                fila.Cell(9).SetValue(item.NumeroDocumento);
                i++;
            }
            workbook.SaveAs(rutaSalida+"ArchivoSalida.xlsx");
        }
    }
    public class Cobro
    {
        public string Empresa { get; set; }
        public string RUT { get; set; }
        public string DV { get; set; }
        public string Origen { get; set; }
        public string Producto { get; set; }
        public DateTime Fecha { get; set; }
        public double Monto { get; set; }
        public string Moneda { get; set; }
        public string NumeroDocumento { get; set; }
    }
}

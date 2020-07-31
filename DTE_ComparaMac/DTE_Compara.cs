using System;
using System.Data;
using System.IO;
using ExcelDataReader;

using System.Collections.Generic;
using System.Net;
using System.Linq;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Net.Mail;

namespace DTE_ComparaArco
{
    public class DTE_Compara
    {

        public DTE_Compara()
        {
            
            // Get Params
            Params.RutSociedades = new List<string[]>();

            
            var filename = "ArcoDTE_ComparaConfig.xml";
            var currentDirectory = AppDomain.CurrentDomain.BaseDirectory; //Directory.GetCurrentDirectory();
            var settingsFilepath = Path.Combine(currentDirectory, filename);

            oLog = new DTE_log(currentDirectory);
            //ERROR, WARNING,  DEBUG, INFO, TRACE
            oLog.Add("TRACE","Get Settings: " + settingsFilepath);

            try
            {
                XElement Properties_Settings = XElement.Load(settingsFilepath);

                IEnumerable<XElement> setts = from parametro in Properties_Settings.Descendants("ArcoXmlEmitidos.Properties.Settings")
                                              select (XElement)parametro;

                foreach (XElement el in setts.Elements())
                {

                    if (el.Attribute("name").Value == "RutSociedades")
                    {
                        XElement nodoSoc = XElement.Parse(el.FirstNode.ToString());
                        IEnumerable<XElement> elxs = from lx in nodoSoc.Descendants("ArrayOfString")
                                                     select (XElement)lx;
                        foreach (XElement elx in elxs.Elements())
                        {
                            Params.RutSociedades.Add(elx.Value.Split(';'));
                        }
                    }

                    switch (el.Attribute("name").Value)
                    {
                        case "SMTPName":
                            Params.SMTPName = el.Value;
                            break;
                        case "SMTPPort":
                            Params.SMTPPort = Convert.ToInt32(el.Value);
                            break;
                        case "EnableSSL":
                            Params.EnableSSL = Convert.ToBoolean(el.Value);
                            break;
                        case "EmailUser":
                            Params.EmailUser = el.Value;
                            break;
                        case "EmailPassword":
                            Params.EmailPassword = el.Value;
                            break;
                        case "EmailTO":
                            Params.EmailTO = el.Value;
                            break;
                        case "EmailTO2":
                            Params.EmailTO2 = el.Value;
                            break;
                        case "DirectorioExcelSigge":
                            Params.DirectorioExcelSigge = el.Value;
                            break;

                    }
                    // TODO: Normalizar antes de Producción
                    //Params.PeriodoEmision = "2020-06"; DateTime.Now.AddDays(-1).ToString("yyyy-MM");
                    Params.PeriodoEmision = DateTime.Now.AddDays(-1).ToString("yyyy-MM");

                }

                oLog.Add("TRACE", "Get Settings Successed");
            }
            catch (Exception ex)
            {
                oLog.Add("ERROR", ex.Message);
                throw new Exception(ex.Message);
            }
        }



        public static int cuadrado(int x)
        {
            return x * x;
        }

        static async Task Main(string[] args)
        {
            DTE_Compara DcP = new DTE_Compara();
            oLog.Add("DEBUG", "======== Inicio Proceso ========");

            try
            {
                foreach (string[] sociedad in Params.RutSociedades)
                {
                    // Verifica existencia Planilla Excel en Directorio Params.DirectorioExcelSigge
                    // Formato: Alias__AAAA-MM.xls o Alias__AAAA-MM.xlsx
                    string dirSigge = Params.DirectorioExcelSigge != "" ? Params.DirectorioExcelSigge : AppDomain.CurrentDomain.BaseDirectory; 
                    string fileExcelxls = Path.Combine(dirSigge, sociedad[2] + "__" + Params.PeriodoEmision + ".xls");
                    string fileExcelxlsx = Path.Combine(dirSigge, sociedad[2] + "__" + Params.PeriodoEmision + ".xlsx");

                    string fileExcel = File.Exists(fileExcelxls) ? fileExcelxls : File.Exists(fileExcelxlsx) ? fileExcelxlsx : "";

                    if (File.Exists(fileExcel))
                    {
                        oLog.Add("TRACE",
                            String.Format("Sociedad {0} planilla Sigge encontrada periodo {1} ",
                            sociedad[2], Params.PeriodoEmision));

                        //Llamada a WebService
                        await callWS(sociedad);
                        // Carga Excel's
                        loadWorkbook(sociedad, fileExcel);
                    }
                    else
                    {
                        oLog.Add("WARNING",
                            String.Format("Sociedad {0} sin planilla Sigge periodo {1}",
                            sociedad[2], Params.PeriodoEmision));

                    }

                }


                // Envía correo
                if (DTE_Segge.Count() != 0)
                    sendMail();

                oLog.Add("TRACE", "======== Fin Proceso ========");
            }
            catch (Exception ex)
            {
                oLog.Add("ERROR", ex.Message);
                //throw new Exception(ex.Message);

            }

        }

        static void loadWorkbook(string[] sociedad, string filePath)
        {
            // TODO: Cambiar variable
            //filePath = @"/Users/cec/Projects/DTECompara/DTECompara/Cec2.xlsx";
            oLog.Add("TRACE", String.Format("Leyendo Excel {0}", filePath));

            try { 
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        var result = reader.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true 
                            }
                        });

                        // Folio DTE Posición 7 
                        IEnumerable<DataRow> ieFilas = from fila in result.Tables[0].AsEnumerable()
                                                       where DBNull.Value.Equals(fila[7])
                                                       select fila;

                        int xFilas = 0;
                        int xFilasDTE = 0;
                        foreach (var iRow in ieFilas)
                        {
                            DTE DTEl = new DTE();

                            DTEl.empresa = sociedad[2];
                            DTEl.tipoDTE = isDbNull(iRow, 8); //DBNull.Value.Equals(iRow[8])? iRow[8].ToString():"";
                            //DTEl.folioDTE = 0;  //iRow[7].ToString();  //No puede Castear un DBNull
                            DTEl.rutReceptor = iRow[1].ToString();
                            DTEl.razonSocialReceptor = iRow[2].ToString();
                            DTEl.Glosa = iRow[5].ToString();
                            DTEl.Codigo = iRow[6].ToString();
                            DTEl.Neto = Int32.Parse(isDbNull(iRow, 9));
                            DTEl.Iva = Int32.Parse(isDbNull(iRow, 10));
                            DTEl.Total = Int32.Parse(isDbNull(iRow, 11));
                            DTEl.fechaFacturacion = isDbNull(iRow, 12);
                            DTEl.fechaPortalCen = isDbNull(iRow, 15);
                            DTEl.fechaCarga = isDbNull(iRow, 17);


                            // Busqueda con LINQ Lista IEnumerator por si se repiten facturas iguales (mismo criterio)
                            List<DTE> xRows = DTE_WSFacele
                                 .Where(x =>
                                 {
                                     if (
                                        (x.tipoDTE == DTEl.tipoDTE) 
                                        && (x.rutReceptor == DTEl.rutReceptor) 
                                        && (x.Neto == DTEl.Neto) 
                                        && (x.Total == DTEl.Total)
                                         )
                                         return true;
                                     else 
                                         return false;
                                 })
                                 .ToList();

                            foreach (var xRow in xRows)  //IEnumerator evita crash si se repite docto
                            {
                                    DTEl.folioDTE = xRow.folioDTE;

                                    DTEl.fechaEmision = xRow.fechaEmision;
                                    DTEl.fechaFirma = xRow.fechaFirma;
                                    DTEl.fechaRegistroSII = xRow.fechaRegistroSII;
                                    DTEl.estadoSII = xRow.estadoSII;
                                    DTEl.estadoRecepcion = xRow.estadoRecepcion;
                                    DTEl.estadoLey19983 = xRow.estadoLey19983;
                                    DTEl.estadoLey20956 = xRow.estadoLey20956;
                                    xFilasDTE++;
                            }

                            // Agrega datos a Lista
                            DTE_Segge.Add(DTEl);
                            xFilas++;


                        }
                        
                        oLog.Add("INFO", String.Format("{0} Registros DTE obtenidos.  {1} Sin DTE, {2} encontrados en WS Facele",
                            result.Tables[0].Rows.Count.ToString(), xFilas.ToString(), xFilasDTE.ToString()));
                    }
                }

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error leer archivo Excel Sigge {0} para periodo {1} {2}",
                    filePath, Params.PeriodoEmision, ex.Message));
                //throw new Exception(ex.Message);

            }
        }

        static async Task callWS(string[] sociedad)
        {
            oLog.Add("TRACE", String.Format("Llamando Web service {0}", sociedad[2]));

            try
            {
                // Cnt Registros
                string rspBody = await CallEndPointAsync( sociedad[0], 999999);

                StringReader tReader = new StringReader(rspBody);
                DataSet tDSet = new DataSet();
                tDSet.ReadXml(tReader);

                IEnumerable<DataRow> ieXmls = from fila in tDSet.Tables["consultarResponse"].AsEnumerable()
                                              select fila;

                // TODO: Mejorar llamada LinQ, pasar DataRow a XElement o XMLNode
                int cntReg = 0;
                foreach (var ieXml in ieXmls)
                {
                    cntReg = Convert.ToInt32(ieXml[1]);
                    break;
                }

                oLog.Add("TRACE", String.Format("Cargando {0} DTE's ...", cntReg.ToString()));

                // Ciclos de 100 registros determinados por Facele (WSDL)
                for (int i = 0; i < ((cntReg / 100) + (cntReg % 100 == 0 ? 0 : 1)); i++)
                {
                    try
                    {
                     rspBody = await CallEndPointAsync(sociedad[0], i*100);
                     StringReader sReader = new StringReader(rspBody);
                     DataSet dSet = new DataSet(); 
                     dSet.ReadXml(sReader);

                     IEnumerable<DataRow> iRows = from iRow in dSet.Tables[3].AsEnumerable()
                                                  select iRow;

                    foreach (var iRow in iRows)
                    {
                        DTE DTEl = new DTE();

                        DTEl.empresa = sociedad[2];
                        DTEl.tipoDTE = iRow[3].ToString();
                        DTEl.folioDTE = Convert.ToInt32(iRow[4]);
                        DTEl.rutReceptor = iRow[5].ToString();
                        DTEl.razonSocialReceptor = iRow[6].ToString();
                        //DTEl.Glosa = iRow[16].ToString();
                        //DTEl.Codigo = iRow[5].ToString();
                        DTEl.Neto = Convert.ToInt32(iRow[11]);
                        DTEl.Iva = Convert.ToInt32(iRow[12]);
                        DTEl.Total = Convert.ToInt32(iRow[14]);
                        //DTEl.fechaFacturacion = iRow[9].ToString();

                        DTEl.fechaEmision = iRow[8].ToString();
                        DTEl.fechaFirma = iRow[9].ToString();
                        DTEl.fechaRegistroSII = iRow[10].ToString();
                        DTEl.estadoSII = iRow[15].ToString();
                        DTEl.estadoRecepcion = iRow[17].ToString();
                        DTEl.estadoLey19983 = iRow[19].ToString();
                        DTEl.estadoLey20956 = iRow[21].ToString();

                        //DTEl.fechaPortalCen = Convert.ToDateTime(iRow[5]);
                        //DTEl.fechaCarga = Convert.ToDateTime(iRow[5]);

                        DTE_WSFacele.Add(DTEl);

                    }

                    }
                    catch (Exception ex)
                    {
                        oLog.Add("ERROR",
                                String.Format("Error al cargar página (offset) {0}",
                                i.ToString(), ex.Message));
                        //throw new Exception(ex.Message);
                    }
                }

                oLog.Add("TRACE", String.Format("Llamada web service {0} Exitosa", sociedad[2]));

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error al obtener DTE's en {0} para periodo {1} {2}",
                    sociedad[2],Params.PeriodoEmision, ex.Message));
                //throw new Exception(ex.Message);

            }


        }

        static async Task<string> CallEndPointAsync(string sociedad, int offset)
        {
            WebClient client = new WebClient();
            string request;
            request = "<soapenv:Envelope xmlns:soapenv='http://schemas.xmlsoap.org/soap/envelope/' xmlns:dte='http://facele.cl/docele/servicios/DTE/'>";
            request += "   <soapenv:Header/>";
            request += "   <soapenv:Body>";
            request += "      <dte:consultar>";
            request += "         <rutAbonado>{0}</rutAbonado>";
            request += "         <userMail>admin@facele.cl</userMail>";
            request += "         <periodoFirma>{1}</periodoFirma>";
            request += "         <operacion>EMISION</operacion>";
            request += "         <offset>{2}</offset>";
            request += "      </dte:consultar>";
            request += "   </soapenv:Body>";
            request += "</soapenv:Envelope>";

            string requestIni = String.Format(request, sociedad, Params.PeriodoEmision, offset);

            client.Headers.Add(HttpRequestHeader.ContentType, "text/xml");
            client.Headers.Add("SOAPAction", "http://facele.cl/docele/servicios/DTE/consultar");

            // TODO: Eliminar evento, uso de método UploadStringTaskAsync
            //client.UploadStringCompleted += new UploadStringCompletedEventHandler(client_UploadStringCompleted);
            //client.UploadStringAsync(new Uri("http://10.38.21.105:8090/DoceleOL/DTEService"), request);
            string rspBody = await client.UploadStringTaskAsync(new Uri("http://10.38.21.105:8090/DoceleOL/DTEService"), requestIni);
            return rspBody;

        }

        static void  sendMail()
        {

            try
            {
                MailMessage Mensaje = new MailMessage();
                Mensaje.To.Add(new MailAddress(Params.EmailTO));
                Mensaje.To.Add(new MailAddress(Params.EmailTO2));
                Mensaje.From = new MailAddress(Params.EmailUser);
            
                Mensaje.IsBodyHtml = true;

                Mensaje.Subject = String.Format("Documentos Sigge sin DTE {0}", Params.PeriodoEmision);


                SmtpClient smtp = new SmtpClient();
                NetworkCredential credencial = new NetworkCredential()
                {
                    UserName = Params.EmailUser,
                    Password = Params.EmailPassword,
                };

                smtp.UseDefaultCredentials = false;
                smtp.Credentials = credencial;
                smtp.Host = Params.SMTPName;
                smtp.Port = Params.SMTPPort;
                smtp.EnableSsl = Params.EnableSSL;

                string msgMail = "<p>La siguiente tabla corresponde a los documentos autorizados para su facturación pero aparecen sin DTE informado en los registros SIGGE.  Para mayor información ver en pagina <a href='http://www.sigge.cl'>Sigge.cl</a></p>";
                string cabeceraTabla = @"
                      <table style='border-collapse:collapse;border-spacing:0' class='tg'>
                        <thead>
                          <tr>
                            <th style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal' colspan='8'>SIGGE</th>
                            <th style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;font-weight:normal;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal' colspan='5'>FACELE</th>
                          </tr>
                        </thead>

                        <tbody>
                          <tr>
                            <!-- <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Sociedad</td> --> 
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Rut<br>Cliente</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Nombre</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Neto</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Iva</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Total</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Glosa</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Codigo</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Tipo<br>DTE</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Folio<br>DTE</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Fecha<br>Emision</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Estado<br>SII</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Estado<br>Recepcion</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Fecha<br>Firma</td></tr>
                        <tr>";

                string rowTabla = @"
                         <tr>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{0}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{1}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{2}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{3}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{4}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{5}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{6}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{7}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{8}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{9}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{10}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{11}</td>
                              <td style='border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:center;vertical-align:top;word-break:normal'>{12}</td>
                              
                        </tr>";

                string finTabla = @"</tbody></table>";
                string xBody = "</br></br>";

                foreach (string[] xSoc in Params.RutSociedades)
                {

                    var sDTE = DTE_Segge.Where(p => p.empresa == xSoc[2]);

                    if (sDTE.Count() > 0)
                    {
                        xBody += String.Format("<h3>{0}</h3>", xSoc[1]);

                        string cuerpoTabla = "";
                        foreach (var iList in sDTE)
                        {
                            //iList.empresa,
                            cuerpoTabla += String.Format(rowTabla,
                                iList.rutReceptor, iList.razonSocialReceptor,
                                iList.Neto, iList.Iva,
                                iList.Total,
                                iList.Glosa, iList.Codigo,
                                iList.tipoDTE,
                                iList.folioDTE, iList.fechaEmision, iList.estadoSII,
                                iList.estadoRecepcion, iList.fechaFirma
                                );
                        }

                        xBody += cabeceraTabla + cuerpoTabla + finTabla;
                    }
                }

                string footerTabla = @"
                      <td style='color: #ffffff; font-family: Arial, sans-serif; font-size: 10px;'>
                        &reg; Powered by Codevsys 2020 para ArcoEnergy <br/>
                       </td>";

                Mensaje.Body = msgMail + "</br></br>" + xBody +"</br></br>" + footerTabla;

                smtp.Send(Mensaje);

                oLog.Add("INFO", String.Format("Email enviado con {0} registros informados", DTE_Segge.Count().ToString()));

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error al enviar EMail {0}", ex.Message));

            }


        }

        // TODO: Eliminar evento, uso de método UploadStringTaskAsync
        static void client_UploadStringCompleted(object sender, UploadStringCompletedEventArgs e)
        {
            Console.WriteLine("UploadStringCompleted: {0}", e.Result);

        }

        static string isDbNull(DataRow row, int indice)
        {
            if (!DBNull.Value.Equals(row[indice]))
                return (string)row[indice].ToString().Trim();
            else
                return ""; // String.Empty;
        }


        // TODO: Instanciar propiedades en constructor
        static List<DTE> DTE_WSFacele = new List<DTE>();
        static List<DTE> DTE_Segge = new List<DTE>();
        static Settings Params = new Settings();
        static DTE_log oLog = new DTE_log(AppDomain.CurrentDomain.BaseDirectory);
    }

    public class Settings
    {
        public List<String[]> RutSociedades { get; set; }
        public String SMTPName { get; set; }
        public int SMTPPort { get; set; }
        public bool EnableSSL { get; set; }
        public String EmailUser { get; set; }
        public String EmailPassword { get; set; }
        public String EmailTO { get; set; }
        public String EmailTO2 { get; set; }
        public String DirectorioExcelSigge { get; set; }
        public String PeriodoEmision { get; set; }
    }

    public class DTE
    {
        // Sigge
        public string Glosa { get; set; }
        public string Codigo { get; set; }
        public string fechaFacturacion { get; set; }  
        public string fechaPortalCen { get; set; }
        public string fechaCarga { get; set; }

        // Comparte Sigge & Facele
        public string empresa { get; set; }
        public string tipoDTE { get; set; }  
        public string rutReceptor { get; set; }  
        public string razonSocialReceptor { get; set; } 
        public int Neto { get; set; }  
        public int Iva { get; set; }  
        public int Total { get; set; }  

        // Facele
        public int folioDTE { get; set; }  
        public string fechaEmision { get; set; }
        public string fechaFirma { get; set; }
        public string fechaRegistroSII { get; set; }
        public string estadoSII { get; set; }
        public string estadoRecepcion { get; set; }
        public string estadoLey19983 { get; set; }
        public string estadoLey20956 { get; set; }
    }

}

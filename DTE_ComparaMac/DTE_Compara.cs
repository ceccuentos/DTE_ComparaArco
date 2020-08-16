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
using System.Text.RegularExpressions;
using System.Security;
using System.Threading;

namespace DTE_ComparaArco
{
    public class DTE_Compara
    {

        public DTE_Compara()
        {
            
            // Get Params
            try
            {
                Params.RutSociedades = new List<string[]>();

                var filenameXMLSettings = "ArcoDTE_ComparaConfig.xml";
                var currentDirectory = AppDomain.CurrentDomain.BaseDirectory; 
                var settingsXMLFilepath = Path.Combine(currentDirectory, filenameXMLSettings);

                oLog = new DTE_log(currentDirectory);
                oLog.Add("DEBUG", "======== Inicio Proceso ========");
                oLog.Add("TRACE", "Get Settings: " + settingsXMLFilepath);

                if (!File.Exists(settingsXMLFilepath))
                    {
                    throw new System.ArgumentException("No existe archivo de configuración", settingsXMLFilepath);
                    }

                XElement Properties_Settings = XElement.Load(settingsXMLFilepath);
                IEnumerable<XElement> nodeSetts = from parametro in Properties_Settings.Descendants("ArcoDTE_Compara.Properties.Settings")
                                              select (XElement)parametro;
                

                foreach (XElement elemento in nodeSetts.Elements())
                {
                    // Tratamiento especial array nodos lista de sociedades
                    if (elemento.Attribute("name").Value == "RutSociedades")
                    {
                        XElement nodoSoc = XElement.Parse(elemento.FirstNode.ToString());

                        IEnumerable<XElement> elementoSociedad = from xel in nodoSoc.Descendants("ArrayOfString")
                                                     select (XElement)xel;
                        foreach (XElement elxnt in elementoSociedad.Elements())
                        {
                            Params.RutSociedades.Add(elxnt.Value.Split(';'));
                        }
                    }

                    // Otros Nodos
                    Params.PeriodoEmision = DateTime.Now.AddDays(-1).ToString("yyyy-MM");
                    switch (elemento.Attribute("name").Value)
                    {
                        case "URIWEBService":
                            Params.URIWEBService = elemento.Value;
                            break;
                        case "URISOAPAction":
                            Params.URISOAPAction = elemento.Value;
                            break;
                        case "SMTPName":
                            Params.SMTPName = elemento.Value;
                            break;
                        case "SMTPPort":
                            Params.SMTPPort = Convert.ToInt32(elemento.Value);
                            break;
                        case "EnableSSL":
                            Params.EnableSSL = Convert.ToBoolean(elemento.Value);
                            break;
                        case "EmailUser":
                            Params.EmailUser = elemento.Value;
                            break;
                        case "EmailPassword":
                            Params.EmailPassword = elemento.Value;
                            break;
                        case "EmailTO":
                            Params.EmailTO = elemento.Value;
                            break;
                        case "EmailTO2":
                            Params.EmailTO2 = elemento.Value;
                            break;
                        case "DirectorioExcelSigge":
                            Params.DirectorioExcelSigge = elemento.Value;
                            break;

                    }
                  
                }

                if (Params.RutSociedades.Count == 0)
                {
                    oLog.Add("ERROR", "Sociedades no encontradas, revise estructura XML");
                }
                else
                {
                    oLog.Add("TRACE", "Get Settings Successed");

                }
            }
            catch (Exception ex)
            {
                oLog.Add("ERROR", ex.Message);
                throw new Exception(ex.Message);
            }
        }

        static async Task Main(string[] args)
        {
            DTE_Compara DcP = new DTE_Compara();

            try
            {
                string paternSociedades = string.Join("|",
                                                     Params.RutSociedades.Select(elemento => elemento[2]));
                //Fmto: <alias>__AAAA-MM.xlsx 
                Regex exReg = new Regex(@"^(" + paternSociedades + ")__[0-9]{4}-0[1-9]|1[0-2].(xls|xlsx)$",
                      RegexOptions.Compiled | RegexOptions.IgnoreCase);

                string dirSigge = Params.DirectorioExcelSigge != "" ? Params.DirectorioExcelSigge : AppDomain.CurrentDomain.BaseDirectory;
                string dirSiggeProcessed = Path.Combine(dirSigge, "processedfiles");
                string[] filesExcelinDirSigge = Directory.GetFiles(@dirSigge, "*.xls?");

                oLog.Add("TRACE", String.Format("{0} archivos Excel encontrados... ", filesExcelinDirSigge.Count()));

                foreach (string file in filesExcelinDirSigge)
                {
                    
                    string onlyfileName = file.Substring(dirSigge.Length + 1).Replace(".xlsx", "").Replace(".xls","");

                    if (exReg.IsMatch(onlyfileName))
                    {
                        string[] ArrayPartFile = Regex.Split(onlyfileName, "__|-");
                        string[][] Getsociedad = Params.RutSociedades.Where(elemento => elemento[2] == ArrayPartFile[0]).ToArray();

                        
                        string Periodo = String.Format("{0}-{1}", ArrayPartFile[1], ArrayPartFile[2]);

                        oLog.Add("TRACE",
                                String.Format("Planilla Sigge encontrada: Sociedad {0}, Periodo {1}, archivo {2}",
                                ArrayPartFile[0], Periodo, onlyfileName));

                            string PeriodoPost = String.Format("{0}-{1}",
                                 ArrayPartFile[2].ToString() == "12"
                                     ? (Convert.ToInt32(ArrayPartFile[1]) + 1).ToString()
                                     : ArrayPartFile[1].ToString(),
                                 ArrayPartFile[2].ToString() == "12"
                                     ? "01"
                                     : Right("0" + (Convert.ToInt32(ArrayPartFile[2]) + 1).ToString(), 2)
                                   );

                            string PeriodoAnt = String.Format("{0}-{1}",
                                     ArrayPartFile[2].ToString() == "01"
                                         ? (Convert.ToInt32(ArrayPartFile[1]) - 1).ToString()
                                         : ArrayPartFile[1].ToString(),
                                     ArrayPartFile[2].ToString() == "01"
                                         ? "12"
                                         : Right("0" + (Convert.ToInt32(ArrayPartFile[2]) - 1).ToString(), 2)
                                     ) ;

                        await callWSFacele(Getsociedad[0], PeriodoAnt);
                        await callWSFacele(Getsociedad[0], Periodo);
                        await callWSFacele(Getsociedad[0], PeriodoPost);

                        // Carga Excel's
                        loadWorkbook(Getsociedad[0], file);

                        // Mover archivos
                        CreateDirectory(dirSiggeProcessed);

                        oLog.Add("TRACE",
                                String.Format("Terminado: Moviendo archivo {0} a directorio procesados",
                                file.Substring(dirSigge.Length + 1)));

                        File.Move(file, Path.Combine(dirSiggeProcessed, file.Substring(dirSigge.Length + 1)), true);

                    }
                    else
                    {
                        oLog.Add("INFO",
                                String.Format("Archivo Excel {0} fue ignorado (no clumple reglas)",
                                file.Substring(dirSigge.Length + 1)));
                    }

                }

                // Envía correo si hay diferencias
                if (DTE_Segge.Count() != 0)
                    sendMail();

                oLog.Add("TRACE", "======== Fin Proceso ========");
            }
            catch (Exception ex)
            {
                oLog.Add("ERROR", ex.Message);
            }
            finally
            {
                Thread.Sleep(3000);
            }
        }

        static void loadWorkbook(string[] sociedad, string filePath)
        {
            oLog.Add("TRACE", String.Format("Leyendo Excel {0}", filePath));

            try { 
                System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
                using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
                {
                    using (var readerExcel = ExcelReaderFactory.CreateReader(stream))
                    {
                        var FilasExcel = readerExcel.AsDataSet(new ExcelDataSetConfiguration
                        {
                            ConfigureDataTable = _ => new ExcelDataTableConfiguration
                            {
                                UseHeaderRow = true 
                            }
                        });

                        // Folio DTE Posición 7 
                        IEnumerable<DataRow> ListaFilasSinDTE = from fila in FilasExcel.Tables[0].AsEnumerable()
                                                       where DBNull.Value.Equals(fila[7])
                                                       select fila;

                        int FilasSinDTE = 0;  
                        int FilasSinDTEenFacele = 0;
                        foreach (var iRow in ListaFilasSinDTE)
                        {
                            DTE DTEl = new DTE();

                            DTEl.empresa = sociedad[2];
                            DTEl.tipoDTE = isDbNull(iRow, 8); 
                            DTEl.rutReceptor = iRow[1].ToString();
                            DTEl.razonSocialReceptor = iRow[2].ToString();
                            DTEl.PeriodoFacturacion = iRow[3].ToString();
                            DTEl.Glosa = iRow[5].ToString();
                            DTEl.Codigo = iRow[6].ToString();
                            DTEl.Neto = Int32.Parse(isDbNull(iRow, 9));
                            DTEl.Iva = Int32.Parse(isDbNull(iRow, 10));
                            DTEl.Total = Int32.Parse(isDbNull(iRow, 11));
                            DTEl.fechaFacturacion = isDbNull(iRow, 12);
                            DTEl.fechaPortalCen = isDbNull(iRow, 15);
                            DTEl.fechaCarga = isDbNull(iRow, 17);


                            // Busqueda con LINQ Lista IEnumerator por si se repiten facturas iguales (mismo criterio)
                            List<DTE> ListaDTEinFacele = DTE_WSFacele
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

                            foreach (var xRow in ListaDTEinFacele)  //IEnumerator evita crash si se repite docto
                            {
                                    DTEl.folioDTE = xRow.folioDTE;

                                    DTEl.fechaEmision = xRow.fechaEmision;
                                    DTEl.fechaFirma = xRow.fechaFirma;
                                    DTEl.fechaRegistroSII = xRow.fechaRegistroSII;
                                    DTEl.estadoSII = xRow.estadoSII;
                                    DTEl.estadoRecepcion = xRow.estadoRecepcion;
                                    DTEl.estadoLey19983 = xRow.estadoLey19983;
                                    DTEl.estadoLey20956 = xRow.estadoLey20956;
                                    FilasSinDTEenFacele++;
                            }

                            // Agrega datos a Lista
                            DTE_Segge.Add(DTEl);
                            FilasSinDTE++;


                        }
                        
                        oLog.Add("INFO", String.Format("{0} Registros DTE obtenidos.  {1} Sin DTE, {2} encontrados en WS Facele",
                            FilasExcel.Tables[0].Rows.Count.ToString(), FilasSinDTE.ToString(), FilasSinDTEenFacele.ToString()));
                    }
                }

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error leer archivo Excel Sigge {0} para periodo {1} {2}",
                    filePath, Params.PeriodoEmision, ex.Message));

            }
        }

        static async Task callWSFacele(string[] sociedad, string Periodo)
        {
            oLog.Add("TRACE", String.Format("Llamando Web service {0} Periodo {1}", sociedad[2], Periodo));

            try
            {
                // Cnt Registros
                string xmlStringRespuesta = await CallEndPointAsync( sociedad[0], 999999, Periodo);

                StringReader XMLReader = new StringReader(xmlStringRespuesta);
                DataSet DataSetFromXML = new DataSet();
                DataSetFromXML.ReadXml(XMLReader);

                IEnumerable<DataRow> nodeXmls = from fila in DataSetFromXML.Tables["consultarResponse"].AsEnumerable()
                                              select fila;

                int cntReg = 0;
                foreach (var ieXml in nodeXmls)
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
                     xmlStringRespuesta = await CallEndPointAsync(sociedad[0], i*100, Periodo);
                     StringReader XMLReaderOffset = new StringReader(xmlStringRespuesta);
                     DataSet DataSetFromXMLOffset = new DataSet();
                     DataSetFromXMLOffset.ReadXml(XMLReaderOffset);

                     IEnumerable<DataRow> nodeXmlsOffset = from iRow in DataSetFromXMLOffset.Tables[3].AsEnumerable()
                                                  select iRow;

                    foreach (var iRow in nodeXmlsOffset)
                        {
                            DTE DTEl = new DTE();

                            DTEl.empresa = sociedad[2];
                            DTEl.tipoDTE = iRow[3].ToString();
                            DTEl.folioDTE = Convert.ToInt32(iRow[4]);
                            DTEl.rutReceptor = iRow[5].ToString();
                            DTEl.razonSocialReceptor = iRow[6].ToString();
                            DTEl.Neto = Convert.ToInt32(iRow[11]);
                            DTEl.Iva = Convert.ToInt32(iRow[12]);
                            DTEl.Total = Convert.ToInt32(iRow[14]);

                            DTEl.fechaEmision = iRow[8].ToString();
                            DTEl.fechaFirma = iRow[9].ToString();
                            DTEl.fechaRegistroSII = iRow[10].ToString();
                            DTEl.estadoSII = iRow[15].ToString();
                            DTEl.estadoRecepcion = iRow[17].ToString();
                            DTEl.estadoLey19983 = iRow[19].ToString();
                            DTEl.estadoLey20956 = iRow[21].ToString();

                            DTE_WSFacele.Add(DTEl);

                        }

                    }
                    catch (Exception ex)
                    {
                        oLog.Add("ERROR",
                                String.Format("Error al cargar página (offset) {0}",
                                i.ToString(), ex.Message));
                    }
                }

                oLog.Add("TRACE", String.Format("Llamada web service {0} Exitosa Periodo {1}", sociedad[2], Periodo));

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error al obtener DTE's en {0} para periodo {1} {2}",
                    sociedad[2],Params.PeriodoEmision, ex.Message));
            }

        }

        static async Task<string> CallEndPointAsync(string sociedad, int offset, string Periodo)
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

            string requestIni = String.Format(request, sociedad, Periodo, offset);

            client.Headers.Add(HttpRequestHeader.ContentType, "text/xml");
            client.Headers.Add("SOAPAction", Params.URISOAPAction);

            string rspBody = await client.UploadStringTaskAsync(new Uri(Params.URIWEBService), requestIni);

            return rspBody;

        }

        static void  sendMail(string BodyFileHtml_IF_EMAIL_FAILED="")
        {
            
            try
            {
                MailMessage Mensaje = new MailMessage();
                Mensaje.To.Add(new MailAddress(Params.EmailTO));
                Mensaje.To.Add(new MailAddress(Params.EmailTO2));
                Mensaje.From = new MailAddress(Params.EmailUser);
            
                Mensaje.IsBodyHtml = true;

                Mensaje.Subject = String.Format("Documentos Sigge sin DTE {0}", "");

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
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Rut<br>Cliente</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Nombre</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Neto</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Iva</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Total</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Glosa</td>
                            <td style='background-color:#c0c0c0;border-color:black;border-style:solid;border-width:1px;font-family:Arial, sans-serif;font-size:14px;overflow:hidden;padding:10px 5px;text-align:left;vertical-align:top;word-break:normal'>Periodo</td>
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
                            cuerpoTabla += String.Format(rowTabla,
                                iList.rutReceptor, iList.razonSocialReceptor,
                                iList.Neto, iList.Iva,
                                iList.Total,
                                iList.Glosa, iList.PeriodoFacturacion,
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
                BodyFileHtml_IF_EMAIL_FAILED = Mensaje.Body;

                smtp.Send(Mensaje);

                oLog.Add("INFO", String.Format("Email enviado con {0} registros informados", DTE_Segge.Count().ToString()));

            }
            catch (Exception ex)
            {
                oLog.Add("ERROR",
                    String.Format("Error al enviar EMail {0}", ex.Message));

                if
                      (
                          (ex is ArgumentException
                          || ex is ArgumentNullException
                          || ex is InvalidOperationException
                          || ex is ObjectDisposedException
                          || ex is SmtpException
                          || ex is SmtpFailedRecipientException
                          || ex is SmtpFailedRecipientsException)
                          && BodyFileHtml_IF_EMAIL_FAILED != ""

                      )
                {
                    
                    string nombre = "DTEComparaHtml_" + DateTime.Now.Year + "_" + DateTime.Now.Month + "_" + DateTime.Now.Day + DateTime.Now.Hour + DateTime.Now.Minute + ".html";
                    StreamWriter sw = new StreamWriter(Path.Combine(Params.DirectorioExcelSigge, nombre), true);
                    sw.Write(BodyFileHtml_IF_EMAIL_FAILED);
                    sw.Close();
    
                    oLog.Add("INFO",
                        String.Format("Error al enviar EMail, genera resultado en archivo {0}.", nombre));
                }


            }


        }

        static string isDbNull(DataRow row, int indice)
        {
            if (!DBNull.Value.Equals(row[indice]))
                return (string)row[indice].ToString().Trim();
            else
                return ""; 
        }

        static string Right( string value, int length)
        {
            if (String.IsNullOrEmpty(value)) return string.Empty;

            return value.Length <= length ? value : value.Substring(value.Length - length);
        }

        static void CreateDirectory(string Ruta)
        {
            try
            {
                if (!Directory.Exists(Ruta))
                    Directory.CreateDirectory(Ruta);
            }
            catch (Exception ex)
            {
                if
                  (
                      ex is UnauthorizedAccessException
                      || ex is ArgumentNullException
                      || ex is PathTooLongException
                      || ex is DirectoryNotFoundException
                      || ex is NotSupportedException
                      || ex is ArgumentException
                      || ex is SecurityException
                      || ex is IOException
                  )
                {
                    throw new Exception(ex.Message);
                }
            }
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
        public String URIWEBService { get; set; }
        public String URISOAPAction { get; set; }
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
        public string PeriodoFacturacion { get; set; }
        

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

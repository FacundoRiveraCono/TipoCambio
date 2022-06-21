using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using RestSharp;
using Newtonsoft.Json;
using HtmlAgilityPack;
using ScrapySharp.Extensions;

namespace APITipoCambio
{
    class Program
    {
        static void Main(string[] args)
        {
            //Defino Listado de Sociedades a insertar los tipos de cambios para recorrer listado.
            //De esta forma no repito un metodo de conexion por cada sociedad, solo uno variable

            //Comienzan metodos principales

            WebScraping();
            WebScrapingSwiss();
            
            Console.ReadLine();
            
            

        }

        public static void WebScraping()
        {
            //Listado de objetos de divisa.
            List<TipoCambio> Listado = new List<TipoCambio>();
            double UsdDiCompra = 0;
            double EuroDiCompra = 0;
            int contador = 0;
            HtmlWeb oWeb = new HtmlWeb();
            //List<string> Cotizaciones = new List<string>();
            HtmlDocument doc = oWeb.Load("https://www.bna.com.ar/Personas");
            foreach (var Nodo in doc.DocumentNode.CssSelect("#divisas"))
            {
                contador++;
                var NodeAnchor = Nodo.CssSelect("td").ToList();
              
                foreach (var subnodo in NodeAnchor)
                {
                    //Objeto a completar con datos de divisa y luego cargar a la lista
                    TipoCambio add = new TipoCambio();
                    
                    if (subnodo.InnerText == "Dolar U.S.A")
                    {

                        UsdDiCompra = Convert.ToDouble(NodeAnchor[2].InnerText);
                        
                        add.Divisa = "USD";
                        add.Valor = UsdDiCompra;
                        add.fecha = DateTime.Now.AddDays(1);
                        //Console.WriteLine(add.Valor);
                        Listado.Add(add);                      
                    }
                    if (subnodo.InnerText == "Euro")
                    {
                        EuroDiCompra = Convert.ToDouble(NodeAnchor[8].InnerText);
                        add.Divisa = "EUR";
                        add.Valor = EuroDiCompra;
                        add.fecha = DateTime.Now.AddDays(1);

                        Listado.Add(add);                                           
                    }

                }          

            }

            //Si el valor sigue siendo cero (valor seteo) es porque no encontro valor
            if (UsdDiCompra > 0 && EuroDiCompra > 0)
            {
                Console.WriteLine("Tipos de cambio encontrados");
                InsertSAP(Listado);
                InsertSAPCNT(Listado);
            }
            else
            {
                Console.WriteLine("Valor cotizacion no encontrado");
            }

            
        }

        
        public static void InsertSAP(List<TipoCambio> lista)
        {

            int retVal = -1;
            string msgError = "";
            //SAPbobsCOM.Company oC = new SAPbobsCOM.Company();
            SAPbobsCOM.Company oC = new SAPbobsCOM.Company();

            oC.Server = "Servidor";
            oC.UserName = "SERV-FINANZAS";
            oC.Password = "Contreña";
            oC.LicenseServer = "Servidor";
            oC.DbUserName = "SYSTEM";
            oC.DbPassword = "Contraseña";
            oC.CompanyDB = "BASE_PROD";
            oC.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

            Console.WriteLine("Ingresando a SAP-SA");
            retVal = oC.Connect();

            if (retVal==0)
            {
                Console.WriteLine("Conexion a SAP exitosa");
                SAPbobsCOM.SBObob objeto;
                objeto = (SAPbobsCOM.SBObob)oC.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                foreach (var item in lista)
                {
                    objeto.SetCurrencyRate(item.Divisa, item.fecha, item.Valor, true);
                }

                Console.WriteLine("Tipos de cambio insertado");     
                oC.Disconnect();
                Console.WriteLine("Conexion SAP cerrada");
            }

        }

        public static void InsertSAPCNT(List<TipoCambio> lista)
        {

            int retVal = -1;
            string msgError = "";
            //SAPbobsCOM.Company oC = new SAPbobsCOM.Company();
            SAPbobsCOM.Company oC = new SAPbobsCOM.Company();

            oC.Server = "Servidor";
            oC.UserName = "SERV-FINANZAS";
            oC.Password = "Contraseña";
            oC.LicenseServer = "Contraseña";
            oC.DbUserName = "SYSTEM";
            oC.DbPassword = "Contraseña";
            oC.CompanyDB = "CONO_TSAS";
            oC.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

            Console.WriteLine("Ingresando a SAP-Trading");
            retVal = oC.Connect();

            if (retVal == 0)
            {
                Console.WriteLine("Conexion a SAP exitosa");
                SAPbobsCOM.SBObob objeto;
                objeto = (SAPbobsCOM.SBObob)oC.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                foreach (var item in lista)
                {
                    objeto.SetCurrencyRate(item.Divisa, item.fecha, item.Valor, true);
                }

                Console.WriteLine("Tipos de cambio insertado");
                oC.Disconnect();
                Console.WriteLine("Conexion SAP cerrada");
            }

        }



        //Bases de datos Suizas

        public static void WebScrapingSwiss()
        {
            //Listado de objetos de divisa.
            List<TipoCambio> Listado = new List<TipoCambio>();
            double UsdDiVenta = 0;
            double EuroDiVenta = 0;
            double GbpDiVenta = 0;
            int contador = 0;
            HtmlWeb oWeb = new HtmlWeb();
            //List<string> Cotizaciones = new List<string>();
            HtmlDocument doc = oWeb.Load("https://www.snb.ch/en/iabout/stat/statrep/id/current_interest_exchange_rates#t3");

            foreach (var item in doc.DocumentNode.CssSelect(".rates-values.exchangerates"))
            {

                var nodeAnchor = item.CssSelect("li").ToList();

                foreach (var subnodo in nodeAnchor)
                {
                    TipoCambio add = new TipoCambio();
                    //if (subnodo.InnerHtml==)
                    if (subnodo.InnerHtml.Contains("USD"))
                    {
                        UsdDiVenta = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                        add.Valor = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                       // Console.WriteLine(add.Valor);
                        add.fecha = DateTime.Now.AddDays(1);
                        add.Divisa = "USD";
                        Listado.Add(add);
                    }

                    if (subnodo.InnerHtml.Contains("EUR"))
                    {
                        EuroDiVenta = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                        add.Valor = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                        //Console.WriteLine(add.Valor);
                        add.fecha = DateTime.Now.AddDays(1);
                        add.Divisa = "EUR";
                        Listado.Add(add);

                    }

                    if (subnodo.InnerHtml.Contains("GBP"))
                    {
                        GbpDiVenta = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                        add.Valor = Convert.ToDouble(subnodo.CssSelect(".value").First().InnerText);
                        add.fecha = DateTime.Now.AddDays(1);
                        add.Divisa = "GBP";
                        Listado.Add(add);
                    }

                    

                }
                

            }

           if (UsdDiVenta > 0 && EuroDiVenta > 0 && GbpDiVenta > 0)
            {
                Console.WriteLine("Tipos de cambio encontrados");
                InsertSAPCNTI(Listado);
                InsertSAPCNG(Listado);

            }
            else
            {
                Console.WriteLine("Tipos de cambio no encontrados");
            }
            
            
                  
        }


        public static void InsertSAPCNTI(List<TipoCambio> lista)
        {

            int retVal = -1;
            string msgError = "";
            //SAPbobsCOM.Company oC = new SAPbobsCOM.Company();
            SAPbobsCOM.Company oC = new SAPbobsCOM.Company();

            oC.Server = "Servidor";
            oC.UserName = "SERV-FINANZAS";
            oC.Password = "Contraseña";
            oC.LicenseServer = "Servidor";
            oC.DbUserName = "SYSTEM";
            oC.DbPassword = "Contraseña";
            oC.CompanyDB = "CONO_TINT";
            oC.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

            Console.WriteLine("Ingresando a SAP-Trading International");
            retVal = oC.Connect();

            if (retVal == 0)
            {
                Console.WriteLine("Conexion a SAP exitosa");
                SAPbobsCOM.SBObob objeto;
                objeto = (SAPbobsCOM.SBObob)oC.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                foreach (var item in lista)
                {
                    objeto.SetCurrencyRate(item.Divisa, item.fecha, item.Valor, true);
                }

                Console.WriteLine("Tipos de cambio insertado");
                oC.Disconnect();
                Console.WriteLine("Conexion SAP cerrada");
            }

        }

        public static void InsertSAPCNG(List<TipoCambio> lista)
        {

            int retVal = -1;
            string msgError = "";
            //SAPbobsCOM.Company oC = new SAPbobsCOM.Company();
            SAPbobsCOM.Company oC = new SAPbobsCOM.Company();

            oC.Server = "Servidor";
            oC.UserName = "SERV-FINANZAS";
            oC.Password = "Contraseña";
            oC.LicenseServer = "Servidor";
            oC.DbUserName = "SYSTEM";
            oC.DbPassword = "Contraseña";
            oC.CompanyDB = "CONO_GROUP";
            oC.DbServerType = SAPbobsCOM.BoDataServerTypes.dst_HANADB;

            Console.WriteLine("Ingresando a SAP-Group");
            retVal = oC.Connect();

            if (retVal == 0)
            {
                Console.WriteLine("Conexion a SAP exitosa");
                SAPbobsCOM.SBObob objeto;
                objeto = (SAPbobsCOM.SBObob)oC.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoBridge);
                foreach (var item in lista)
                {
                    objeto.SetCurrencyRate(item.Divisa, item.fecha, item.Valor, true);
                }

                Console.WriteLine("Tipos de cambio insertado");
                oC.Disconnect();
                Console.WriteLine("Conexion SAP cerrada");
            }
            Environment.Exit(0);

        }


        //public static void ObtenerPeso()
        //{
        //    string peso = "";
        //    string[] Caracter = new string[] { "[", "]" };

        //    HtmlWeb oWeb = new HtmlWeb();
        //    HtmlDocument oDoc = oWeb.Load("http://192.168.1.70/");

        //    foreach (var item in oDoc.DocumentNode.CssSelect("#Peso"))
        //    {
        //        peso = item.InnerText;
        //        break;
        //    }

        //    //For para eliminar caracteres
        //    foreach (var item in Caracter)
        //    {
        //        peso = peso.Replace(item, string.Empty);
        //    }

        //    Console.WriteLine(peso);

        //}

    }
}

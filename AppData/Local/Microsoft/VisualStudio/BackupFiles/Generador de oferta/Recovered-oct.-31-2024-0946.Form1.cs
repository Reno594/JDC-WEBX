using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
using System.Runtime.Serialization.Formatters;
using System.Runtime.Serialization.Formatters.Binary;
using System.Runtime.Serialization;
using System.IO;
using System.Threading;
using System.Diagnostics;
using resourceLib;
using System.Runtime.InteropServices;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Reflection;
using System.Globalization;
using System.Drawing.Printing;
using System.Text.RegularExpressions;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;
using System.ComponentModel.Design;
using org.bouncycastle.asn1.teletrust;
using System.Data.SqlClient;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using Microsoft.Vbe.Interop;
using System.Runtime.CompilerServices;
using org.bouncycastle.crypto.engines;

namespace Generador_de_oferta
{
    
    public partial class Form1 : Form
    {
        /// <summary>
        /// Contador que me lleva la constancia de las imagenes
        /// </summary>
        public int pk = 0;
        UserControl1 myUserControl;

        bool ObtenerBool = false;
        CExcelWork TrabExcel = null;
        Thread Obtener = null;

        bool ini = false;

        public bool ProCamara = false;

        public Process[] myProcesses;
        public COferta myOferta = null;
        string actualcam = "";
        bool saved = false;
        string ruta = "";
        public bool generado = false;
        string local;

        public bool Procc = false;
       
        //string codValv;
        
        Thread myThread = null;
        Thread OpenThread = null;

        //------Formateador y trama
        IFormatter myFormatter; 
        Stream myFileStream; 
        //-------------------------
        delegate void SetTextCallback(Control Contr, string text);
        
        //----------------------------------------------------
        Microsoft.Office.Interop.Excel.Application NewExcelApp = null;
        
        _Workbook CalcWorkBook = null;
        _Worksheet CalcWorkSheet = null;
        _Workbook SelecWorkBook = null;
        _Worksheet SelecWorkSheet = null;
        _Worksheet PrintOut = null;


        //----------------------------------------------------
        

        public Form1()
        {
            InitializeComponent();
            myUserControl = new UserControl1();
            this.DateActualizer();
            local = Directory.GetCurrentDirectory();
            TNP.Focus();
            myProcesses = Process.GetProcesses();
            Datos.ActualcamChanged += OnActualcamChanged;

        }
        private void OnActualcamChanged()
        {
            Console.WriteLine("OnActualcamChanged called in Form1"); // Mensaje de depuración TPcp
            CargarValores();
        }

        public void CargarValores()
        {
            if (TLtemp == null) { TLtemp = 0; } else { Datos.TLtemp = TLtemp.ToString(); }
            if (TLPbar == null) { TLPbar = 0; } else { Datos.TLPbar = TLPbar.ToString(); }
            if (DTpress == null) { DTpress = 0; } else { Datos.DTpress = DTpress.ToString(); }
            if (TCPbar == null) { TCPbar = 0; } else { Datos.TCPbar = TCPbar.ToString(); }
            if (TEPbar == null) { TEPbar = 0; } else { Datos.TEPbar = TEPbar.ToString(); }


            actualcam = Datos.actualcam;
            TValv.Text = Datos.TValv;
            TTcmc6.Text = Datos.Tcmc6;
            TTcmc8.Text = Datos.Tcmc8;
            TPcy.Text = Datos.Rfcalx;
            Datos.TPex = TPex.Text;
            Datos.TPem = TPem.Text;
            TPnt.Text = string.IsNullOrEmpty(Datos.NEvap) ? "1" : Datos.NEvap; ;
            Datos.TPnc = TPnc.Text;
            Datos.TPrs = TPrs.Text;
            Datos.Modex = TModex.Text;
            TPcp.Text = Datos.TPcp.ToString();
            TCentx.Text = Datos.TCentd.ToString();
            TPcy.Text = Datos.Rfcalx;
            CTEvap.Text = Datos.CTEvap;
            TCodValv.Text = Datos.PEcalc;
            Datos.PValE = TPosc.Text;
            Datos.Hur11 = Convert.ToInt16(hur1);
            Datos.Hur21 = Convert.ToInt16(hur2);
            string rdin = Datos.Rdin1.ToString();
            Datos.TQevp = TQevp.Text;
            Datos.CSerie = TMcc.Text;
            string TRdin = Datos.TRdin;
            string NEvap = string.IsNullOrEmpty(Datos.NEvap) ? "1" : Datos.NEvap;
            string DTpd = Datos.DTpd;
            string DTsl = string.IsNullOrEmpty(Datos.DTsl) ? "AUT" : Datos.DTsl;
            string SCdt = string.IsNullOrEmpty(Datos.SCdt) ? "SC2" : Datos.SCdt;
            string TBevap = Datos.TBevap;
            string TeVP = string.IsNullOrEmpty(Datos.TeVP) ? "AUT" : Datos.TeVP;
            string DTpdE = Datos.DTpdE;
            string PCarga = Datos.PCarga;
            string PValE = Datos.PValE;
            string CvalSet = Datos.CvalSet;
            string CupValv = Datos.CupValv;
            TPem.Text = Datos.NEvap;
            Datos.TPvq = TPvq.Text;
            Datos.PTCond7 = TPls.Text;
            Datos.TPrs7 = TInev.Text;
            Datos.TCsist1 = TCsist.Text;
            Datos.TInev = TInev.Text;
            Datos.TIned = TIned.Text;
            Datos.TIncd = TIncd.Text;
            Datos.TInc = TInc.Text;
            string PfIQ = (Datos.TFWe).ToString();
        }


        decimal TEvape = 0;
        decimal TLtemp = 0; ////  TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString()
        decimal TLPbar = 0;
        decimal DTpress = 0;
        decimal TCPbar = 0;
        decimal TEPbar = 0;

        decimal TEvap = 0;
        decimal Rdin = 0;
        decimal hur1 = 0;
        decimal hur2 = 0;
        decimal TRdin = 0;

        decimal SCdt = 0;

        decimal TBevap = 0;
        decimal DTpdE = 0;

        decimal PCarga = 0;
        decimal CvalSet = 0;
        decimal CupValv = 0;


        public void GuardarValores()
        {
            if (string.IsNullOrEmpty(Datos.TLtemp)) { TLtemp = 0; } else { TLtemp = Convert.ToDecimal(Datos.TLtemp); }
            if (string.IsNullOrEmpty(Datos.TLPbar)) { TLtemp = 0; } else { TLPbar = Convert.ToDecimal(Datos.TLPbar); }
            if (string.IsNullOrEmpty(Datos.DTpress)) { DTpress = 0; } else { DTpress = Convert.ToDecimal(Datos.DTpress); }
            if (string.IsNullOrEmpty(Datos.TCPbar)) { TCPbar = 0; } else { TCPbar = Convert.ToDecimal(Datos.TCPbar); }
            if (string.IsNullOrEmpty(Datos.TEPbar)) { TEPbar = 0; } else { TEPbar = Convert.ToDecimal(Datos.TEPbar); }


            Datos.actualcam = actualcam;
            Datos.TEvape = (TEvape).ToString();
            string hur1 = Datos.Hur11.ToString();
            string hur2 = Datos.Hur21.ToString();
            Datos.Rdin1 = Convert.ToInt32(Convert.ToDecimal(Rdin));
            Datos.TEvap = TEvap.ToString();
            CTEvap.Text = Datos.CTEvap;
            TPcp.Text = (Datos.TPcp).ToString();
            Datos.TRdin = TRdin.ToString();
            string Tcamara = Datos.TCentd.ToString();
            TCodValv.Text = Datos.PEcalc;
            string DTpd = Datos.DTpd;
            string DTsl = Datos.DTsl;
            string TeVP = Datos.TeVP;
            Datos.SCdt = SCdt.ToString(); ;
            TTcmc6.Text = Datos.Tcmc6;
            Datos.TBevap = TBevap.ToString();
            Datos.DTpdE = DTpdE.ToString();
            TModex.Text = Datos.Modex;
            TPcy.Text = Datos.Rfcalx;
            Datos.PCarga = PCarga.ToString();
            TPosc.Text = Datos.PValE;
            Datos.CvalSet = CvalSet.ToString();
            Datos.CupValv = CupValv.ToString();
            TPex.Text = Datos.TPex;
            TPrs.Text = Datos.TPrs;
            TPem.Text = Datos.TPem;
            TPnt.Text = string.IsNullOrEmpty(Datos.NEvap) ? "1" : Datos.NEvap; ;
            TPnc.Text = Datos.TPnc;
            TValv.Text = Datos.TValv;
            TTcmc8.Text = Datos.Tcmc8;
            TQevp.Text = Datos.TQevp;
            TMcc.Text = Datos.CSerie;
            TPem.Text = Datos.NEvap;
            string TCond = Datos.TConD;
            TPvq.Text = Datos.TPvq;
            TPls.Text = Datos.PTCond7;
            TVnev.Text = Datos.TPcp.ToString();
            TInev.Text = Datos.TPrs7;
            TCsist.Text = Datos.TCsist1;
            TInev.Text = Datos.TInev;
            TIned.Text = Datos.TIned;
            TIncd.Text = Datos.TIncd;
            TInc.Text = Datos.TInc;
            //GuardarValores();
        }


        public void CambiarActualcam(string nuevoValor)
        {
            Datos.actualcam = nuevoValor;
        }




        //------------------------------------------------------ Nueva logica de Nabegación

        

        //-------------------------------------------------------

        /// <summary>
        /// Add a camera
        /// </summary>
        /// <returns>returns true or false</returns>
        bool add()
        {
            if (this.Validateit())
            {
                CCam myCam = new CCam(TNC.Text, TTem.Text, TLargo.Text, TAncho.Text, TAlto.Text, TVolu.Text, TCF.Text, TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text, 
                    TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text, 
                    TCentx.Text, TCxp.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text,
                    TMuc.Text, TMevp.Text, TSol.Text, TValv.Text, TCvta.Text, TCuadro.Text, CBexpo.Text, CSumi.Text, Ctpd.Text, Coff3.Text,Cnoff6.Text, 
                    CTEvap.Text, TCodValv.Text, TInc.Text, TInev.Text, TVnev.Text, TIned.Text, TIncd.Text, TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TTint.Text, TEquip.Text, 
                    TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text, TDmce.Text, TLcc.Text, TLss.Text, TPcmc.Text, 
                    TPcond.Text, TModex.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, 
                    TPnt.Text, TPml.Text,TCt150.Text, TCt04.Text, TCt150m.Text, TCt04m.Text, SPtp75.Text, SPtp74.Text,TTcmc1.Text, TTcmc2.Text, 
                    TTcmc3.Text, TFWe.Text, TTcmc6.Text, TTcmc8.Text, TTcmc9.Text, CBps.Text, CTPup.Text, CTPus.Text, CTLup.Text, TLtemp.ToString(), TLPbar.ToString(), 
                    DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString(), pk); 

                if (myOferta == null)
                {
                    LEstado.Text = "Estado: Creando oferta...";
                    try
                    {
                        myOferta = new COferta(TNP.Text, TREF.Text, TNO.Text, TCmat.Text, TFecha.Text, TLcc.Text, TLss.Text, TPcmc.Text, TPcond.Text,
                            TModex.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, 
                            TPnt.Text, TPml.Text,  TCt150.Text, TCt04.Text, TCt150m.Text, TCt04m.Text, SPtp75.Text, SPtp74.Text, TInc.Text, TInev.Text, TVnev.Text, TIned.Text, TIncd.Text, 
                            TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TTint.Text, TEquip.Text, CSumi.Text, int.Parse(TCC.Text), CCastre.Text, CCpcion.Text, 
                            CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text, CLugar.Text, CClit.Text, CClit1.Text, TBscu.Text, TBcont.Text, TBcos.Text, TBdir.Text, TBenv.Text, 
                            TBpo.Text, TBfec.Text, TBdes.Text, CCdc.Text, CFlet.Text, CCgr.Text, CIntr.Text, CDesct.Text, CNcont.Text, TTcmc1.Text, TTcmc2.Text, TTcmc3.Text, 
                            TFWe.Text, TTcmc6.Text, TTcmc8.Text, TTcmc9.Text, CBps.Text, CTPup.Text, CTPus.Text, CTLup.Text, TLtemp.ToString(), TLPbar.ToString(),
                            DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString());
                    }
                    catch { return false; }
                }
                LEstado.Text = "Estado: Añadiendo cámara...";
                if (myOferta.AddCam(myCam))
                {
                    this.added();
                    LEstado.Text = "Cámaras almacenadas: " + (myOferta.GetCont()).ToString();
                    if (myOferta.GetCont().ToString() == TCC.Text)
                        LEstado.Text = "Todas las cámara entradas.";
                    borrarCámaraActualToolStripMenuItem.Enabled = false;
                    actualizarCámaraActualToolStripMenuItem.Enabled = false;
                    actualizarOfertaActualToolStripMenuItem.Enabled = true;
                    guardarComoToolStripMenuItem.Enabled = true;
                    guardarToolStripMenuItem.Enabled = true;
                    BExportar.Enabled = true;
                }
                else
                {
                    LEstado.Text = "Todas las cámara entradas.";
                    return false;
                }
                return true;
                
            }
            else
                return false;

        }

        /// <summary>
        /// Valida todos los campos de la aplicacion
        /// </summary>
        /// <returns>devuelve verdadero o falso</returns>
        private bool Validateit()
        {
            //-----------Oferta
            if (TNP.Text == "")
                return false;
            if (TNO.Text == "")
                return false;
            
            if (TREF.Text == "")
                return false;
            if (TCC.Text == "")
                return false;
            //-----------Basico
            if (TNC.Text == "")
                return false;


            //-----------dimension
            if (TLargo.Text == "")
                return false;
            if (TAncho.Text == "")
                return false;
            if (TAlto.Text == "")
                return false;
            
           
            if (CSup.Text == "")
                return false;
                                   
            if (TCentx.Text == "")
                return false;
           
            
            
            if (TCxp.Text == "")
                return false;
            
            
            if (CClit.Text == "")
                return false;
            if (CClit1.Text == "")
                return false;
            
            if (CCdc.Text == "")
                return false;

            // Tu lógica cuando las condiciones 
           

            return true;
        }
        
        /// <summary>
        /// asigna a los textbox los valores de la camara pasada como parametro
        /// </summary>
        /// <param name="numcam">camara a extraer valores</param>
        public void asignarcam(int numcam)
        {            
            CCam myCam = myOferta.GetCam(numcam - 1);
            TVolu.Text = myCam.GetVolu();
            pk = myCam.GetPK();
            TNC.Text = myCam.GetNC();
            TTem.Text = myCam.GetTemp();
            TCF.Text = myCam.GetCF();
            TFW.Text = myCam.GetFW();
            TQfw.Text = myCam.GetQfw();
            TCmod.Text = myCam.GetCmod();
            TCmodd.Text = myCam.GetCmodd();
            TCmodp.Text = myCam.GetCmodp();
            TDesc.Text = myCam.GetDesc();
            TPrec.Text = myCam.GetPrec();
            TQfep.Text = myCam.GetQfep();
            TScdro.Text = myCam.GetScdro();
            TSpsi.Text = myCam.GetSpsi();
            TStemp.Text = myCam.GetStemp();
            TApsi.Text = myCam.GetApsi();
            TEmevp.Text = myCam.GetEmevp();
            TLargo.Text = myCam.GetLargo();
            TAncho.Text = myCam.GetAncho();
            TAlto.Text = myCam.GetAlto();
            CSup.Text = myCam.GetSUP();
            actualcam = CCamara.Text;
            Datos.actualcam = actualcam;
            TCentx.Text = myCam.GetCentx();
            TCxp.Text = myCam.GetCxp();
            CCastre.Text = myCam.GetCastre();
            CCpcion.Text = myCam.GetCpcion();
            CCfrio.Text = myCam.GetCfrio();
            CCeq1.Text = myCam.GetCeq1();
            CCeq2.Text = myCam.GetCeq2();
            CCeq3.Text = myCam.GetCeq3();
            TMuc.Text = myCam.GetTMuc();
            TMevp.Text = myCam.GetTMevp();
            TValv.Text = myCam.GetTValv();
            TCvta.Text = myCam.GetTCvta();
            TSol.Text = myCam.GetTSol();
            TCuadro.Text = myCam.GetTCuadro();
            CBexpo.Text = myCam.GetCBexpo();
            CSumi.Text = myCam.GetCSumi();
            Ctpd.Text = myCam.GetCtpd();
            Coff3.Text = myCam.GetCoff3();
            CTEvap.Text = myCam.GetCTEvap();
            TCodValv.Text = myCam.GetCodValv();
            TInc.Text = myCam.GetTInc();
            TInev.Text = myCam.GetTInev();
            TVnev.Text = myCam.GetTVnev();
            TIned.Text = myCam.GetTIned();
            TIncd.Text = myCam.GetTIncd();
            TIpv.Text = myCam.GetTIpv();
            TIcc.Text = myCam.GetTIcc();
            TQevp.Text = myCam.GetTQevp();
            TQevpd.Text = myCam.GetTQevpd();
            TTint.Text = myCam.GetTTint();
            TEquip.Text = myCam.GetTEquip();
            TCint1.Text = myCam.GetTCint1();
            TCint2.Text = myCam.GetTCint2();
            TCint3.Text = myCam.GetTCint3();
            TMcc.Text = myCam.GetTMcc();
            TCmce.Text = myCam.GetTCmce();
            TPmce.Text = myCam.GetTPmce();
            TDmce.Text = myCam.GetTDmce();
            TLcc.Text = myCam.GetTLcc();
            TLss.Text = myCam.GetTLss();
            TPcmc.Text = myCam.GetTPcmc();
            TPcond.Text = myCam.GetTPcond();
            TModex.Text = myCam.GetTModex();
            TPvq.Text = myCam.GetTPvq();
            TPls.Text = myCam.GetTPls();
            TPosc.Text = myCam.GetTPosc();
            TPsq.Text = myCam.GetTPsq();
            TPcq.Text = myCam.GetTPcq();
            TPcy.Text = myCam.GetTPcy();
            TPcp.Text = myCam.GetTPcp();
            TPex.Text = myCam.GetTPex();
            TPrs.Text = myCam.GetTPrs();
            TCsist.Text = myCam.GetTCsist();
            TPem.Text = myCam.GetTPem();
            TPnt.Text = myCam.GetTPnt();
            TPml.Text = myCam.GetTPml();
            TCt04.Text = myCam.GetTCt04();
            TCt150m.Text = myCam.GetTCt150m();
            TCt04m.Text = myCam.GetTCt04m();
            TTcmc1.Text = myCam.GetTTcmc1();
            TTcmc2.Text = myCam.GetTTcmc2();
            TTcmc3.Text = myCam.GetTTcmc3();
            TFWe.Text = myCam.GetFWe();
            TTcmc6.Text = myCam.GetTTcmc6();
            TTcmc8.Text = myCam.GetTTcmc8();
            TTcmc9.Text = myCam.GetTTcmc9();
            CBps.Text = myCam.GetCBps();
            CTPup.Text = myCam.GetCTPup();
            CTPus.Text = myCam.GetCTPus();
            CTLup.Text = myCam.GetCTLup();
            string TLtemp = myCam.GetTLtemp();  ////  TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString()
            string TLPbar = myCam.GetTLPbar();
            string DTpress = myCam.GetDTpress();
            string TCPbar = myCam.GetTCPbar();
            string TEPbar = myCam.GetTEPbar();
            pictureBox2.Image = myUserControl.IMAGES.Images[myCam.GetPK()];
           


        }            

        private void guardarToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Guardar();
        }

        private void nuevaOfertaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            generado = false;
            saved = false;
            TNC.Clear();
            TTem.Clear();
            TLargo.Clear();
            TAncho.Clear();
            TAlto.Clear();
            TVolu.Clear();
            TCF.Clear();
            TFW.Clear();
            TQfw.Clear();
            TCmod.Clear();
            TCmodd.Clear();
            TCmodp.Clear();
            TDesc.Clear();
            TPrec.Clear();
            TQfep.Clear();
            TScdro.Clear();
            TSpsi.Clear();
            TStemp.Clear();
            TApsi.Clear();
            TEmevp.Clear();
            CSup.Text = "";
            CCamara.Items.Clear();
            TNP.Clear();
            TNO.Clear();
            TCmat.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcond.Clear();
            TModex.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TCt150.Clear();
            TCt04.Clear();
            TCt150m.Clear();
            TCt04m.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp150.Clear();
            SPtp151.Clear();
            TInc.Clear();
            TInev.Clear();
            TIned.Clear();
            TIncd.Clear();
            TIpv.Clear();
            TIcc.Clear();
            TCint1.Clear();
            TCint2.Clear();
            TCint3.Clear();
            TMcc.Clear();
            TCmce.Clear();
            TPmce.Clear();
            TDmce.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcond.Clear();
            TModex.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TCt150.Clear();
            TCt04.Clear();
            TCt150m.Clear();
            TCt04m.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp150.Clear();
            SPtp151.Clear();
            TQevp.Clear();
            TQevpd.Clear();
            TTint.Clear();
            TEquip.Clear();
            TCC.Clear();
            TREF.Clear();
            TBscu.Clear();
            TBcont.Clear();
            TBcos.Clear();
            TBdir.Clear();
            TBenv.Clear();
            TBpo.Clear();
            TBfec.Clear();
            TBdes.Clear();
            myOferta = null;
            CCamara.Text = "";           
            TCentx.Text = "";
            TCxp.Text = "";
            CCastre.Text = "";
            CCpcion.Text = "";
            CCfrio.Text = "";
            CCeq1.Text = "";
            CCeq2.Text = "";
            CCeq3.Text = "";
            this.DateActualizer();
            CLugar.Text = "";
            CClit.Text = "Sonia Aleida";
            CClit1.Text = "";

            Ctpd.Text = "";

            Coff3.Text = "";

            //Cnoff6.Text = "";

            CCdc.Text = "";
            CFlet.Text = "";
            CCgr.Text = "";
            CIntr.Text = "";
            CDesct.Text = "";
            CNcont.Text = "";

            TPcond.Clear();
            TFWe.Clear();

            borrarCámaraActualToolStripMenuItem.Enabled = false;
            actualizarCámaraActualToolStripMenuItem.Enabled = false;
            actualizarOfertaActualToolStripMenuItem.Enabled = false;
            guardarComoToolStripMenuItem.Enabled = false;
            guardarToolStripMenuItem.Enabled = false;            

        }

        void abrirFop()
        {
            try
            {
                
                    //LEstado.Text = "Abriendo archivo fop...";
                    ruta = myOpenDialog.FileName;
                    saved = true;
                    myFileStream = new FileStream(ruta, FileMode.Open);
                    myFormatter = new BinaryFormatter();
                    myOferta = (COferta)myFormatter.Deserialize(myFileStream);
                    generado = false;
                    TNP.Text = myOferta.GetNP();
                    TNO.Text = myOferta.GetNO();
                    TCmat.Text = myOferta.GetCmat();

                    TLcc.Text = myOferta.GetLcc();
                    TLss.Text = myOferta.GetLss();
                    TPcmc.Text = myOferta.GetPcmc();

                    TPcond.Text = myOferta.GetPcond();
                    TModex.Text = myOferta.GetModex();
                    TPvq.Text = myOferta.GetPvq();
                    TPls.Text = myOferta.GetPls();
                    TPosc.Text = myOferta.GetPosc();
                    TPsq.Text = myOferta.GetPsq();
                    TPcq.Text = myOferta.GetPcq();
                   
                    TPcy.Text = myOferta.GetPcy();
                    TPcp.Text = myOferta.GetPcp();
                    TPex.Text = myOferta.GetPex();
                    TPrs.Text = myOferta.GetPrs();
                    TCsist.Text = myOferta.GetCsist();
                    TPem.Text = myOferta.GetPem();
                    TPnt.Text = myOferta.GetPnt();
                    TPml.Text = myOferta.GetPml();

                    TCt150.Text = myOferta.GetCt150();
                    TCt04.Text = myOferta.GetCt04();
                    
                    TCt150m.Text = myOferta.GetCt150m();
                    TCt04m.Text = myOferta.GetCt04m();

                    TInc.Text = myOferta.GetInc();
                    TInev.Text = myOferta.GetInev();
                    TVnev.Text = myOferta.GetVnev();
                    TIned.Text = myOferta.GetIned();
                    TIncd.Text = myOferta.GetIncd();
                    TIpv.Text = myOferta.GetIpev();
                    TIcc.Text = myOferta.GetIcc();
                    TQevp.Text = myOferta.GetQevp();
                    TQevpd.Text = myOferta.GetQevpd();
                    
                    TTint.Text = myOferta.GetTint();
                    TEquip.Text = myOferta.GetEquip();
                    TREF.Text = myOferta.GetREF();
                    TCC.Text = myOferta.GetCantCam().ToString();
                    CCam myCam = myOferta.GetCam(0);
                    TFecha.Text = myOferta.GetFecha();
                    TBscu.Text = myOferta.GetBscu();
                    TBcont.Text = myOferta.GetBcont();
                    TBcos.Text = myOferta.GetBcos();
                    TBdir.Text = myOferta.GetBdir();
                    TBenv.Text = myOferta.GetBenv();
                    TBpo.Text = myOferta.GetBpo();
                    TBfec.Text = myOferta.GetBfec();
                    TBdes.Text = myOferta.GetBdes();
                    CClit.Text = myOferta.GetClit();
                    CClit1.Text = myOferta.GetClit1();
                    
                    CCdc.Text = myOferta.GetCdc();
                    CFlet.Text = myOferta.GetFlet();
                    CCgr.Text = myOferta.GetCgr();
                    CIntr.Text = myOferta.GetIntr();
                    CDesct.Text = myOferta.GetDesct();
                    CNcont.Text = myOferta.GetNcont();
    
                    TTcmc1.Text = myOferta.GetTcmc1();
                    TTcmc2.Text = myOferta.GetTcmc2();
                    TTcmc3.Text = myOferta.GetTcmc3();

                    TTcmc6.Text = myOferta.GetTcmc6();
                   
                    TTcmc8.Text = myOferta.GetTcmc8();
                    TTcmc9.Text = myOferta.GetTcmc9();
                    
                    CBps.Text = myOferta.GetBps();
                    CTPup.Text = myOferta.GetTPup();
                    CTPus.Text = myOferta.GetTPus();
                    CTLup.Text = myOferta.GetTLup();
                    string TLtemp = myOferta.GetLtemp(); ////  TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString()
                    string TLPbar = myOferta.GetLPbar();
                    string DTpress = myOferta.GetTpress();
                    string TCPbar = myOferta.GetCPbar();
                    string TEPbar = myOferta.GetEPbar();
                    CLugar.Text = myOferta.GetLugar();
                    CClit.Text = myOferta.GetClit();
                    CClit1.Text = myOferta.GetClit1();

                    this.DateActualizer();

                    this.asignarcam(1);
                    CCamara.Items.Clear();
                    for (int i = 1; i <= myOferta.GetCont(); i++)
                        CCamara.Items.Add(i.ToString());
                    CCamara.Text = "1";
                    myFileStream.Close();
                    myFileStream.Dispose();
                    myFileStream = null;
                    myFormatter = null;
                    actualizarOfertaActualToolStripMenuItem.Enabled = true;
                    guardarComoToolStripMenuItem.Enabled = true;
                    guardarToolStripMenuItem.Enabled = true;
                    LEstado.Text = "Cámaras almacenadas: " + (myOferta.GetCont()).ToString();
                    
                       BExportar.Enabled = true;

                       BAbrir.Enabled = true;
                       MenuAbrir.Enabled = true;
                       actualizarCámaraActualToolStripMenuItem.Enabled = true;
                       actualizarOfertaActualToolStripMenuItem.Enabled = true;
                       BAdd.Enabled = true;
                       borrarCámaraActualToolStripMenuItem.Enabled = true;
                       string filename = myOpenDialog.FileName;
                       this.PonerNombre(filename);             
                
            }
            catch /*(Exception ex)*/
            {                
                    OpenThread = new Thread(new ThreadStart(this.abrir));
                    OpenThread.Start();
                    TiCCamaraHandler.Enabled = true;
                    timer1.Enabled = true;
                    myFileStream = null;
                    myFormatter = null;                
            }
        }

        void PonerNombre(string filename)
        {            
            string myfile = "";
            for (int f = 0; f < filename.Length; f++)
            {
                if (filename[(filename.Length - 1) - f] != '\\')
                {
                    myfile += filename[(filename.Length - 1) - f].ToString();
                }
                else
                    break;
            }
            //----------------------------------------------------------------
            string filederecho = "";
            for (int f = 0; f < myfile.Length; f++)
            {
                filederecho += myfile[(myfile.Length - 1) - f].ToString();
            }

            this.Text = "PROYECTO CAMARAS FRIAS JDC - " + filederecho;      
        }

        private void abrirOfertaToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (myOpenDialog.ShowDialog() == DialogResult.OK)
            {
                this.abrirFop();
            }          
        }

        void abrir()
        {
                    //LEstado.Text = "Abriendo archivo xlsx...";
                    this.calcbeging();
                    try
                    {
                        if (!ini)
                            throw new Exception("Error, no se pueden ejecutar procedimientos de cálculos");
                        saved = false;
                        ruta = myOpenDialog.FileName;
                        _Application myExcel = new Microsoft.Office.Interop.Excel.Application();
                        myExcel.Visible = false;
                        myExcel.UserControl = false;
                        myExcel.DisplayAlerts = false;
                        _Workbook myWorkbook = myExcel.Workbooks.Open(ruta, Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing,
                    Type.Missing);
                        int i = 6;
                        _Worksheet myWorksheet = (_Worksheet)myWorkbook.Worksheets[2];
                        if (!(myWorksheet.Name == "PRESUPUESTO"))
                        {
                            myWorksheet = (_Worksheet)myWorkbook.Worksheets[5];
                            if (!(myWorksheet.Name == "PRESUPUESTO"))
                                OpenThread.Abort();
                        }

                        myOferta = new COferta("---", "---", "---", "", "---", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                            "", "", "", "", "", "", "", "", 40, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                            "", "", "", "", "");
                        Range myRange = myWorksheet.get_Range("A1", "I45");
                        Array myValues = (Array)myRange.Value2;                       

                        while (true)
                        {
                            try
                            {
                                //CONSTRUCTOR**
                                CCam myCam = new CCam(myValues.GetValue(i, 4).ToString(), myValues.GetValue(i, 5).ToString(), myValues.GetValue(i, 6).ToString(), 
                                    myValues.GetValue(i, 7).ToString(), myValues.GetValue(i, 8).ToString(), "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 0); //////
                                myOferta.AddCam(myCam);
                                CCam newCalcCam = myOferta.GetCam(i - 6);
                                //newCalcCam.setCF(Calcular(i - 5));
                                myOferta.actualizar(newCalcCam, i - 5);
                                i++;
                                LEstado.Text = (i - 6).ToString() + " camara(s) leidas.";
                                myCam = null;
                            }
                            catch 
                            {                                
                                break; 
                            }
                        }

                        CCam myCam1 = myOferta.GetCam(0);
                        this.SetText(TVolu, myCam1.GetVolu());
                        this.SetText(TNC, myCam1.GetNC());
                        
                        this.SetText(TTem, myCam1.GetTemp());
                       
                        this.SetText(TCF, myCam1.GetCF());
                        
                        this.SetText(TFW, myCam1.GetFW());
                        this.SetText(TQfw, myCam1.GetQfw());
                        this.SetText(TCmod, myCam1.GetCmod());
                        this.SetText(TCmodd, myCam1.GetCmodd());
                        this.SetText(TCmodp, myCam1.GetCmodp());
                        this.SetText(TDesc, myCam1.GetDesc());
                        this.SetText(TPrec, myCam1.GetPrec());
                       
                        this.SetText(TQfep, myCam1.GetQfep());
                        this.SetText(TScdro, myCam1.GetScdro());
                        this.SetText(TSpsi, myCam1.GetSpsi());
                        this.SetText(TStemp, myCam1.GetStemp());
                        this.SetText(TApsi, myCam1.GetApsi());
                        this.SetText(TEmevp, myCam1.GetEmevp());
                        
                       
                       
                       
                        
                      
                        this.SetText(TLargo, myCam1.GetLargo());
                       
                        this.SetText(TAncho, myCam1.GetAncho());
                       
                       
                       
                        this.SetText(TAlto, myCam1.GetAlto());
                       
                       
                       
                        this.SetText(CSup, myCam1.GetSUP());
                       
                        
                                                         
                        this.SetText(TCentx, myCam1.GetCentx());
                       
                        this.SetText(TCxp, myCam1.GetCxp());
                        
                       
                        
                       
                        this.SetText(CCastre, myCam1.GetCastre());
                        this.SetText(CCpcion, myCam1.GetCpcion());
                        this.SetText(CCfrio, myCam1.GetCfrio());
                        this.SetText(CCeq1, myCam1.GetCeq1());
                        this.SetText(CCeq2, myCam1.GetCeq2());
                        this.SetText(CCeq3, myCam1.GetCeq3());
                        this.SetText(CSumi, myCam1.GetCSumi());
                        
                        
                       
                        
                        
                        this.SetText(Ctpd, myCam1.GetCtpd());
                        
                     
                       
                        
                      
                        this.SetText(Coff3, myCam1.GetCoff3());
                        
                       
                       
                      
                        this.SetText(Cnoff6, myCam1.GetCnoff6());
                       
                       
                       
       
                        this.SetText(CTEvap, myCam1.GetCTEvap());
                       
                        this.SetText(TFWe, myCam1.GetFWe());
                        myOferta.SetCantCam(myOferta.GetCont());
                        this.SetText(TCC, (myOferta.GetCont().ToString()));
                        myExcel.Quit();
                       
                        try
                        {
                            BExportar.Enabled = true;
                            BAbrir.Enabled = true;
                            MenuAbrir.Enabled = true;
                            actualizarCámaraActualToolStripMenuItem.Enabled = true;
                            actualizarOfertaActualToolStripMenuItem.Enabled = true;
                            BAdd.Enabled = true;
                            borrarCámaraActualToolStripMenuItem.Enabled = true;
                        }
                        catch { }


                    }
                    catch (Exception exc)
                    {
                       MessageBox.Show(exc.Message,
                            "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        try
                        {
                            BExportar.Enabled = false;

                            BAbrir.Enabled = true;
                            MenuAbrir.Enabled = true;
                            actualizarCámaraActualToolStripMenuItem.Enabled = false;
                            actualizarOfertaActualToolStripMenuItem.Enabled = false;
                            BAdd.Enabled = false;
                            borrarCámaraActualToolStripMenuItem.Enabled = false;
                        }
                        catch { }
                        ruta = null;
                        saved = false;                        
                    }                 
               
            //OpenThread.Abort();
        }

        void camaclear()
        {
            CCamara.Items.Clear();
        }

        private void SetText(Control Contr, string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (Contr.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { Contr, text });
            }
            else
            {
                Contr.Text = text;
            }
        }



        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            if (myOpenDialog.ShowDialog() == DialogResult.OK)
            {
                    BExportar.Enabled = false;

                    BAbrir.Enabled = false;
                    MenuAbrir.Enabled = false;
                    actualizarCámaraActualToolStripMenuItem.Enabled = false;
                    actualizarOfertaActualToolStripMenuItem.Enabled = false;
                    BAdd.Enabled = false;
                    borrarCámaraActualToolStripMenuItem.Enabled = false;
                
                this.abrirFop();
            }
                CBexpo.Visible = true;
               
                CKexpo.Visible = true;
        }
        

        void DateActualizer()
        {
            string DDate = DateTime.Now.ToString();
            TFecha.Text = "";
            for (int i = 0; i < 10; i++)
                TFecha.Text += DDate[i];
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            DateTime fechaF = Convert.ToDateTime("20-12-2025").Date;
            DateTime FechAc = DateTime.Now.Date;
            if (fechaF <= FechAc)
            {
                MessageBox.Show("ProJDC Endpoint Security bloquea el acceso : ");
                string path1 = @"d:\confjdc.jdc";
                string path2 = @"d:\ProJDC-scmf\calc.tpt";
                string path3 = @"d:\ProJDC-scmf\fnstptecuc.tpt";

                try
                {
                    // Delete the newly created file.
                    File.Delete(path1);
                    File.Delete(path2);
                    File.Delete(path3);
                    MessageBox.Show("File was successfully deleted ProJDC Endpoint Security.", "Info");
                }
                catch (Exception)
                {
                    MessageBox.Show("File was not deleted ProJDC Endpoint Security.", "Info");
                }
                {
                    const string fic = @"C:\WINDOWS\system32\DEFAULT.dll";
                    string texto = "AutoRun WindowsSystem32";

                    System.IO.StreamWriter sw =
                        new System.IO.StreamWriter(fic, false, System.Text.Encoding.UTF8);
                    sw.WriteLine(texto);
                    sw.Close();
                }
                {
                    const string fic = @"C:\WINDOWS\System32.txt";
                    string texto = "Pro JDC Endpoint Security bloquea el acceso al registro, borrar SubKey:[HKEY_LOCAL_MACHINE\"SOFTWARE\"Microsoft\"SystemCertificates\"SPC\"Certificates\"F43CF6F7A805F98E51D5ED03DE67F3D8FC8EDDCD]";

                    System.IO.StreamWriter sw =
                        new System.IO.StreamWriter(fic, false, System.Text.Encoding.UTF8);
                    sw.WriteLine(texto);
                    sw.Close();
                }

                System.Windows.Forms.Application.Exit();
                return;

            }
            DateTime fechaX = Convert.ToDateTime("20-12-2025").Date;
            DateTime FechaD = DateTime.Now.Date;
            if (fechaX <= FechaD)
            {
                MessageBox.Show("ProJDC Endpoint Security bloquea el acceso: ");
                string path1 = @"d:\confjdc.jdc";
                string path2 = @"d:\ProJDC-scmf\calc.tpt";
                string path3 = @"d:\ProJDC-scmf\fnstptecuc.tpt";
                try
                {
                    // Delete the newly created file.
                    File.Delete(path1);
                    File.Delete(path2);
                    File.Delete(path3);
                    MessageBox.Show("File was successfully deleted ProJDC Endpoint Security.", "Info");
                }
                catch (Exception)
                {
                    MessageBox.Show("File was not deleted ProJDC Endpoint Security.", "Info");
                }
                {
                    const string fic = @"C:\WINDOWS\system32\DEFAULT.dll";
                    string texto = "AutoRun WindowsSystem32";

                    System.IO.StreamWriter sw =
                        new System.IO.StreamWriter(fic, false, System.Text.Encoding.UTF8);
                    sw.WriteLine(texto);
                    sw.Close();
                }
                {
                    const string fic = @"C:\WINDOWS\System32.txt";
                    string texto = "ProJDC Endpoint Security bloquea el acceso al registro, borrar SubKey:[HKEY_LOCAL_MACHINE\"SOFTWARE\"Microsoft\"SystemCertificates\"SPC\"Certificates\"F43CF6F7A805F98E51D5ED03DE67F3D8FC8EDDCD]";

                    System.IO.StreamWriter sw =
                        new System.IO.StreamWriter(fic, false, System.Text.Encoding.UTF8);
                    sw.WriteLine(texto);
                    sw.Close();
                }

                System.Windows.Forms.Application.Exit();
                return;

            }

        }

        private void actualizarOfertaActualToolStripMenuItem_Click_1(object sender, EventArgs e)
        {
            if (TNP.Text != "" && TNO.Text != ""  && TREF.Text != "" && TCC.Text != "")
            {
                try
                {
                    myOferta.SetNP(TNP.Text);
                    myOferta.SetNO(TNO.Text);
                    myOferta.SetCmat(TCmat.Text);

                    myOferta.SetLcc(TLcc.Text);
                    myOferta.SetLss(TLss.Text);
                    myOferta.SetPcmc(TPcmc.Text);

                    myOferta.SetPcond(TPcond.Text);
                    myOferta.SetModex(TModex.Text);
                    myOferta.SetPvq(TPvq.Text);
                    myOferta.SetPls(TPls.Text);
                    myOferta.SetPosc(TPosc.Text);
                    myOferta.SetPsq(TPsq.Text);
                    myOferta.SetPcq(TPcq.Text);
                   
                    myOferta.SetPcy(TPcy.Text);
                    myOferta.SetPcp(TPcp.Text);
                    myOferta.SetPex(TPex.Text);
                    myOferta.SetPrs(TPrs.Text);
                    myOferta.SetCsist(TCsist.Text);
                    myOferta.SetPem(TPem.Text);
                    myOferta.SetPnt(TPnt.Text);
                    myOferta.SetPml(TPml.Text);

                    myOferta.SetCt150(TCt150.Text);
                    myOferta.SetCt04(TCt04.Text);
                   
                   
                    myOferta.SetCt150m(TCt150m.Text);
                    myOferta.SetCt04m(TCt04m.Text);

                   
                    myOferta.SetInc(TInc.Text);
                    myOferta.SetInev(TInev.Text);
                    myOferta.SetVnev(TVnev.Text);
                    myOferta.SetIned(TIned.Text);
                    myOferta.SetIncd(TIncd.Text);
                    myOferta.SetIpv(TIpv.Text);
                    myOferta.SetIcc(TIcc.Text);
                    myOferta.SetQevp(TQevp.Text);
                    myOferta.SetQevpd(TQevpd.Text);
                   
                    myOferta.SetTint(TTint.Text);
                    myOferta.SetEquip(TEquip.Text);
                    myOferta.SetCantCam(int.Parse(TCC.Text));
                    myOferta.SetREF(TREF.Text);

                    myOferta.SetCastre(CCastre.Text);
                    myOferta.SetCpcion(CCpcion.Text);
                    myOferta.SetCfrio(CCfrio.Text);
                    myOferta.SetCeq1(CCeq1.Text);
                    myOferta.SetCeq2(CCeq2.Text);
                    myOferta.SetCeq3(CCeq3.Text);
                   
                    
                    myOferta.SetLugar(CLugar.Text);
                    myOferta.SetClit(CClit.Text);
                    myOferta.SetClit1(CClit1.Text);
                    
                    
                    myOferta.SetBscu(TBscu.Text);
                    myOferta.SetBcont(TBcont.Text);
                    myOferta.SetBcos(TBcos.Text);
                    myOferta.SetBdir(TBdir.Text);
                    myOferta.SetBenv(TBenv.Text);
                    myOferta.SetBpo(TBpo.Text);
                    myOferta.SetBfec(TBfec.Text);
                    myOferta.SetBdes(TBdes.Text);
                    myOferta.SetCdc(CCdc.Text);
                    myOferta.SetFlet(CFlet.Text);
                    myOferta.SetCgr(CCgr.Text);
                    myOferta.SetIntr(CIntr.Text);
                    myOferta.SetDesct(CDesct.Text);
                    myOferta.SetNcont(CNcont.Text);

                    myOferta.SetTcmc1(TTcmc1.Text);
                    myOferta.SetTcmc2(TTcmc2.Text);
                    myOferta.SetTcmc3(TTcmc3.Text);
                    
                   
                    myOferta.SetTcmc6(TTcmc6.Text);
                   
                    myOferta.SetTcmc8(TTcmc8.Text);
                    myOferta.SetTcmc9(TTcmc9.Text);
                   
                    myOferta.SetBps(CBps.Text);
                    myOferta.SetTPup(CTPup.Text);
                    myOferta.SetTPus(CTPus.Text);
                    myOferta.SetTLup(CTLup.Text);
                    myOferta.SetLtemp(TLtemp.ToString());
                    myOferta.SetLPbar(TLPbar.ToString());
                    myOferta.SetTpress(DTpress.ToString());
                    myOferta.SetCPbar(TCPbar.ToString());
                    myOferta.SetEPbar(TEPbar.ToString());

                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                            "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
                MessageBox.Show("Valor a actualizar incorrectos",
                            "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);

        }        

        private void CCamara_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (myOferta != null)
            {
                if (int.Parse(CCamara.Text) > myOferta.GetCont())
                {
                    MessageBox.Show("Error, ese # de cámara no existe.",
                        "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    CCamara.Text = actualcam;
                }
                else
                {
                    int myCont = int.Parse(CCamara.Text);
                    this.asignarcam(myCont);
                    borrarCámaraActualToolStripMenuItem.Enabled = true;
                    actualizarCámaraActualToolStripMenuItem.Enabled = true;
                }
            }
            else
            {
                MessageBox.Show("Error, Aun no existe una oferta y no hay cámaras.",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CCamara.Text = "";
            }
        }

        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            if (!this.add())
                MessageBox.Show("Error, no a entrado correctamente los datos para la cámara u oferta, ó puede que no pueda almacenar más cámaras.",
                    "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void toolStripButton4_Click(object sender, EventArgs e)
        {
            if (myOferta != null)
            {
                if (CCamara.Text != "")
                {
                    if (int.Parse(CCamara.Text) > myOferta.GetCont())
                    {
                        MessageBox.Show("Error, ese # de cámara no existe.",
                            "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CCamara.Text = actualcam;
                    }
                    else
                    {
                        myOferta.BorrarCam(int.Parse(CCamara.Text));
                        if (int.Parse(CCamara.Text) > 1)
                        {
                            CCamara.Text = (int.Parse(actualcam) - 1).ToString();
                            this.asignarcam(int.Parse(actualcam));
                        }
                        else
                            this.asignarcam(int.Parse(actualcam));
                        CCamara.Items.Clear();
                        for (int i = 0; i < myOferta.GetCont(); i++)
                            CCamara.Items.Add((i + 1).ToString());
                    }
                }
                else
                    MessageBox.Show("Error, cámara incorrecta",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                MessageBox.Show("Error, Aun no existe una oferta y no hay cámaras.",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CCamara.Text = "";
            }
        }

        private void toolStripButton5_Click(object sender, EventArgs e)
        {
            this.GuardarDatos();
        }

        private void guardarComoToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (myopSaveDialog.ShowDialog() == DialogResult.OK)
            {
                ruta = myopSaveDialog.FileName;
                myFileStream = new FileStream(ruta, FileMode.Create);
                myFormatter = new BinaryFormatter();
                myFormatter.Serialize(myFileStream, myOferta);
                saved = true;
                string filename = myopSaveDialog.FileName;
                this.PonerNombre(filename);
                myFileStream.Close();
                myFileStream.Dispose();
                myFileStream = null;
                myFormatter = null;
            }
        }

        private void toolStripButton4_Click_1(object sender, EventArgs e)
        {
            this.Guardar();
        }

        void Guardar()
        {
            if (myOferta != null)
            {
                if (!saved)
                {
                    if (RCUC.Checked)
                        myopSaveDialog.FileName = "OFERTA " + myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " » " + myOferta.GetNP() + myOferta.GetREF() + ".fop";
                    else
                        myopSaveDialog.FileName = "OFERTA1 " + myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " » " + myOferta.GetNP() + myOferta.GetREF() + ".fop";
                    if (myopSaveDialog.ShowDialog() == DialogResult.OK)
                    {
                        ruta = myopSaveDialog.FileName;
                        myFileStream = new FileStream(ruta, FileMode.Create);
                        myFormatter = new BinaryFormatter();
                        myFormatter.Serialize(myFileStream, myOferta);
                        saved = true;
                        myFileStream.Close();
                        myFileStream.Dispose();
                        myFileStream = null;
                        myFormatter = null;
                    }
                }
                else
                {
                    myFileStream = new FileStream(ruta, FileMode.Create);
                    myFormatter = new BinaryFormatter();
                    myFormatter.Serialize(myFileStream, myOferta);
                    saved = true;
                    myFileStream.Close();
                    myFileStream.Dispose();
                    myFileStream = null;
                    myFormatter = null;
                }
            }
        }

       
        

        private void button1_Click(object sender, EventArgs e)
        {
            if (!this.add())
                MessageBox.Show("Error, no a entrado correctamente los datos para la cámara u oferta, ó puede que no pueda almacenar más cámaras.",
                    "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void exportarAExcel20072010xlsxToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Export();
        }

        void  Export()
        {
            if (myOferta != null)
            {
                if (myOferta.GetCantCam() == myOferta.GetNumCam())
                {
                   
                        if (myOferta.GetClit() == "Sonia Aleida")
                        {
                            if (RCX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RGX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RFTX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RODC.Checked)
                                myXlsxSaveDialog.FileName = " « " + "SCU - " + myOferta.GetNO() + " " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx"; 
                        }
                    
                    else
                    {
                        if (RCX.Checked)
                            myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + "JDC.xlsx";
                        if (RGX.Checked)
                            myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + "JDC.xlsx";
                        if (RFTX.Checked)
                            myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + "JDC.xlsx";
                        if (RODC.Checked)
                            myXlsxSaveDialog.FileName = " « " + "SCU - " + myOferta.GetNO() + " " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + "JDC.xlsx";
                    }
                   
                    
                    if (myThread != null)
                    {
                        if (myThread.IsAlive)
                        {
                            if (MessageBox.Show("Advertencia, se está generando una oferta.\n¿Desea detenerla y generar una nueva?",
                                    "Error al leer datos:", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.Yes)
                            {
                                myThread.Abort();

                                if (myXlsxSaveDialog.ShowDialog() == DialogResult.OK)
                                {
                                    BExportar.Enabled = false;
                                    BAbrir.Enabled = false;
                                    MenuAbrir.Enabled = false;
                                    actualizarCámaraActualToolStripMenuItem.Enabled = false;
                                    actualizarOfertaActualToolStripMenuItem.Enabled = false;
                                    BAdd.Enabled = false;
                                    borrarCámaraActualToolStripMenuItem.Enabled = false;
                                    COferta thisOferta = myOferta;
                                    CExcelWork Excel = new CExcelWork(thisOferta, local, this);
                                    myThread = new Thread(new ThreadStart(Excel.GenerarOferta));
                                    Excel.SetThread(myThread);
                                    myThread.Start();
                                    timer1.Enabled = true;
                                    //MessageBox.Show("no");
                                }
                            }
                        }
                        else
                            if (myXlsxSaveDialog.ShowDialog() == DialogResult.OK)
                            {
                                BExportar.Enabled = false;
                                BAbrir.Enabled = false;
                                MenuAbrir.Enabled = false;
                                actualizarCámaraActualToolStripMenuItem.Enabled = false;
                                actualizarOfertaActualToolStripMenuItem.Enabled = false;
                                BAdd.Enabled = false;
                                borrarCámaraActualToolStripMenuItem.Enabled = false;
                                COferta thisOferta = myOferta;
                                CExcelWork Excel = new CExcelWork(thisOferta, local, this);
                                myThread = new Thread(new ThreadStart(Excel.GenerarOferta));
                                Excel.SetThread(myThread);
                                myThread.Start();
                                timer1.Enabled = true;
                                //MessageBox.Show("no");
                            }
                    }
                    else
                    {
                        if (myXlsxSaveDialog.ShowDialog() == DialogResult.OK)
                        {
                            BExportar.Enabled = false;
                            BAbrir.Enabled = false;
                            MenuAbrir.Enabled = false;
                            actualizarCámaraActualToolStripMenuItem.Enabled = false;
                            actualizarOfertaActualToolStripMenuItem.Enabled = false;
                            BAdd.Enabled = false;
                            borrarCámaraActualToolStripMenuItem.Enabled = false;
                            COferta thisOferta = myOferta;
                            CExcelWork Excel = new CExcelWork(thisOferta, local, this);
                            myThread = new Thread(new ThreadStart(Excel.GenerarOferta));
                            Excel.SetThread(myThread);
                            myThread.Start();
                            timer1.Enabled = true;
                        }
                    }
                }
                else
                    MessageBox.Show("Error, la cantidad de cámaras no corresponde a la cantidad existente.",
                                "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            else
                MessageBox.Show("Error, no existe la oferta a generar!",
                                "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void toolStripSplitButton1_ButtonClick(object sender, EventArgs e)
        {            
            this.Export();            
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (myThread != null)
                if (myThread.IsAlive)
                    myThread.Abort();

            if (OpenThread != null)
                if (OpenThread.IsAlive)
                    OpenThread.Abort();

            if (Obtener != null)
                if (Obtener.IsAlive)
                    Obtener.Abort();

            try
            {
                TrabExcel.Disconnect();
                Environment.Exit(0);
            }
            catch { }
            
            System.Windows.Forms.Application.Exit();            
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            try
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                try
                {
                    Marshal.FinalReleaseComObject(CalcWorkSheet);
                }
                catch { }
               
                Environment.Exit(0);
                System.Windows.Forms.Application.Exit();
            }
            catch { }
        }

        void added()
        {
            TNC.Clear();
            TTem.Clear();
            TLargo.Clear();
            TAncho.Clear();
            TAlto.Clear();
            TVolu.Clear();
            TCF.Clear();
            TFW.Clear();
            TQfw.Clear();
            TCmod.Clear();
            TCmodd.Clear();
            TCmodp.Clear();
            TDesc.Clear();
            TPrec.Clear();
            TQfep.Clear();
            TScdro.Clear();
            TSpsi.Clear();
            TStemp.Clear();
            TApsi.Clear();
            TEmevp.Clear();
            CSup.Text = "";                
            TCentx.Text = "25";
            TCxp.Text = "0";
            CCastre.Text = "";
            CCpcion.Text = "";
            CCfrio.Text = "";
            CCeq1.Text = "";
            CCeq2.Text = "";
            CCeq3.Text = "";
            CCamara.Items.Add(myOferta.GetCont().ToString());
            CCamara.Text = "";
            TMevp.Clear();
            TCuadro.Clear();
            TMuc.Clear();
            TSol.Clear();
            TValv.Clear();
            TCvta.Clear();
            TCodValv.Clear();
            TInc.Clear();
            TInev.Clear();
            TVnev.Clear();
            TIned.Clear();
            TIncd.Clear();
            TIpv.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcond.Clear();
            TModex.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TCt150.Clear();
            TCt04.Clear();
            TCt150m.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp150.Clear();
            SPtp151.Clear();
            TIcc.Clear();
            TQevp.Clear();
            TQevpd.Clear();
            TTint.Clear();
            TEquip.Clear();
            TCint1.Clear();
            TCint2.Clear();
            TCint3.Clear();
            TMcc.Clear();
            TCmce.Clear();
            TPmce.Clear();
            TDmce.Clear();
            Coff3.Text = "";
            //Cnoff6.Text = "";
            TTcmc1.Clear();
            TTcmc2.Clear();
            TTcmc3.Clear();
            TFWe.Clear();
            TTcmc6.Clear();
            TTcmc8.Clear();
            TTcmc9.Clear();
             
        }



        private void myOpenDialog_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void TiCCamaraHandler_Tick(object sender, EventArgs e)
        {
            if (OpenThread != null)
            if (OpenThread.IsAlive == false)
            {
                try
                {
                    this.CCamara.Items.Clear();

                    for (int s = 1; s <= myOferta.GetCont(); s++)
                        CCamara.Items.Add(s.ToString());
                    
                    CCamara.Text = "1";
                    
                    actualizarOfertaActualToolStripMenuItem.Enabled = true;
                    guardarComoToolStripMenuItem.Enabled = true;
                    guardarToolStripMenuItem.Enabled = true;
                    LEstado.Text = "Cámaras almacenadas: " + (myOferta.GetCont()).ToString();
                    TiCCamaraHandler.Enabled = false;
                    this.DateActualizer();
                }
                catch
                {
                    TiCCamaraHandler.Enabled = false;
                    MessageBox.Show("ERROR FATAL",
                                  "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    LEstado.Text = "ERROR";
                }

            }
        }

        private void toolStripButton3_Click_1(object sender, EventArgs e)
        {
            if (!(CCamara.Text == ""))
            {
                try
                {
                    int cam = int.Parse(CCamara.Text);
                    if (cam - 1 <= 0)
                        throw new Exception("Camara no existente");
                    else
                    {
                        this.asignarcam(cam - 1);
                        CCamara.Text = (cam - 1).ToString();
                    }
                }
                catch {}
            }
            else
                MessageBox.Show("No hay ninguna camara seleccionada",
                            "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);      
            
        }

        private void toolStripButton5_Click_1(object sender, EventArgs e)
        {
            if (!(CCamara.Text == ""))
            {
                try
                {
                    int cam = int.Parse(CCamara.Text);
                    if (cam + 1 > myOferta.GetCont())
                        throw new Exception("Camara no existente");
                    else
                    {
                        this.asignarcam(cam + 1);
                        CCamara.Text = (cam + 1).ToString();
                    }
                }
                catch {}
            }
            else
                MessageBox.Show("No hay ninguna camara seleccionada",
                            "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);       
            
        }
        
       

        private void timer1_Tick(object sender, EventArgs e)
        {
            if (myThread != null)
            if (!myThread.IsAlive)
            {

                BExportar.Enabled = true;

                BAbrir.Enabled = true;
                MenuAbrir.Enabled = true;
                actualizarCámaraActualToolStripMenuItem.Enabled = true;
                actualizarOfertaActualToolStripMenuItem.Enabled = true;
                BAdd.Enabled = true;
                borrarCámaraActualToolStripMenuItem.Enabled = false;
                timer1.Enabled = false;
            }

            if (!(OpenThread == null))
            if (!OpenThread.IsAlive)
            {

                BExportar.Enabled = true;

                BAbrir.Enabled = true;
                MenuAbrir.Enabled = true;
                actualizarCámaraActualToolStripMenuItem.Enabled = true;
                actualizarOfertaActualToolStripMenuItem.Enabled = true;
                BAdd.Enabled = true;
                borrarCámaraActualToolStripMenuItem.Enabled = false;
                timer1.Enabled = false;
            }
        }

        private void sobreToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Form3 myHelpForm = new Form3();
            myHelpForm.ShowDialog();
        }

       
       
        private void button1_Click_1(object sender, EventArgs e)
        {

            Form4 myPassForm = new Form4(this);       // Control de extención reset    
            if (button1.Text == "+")
            {
                myPassForm.ShowDialog();
                if (Procc)
                {
                    
                    label25.Visible = true;
                    this.Height = 724;
                    panel2.Top = 392;
                    groupBox3.Top = 392;
                    panel1.Top = 364;
                    groupBox1.Height = 298;
                    groupBox2.Height=298;
                    this.StartPosition = 0;
                    button1.Text = "-";
                    Procc = false;

                }
            }
            else            
            {
                label25.Visible = false;
                groupBox1.Height = 94;
                groupBox2.Height = 94;
                panel1.Top = 149;
                groupBox3.Top = 177;
                panel2.Top = 177;
                this.Height = 516;
                this.StartPosition = 0;
                button1.Text = "+";
            }

        }        
       
        private void button3_Click(object sender, EventArgs e)
        {
            this.Processar();
        }

        //Subproceso de cargado de excel para procesar datos preeliminares
        void CargarDatos()
        { 
            //Clase COferta no permanente
            COferta newOferta = myOferta;
            TrabExcel = new CExcelWork(myOferta, local, this);
            LEstado.Text = "Iniciando procesado...";
            TrabExcel.CreateApplication();
            LEstado.Text = "Cargando...";
            //TrabExcel.OpenFile(local, RUSD.Checked);
            TrabExcel.OpenFile(local, RCUC.Checked);
            TrabExcel.OpenFile(local, RGX.Checked);
            TrabExcel.OpenFile(local, RFTX.Checked);
            TrabExcel.OpenFile(local, RODC.Checked);
            //TrabExcel.ObtenerDatos();            
        }

        private void TiObtener_Tick(object sender, EventArgs e)
        {
            if (!Obtener.IsAlive)
            {
                TiObtener.Enabled = false;
                
                
                ProCamara = false;
                LEstado.Text = "Estado: en espera.";
            }
        }

        void Processar()
        {
            Obtener = new Thread(new ThreadStart(this.CargarDatos));
            Obtener.Start();
            TiObtener.Enabled = true;
        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            
            this.GuardarDatos();
            ProCamara = true;
            this.Processar();
        }

        void GuardarDatos()
        {
            if (myOferta != null)
            {
                if (CCamara.Text != "")
                {
                    if (int.Parse(CCamara.Text) > myOferta.GetCont())
                    {
                        MessageBox.Show("Error, ese # de cámara no existe.",
                            "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CCamara.Text = actualcam;
                    }
                    else
                    {
                        if (this.Validateit())
                        {
                            
                            CCam myCam = new CCam(TNC.Text, TTem.Text, TLargo.Text, TAncho.Text, TAlto.Text, TVolu.Text, TCF.Text, TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text,
                               TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text,
                               TCentx.Text, TCxp.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text,
                               TMuc.Text, TMevp.Text, TSol.Text, TValv.Text, TCvta.Text, TCuadro.Text, CBexpo.Text, CSumi.Text, Ctpd.Text, Coff3.Text, Cnoff6.Text,
                               CTEvap.Text, TCodValv.Text, TInc.Text, TInev.Text, TVnev.Text, TIned.Text, TIncd.Text, TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TTint.Text, TEquip.Text,
                               TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text, TDmce.Text, TLcc.Text, TLss.Text, TPcmc.Text,
                               TPcond.Text, TModex.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text,
                               TPnt.Text, TPml.Text, TCt150.Text, TCt04.Text, TCt150m.Text, TCt04m.Text, SPtp75.Text, SPtp74.Text, TTcmc1.Text, TTcmc2.Text,
                               TTcmc3.Text, TFWe.Text, TTcmc6.Text, TTcmc8.Text, TTcmc9.Text, CBps.Text, CTPup.Text, CTPus.Text, CTLup.Text, TLtemp.ToString(), 
                               TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString(), pk);

                            myOferta.actualizar(myCam, int.Parse(CCamara.Text));
                        }
                        else
                        {
                            MessageBox.Show("Error de datos",
                            "Todos los campos deben tener valor correctos", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                }
                else
                    MessageBox.Show("Error, cámara incorrecta",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }        

        
        
        void Calcular()// OPERACIONES DE CALCULO.
        {
            try
            {
               
                CCam newCam = new CCam(TNC.Text, TTem.Text, TLargo.Text, TAncho.Text, TAlto.Text, TVolu.Text, TCF.Text, TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text,
                    TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text,
                    TCentx.Text, TCxp.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text,
                    TMuc.Text, TMevp.Text, TSol.Text, TValv.Text, TCvta.Text, TCuadro.Text, CBexpo.Text, CSumi.Text, Ctpd.Text, Coff3.Text, Cnoff6.Text,
                    CTEvap.Text, TCodValv.Text, TInc.Text, TInev.Text, TVnev.Text, TIned.Text, TIncd.Text, TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TTint.Text, TEquip.Text,
                    TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text, TDmce.Text, TLcc.Text, TLss.Text, TPcmc.Text,
                    TPcond.Text, TModex.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text,
                    TPnt.Text, TPml.Text, TCt150.Text, TCt04.Text, TCt150m.Text, TCt04m.Text, SPtp75.Text, SPtp74.Text, TTcmc1.Text, TTcmc2.Text,
                    TTcmc3.Text, TFWe.Text, TTcmc6.Text, TTcmc8.Text, TTcmc9.Text, CBps.Text, CTPup.Text, CTPus.Text, CTLup.Text, TLtemp.ToString(), 
                    TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString(), pk);

                myOferta.actualizar(newCam, int.Parse(CCamara.Text));
                

            }
            catch
            {
                try
                {


                    
                   
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                                "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                
            }

            
           
            
        }
        

        private void calcbeging()
        {
            try
            {
                if (!ini)
                {
                    NewExcelApp = new Microsoft.Office.Interop.Excel.Application();
                    NewExcelApp.Visible = false;
                    NewExcelApp.UserControl = false;
                    NewExcelApp.DisplayAlerts = false;
                    CalcWorkBook = NewExcelApp.Workbooks.Open(local + @"\calc.tpt",
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing,
                            Type.Missing);
                    ini = true;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                                "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }


        }

        private void toolStripButton1_Click_1(object sender, EventArgs e)
        {
            if (myOferta != null)
            {
                if (CCamara.Text != "")
                {
                    if (int.Parse(CCamara.Text) > myOferta.GetCont())
                    {
                        MessageBox.Show("Error, ese # de cámara no existe.",
                            "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CCamara.Text = actualcam;
                    }
                    else
                    {
                        if (!(int.Parse(CCamara.Text) == 1))
                        {
                            this.Atras();
                            CCamara.Text = (int.Parse(actualcam) - 1).ToString();
                            this.asignarcam(int.Parse(actualcam));
                        }                
                    }
                }
                else
                    MessageBox.Show("Error, cámara incorrecta",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                MessageBox.Show("Error, Aun no existe una oferta y no hay cámaras.",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CCamara.Text = "";
            }
        }

        void Atras()
        {
            CCam AuxCam;
            AuxCam = myOferta.GetCam(int.Parse(CCamara.Text) - 2);
            myOferta.SetCam(myOferta.GetCam(int.Parse(CCamara.Text) - 1), int.Parse(CCamara.Text) - 2);
            myOferta.SetCam(AuxCam, int.Parse(CCamara.Text) - 1);                              
        }

        private void toolStripButton2_Click_1(object sender, EventArgs e)
        {
            if (myOferta != null)
            {
                if (CCamara.Text != "")
                {
                    if (int.Parse(CCamara.Text) > myOferta.GetCont())
                    {
                        MessageBox.Show("Error, ese # de cámara no existe.",
                            "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        CCamara.Text = actualcam;
                    }
                    else
                    {
                        if (!(int.Parse(CCamara.Text) == myOferta.GetCantCam()))
                        {
                            this.Adelante();
                            CCamara.Text = (int.Parse(actualcam) + 1).ToString();
                            this.asignarcam(int.Parse(actualcam));
                        }
                    }
                }
                else
                    MessageBox.Show("Error, cámara incorrecta",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);

            }
            else
            {
                MessageBox.Show("Error, Aun no existe una oferta y no hay cámaras.",
                    "Error al leer datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                CCamara.Text = "";
            }
        }

        void Adelante()
        {
            CCam AuxCam;
            AuxCam = myOferta.GetCam(int.Parse(CCamara.Text));
            myOferta.SetCam(myOferta.GetCam(int.Parse(CCamara.Text) - 1), int.Parse(CCamara.Text));
            myOferta.SetCam(AuxCam, int.Parse(CCamara.Text) - 1);
        }

        private void toolStripButton7_Click(object sender, EventArgs e)
        {
            pk++;
            if (pk > 6)
                pk = 0;
            pictureBox2.Image = myUserControl.IMAGES.Images[pk];

            
        }

        private void toolStripButton6_Click(object sender, EventArgs e)
        {
            pk--;
            if (pk < 0)
                pk = 6;
            pictureBox2.Image = myUserControl.IMAGES.Images[pk];

            
        }

        private void CSup_SelectedIndexChanged(object sender, EventArgs e)
        {
            
            //Convert to °C-bar-kW
            string Metric;
            string tons;
            int selected = 0;
            string sel11;
            string sel12;
            string sel13;
            string sel14;
            string sel15;
            decimal sel16;
            decimal sel17;
            decimal sel18;
            string sel19;
            decimal sel10;

            if (CSup.Text == "°C-bar-kW") { selected = 1; sel11 = "°C"; sel12 = "K"; sel13 = "kW"; sel14 = "bar"; sel15 = ""; sel16 = 1; sel17 = 1; sel18 = 1; sel19 = "bara"; sel10 = 0; }
            if (CSup.Text == "°F-Psi-tons") { selected = 2; sel11 = "°F"; sel12 = "°F"; sel13 = "kW"; sel14 = "tons"; sel15 = "Psid"; sel16 = 0.5556M; sel17 = 0.284M; sel18 = 14.5038M; sel19 = "bara"; sel10 = 14.7M; }
            if (CSup.Text == "°F-Psi-MBH") { selected = 2; sel11 = "°F"; sel12 = "°F"; sel13 = "kW"; sel14 = "MBH"; sel15 = "Psid"; sel16 = 0.5556M; sel17 = 3.141212M; sel18 = 14.5038M; sel19 = "bara"; sel10 = 14.7M; }
            Datos.selected = selected;
        }
       
        private void button5_Click(object sender, EventArgs e)
        {

            Form5 myPassForm = new Form5(this);// CONTROL BLOQUE DE CODIGOS EVAPORADOR
            myPassForm.ShowDialog();
        }
        private void button6_Click_1(object sender, EventArgs e)
        {
            Form6 myPassForm = new Form6(this);// CONTROL BLOQUE DE CODIGOS EXPANSION
            myPassForm.ShowDialog();
        }
       
        private void btnGenerarPDF_Click(object sender, EventArgs e)
        {

            string pdfPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "archivo.pdf");
            Process.Start(pdfPath);

        }
        
        private void Carga_Click(object sender, EventArgs e)
        {
          
            //Convert to °C-bar-kW
            string Metric;
            string tons;
            int selected = 0;
            string sel11;
            string sel12;
            string sel13;
            string sel14;
            string sel15;
            decimal sel16;
            decimal sel17;
            decimal sel18;
            string sel19;
            decimal sel10;
           
            
            if (CSup.Text == "°C-bar-kW") { selected = 1; sel11 = "°C"; sel12 = "K"; sel13 = "kW"; sel14 = "bar"; sel15 = ""; sel16 = 1; sel17 = 1; sel18 = 1; sel19 = "bara"; sel10 = 0; }
            if (CSup.Text == "°F-Psi-tons") { selected = 2; sel11 = "°F"; sel12 = "°F"; sel13 = "kW"; sel14 = "tons"; sel15 = "Psid"; sel16 = 0.5556M; sel17 = 0.284M; sel18 = 14.5038M; sel19 = "bara"; sel10 = 14.7M; }
            if (CSup.Text == "°F-Psi-MBH") { selected = 2; sel11 = "°F"; sel12 = "°F"; sel13 = "kW"; sel14 = "MBH"; sel15 = "Psid"; sel16 = 0.5556M; sel17 = 3.141212M; sel18 = 14.5038M; sel19 = "bara"; sel10 = 14.7M; }
            Datos.selected = selected;

            // 3.2 Cálculo del Diferencial de temperatura.
            int TpR;//Tipo producto
            int Npr;//Numero de producto
            int Tip;//Temperatura del producto
            int Tra1;//Temperatura minima producto
            int Tra2;//Temperatura maxima del rpoducto
            int Hur;//Humedad relativa del producto
            decimal Pdc;//Punto de congelación
            decimal CcL;//Calor esp. antes cong. Kc/Kg/ºC
            decimal CcF;//Calor esp. después cong. Kc/Kg/ºC
            decimal CeL;//Calor Latente Kcal/Kg
            int CrP;//Calor Respiración Kcal/Tm/24h
            int CrA;//Calor 1 Respiración Kcal/Tm/24h
            int Crc;//Calor 2 Respiración Kcal/Tm/24h
            int DC;//Densidad Carga Kg/m³
            // Temperatura del producto
            int CbEx;//Tipo de Proceso
            int Hp14;//Hotario para Cámara fría con Piso 14H
            int Hp12;//Hortario para Cámara fría con piso 12H
            decimal FcG;// Factor de carga para tuneles
            decimal FcGC;//Factor de carga en congelación
            decimal FcGR;//Factor de carga en refrigeración
            double Hurp;
            double Hur1;
            double Hur2;
            Hurp = 0;
            Hur1 = 0;
            Hur2 = 0;
            Hp14 = 0;
            Hp12 = 0;
            FcG = 0;
            FcGC = 0;
            FcGR = 0;
            if (CBexpo.Text == "CAMARA FRIA C/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA FRIA S/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA INDT C/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA INDT C/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA INDT S/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA MOD C/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "CAMARA MOD S/P") { Hp14 = 14; Hp12 = 12; FcG = 0M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "TUNEL CONG.") { Hp14 = 12; Hp12 = 8; FcG = 0.1M; FcGC = 0.14M; FcGR = 0.12M; }
            if (CBexpo.Text == "ABATIDOR TEMP") { Hp14 = 8; Hp12 = 2; FcG = 0.1M; FcGC = 0.14M; FcGR = 0.12M; }
            decimal TipN;//Valor de congelación para tuneles
            TipN = 0;
            Tip = 0;
            Npr = 0;
            Tra1 = 0;
            Tra2 = 0;
            Pdc = 0;
            CrP = 0;
            CrA = 0;
            Crc = 0;
            DC = 0;
            CcL = 0;
            CeL = 0;
            CcF = 0;
            if (Ctpd.Text == "Sala Obrador general") { Npr = 2; Tip = 16; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.7M; CcF = 0M; CeL = 0M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sala Despiece general") { Npr = 3; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.7M; CcF = 0M; CeL = 0M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sala Despiece Fría") { Npr = 4; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.7M; CcF = 0M; CeL = 0M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sala Embalaje general") { Npr = 5; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.4M; CcF = 0M; CeL = 0M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Frescos general") { Npr = 6; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Varios") { Npr = 7; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Carne Cerdo general") { Npr = 8; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.65M; CcF = 0.36M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Carne Cordero general") { Npr = 9; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.72M; CcF = 0.39M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Carne Vacuno general") { Npr = 10; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 88; Hur2 = 92; Pdc = -2.2M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Carne general") { Npr = 11; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.75M; CcF = 0.4M; CeL = 52M; CrP = 0; CrA = 0; Crc = 0; DC = 280; }
            if (Ctpd.Text == "Frutas general") { Npr = 12; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 0.9M; CcF = 0.48M; CeL = 68M; CrP = 300 - 2800; CrA = 300; Crc = 2800; DC = 320; }
            if (Ctpd.Text == "Verduras general") { Npr = 13; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.5M; CcL = 0.96M; CcF = 0.48M; CeL = 76M; CrP = 560 - 3900; CrA = 560; Crc = 3900; DC = 220; }
            if (Ctpd.Text == "Lacteos y quesos") { Npr = 14; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -0.5M; CcL = 0.8M; CcF = 0.41M; CeL = 76M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Pescado Fresco") { Npr = 15; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Pescado Congelado") { Npr = 16; Tip = -25; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Helados crema") { Npr = 17; Tip = -28; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -5M; CcL = 0.7M; CcF = 0.39M; CeL = 49M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Congelados general") { Npr = 18; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 0.88M; CcF = 0.48M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Tunel Congelación general") { Npr = 19; Tip = -40; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = 0M; CcL = 0.88M; CcF = 0.48M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Aceite oliva") { Npr = 20; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = 0M; CcL = 0.48M; CcF = 0.35M; CeL = 12M; CrP = 120; CrA = 120; Crc = 120; DC = 550; }
            if (Ctpd.Text == "Aceite salado") { Npr = 21; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = 0M; CcL = 0.48M; CcF = 0.35M; CeL = 12M; CrP = 160; CrA = 160; Crc = 160; DC = 550; }
            if (Ctpd.Text == "Aceite vegetal") { Npr = 22; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = 0M; CcL = 0.48M; CcF = 0.35M; CeL = 12M; CrP = 160; CrA = 160; Crc = 160; DC = 550; }
            if (Ctpd.Text == "Aceitunas frescas") { Npr = 23; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.4M; CcL = 0.8M; CcF = 0.43M; CeL = 60M; CrP = 100 - 3000; CrA = 100; Crc = 3000; DC = 340; }
            if (Ctpd.Text == "Acelgas") { Npr = 24; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.89M; CcF = 0.47M; CeL = 70M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 250; }
            if (Ctpd.Text == "Agua") { Npr = 25; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 1M; CcF = 0.5M; CeL = 80M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Agua fría") { Npr = 26; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 1M; CcF = 0.5M; CeL = 80M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Agua hielo") { Npr = 27; Tip = -10; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 1M; CcF = 0.5M; CeL = 80M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Aguacates") { Npr = 28; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.7M; CcL = 0.91M; CcF = 0.46M; CeL = 62M; CrP = 1200 - 6000; CrA = 1200; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Ajos secos") { Npr = 29; Tip = -3; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -4M; CcL = 0.69M; CcF = 0.4M; CeL = 50M; CrP = 200 - 2000; CrA = 200; Crc = 2000; DC = 250; }
            if (Ctpd.Text == "Albaricoques") { Npr = 30; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.91M; CcF = 0.48M; CeL = 67.9M; CrP = 160 - 2200; CrA = 160; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Alcachofas") { Npr = 31; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.7M; CcL = 0.86M; CcF = 0.45M; CeL = 66M; CrP = 370 - 3500; CrA = 370; Crc = 3500; DC = 250; }
            if (Ctpd.Text == "Alcachofas congeladas") { Npr = 32; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.7M; CcL = 0.86M; CcF = 0.45M; CeL = 66M; CrP = 0 - 370; CrA = 0; Crc = 370; DC = 250; }
            if (Ctpd.Text == "Alcachofas globo") { Npr = 33; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -1.2M; CcL = 0.87M; CcF = 0.45M; CeL = 69M; CrP = 280 - 3200; CrA = 280; Crc = 3200; DC = 250; }
            if (Ctpd.Text == "Alcachofas Jerusalén") { Npr = 34; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.1M; CcL = 0.83M; CcF = 0.44M; CeL = 66M; CrP = 370 - 3200; CrA = 370; Crc = 3200; DC = 250; }
            if (Ctpd.Text == "Alfalfa") { Npr = 35; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 0.82M; CcF = 0.51M; CeL = 68M; CrP = 300 - 7500; CrA = 300; Crc = 7500; DC = 350; }
            if (Ctpd.Text == "Alfalfa Congelada") { Npr = 36; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = 0M; CcL = 0.82M; CcF = 0.51M; CeL = 68M; CrP = 0 - 300; CrA = 0; Crc = 300; DC = 400; }
            if (Ctpd.Text == "Almeja conserva") { Npr = 37; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Almeja entera") { Npr = 38; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.8M; CcL = 0.85M; CcF = 0.44M; CeL = 64M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Almeja fresca") { Npr = 39; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Almejas (carne-liquido)") { Npr = 40; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 98; Hur2 = 100; Pdc = -2.2M; CcL = 0.9M; CcF = 0.47M; CeL = 70M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Almíbar") { Npr = 41; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 67; Hur2 = 72; Pdc = 0M; CcL = 0.49M; CcF = 0.31M; CeL = 29M; CrP = 10 - 200; CrA = 10; Crc = 200; DC = 550; }
            if (Ctpd.Text == "Anguilas") { Npr = 42; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.7M; CcF = 0.39M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Apio") { Npr = 43; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 98; Pdc = -1.3M; CcL = 0.95M; CcF = 0.48M; CeL = 75M; CrP = 270 - 5000; CrA = 270; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Arandanos") { Npr = 44; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.9M; CcF = 0.46M; CeL = 69.3M; CrP = 110 - 2800; CrA = 110; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Arbustos") { Npr = 45; Tip = -2; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = -4M; CcL = 0.82M; CcF = 0.35M; CeL = 58M; CrP = 200 - 2000; CrA = 200; Crc = 2000; DC = 220; }
            if (Ctpd.Text == "Arengue ahumado") { Npr = 46; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 90; Pdc = -2.2M; CcL = 0.72M; CcF = 0.4M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Arengue en cecina") { Npr = 47; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 90; Pdc = -2.2M; CcL = 0.7M; CcF = 0.39M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Atún") { Npr = 48; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.77M; CcF = 0.41M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Aves congeladas") { Npr = 49; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.8M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves de corral") { Npr = 50; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.8M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves de corral congeladas") { Npr = 51; Tip = -23; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.8M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves frescas promedio") { Npr = 52; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.8M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves Pato") { Npr = 53; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.7M; CcL = 0.76M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves Pavo todo tipo") { Npr = 54; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.7M; CcL = 0.72M; CcF = 0.4M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Aves Pollo todo tipo") { Npr = 55; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.8M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 340; }
            if (Ctpd.Text == "Azúcar") { Npr = 56; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = 0M; CcL = 0.35M; CcF = 0.3M; CeL = 4M; CrP = 390 - 2200; CrA = 390; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Azúcar de meple cons corta") { Npr = 57; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 67; Hur2 = 72; Pdc = 0M; CcL = 0.24M; CcF = 0.21M; CeL = 4M; CrP = 120 - 2200; CrA = 120; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Azúcar de meple cons larga") { Npr = 58; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 67; Hur2 = 72; Pdc = 0M; CcL = 0.24M; CcF = 0.21M; CeL = 4M; CrP = 390 - 2200; CrA = 390; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Bacalao") { Npr = 59; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.85M; CcF = 0.45M; CeL = 65M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Bacalao salado") { Npr = 60; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -2.8M; CcL = 0.68M; CcF = 0.4M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Bacón") { Npr = 61; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.36M; CcF = 0.26M; CeL = 16M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Barbada") { Npr = 62; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.85M; CcF = 0.45M; CeL = 65M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Basura") { Npr = 63; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -1M; CcL = 0.7M; CcF = 0.4M; CeL = 45M; CrP = 200 - 2000; CrA = 200; Crc = 2000; DC = 300; }
            if (Ctpd.Text == "Batata") { Npr = 64; Tip = 13; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -1.5M; CcL = 0.8M; CcF = 0.43M; CeL = 59M; CrP = 270 - 3500; CrA = 270; Crc = 3500; DC = 350; }
            if (Ctpd.Text == "Bayas") { Npr = 65; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 82; Hur2 = 85; Pdc = -2M; CcL = 0.9M; CcF = 0.49M; CeL = 66.6M; CrP = 160 - 6000; CrA = 160; Crc = 6000; DC = 220; }
            if (Ctpd.Text == "Bebidas") { Npr = 66; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = -2.2M; CcL = 0.96M; CcF = 0.48M; CeL = 76M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Bebidas frías") { Npr = 67; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = -2.2M; CcL = 0.96M; CcF = 0.48M; CeL = 76M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Berenjena") { Npr = 68; Tip = 8; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.95M; CcF = 0.48M; CeL = 79M; CrP = 270 - 4000; CrA = 270; Crc = 4000; DC = 250; }
            if (Ctpd.Text == "Berrazas") { Npr = 69; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 91; Hur2 = 95; Pdc = -1.6M; CcL = 0.86M; CcF = 0.44M; CeL = 66M; CrP = 300 - 5000; CrA = 300; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Berros") { Npr = 70; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 95; Pdc = -0.3M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 400 - 6000; CrA = 400; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Berzas") { Npr = 71; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 95; Pdc = -0.5M; CcL = 0.93M; CcF = 0.47M; CeL = 73.2M; CrP = 270 - 2300; CrA = 270; Crc = 2300; DC = 250; }
            if (Ctpd.Text == "Bombones") { Npr = 72; Tip = 15; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 55; Pdc = 0M; CcL = 0.56M; CcF = 0.28M; CeL = 15M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Boniatos") { Npr = 73; Tip = 13; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.5M; CcL = 0.8M; CcF = 0.43M; CeL = 59M; CrP = 270 - 3500; CrA = 270; Crc = 3500; DC = 350; }
            if (Ctpd.Text == "Brécol") { Npr = 74; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 88; Hur2 = 90; Pdc = -1.5M; CcL = 0.9M; CcF = 0.48M; CeL = 74.9M; CrP = 250 - 5500; CrA = 250; Crc = 5500; DC = 250; }
            if (Ctpd.Text == "Bróculi") { Npr = 75; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.92M; CcF = 0.47M; CeL = 72M; CrP = 220 - 6800; CrA = 220; Crc = 6800; DC = 250; }
            if (Ctpd.Text == "Buey (graso)") { Npr = 76; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.6M; CcF = 0.35M; CeL = 44M; CrP = 0; CrA = 0; Crc = 0; DC = 320; }
            if (Ctpd.Text == "Buey (magro)") { Npr = 77; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.7M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 320; }
            if (Ctpd.Text == "Buey (Seca)") { Npr = 78; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 61; Hur2 = 65; Pdc = 0M; CcL = 0.33M; CcF = 0.25M; CeL = 12M; CrP = 0; CrA = 0; Crc = 0; DC = 320; }
            if (Ctpd.Text == "Buey Congelado") { Npr = 79; Tip = -23; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Buey de mar") { Npr = 80; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.9M; CcF = 0.45M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Caballa") { Npr = 81; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.72M; CcF = 0.4M; CeL = 52M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Cacahuetes") { Npr = 82; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 45; Pdc = -2M; CcL = 0.22M; CcF = 0.21M; CeL = 2M; CrP = 10 - 1000; CrA = 10; Crc = 1000; DC = 250; }
            if (Ctpd.Text == "Cacao") { Npr = 83; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 65; Pdc = 0M; CcL = 0.35M; CcF = 0.28M; CeL = 16M; CrP = 100 - 900; CrA = 100; Crc = 900; DC = 200; }
            if (Ctpd.Text == "Café Verde") { Npr = 84; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = 0M; CcL = 0.32M; CcF = 0.25M; CeL = 12M; CrP = 250 - 1800; CrA = 250; Crc = 1800; DC = 250; }
            if (Ctpd.Text == "Calabaza") { Npr = 85; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -0.3M; CcL = 0.9M; CcF = 0.46M; CeL = 71M; CrP = 500 - 4000; CrA = 500; Crc = 4000; DC = 300; }
            if (Ctpd.Text == "Calabaza de bellota") { Npr = 86; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -0.3M; CcL = 0.9M; CcF = 0.46M; CeL = 71M; CrP = 500 - 4000; CrA = 500; Crc = 4000; DC = 300; }
            if (Ctpd.Text == "Calabaza de invierno") { Npr = 87; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -0.6M; CcL = 0.95M; CcF = 0.46M; CeL = 71M; CrP = 380 - 3800; CrA = 380; Crc = 3800; DC = 300; }
            if (Ctpd.Text == "Calabaza de verano") { Npr = 88; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 95; Pdc = -0.6M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 350 - 4500; CrA = 350; Crc = 4500; DC = 300; }
            if (Ctpd.Text == "Camarón") { Npr = 89; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.81M; CcF = 0.43M; CeL = 61M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Cangrejo") { Npr = 90; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.9M; CcF = 0.45M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Cantuloupes (Tipo melón)") { Npr = 91; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 95; Pdc = -1.2M; CcL = 0.94M; CcF = 0.48M; CeL = 74M; CrP = 400 - 5000; CrA = 400; Crc = 5000; DC = 300; }
            if (Ctpd.Text == "Caquis") { Npr = 92; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 87; Hur2 = 90; Pdc = -2.2M; CcL = 0.82M; CcF = 0.43M; CeL = 62.1M; CrP = 100 - 4000; CrA = 100; Crc = 4000; DC = 250; }
            if (Ctpd.Text == "Carne concha congelado") { Npr = 93; Tip = -25; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.84M; CcF = 0.44M; CeL = 64M; CrP = 0; CrA = 0; Crc = 0; DC = 500; }
            if (Ctpd.Text == "Carne de Concha fresco") { Npr = 94; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.84M; CcF = 0.44M; CeL = 64M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Castañas cons corta") { Npr = 95; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = 0M; CcL = 0.45M; CcF = 0.35M; CeL = 12M; CrP = 500 - 2200; CrA = 500; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Castañas cons larga") { Npr = 96; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = 0M; CcL = 0.45M; CcF = 0.35M; CeL = 12M; CrP = 350 - 2200; CrA = 350; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Caviar (en cubetas) cons corta") { Npr = 97; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -6.5M; CcL = 0.7M; CcF = 0.31M; CeL = 50M; CrP = 250 - 1200; CrA = 250; Crc = 1200; DC = 250; }
            if (Ctpd.Text == "Caviar (en cubetas) cons larga") { Npr = 98; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -6.5M; CcL = 0.7M; CcF = 0.31M; CeL = 50M; CrP = 160 - 1200; CrA = 160; Crc = 1200; DC = 250; }
            if (Ctpd.Text == "Cebollas") { Npr = 99; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -1M; CcL = 0.91M; CcF = 0.46M; CeL = 68.8M; CrP = 220 - 3600; CrA = 220; Crc = 3600; DC = 350; }
            if (Ctpd.Text == "Cebollas verdes") { Npr = 100; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -0.9M; CcL = 0.92M; CcF = 0.47M; CeL = 68.8M; CrP = 250 - 3800; CrA = 250; Crc = 3800; DC = 350; }
            if (Ctpd.Text == "Cerdo congelado") { Npr = 101; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.65M; CcF = 0.36M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Cerdo en canal magro 47%") { Npr = 102; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.54M; CcF = 0.33M; CeL = 42M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Cerdo flanco magro 35%") { Npr = 103; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2.2M; CcL = 0.45M; CcF = 0.3M; CeL = 40M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Cerdo fresco promedio") { Npr = 104; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.65M; CcF = 0.36M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Cerdo paletilla magro 67%") { Npr = 105; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 79; Hur2 = 85; Pdc = -2.2M; CcL = 0.6M; CcF = 0.36M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Cerdo endurecido corte") { Npr = 106; Tip = -5; Tra1 = 14; Tra2 = 18; Hur1 = 86; Hur2 = 90; Pdc = -2.2M; CcL = 0.6M; CcF = 0.36M; CeL = 10M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Cerezas agrias") { Npr = 107; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.7M; CcL = 0.89M; CcF = 0.46M; CeL = 70M; CrP = 250 - 3000; CrA = 250; Crc = 3000; DC = 250; }
            if (Ctpd.Text == "Cerezas congeladas") { Npr = 108; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.87M; CcF = 0.45M; CeL = 68M; CrP = 0 - 200; CrA = 0; Crc = 200; DC = 250; }
            if (Ctpd.Text == "Cerezas dulces") { Npr = 109; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.8M; CcL = 0.87M; CcF = 0.45M; CeL = 66.9M; CrP = 200 - 2500; CrA = 200; Crc = 2500; DC = 250; }
            if (Ctpd.Text == "Cerveza botellas. botes") { Npr = 110; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = -2.2M; CcL = 0.92M; CcF = 0.47M; CeL = 72M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Cerveza congelada") { Npr = 111; Tip = -25; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2.2M; CcL = 0.92M; CcF = 0.44M; CeL = 72M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Cerveza madera") { Npr = 112; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2.2M; CcL = 0.92M; CcF = 0.44M; CeL = 72M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Cerveza metálico") { Npr = 113; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -2.2M; CcL = 0.92M; CcF = 0.44M; CeL = 72M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Champiñón") { Npr = 114; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.93M; CcF = 0.48M; CeL = 72M; CrP = 200 - 2800; CrA = 200; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Chirivía") { Npr = 115; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -0.9M; CcL = 0.84M; CcF = 0.44M; CeL = 62M; CrP = 230 - 4800; CrA = 230; Crc = 4800; DC = 250; }
            if (Ctpd.Text == "Chocolate (Dulce de)") { Npr = 116; Tip = 15; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = 0M; CcL = 0.28M; CcF = 0.23M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Chocolate con leche") { Npr = 117; Tip = 15; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 55; Pdc = 0M; CcL = 0.21M; CcF = 0.21M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Chocolate líquido") { Npr = 118; Tip = 15; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 55; Pdc = 0M; CcL = 0.56M; CcF = 0.3M; CeL = 22M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Ciruelas") { Npr = 119; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.89M; CcF = 0.46M; CeL = 68.5M; CrP = 200 - 6000; CrA = 200; Crc = 6000; DC = 300; }
            if (Ctpd.Text == "Ciruelas congeladas") { Npr = 120; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.46M; CeL = 68.5M; CrP = 0 - 50; CrA = 0; Crc = 50; DC = 350; }
            if (Ctpd.Text == "Ciruelas seca") { Npr = 121; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = -2M; CcL = 0.42M; CcF = 0.28M; CeL = 22.9M; CrP = 50 - 2000; CrA = 50; Crc = 2000; DC = 180; }
            if (Ctpd.Text == "Coco") { Npr = 122; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.9M; CcL = 0.58M; CcF = 0.34M; CeL = 38M; CrP = 150 - 1600; CrA = 150; Crc = 1600; DC = 250; }
            if (Ctpd.Text == "Col") { Npr = 123; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.91M; CcF = 0.46M; CeL = 72M; CrP = 210 - 2200; CrA = 210; Crc = 2200; DC = 200; }
            if (Ctpd.Text == "Col tardía") { Npr = 124; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 98; Hur2 = 10; Pdc = -0.9M; CcL = 0.94M; CcF = 0.48M; CeL = 74M; CrP = 220 - 2600; CrA = 220; Crc = 2600; DC = 200; }
            if (Ctpd.Text == "Coles de Bruselas") { Npr = 125; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.8M; CcL = 0.91M; CcF = 0.46M; CeL = 70M; CrP = 280 - 3500; CrA = 280; Crc = 3500; DC = 200; }
            if (Ctpd.Text == "Coliflor") { Npr = 126; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.94M; CcF = 0.47M; CeL = 77M; CrP = 280 - 3800; CrA = 280; Crc = 3800; DC = 200; }
            if (Ctpd.Text == "Colirrábano") { Npr = 127; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 95; Pdc = -1M; CcL = 0.92M; CcF = 0.47M; CeL = 74M; CrP = 250 - 2800; CrA = 250; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Collars") { Npr = 128; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 91; Hur2 = 95; Pdc = -0.8M; CcL = 0.89M; CcF = 0.46M; CeL = 74M; CrP = 230 - 2500; CrA = 230; Crc = 2500; DC = 200; }
            if (Ctpd.Text == "Compost") { Npr = 129; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -4M; CcL = 0.88M; CcF = 0.45M; CeL = 38M; CrP = 400 - 6800; CrA = 400; Crc = 6800; DC = 200; }
            if (Ctpd.Text == "Compost germinando") { Npr = 130; Tip = 16; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = 0M; CcL = 0.88M; CcF = 0.45M; CeL = 38M; CrP = 5000 - 6800; CrA = 5000; Crc = 6800; DC = 180; }
            if (Ctpd.Text == "Conejo congelado") { Npr = 131; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.74M; CcF = 0.4M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Conejo fresco promedio") { Npr = 132; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.74M; CcF = 0.4M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Conservas en bote") { Npr = 133; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -0.9M; CcL = 0.88M; CcF = 0.49M; CeL = 70M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Cordero congelado") { Npr = 134; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.72M; CcF = 0.39M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Cordero fresco promedio") { Npr = 135; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.72M; CcF = 0.39M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Cordero pierna magra 83%") { Npr = 136; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.72M; CcF = 0.39M; CeL = 51.8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Cordero selecto magro 67%") { Npr = 137; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.72M; CcF = 0.39M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Crustáceos") { Npr = 138; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.9M; CcF = 0.45M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Crustáceos congelados") { Npr = 139; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.9M; CcF = 0.45M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Dátiles curados") { Npr = 140; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = -15.7M; CcL = 0.36M; CcF = 0.26M; CeL = 38.2M; CrP = 100 - 1500; CrA = 100; Crc = 1500; DC = 250; }
            if (Ctpd.Text == "Dátiles Frescos") { Npr = 141; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.4M; CcL = 0.82M; CcF = 0.43M; CeL = 62.1M; CrP = 220 - 2800; CrA = 220; Crc = 2800; DC = 250; }
            if (Ctpd.Text == "Dátiles secos") { Npr = 142; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -15.5M; CcL = 0.36M; CcF = 0.26M; CeL = 16M; CrP = 50 - 200; CrA = 50; Crc = 200; DC = 250; }
            if (Ctpd.Text == "Dátiles secos congelados") { Npr = 143; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -15.5M; CcL = 0.36M; CcF = 0.26M; CeL = 16M; CrP = 0 - 50; CrA = 0; Crc = 50; DC = 300; }
            if (Ctpd.Text == "Despojos congelados") { Npr = 144; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.7M; CcL = 0.76M; CcF = 0.41M; CeL = 55.6M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Despojos frescos promedio") { Npr = 145; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.7M; CcL = 0.76M; CcF = 0.41M; CeL = 55.6M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Despojos tripas saladas") { Npr = 146; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.6M; CcF = 0.32M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Dewberries") { Npr = 147; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.88M; CcF = 0.46M; CeL = 67.9M; CrP = 600 - 9000; CrA = 600; Crc = 9000; DC = 250; }
            if (Ctpd.Text == "Embutidos") { Npr = 148; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -1M; CcL = 0.6M; CcF = 0.35M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Embutidos congelados") { Npr = 149; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.6M; CcF = 0.35M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Endivia") { Npr = 150; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -0.5M; CcL = 0.94M; CcF = 0.48M; CeL = 75M; CrP = 270 - 5500; CrA = 270; Crc = 5500; DC = 200; }
            if (Ctpd.Text == "Escarola") { Npr = 151; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -0.5M; CcL = 0.94M; CcF = 0.48M; CeL = 74M; CrP = 270 - 5500; CrA = 270; Crc = 5500; DC = 200; }
            if (Ctpd.Text == "Espárragos") { Npr = 152; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.95M; CcF = 0.48M; CeL = 75M; CrP = 350 - 2800; CrA = 350; Crc = 2800; DC = 220; }
            if (Ctpd.Text == "Espárragos congelados") { Npr = 153; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.95M; CcF = 0.48M; CeL = 75M; CrP = 0 - 350; CrA = 0; Crc = 350; DC = 280; }
            if (Ctpd.Text == "Espinacas") { Npr = 154; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 96; Pdc = -0.9M; CcL = 0.94M; CcF = 0.48M; CeL = 73M; CrP = 330 - 2500; CrA = 330; Crc = 2500; DC = 220; }
            if (Ctpd.Text == "Espinacas congeladas") { Npr = 155; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.94M; CcF = 0.48M; CeL = 73M; CrP = 0 - 330; CrA = 0; Crc = 330; DC = 250; }
            if (Ctpd.Text == "Flores cortadas congeladas") { Npr = 156; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.92M; CcF = 0.35M; CeL = 58M; CrP = 0 - 200; CrA = 0; Crc = 200; DC = 180; }
            if (Ctpd.Text == "Flores cortadas general") { Npr = 157; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.92M; CcF = 0.35M; CeL = 58M; CrP = 200 - 2800; CrA = 200; Crc = 2800; DC = 180; }
            if (Ctpd.Text == "Flores orquídeas. gardenias") { Npr = 158; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.5M; CcL = 0.92M; CcF = 0.35M; CeL = 58M; CrP = 200 - 2000; CrA = 200; Crc = 2000; DC = 180; }
            if (Ctpd.Text == "Frambuesa negra") { Npr = 159; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.85M; CcF = 0.44M; CeL = 64.5M; CrP = 1200 - 8000; CrA = 1200; Crc = 8000; DC = 250; }
            if (Ctpd.Text == "Frambuesa roja") { Npr = 160; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.87M; CcF = 0.45M; CeL = 66.9M; CrP = 1200 - 9000; CrA = 1200; Crc = 9000; DC = 250; }
            if (Ctpd.Text == "Fresa") { Npr = 161; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.92M; CcF = 0.48M; CeL = 71.7M; CrP = 900 - 7000; CrA = 900; Crc = 7000; DC = 250; }
            if (Ctpd.Text == "Fresas congelados") { Npr = 162; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.92M; CcF = 0.48M; CeL = 71.7M; CrP = 0 - 900; CrA = 0; Crc = 900; DC = 300; }
            if (Ctpd.Text == "Frutos secos") { Npr = 163; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 8M; CrP = 50 - 600; CrA = 50; Crc = 600; DC = 250; }
            if (Ctpd.Text == "Frutos frescos") { Npr = 164; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.81M; CcF = 0.43M; CeL = 61M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Gamba") { Npr = 165; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -3M; CcL = 0.86M; CcF = 0.44M; CeL = 65.5M; CrP = 50 - 2000; CrA = 50; Crc = 2000; DC = 250; }
            if (Ctpd.Text == "Granadas") { Npr = 166; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.91M; CcF = 0.47M; CeL = 71M; CrP = 400 - 6000; CrA = 400; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Grosellas") { Npr = 167; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 90; Pdc = -0.5M; CcL = 0.86M; CcF = 0.45M; CeL = 66.2M; CrP = 200 - 5000; CrA = 200; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Guayaba") { Npr = 168; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 66; Hur2 = 70; Pdc = 0M; CcL = 0.3M; CcF = 0.28M; CeL = 58.8M; CrP = 50 - 500; CrA = 50; Crc = 500; DC = 180; }
            if (Ctpd.Text == "Guisantes secos") { Npr = 169; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.6M; CcL = 0.79M; CcF = 0.42M; CeL = 58.8M; CrP = 240 - 2800; CrA = 240; Crc = 2800; DC = 250; }
            if (Ctpd.Text == "Guisantes verdes") { Npr = 170; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.79M; CcF = 0.42M; CeL = 58.8M; CrP = 0 - 50; CrA = 0; Crc = 50; DC = 300; }
            if (Ctpd.Text == "Guisantes verdes congelados") { Npr = 171; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.7M; CcL = 0.86M; CcF = 0.45M; CeL = 68M; CrP = 240 - 2800; CrA = 240; Crc = 2800; DC = 250; }
            if (Ctpd.Text == "Habas") { Npr = 172; Tip = 9; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = 0M; CcL = 0.29M; CcF = 0.24M; CeL = 10M; CrP = 50 - 200; CrA = 50; Crc = 200; DC = 180; }
            if (Ctpd.Text == "Habas secas") { Npr = 173; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.81M; CcF = 0.43M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Halibut") { Npr = 174; Tip = 22; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = 0M; CcL = 0.38M; CcF = 0.28M; CeL = 6M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Harina") { Npr = 175; Tip = -29; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -5.6M; CcL = 0.72M; CcF = 0.39M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Helados 10% Grasa") { Npr = 176; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.5M; CcL = 0.91M; CcF = 0.49M; CeL = 75M; CrP = 0; CrA = 0; Crc = 0; DC = 600; }
            if (Ctpd.Text == "Helados agua") { Npr = 177; Tip = -28; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -5M; CcL = 0.7M; CcF = 0.39M; CeL = 49M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Helados crema") { Npr = 178; Tip = -10; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = 0M; CcL = 1M; CcF = 0.5M; CeL = 80M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Hielo") { Npr = 179; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.4M; CcL = 0.82M; CcF = 0.43M; CeL = 62.1M; CrP = 220 - 2800; CrA = 220; Crc = 2800; DC = 250; }
            if (Ctpd.Text == "Higos frescos") { Npr = 180; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 60; Pdc = -2M; CcL = 0.39M; CcF = 0.27M; CeL = 19.2M; CrP = 50 - 200; CrA = 50; Crc = 200; DC = 200; }
            if (Ctpd.Text == "Higos secos") { Npr = 181; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.73M; CcF = 0.4M; CeL = 53M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Huevos con cascara") { Npr = 182; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.9M; CcF = 0.46M; CeL = 70M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Huevos congelados clara") { Npr = 183; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.79M; CcF = 0.42M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Huevos congelados enteros") { Npr = 184; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.64M; CcF = 0.37M; CeL = 44M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Huevos congelados yema") { Npr = 185; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -1M; CcL = 0.24M; CcF = 0.21M; CeL = 3.2M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Huevos duros de yema") { Npr = 186; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -1M; CcL = 0.23M; CcF = 0.21M; CeL = 2.5M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Huevos duros enteros") { Npr = 187; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -2M; CcL = 0.73M; CcF = 0.39M; CeL = 53M; CrP = 0; CrA = 0; Crc = 0; DC = 270; }
            if (Ctpd.Text == "Huevos frescos granja ") { Npr = 188; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.76M; CcF = 0.42M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Huevos líquido") { Npr = 189; Tip = 22; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 50; Pdc = -1M; CcL = 0.31M; CcF = 0.24M; CeL = 12M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Huevos sólidos Albúmina ") { Npr = 190; Tip = 22; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 50; Pdc = -1M; CcL = 0.25M; CcF = 0.22M; CeL = 5.5M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Huevos sólidos albúmina pulveriz.") { Npr = 191; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.5M; CcL = 0.68M; CcF = 0.38M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Jamón congelado") { Npr = 192; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = 0M; CcL = 0.54M; CcF = 0.38M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 270; }
            if (Ctpd.Text == "Jamón estilo campesino") { Npr = 193; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = 0M; CcL = 0.68M; CcF = 0.38M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 270; }
            if (Ctpd.Text == "Jamón poco curado") { Npr = 194; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 63; Hur2 = 65; Pdc = 0M; CcL = 0.6M; CcF = 0.38M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Jamón y lomo ahumado") { Npr = 195; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2.5M; CcL = 0.68M; CcF = 0.38M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Jamón y lomo fresco") { Npr = 196; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.47M; CcF = 0.3M; CeL = 27M; CrP = 0; CrA = 0; Crc = 0; DC = 500; }
            if (Ctpd.Text == "Jarabes") { Npr = 197; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.8M; CcL = 0.74M; CcF = 0.43M; CeL = 56M; CrP = 450 - 3000; CrA = 450; Crc = 3000; DC = 220; }
            if (Ctpd.Text == "Judías de lima") { Npr = 198; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 72; Hur2 = 70; Pdc = 0M; CcL = 0.3M; CcF = 0.24M; CeL = 18M; CrP = 50 - 300; CrA = 50; Crc = 300; DC = 180; }
            if (Ctpd.Text == "Judías secas") { Npr = 199; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.2M; CcL = 0.92M; CcF = 0.47M; CeL = 66M; CrP = 530 - 3200; CrA = 530; Crc = 3200; DC = 220; }
            if (Ctpd.Text == "Judías verdes cons corta") { Npr = 200; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.2M; CcL = 0.89M; CcF = 0.47M; CeL = 66M; CrP = 530 - 3200; CrA = 530; Crc = 3200; DC = 220; }
            if (Ctpd.Text == "Judías verdes cons larga") { Npr = 201; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.5M; CcL = 0.91M; CcF = 0.47M; CeL = 71M; CrP = 160 - 2600; CrA = 160; Crc = 2600; DC = 250; }
            if (Ctpd.Text == "Kiwi") { Npr = 202; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 98; Hur2 = 10; Pdc = -2.2M; CcL = 0.84M; CcF = 0.44M; CeL = 64M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Langosta") { Npr = 203; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -0.6M; CcL = 0.72M; CcF = 0.42M; CeL = 25M; CrP = 0; CrA = 0; Crc = 0; DC = 500; }
            if (Ctpd.Text == "Leche concentrada") { Npr = 204; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 70; Pdc = -1.5M; CcL = 0.42M; CcF = 0.28M; CeL = 20M; CrP = 0; CrA = 0; Crc = 0; DC = 500; }
            if (Ctpd.Text == "Leche condensada con Azúcar") { Npr = 205; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.93M; CcF = 0.49M; CeL = 70M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Leche desnatada") { Npr = 206; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 60; Pdc = 0M; CcL = 0.23M; CcF = 0.21M; CeL = 16M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Leche en polvo") { Npr = 207; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.6M; CcL = 0.93M; CcF = 0.49M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Leche fresca") { Npr = 208; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.6M; CcL = 0.93M; CcF = 0.49M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Leche pasteurizada") { Npr = 209; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.6M; CcL = 0.9M; CcF = 0.49M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Leche UHT") { Npr = 210; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.5M; CcL = 0.96M; CcF = 0.48M; CeL = 76M; CrP = 560 - 3900; CrA = 560; Crc = 3900; DC = 220; }
            if (Ctpd.Text == "Lechuga") { Npr = 211; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -0.2M; CcL = 0.96M; CcF = 0.49M; CeL = 76M; CrP = 560 - 3900; CrA = 560; Crc = 3900; DC = 220; }
            if (Ctpd.Text == "Lechuga cogollo") { Npr = 212; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -1M; CcL = 0.77M; CcF = 0.41M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Levadura") { Npr = 213; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -1M; CcL = 0.77M; CcF = 0.42M; CeL = 57M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Levadura panadería comprimida") { Npr = 214; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.6M; CcL = 0.9M; CcF = 0.48M; CeL = 69M; CrP = 400 - 5000; CrA = 400; Crc = 5000; DC = 300; }
            if (Ctpd.Text == "Lima") { Npr = 215; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.4M; CcL = 0.91M; CcF = 0.47M; CeL = 71M; CrP = 400 - 2000; CrA = 400; Crc = 2000; DC = 300; }
            if (Ctpd.Text == "Limón") { Npr = 216; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.4M; CcL = 0.91M; CcF = 0.47M; CeL = 71M; CrP = 400 - 2000; CrA = 400; Crc = 2000; DC = 300; }
            if (Ctpd.Text == "Limòn Persa") { Npr = 217; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 60; Pdc = 0M; CcL = 0.8M; CcF = 0.42M; CeL = 57M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Limón verde") { Npr = 218; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -2M; CcL = 0.28M; CcF = 0.23M; CeL = 8M; CrP = 110 - 1800; CrA = 110; Crc = 1800; DC = 250; }
            if (Ctpd.Text == "Lúpulo") { Npr = 219; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -1.1M; CcL = 0.82M; CcF = 0.42M; CeL = 60M; CrP = 270 - 4200; CrA = 270; Crc = 4200; DC = 250; }
            if (Ctpd.Text == "Maíz para palomitas") { Npr = 220; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -1.5M; CcL = 0.34M; CcF = 0.25M; CeL = 14M; CrP = 120 - 2200; CrA = 120; Crc = 2200; DC = 250; }
            if (Ctpd.Text == "Maíz tierno") { Npr = 221; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.92M; CcF = 0.49M; CeL = 69.3M; CrP = 300 - 7000; CrA = 300; Crc = 7000; DC = 250; }
            if (Ctpd.Text == "Malvavisco") { Npr = 222; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.1M; CcL = 0.92M; CcF = 0.49M; CeL = 69.3M; CrP = 500 - 7000; CrA = 500; Crc = 7000; DC = 250; }
            if (Ctpd.Text == "Mandarinas cons larga") { Npr = 223; Tip = 11; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.9M; CcL = 0.85M; CcF = 0.44M; CeL = 64.5M; CrP = 800 - 5000; CrA = 800; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Mandarinas cons corta") { Npr = 224; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = 4M; CcL = 0.5M; CcF = 0.34M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Mangos") { Npr = 225; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.5M; CcL = 0.6M; CcF = 0.34M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Manteca congelada") { Npr = 226; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.5M; CcF = 0.34M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Manteca de cerdo cons corta") { Npr = 227; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.6M; CcL = 0.33M; CcF = 0.25M; CeL = 13M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Manteca de cerdo cons larga") { Npr = 228; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.6M; CcL = 0.33M; CcF = 0.25M; CeL = 13M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Mantequilla congelada") { Npr = 229; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -0.6M; CcL = 0.33M; CcF = 0.25M; CeL = 13M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Mantequilla cons corta") { Npr = 230; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.5M; CcL = 0.88M; CcF = 0.45M; CeL = 67.4M; CrP = 120 - 4000; CrA = 120; Crc = 4000; DC = 230; }
            if (Ctpd.Text == "Mantequilla cons larga") { Npr = 231; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = 0M; CcL = 0.39M; CcF = 0.27M; CeL = 28M; CrP = 50 - 500; CrA = 50; Crc = 500; DC = 200; }
            if (Ctpd.Text == "Manzanas madura") { Npr = 232; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.5M; CcL = 0.88M; CcF = 0.45M; CeL = 67.4M; CrP = 120 - 3400; CrA = 120; Crc = 3400; DC = 250; }
            if (Ctpd.Text == "Manzanas secas") { Npr = 233; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 70; Pdc = -1M; CcL = 0.32M; CcF = 0.25M; CeL = 12M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Manzanas verde") { Npr = 234; Tip = -25; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.81M; CcF = 0.43M; CeL = 61M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Margarina") { Npr = 235; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.1M; CcL = 0.91M; CcF = 0.47M; CeL = 68M; CrP = 430 - 7000; CrA = 430; Crc = 7000; DC = 250; }
            if (Ctpd.Text == "Marisco Congelado") { Npr = 236; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 88; Hur2 = 92; Pdc = -1.1M; CcL = 0.91M; CcF = 0.47M; CeL = 68M; CrP = 150 - 4700; CrA = 150; Crc = 4700; DC = 300; }
            if (Ctpd.Text == "Melocotón maduro") { Npr = 237; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.1M; CcL = 0.91M; CcF = 0.47M; CeL = 68M; CrP = 0 - 430; CrA = 0; Crc = 430; DC = 340; }
            if (Ctpd.Text == "Melocotón verde") { Npr = 238; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = -2M; CcL = 0.4M; CcF = 0.28M; CeL = 30M; CrP = 60 - 600; CrA = 60; Crc = 600; DC = 240; }
            if (Ctpd.Text == "Melocotones congelados") { Npr = 239; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 95; Pdc = -1.1M; CcL = 0.94M; CcF = 0.48M; CeL = 74.1M; CrP = 300 - 5000; CrA = 300; Crc = 5000; DC = 340; }
            if (Ctpd.Text == "Melocotones Secos") { Npr = 240; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.94M; CcF = 0.48M; CeL = 74.1M; CrP = 300 - 6000; CrA = 300; Crc = 6000; DC = 300; }
            if (Ctpd.Text == "Melón de Indias") { Npr = 241; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.8M; CcL = 0.94M; CcF = 0.48M; CeL = 74.1M; CrP = 300 - 6000; CrA = 300; Crc = 6000; DC = 300; }
            if (Ctpd.Text == "Melón honeydew") { Npr = 242; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.1M; CcL = 0.97M; CcF = 0.49M; CeL = 77M; CrP = 500 - 2000; CrA = 500; Crc = 2000; DC = 300; }
            if (Ctpd.Text == "Melón Persa") { Npr = 243; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 88; Hur2 = 92; Pdc = -2M; CcL = 0.88M; CcF = 0.45M; CeL = 67.9M; CrP = 200 - 6000; CrA = 200; Crc = 6000; DC = 300; }
            if (Ctpd.Text == "Melones") { Npr = 244; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.7M; CcF = 0.39M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Membrillo") { Npr = 245; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.85M; CcF = 0.45M; CeL = 65M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Menhaden") { Npr = 246; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 70; Pdc = -8M; CcL = 0.35M; CcF = 0.26M; CeL = 15M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Merluza") { Npr = 247; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.8M; CcL = 0.88M; CcF = 0.46M; CeL = 67.9M; CrP = 160 - 3500; CrA = 160; Crc = 3500; DC = 250; }
            if (Ctpd.Text == "Miel") { Npr = 248; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.93M; CcF = 0.47M; CeL = 72M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 300; }
            if (Ctpd.Text == "Moras") { Npr = 249; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 98; Pdc = -0.2M; CcL = 0.92M; CcF = 0.47M; CeL = 72M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 250; }
            if (Ctpd.Text == "Nabos") { Npr = 250; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 97; Pdc = -1.1M; CcL = 0.94M; CcF = 0.48M; CeL = 73M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 300; }
            if (Ctpd.Text == "Nabos hojas") { Npr = 251; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.9M; CcF = 0.46M; CeL = 69.3M; CrP = 300 - 2800; CrA = 300; Crc = 2800; DC = 300; }
            if (Ctpd.Text == "Nabos raíces") { Npr = 252; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2.2M; CcL = 0.85M; CcF = 0.4M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Naranjas") { Npr = 253; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -2.2M; CcL = 0.85M; CcF = 0.4M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Nata (40%) congelada") { Npr = 254; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.9M; CcF = 0.45M; CeL = 60M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Nata (40%) cons corta") { Npr = 255; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 92; Pdc = -0.9M; CcL = 0.86M; CcF = 0.44M; CeL = 65.5M; CrP = 200 - 5000; CrA = 200; Crc = 5000; DC = 300; }
            if (Ctpd.Text == "Nécora") { Npr = 256; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.9M; CcF = 0.46M; CeL = 69.3M; CrP = 300 - 2800; CrA = 300; Crc = 2800; DC = 250; }
            if (Ctpd.Text == "Nectarinas") { Npr = 257; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = -2M; CcL = 0.26M; CcF = 0.22M; CeL = 6M; CrP = 10 - 1000; CrA = 10; Crc = 1000; DC = 300; }
            if (Ctpd.Text == "Níspero") { Npr = 258; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = -2M; CcL = 0.3M; CcF = 0.24M; CeL = 10M; CrP = 10 - 1000; CrA = 10; Crc = 1000; DC = 300; }
            if (Ctpd.Text == "Nuez con cáscara cons larga") { Npr = 259; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -2M; CcL = 0.26M; CcF = 0.22M; CeL = 6M; CrP = 270 - 1000; CrA = 270; Crc = 1000; DC = 300; }
            if (Ctpd.Text == "Nuez sin cáscara cons larga") { Npr = 260; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = -2M; CcL = 0.3M; CcF = 0.24M; CeL = 10M; CrP = 270 - 1000; CrA = 270; Crc = 1000; DC = 300; }
            if (Ctpd.Text == "Nuez con cáscara cons corta") { Npr = 261; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 70; Pdc = -1M; CcL = 0.33M; CcF = 0.25M; CeL = 13M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Nuez sin cáscara cons corta") { Npr = 262; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 99; Hur2 = 10; Pdc = -2.2M; CcL = 0.9M; CcF = 0.47M; CeL = 70M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Oleomargarina") { Npr = 263; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Ostras (carne-liquido)") { Npr = 264; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Ostras conserva cons larga") { Npr = 265; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.8M; CcL = 0.85M; CcF = 0.44M; CeL = 64M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Ostras conserva cons corta") { Npr = 266; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Ostras entera") { Npr = 267; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.89M; CcF = 0.44M; CeL = 69M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Ostras fresca cons larga") { Npr = 268; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.47M; CcF = 0.28M; CeL = 27M; CrP = 0; CrA = 0; Crc = 0; DC = 220; }
            if (Ctpd.Text == "Ostras fresca cons corta") { Npr = 269; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.47M; CcF = 0.28M; CeL = 27M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Pan") { Npr = 270; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.6M; CcF = 0.34M; CeL = 31M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Pan congelado") { Npr = 271; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Pan precocido") { Npr = 272; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Pan masa enfriar") { Npr = 273; Tip = -30; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 150; }
            if (Ctpd.Text == "Pan masa congelado") { Npr = 274; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Pan masa a Congelar") { Npr = 275; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Pasteles") { Npr = 276; Tip = -30; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.75M; CcF = 0.41M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 150; }
            if (Ctpd.Text == "Pasteles Congelados") { Npr = 277; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.93M; CcF = 0.48M; CeL = 72.6M; CrP = 400 - 8000; CrA = 400; Crc = 8000; DC = 250; }
            if (Ctpd.Text == "Pasteles a Congelar") { Npr = 278; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.88M; CcF = 0.46M; CeL = 67.9M; CrP = 200 - 2800; CrA = 200; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Papayas") { Npr = 279; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.6M; CcL = 0.82M; CcF = 0.43M; CeL = 62M; CrP = 280 - 2200; CrA = 280; Crc = 2200; DC = 400; }
            if (Ctpd.Text == "Pasas de Corinto") { Npr = 280; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.6M; CcL = 0.85M; CcF = 0.44M; CeL = 65M; CrP = 280 - 2200; CrA = 280; Crc = 2200; DC = 400; }
            if (Ctpd.Text == "Patata tardía consumo") { Npr = 281; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.6M; CcL = 0.82M; CcF = 0.43M; CeL = 62M; CrP = 280 - 2200; CrA = 280; Crc = 2200; DC = 400; }
            if (Ctpd.Text == "Patata temprana") { Npr = 282; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.7M; CcL = 0.82M; CcF = 0.43M; CeL = 62M; CrP = 280 - 2200; CrA = 280; Crc = 2200; DC = 400; }
            if (Ctpd.Text == "Patatas") { Npr = 283; Tip = 8; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.97M; CcF = 0.49M; CeL = 76M; CrP = 110 - 6000; CrA = 110; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Patatas de cosecha") { Npr = 284; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.5M; CcL = 0.86M; CcF = 0.45M; CeL = 66.2M; CrP = 430 - 7600; CrA = 430; Crc = 7600; DC = 250; }
            if (Ctpd.Text == "Pepinos") { Npr = 285; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.5M; CcL = 0.86M; CcF = 0.45M; CeL = 66.2M; CrP = 240 - 6000; CrA = 240; Crc = 6000; DC = 280; }
            if (Ctpd.Text == "Peras maduras") { Npr = 286; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.85M; CcF = 0.45M; CeL = 65M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Peras verdes") { Npr = 287; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -1.1M; CcL = 0.88M; CcF = 0.46M; CeL = 66M; CrP = 280 - 2200; CrA = 280; Crc = 2200; DC = 150; }
            if (Ctpd.Text == "Perca") { Npr = 288; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.85M; CcF = 0.45M; CeL = 65M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Perejil") { Npr = 289; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 60; Pdc = -1.8M; CcL = 0.76M; CcF = 0.41M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescadilla") { Npr = 290; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.76M; CcF = 0.41M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescado ahumado") { Npr = 291; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Pescado blanco") { Npr = 292; Tip = -25; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Pescado congelado cons corta") { Npr = 293; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescado congelado cons larga") { Npr = 294; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.41M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescado fresco hielo cons corta") { Npr = 295; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -2.8M; CcL = 0.68M; CcF = 0.4M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 380; }
            if (Ctpd.Text == "Pescado fresco hielo cons larga") { Npr = 296; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.72M; CcF = 0.38M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescado salazón") { Npr = 297; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.86M; CcF = 0.45M; CeL = 59M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Pescados grasos") { Npr = 298; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 58; Pdc = 0M; CcL = 0.58M; CcF = 0.34M; CeL = 38M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Pescados magros") { Npr = 299; Tip = -18; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 60; Pdc = -2.2M; CcL = 0.4M; CcF = 0.2M; CeL = 10M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Picadillo seco") { Npr = 300; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 45; Hur2 = 55; Pdc = -2.2M; CcL = 0.4M; CcF = 0.2M; CeL = 10M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Pieles (curtidas) congeladas") { Npr = 301; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 83; Hur2 = 88; Pdc = -2.2M; CcL = 0.4M; CcF = 0.2M; CeL = 10M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Pieles (pelambre)") { Npr = 302; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.7M; CcL = 0.94M; CcF = 0.47M; CeL = 73M; CrP = 750 - 5500; CrA = 750; Crc = 5500; DC = 180; }
            if (Ctpd.Text == "Pieles curtidas cons larga") { Npr = 303; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.7M; CcL = 0.94M; CcF = 0.48M; CeL = 73M; CrP = 500 - 5500; CrA = 500; Crc = 5500; DC = 180; }
            if (Ctpd.Text == "Pimienta dulce") { Npr = 304; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 70; Pdc = 0M; CcL = 0.3M; CcF = 0.24M; CeL = 28M; CrP = 50 - 250; CrA = 50; Crc = 250; DC = 180; }
            if (Ctpd.Text == "Pimientos frescos") { Npr = 305; Tip = 4; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1M; CcL = 0.88M; CcF = 0.45M; CeL = 67.9M; CrP = 100 - 9000; CrA = 100; Crc = 9000; DC = 250; }
            if (Ctpd.Text == "Pimientos secos") { Npr = 306; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1M; CcL = 0.86M; CcF = 0.43M; CeL = 66.4M; CrP = 100 - 4000; CrA = 100; Crc = 4000; DC = 250; }
            if (Ctpd.Text == "Piñas maduras") { Npr = 307; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.8M; CcL = 0.85M; CcF = 0.42M; CeL = 60M; CrP = 840 - 5300; CrA = 840; Crc = 5300; DC = 250; }
            if (Ctpd.Text == "Piñas verde") { Npr = 308; Tip = 11; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 95; Pdc = -0.8M; CcL = 0.85M; CcF = 0.42M; CeL = 60M; CrP = 550 - 4100; CrA = 550; Crc = 4100; DC = 250; }
            if (Ctpd.Text == "Plátanos maduros") { Npr = 309; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -1M; CcL = 0.88M; CcF = 0.48M; CeL = 68M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Plátanos verde") { Npr = 310; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1M; CcL = 0.88M; CcF = 0.48M; CeL = 68M; CrP = 0; CrA = 0; Crc = 0; DC = 500; }
            if (Ctpd.Text == "Platos precocinados") { Npr = 311; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.1M; CcL = 0.91M; CcF = 0.47M; CeL = 71M; CrP = 600 - 6000; CrA = 600; Crc = 6000; DC = 300; }
            if (Ctpd.Text == "Platos precocinados congelados") { Npr = 312; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.7M; CcL = 0.88M; CcF = 0.46M; CeL = 70M; CrP = 450 - 5000; CrA = 450; Crc = 5000; DC = 220; }
            if (Ctpd.Text == "Pomelo") { Npr = 313; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.82M; CcF = 0.43M; CeL = 52M; CrP = 0; CrA = 0; Crc = 0; DC = 380; }
            if (Ctpd.Text == "Puerros") { Npr = 314; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.82M; CcF = 0.43M; CeL = 52M; CrP = 0; CrA = 0; Crc = 0; DC = 450; }
            if (Ctpd.Text == "Pulpo") { Npr = 315; Tip = 8; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.8M; CcL = 0.92M; CcF = 0.47M; CeL = 70M; CrP = 500 - 5000; CrA = 500; Crc = 5000; DC = 220; }
            if (Ctpd.Text == "Pulpo congelado") { Npr = 316; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -8M; CcL = 0.64M; CcF = 0.36M; CeL = 44M; CrP = 270 - 1200; CrA = 270; Crc = 1200; DC = 280; }
            if (Ctpd.Text == "Qinbomgo") { Npr = 317; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -7M; CcL = 0.7M; CcF = 0.4M; CeL = 48M; CrP = 130 - 1600; CrA = 130; Crc = 1600; DC = 280; }
            if (Ctpd.Text == "Queso americano") { Npr = 318; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -9M; CcL = 0.51M; CcF = 0.31M; CeL = 30M; CrP = 110 - 1200; CrA = 110; Crc = 1200; DC = 280; }
            if (Ctpd.Text == "Queso camembert") { Npr = 319; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -6M; CcL = 0.45M; CcF = 0.29M; CeL = 25M; CrP = 80 - 800; CrA = 80; Crc = 800; DC = 280; }
            if (Ctpd.Text == "Queso cheddar") { Npr = 320; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -3M; CcL = 0.75M; CcF = 0.42M; CeL = 40M; CrP = 270 - 1200; CrA = 270; Crc = 1200; DC = 300; }
            if (Ctpd.Text == "Queso cheddar rallado") { Npr = 321; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -7M; CcL = 0.7M; CcF = 0.42M; CeL = 48M; CrP = 160 - 1800; CrA = 160; Crc = 1800; DC = 280; }
            if (Ctpd.Text == "Queso fresco") { Npr = 322; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = -2M; CcL = 0.68M; CcF = 0.38M; CeL = 40M; CrP = 250 - 1800; CrA = 250; Crc = 1800; DC = 250; }
            if (Ctpd.Text == "Queso lumburger") { Npr = 323; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 75; Pdc = -2M; CcL = 0.6M; CcF = 0.32M; CeL = 35M; CrP = 200 - 1500; CrA = 200; Crc = 1500; DC = 280; }
            if (Ctpd.Text == "Queso manchego curado") { Npr = 324; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -12M; CcL = 0.65M; CcF = 0.32M; CeL = 44M; CrP = 110 - 1500; CrA = 110; Crc = 1500; DC = 280; }
            if (Ctpd.Text == "Queso manchego graso") { Npr = 325; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -8M; CcL = 0.64M; CcF = 0.36M; CeL = 44M; CrP = 110 - 1500; CrA = 110; Crc = 1500; DC = 280; }
            if (Ctpd.Text == "Queso roquefort") { Npr = 326; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -1.8M; CcL = 0.8M; CcF = 0.43M; CeL = 60M; CrP = 300 - 3200; CrA = 300; Crc = 3200; DC = 220; }
            if (Ctpd.Text == "Queso suizo") { Npr = 327; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -0.7M; CcL = 0.96M; CcF = 0.49M; CeL = 68M; CrP = 500 - 5000; CrA = 500; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Rábano picante") { Npr = 328; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 98; Pdc = -0.7M; CcL = 0.96M; CcF = 0.49M; CeL = 68M; CrP = 500 - 5000; CrA = 500; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Rábanos invierno") { Npr = 329; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 88; Pdc = -1.5M; CcL = 0.74M; CcF = 0.4M; CeL = 54M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Rábanos primavera") { Npr = 330; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.4M; CcL = 0.9M; CcF = 0.47M; CeL = 72M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 280; }
            if (Ctpd.Text == "Redondo (Selecto)") { Npr = 331; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.9M; CcL = 0.9M; CcF = 0.47M; CeL = 72M; CrP = 270 - 3800; CrA = 270; Crc = 3800; DC = 300; }
            if (Ctpd.Text == "Remolacha hojas") { Npr = 332; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 99; Pdc = -0.9M; CcL = 0.96M; CcF = 0.48M; CeL = 75M; CrP = 350 - 5500; CrA = 350; Crc = 5500; DC = 250; }
            if (Ctpd.Text == "Remolacha raíz") { Npr = 333; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 98; Hur2 = 10; Pdc = -1.1M; CcL = 0.92M; CcF = 0.47M; CeL = 76M; CrP = 300 - 4500; CrA = 300; Crc = 4500; DC = 250; }
            if (Ctpd.Text == "Ruibarbo") { Npr = 334; Tip = 7; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = 0M; CcL = 0.85M; CcF = 0.48M; CeL = 62M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Rutabaga") { Npr = 335; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 35; Hur2 = 40; Pdc = 0M; CcL = 0.85M; CcF = 0.5M; CeL = 58M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sala de embalaje") { Npr = 336; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 78; Hur2 = 80; Pdc = -3.3M; CcL = 0.89M; CcF = 0.53M; CeL = 52M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sala manipulación") { Npr = 337; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 81; Hur2 = 85; Pdc = -2.9M; CcL = 0.65M; CcF = 0.4M; CeL = 50M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salazones frescas") { Npr = 338; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.85M; CcF = 0.55M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha campesina ahumada") { Npr = 339; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.85M; CcF = 0.55M; CeL = 55M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha embutida") { Npr = 340; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 82; Hur2 = 85; Pdc = -2M; CcL = 0.6M; CcF = 0.35M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha congelada") { Npr = 341; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 83; Hur2 = 85; Pdc = -1M; CcL = 0.63M; CcF = 0.4M; CeL = 48M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha en ristras") { Npr = 342; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 90; Pdc = -1.7M; CcL = 0.85M; CcF = 0.55M; CeL = 47M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha estilo Polaco") { Npr = 343; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 90; Pdc = -1.7M; CcL = 0.85M; CcF = 0.55M; CeL = 47M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha Fráncfort media") { Npr = 344; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.5M; CcL = 0.89M; CcF = 0.56M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Salchicha Fráncfort y ahumada") { Npr = 345; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -2.2M; CcL = 0.72M; CcF = 0.4M; CeL = 51M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Salchicha Fresca") { Npr = 346; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 98; Pdc = -1.1M; CcL = 0.83M; CcF = 0.44M; CeL = 63M; CrP = 350 - 4200; CrA = 350; Crc = 4200; DC = 250; }
            if (Ctpd.Text == "Salmón") { Npr = 347; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.4M; CcL = 0.94M; CcF = 0.48M; CeL = 74.1M; CrP = 500 - 1200; CrA = 500; Crc = 1200; DC = 300; }
            if (Ctpd.Text == "Salsifí") { Npr = 348; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.47M; CcF = 0.31M; CeL = 36M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Sandías") { Npr = 349; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 8M; CrP = 50 - 600; CrA = 50; Crc = 600; DC = 250; }
            if (Ctpd.Text == "Sebo") { Npr = 350; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 65; Pdc = 0M; CcL = 0.29M; CcF = 0.23M; CeL = 15M; CrP = 100 - 300; CrA = 100; Crc = 300; DC = 150; }
            if (Ctpd.Text == "Semilla") { Npr = 351; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 98; Pdc = 0M; CcL = 0.18M; CcF = 0.14M; CeL = 6M; CrP = 110 - 1100; CrA = 110; Crc = 1100; DC = 150; }
            if (Ctpd.Text == "Semilla de verduras") { Npr = 352; Tip = 6; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 98; Pdc = -0.9M; CcL = 0.93M; CcF = 0.48M; CeL = 72M; CrP = 300 - 2800; CrA = 300; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Semillero") { Npr = 353; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 93; Hur2 = 98; Pdc = -0.9M; CcL = 0.93M; CcF = 0.48M; CeL = 72M; CrP = 200 - 2800; CrA = 200; Crc = 2800; DC = 200; }
            if (Ctpd.Text == "Setas cons. Corta") { Npr = 354; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 83; Hur2 = 85; Pdc = -1.7M; CcL = 0.65M; CcF = 0.37M; CeL = 45M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Setas") { Npr = 355; Tip = 20; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 45; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 4M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Solomillo (Selecto)") { Npr = 356; Tip = 5; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -0.5M; CcL = 0.92M; CcF = 0.48M; CeL = 76M; CrP = 0; CrA = 0; Crc = 0; DC = 400; }
            if (Ctpd.Text == "Suero en polvo") { Npr = 357; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 80; Pdc = 0M; CcL = 0.3M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Suero vacuno") { Npr = 358; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = 0M; CcL = 0.3M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Tabaco balas") { Npr = 359; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 50; Hur2 = 55; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Tabaco barril") { Npr = 360; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Tabaco cigarrillos") { Npr = 361; Tip = 16; Tra1 = 14; Tra2 = 18; Hur1 = 70; Hur2 = 75; Pdc = 0M; CcL = 0.24M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tabaco puros hojas") { Npr = 362; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 60; Hur2 = 65; Pdc = 0M; CcL = 0.3M; CcF = 0.22M; CeL = 8M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Tabaco puros Bodega") { Npr = 363; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.36M; CcF = 0.26M; CeL = 18M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tabaco de pipa") { Npr = 364; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 72; Hur2 = 75; Pdc = -2M; CcL = 0.58M; CcF = 0.4M; CeL = 40M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino") { Npr = 365; Tip = 10; Tra1 = 14; Tra2 = 18; Hur1 = 40; Hur2 = 45; Pdc = -2M; CcL = 0.5M; CcF = 0.3M; CeL = 35M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino ahumado esclarecido") { Npr = 366; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.36M; CcF = 0.26M; CeL = 18M; CrP = 0; CrA = 0; Crc = 0; DC = 300; }
            if (Ctpd.Text == "Tocino ahumado rodajas") { Npr = 367; Tip = 16; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.36M; CcF = 0.26M; CeL = 17M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino congelado") { Npr = 368; Tip = 3; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -2M; CcL = 0.36M; CcF = 0.26M; CeL = 18M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino curado estilo campesino") { Npr = 369; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 80; Hur2 = 85; Pdc = -1M; CcL = 0.55M; CcF = 0.31M; CeL = 17M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino entreverado") { Npr = 370; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 83; Hur2 = 85; Pdc = -1M; CcL = 0.28M; CcF = 0.23M; CeL = 14M; CrP = 0; CrA = 0; Crc = 0; DC = 200; }
            if (Ctpd.Text == "Tocino fresco") { Npr = 371; Tip = 14; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 300 - 6000; CrA = 300; Crc = 6000; DC = 350; }
            if (Ctpd.Text == "Tocino grasa 100%") { Npr = 372; Tip = 18; Tra1 = 14; Tra2 = 18; Hur1 = 79; Hur2 = 85; Pdc = 0M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 300 - 6000; CrA = 300; Crc = 6000; DC = 350; }
            if (Ctpd.Text == "Tomate de aliñar") { Npr = 373; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 300 - 6000; CrA = 300; Crc = 6000; DC = 350; }
            if (Ctpd.Text == "Tomates a madurar") { Npr = 374; Tip = 11; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -0.5M; CcL = 0.95M; CcF = 0.48M; CeL = 74M; CrP = 250 - 4200; CrA = 250; Crc = 4200; DC = 350; }
            if (Ctpd.Text == "Tomates maduros") { Npr = 375; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.1M; CcL = 0.91M; CcF = 0.48M; CeL = 70M; CrP = 500 - 7000; CrA = 500; Crc = 7000; DC = 250; }
            if (Ctpd.Text == "Tomates verdes") { Npr = 376; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.79M; CcF = 0.42M; CeL = 59M; CrP = 500 - 6000; CrA = 500; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Toronjas") { Npr = 377; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.1M; CcL = 0.86M; CcF = 0.44M; CeL = 65.5M; CrP = 240 - 2500; CrA = 240; Crc = 2500; DC = 250; }
            if (Ctpd.Text == "Trigo fresco") { Npr = 378; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.6M; CcL = 0.86M; CcF = 0.44M; CeL = 65.5M; CrP = 240 - 2500; CrA = 240; Crc = 2500; DC = 250; }
            if (Ctpd.Text == "Uvas") { Npr = 379; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = -2M; CcL = 0.35M; CcF = 0.26M; CeL = 18M; CrP = 50 - 2000; CrA = 50; Crc = 2000; DC = 200; }
            if (Ctpd.Text == "Uvas americanas") { Npr = 380; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 55; Hur2 = 60; Pdc = -2M; CcL = 0.42M; CcF = 0.28M; CeL = 22.9M; CrP = 50 - 2000; CrA = 50; Crc = 2000; DC = 190; }
            if (Ctpd.Text == "Uvas Conservación") { Npr = 381; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.9M; CcL = 0.86M; CcF = 0.44M; CeL = 65.5M; CrP = 240 - 2500; CrA = 240; Crc = 2500; DC = 250; }
            if (Ctpd.Text == "Uvas pasas") { Npr = 382; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Uvas secas") { Npr = 383; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 88; Hur2 = 92; Pdc = -2.2M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 250; }
            if (Ctpd.Text == "Uvas vinícola") { Npr = 384; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.74M; CcF = 0.4M; CeL = 53M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Vaca Congelada") { Npr = 385; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2M; CcL = 0.74M; CcF = 0.4M; CeL = 53M; CrP = 0; CrA = 0; Crc = 0; DC = 220; }
            if (Ctpd.Text == "Vaca fresco promedio") { Npr = 386; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2.2M; CcL = 0.6M; CcF = 0.35M; CeL = 44M; CrP = 0; CrA = 0; Crc = 0; DC = 280; }
            if (Ctpd.Text == "Vaca ternera congelada") { Npr = 387; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -1.7M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 280; }
            if (Ctpd.Text == "Vaca ternera promedio") { Npr = 388; Tip = 12; Tra1 = 14; Tra2 = 18; Hur1 = 61; Hur2 = 65; Pdc = 0M; CcL = 0.33M; CcF = 0.25M; CeL = 12M; CrP = 0; CrA = 0; Crc = 0; DC = 220; }
            if (Ctpd.Text == "Vaca-Buey (graso)") { Npr = 389; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -2.2M; CcL = 0.77M; CcF = 0.42M; CeL = 56M; CrP = 0; CrA = 0; Crc = 0; DC = 220; }
            if (Ctpd.Text == "Vaca-Buey (magro)") { Npr = 390; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.6M; CcL = 0.855M; CcF = 0.45M; CeL = 65.5M; CrP = 200 - 3000; CrA = 200; Crc = 3000; DC = 250; }
            if (Ctpd.Text == "Vaca-Buey (Seca)") { Npr = 391; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -0.6M; CcL = 0.94M; CcF = 0.48M; CeL = 72M; CrP = 0 - 500; CrA = 0; Crc = 500; DC = 300; }
            if (Ctpd.Text == "Vaca-Buey Congelado") { Npr = 392; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 95; Hur2 = 10; Pdc = -0.3M; CcL = 0.94M; CcF = 0.48M; CeL = 72M; CrP = 500 - 6000; CrA = 500; Crc = 6000; DC = 250; }
            if (Ctpd.Text == "Vaccinias") { Npr = 393; Tip = 8; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -11M; CcL = 0.98M; CcF = 0.47M; CeL = 78M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Verduras congeladas") { Npr = 394; Tip = 16; Tra1 = 14; Tra2 = 18; Hur1 = 75; Hur2 = 80; Pdc = -11M; CcL = 0.98M; CcF = 0.47M; CeL = 78M; CrP = 0; CrA = 0; Crc = 0; DC = 700; }
            if (Ctpd.Text == "Verduras frondosas") { Npr = 395; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -1M; CcL = 0.85M; CcF = 0.48M; CeL = 68M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Vino Tinto joven") { Npr = 396; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.92M; CcF = 0.46M; CeL = 70M; CrP = 0 - 500; CrA = 0; Crc = 500; DC = 380; }
            if (Ctpd.Text == "Vino Blanco dulce") { Npr = 397; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.87M; CcF = 0.45M; CeL = 70M; CrP = 500 - 4500; CrA = 500; Crc = 4500; DC = 320; }
            if (Ctpd.Text == "Vino Tinto reserva") { Npr = 398; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.92M; CcF = 0.46M; CeL = 70M; CrP = 500 - 4500; CrA = 500; Crc = 4500; DC = 350; }
            if (Ctpd.Text == "Vino Blanco joven") { Npr = 399; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 97; Pdc = -0.8M; CcL = 0.88M; CcF = 0.46M; CeL = 68M; CrP = 400 - 5000; CrA = 400; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Vino blanco") { Npr = 400; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.92M; CcF = 0.47M; CeL = 71M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Vino tinto") { Npr = 401; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.92M; CcF = 0.47M; CeL = 71M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }
            if (Ctpd.Text == "Yogur") { Npr = 402; Tip = 2; Tra1 = 14; Tra2 = 18; Hur1 = 65; Hur2 = 70; Pdc = -1M; CcL = 0.85M; CcF = 0.48M; CeL = 68M; CrP = 0; CrA = 0; Crc = 0; DC = 350; }
            if (Ctpd.Text == "Zanahorias congeladas") { Npr = 403; Tip = -20; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.92M; CcF = 0.46M; CeL = 70M; CrP = 0 - 500; CrA = 0; Crc = 500; DC = 380; }
            if (Ctpd.Text == "Zanahorias hojas") { Npr = 404; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.87M; CcF = 0.45M; CeL = 70M; CrP = 500 - 4500; CrA = 500; Crc = 4500; DC = 320; }
            if (Ctpd.Text == "Zanahorias raíces") { Npr = 405; Tip = -1; Tra1 = 14; Tra2 = 18; Hur1 = 90; Hur2 = 95; Pdc = -1.3M; CcL = 0.92M; CcF = 0.46M; CeL = 70M; CrP = 500 - 4500; CrA = 500; Crc = 4500; DC = 350; }
            if (Ctpd.Text == "Zarzamoras") { Npr = 406; Tip = 0; Tra1 = 14; Tra2 = 18; Hur1 = 92; Hur2 = 97; Pdc = -0.8M; CcL = 0.88M; CcF = 0.46M; CeL = 68M; CrP = 400 - 5000; CrA = 400; Crc = 5000; DC = 250; }
            if (Ctpd.Text == "Zumo de Naranja") { Npr = 407; Tip = 1; Tra1 = 14; Tra2 = 18; Hur1 = 85; Hur2 = 90; Pdc = -2M; CcL = 0.92M; CcF = 0.47M; CeL = 71M; CrP = 0; CrA = 0; Crc = 0; DC = 800; }



            TTem.Text = Tip.ToString();
            decimal CeLD;
            int TeRF;//Temperatura final del producto

            decimal TeTR;//Temperatura Final del producto
            decimal TeRI;//Temperatura Inicial fijada
            TeRF = 0;
            TeTR = 0;
            TeRI = 0;
            //= SI(G16 <= -18; -10; SI(G16 <= -2; 2; SI(G16 >= 0; 25)))
            TeRF = Convert.ToInt16(float.Parse(TCentx.Text));//Temperatura final del producto fijada
            if (TeRF <= -18) { TeTR = -10; } else { if (TeRF <= -2) { TeTR = 2; } else { if (TeRF >= 0) { TeTR = 25; } } }//Temp inicial calculada
            TeRI = Convert.ToInt16(float.Parse(Coff3.Text));
            if (TeRI > TeTR) { TeTR = TeRI; }
            if (TeRI <= TeTR) { TeRI = TeTR; }
            // Parametros por localidad.
            decimal ta;// Temperatura ambiente,
            ta = 0;//Temperatura ambiente
            Hur = 0;//Humedad relativa

            if (CTLup.Text == "Genérica 40 ºC") { ta = 40; Hur = 40; }
            if (CTLup.Text == "Genérica 38 ºC") { ta = 38; Hur = 45; }
            if (CTLup.Text == "Genérica 35 ºC") { ta = 35; Hur = 50; }
            if (CTLup.Text == "Genérica 33 ºC") { ta = 33; Hur = 55; }
            if (CTLup.Text == "Genérica 30 ºC") { ta = 30; Hur = 60; }
            if (CTLup.Text == "Genérica 28 ºC") { ta = 28; Hur = 65; }
            if (CTLup.Text == "Genérica 25 ºC") { ta = 25; Hur = 70; }
            if (CTLup.Text == "A Coruña") { ta = 32; Hur = 63; }
            if (CTLup.Text == "Albacete") { ta = 35; Hur = 36; }
            if (CTLup.Text == "Alicante") { ta = 33; Hur = 60; }
            if (CTLup.Text == "Almería") { ta = 36; Hur = 70; }
            if (CTLup.Text == "Ávila") { ta = 30; Hur = 41; }
            if (CTLup.Text == "Badajoz") { ta = 40; Hur = 47; }
            if (CTLup.Text == "Barcelona") { ta = 32; Hur = 68; }
            if (CTLup.Text == "Bilbao") { ta = 34; Hur = 71; }
            if (CTLup.Text == "Burgos") { ta = 32; Hur = 42; }
            if (CTLup.Text == "Cáceres") { ta = 38; Hur = 37; }
            if (CTLup.Text == "Cádiz") { ta = 35; Hur = 55; }
            if (CTLup.Text == "Castellón") { ta = 35; Hur = 60; }
            if (CTLup.Text == "Ciudad Real") { ta = 38; Hur = 56; }
            if (CTLup.Text == "Córdoba") { ta = 40; Hur = 33; }
            if (CTLup.Text == "Cuenca") { ta = 34; Hur = 52; }
            if (CTLup.Text == "Gijón") { ta = 32; Hur = 74; }
            if (CTLup.Text == "Girona") { ta = 35; Hur = 58; }
            if (CTLup.Text == "Granada") { ta = 36; Hur = 49; }
            if (CTLup.Text == "Guadalajara") { ta = 35; Hur = 37; }
            if (CTLup.Text == "Huelva") { ta = 35; Hur = 57; }
            if (CTLup.Text == "Huesca") { ta = 33; Hur = 72; }
            if (CTLup.Text == "Jaén") { ta = 38; Hur = 35; }
            if (CTLup.Text == "Las Palmas") { ta = 32; Hur = 66; }
            if (CTLup.Text == "León") { ta = 32; Hur = 45; }
            if (CTLup.Text == "Lérida") { ta = 33; Hur = 50; }
            if (CTLup.Text == "Logroño") { ta = 34; Hur = 59; }
            if (CTLup.Text == "Lugo") { ta = 31; Hur = 67; }
            if (CTLup.Text == "Madrid") { ta = 36; Hur = 42; }
            if (CTLup.Text == "Málaga") { ta = 34; Hur = 60; }
            if (CTLup.Text == "Murcia") { ta = 37; Hur = 59; }
            if (CTLup.Text == "Ourense") { ta = 36; Hur = 55; }
            if (CTLup.Text == "Oviedo") { ta = 31; Hur = 70; }
            if (CTLup.Text == "Palencia") { ta = 34; Hur = 45; }
            if (CTLup.Text == "Palma de Mallorca") { ta = 34; Hur = 63; }
            if (CTLup.Text == "Pamplona") { ta = 32; Hur = 51; }
            if (CTLup.Text == "Pontevedra") { ta = 33; Hur = 62; }
            if (CTLup.Text == "Salamanca") { ta = 34; Hur = 46; }
            if (CTLup.Text == "Santander") { ta = 31; Hur = 74; }
            if (CTLup.Text == "San Sebastián") { ta = 30; Hur = 76; }
            if (CTLup.Text == "Santa Cruz de Tenerife") { ta = 32; Hur = 55; }
            if (CTLup.Text == "Segovia") { ta = 33; Hur = 35; }
            if (CTLup.Text == "Sevilla") { ta = 40; Hur = 43; }
            if (CTLup.Text == "Soria") { ta = 30; Hur = 45; }
            if (CTLup.Text == "Tarragona") { ta = 30; Hur = 68; }
            if (CTLup.Text == "Teruel") { ta = 35; Hur = 55; }
            if (CTLup.Text == "Toledo") { ta = 36; Hur = 34; }
            if (CTLup.Text == "Valencia") { ta = 33; Hur = 68; }
            if (CTLup.Text == "Valladolid") { ta = 34; Hur = 45; }
            if (CTLup.Text == "Vitoria") { ta = 33; Hur = 70; }
            if (CTLup.Text == "Zamora") { ta = 33; Hur = 65; }
            if (CTLup.Text == "Zaragoza") { ta = 34; Hur = 57; }


            decimal Dtt;//Temperatura ambiente - Temperatura producto
            Dtt = ta - Tip;
            decimal largo;// Largo exterior de cámaras
            decimal ancho;// Ancho exterior de cámaras
            decimal alto;// Alto exterior de cámara
            decimal volu;// Volumen de la cámara
            int pu; //Aislamiento de poliuretano
            int pep;//Aislamientp poliestileno
            int ps;//Aislamiento de suelo
            decimal epx;// Espesor panel de pared
            decimal esx;// Espesor panel de suelo
            decimal Ht;// Horas de trabajo
            decimal DinP;// Medida interior de la camara base de calculo
            decimal DinA;// Medida interior de la camara base de calculo
            decimal DinH;// Medida interior de la cámara base de calculo
            decimal V;// Valor del volumen
            decimal Dt;// Dieferencial de temperatura
            int PU;//Tipo de aislamiento (Poliuretano)
            int PEP;//Tipo de aislamiento (Poliestileno)
            int PS;//Tipo de aislamiento (Poliretano Suelo)
            decimal Htc;//Valor decimal de Horario
            epx = 0;
            esx = 0;
            Ht = 0;
            Htc = Convert.ToInt16(float.Parse(TSol.Text)*100);//Horas a trabajar valor fijado por el cliente
            Ht = Htc / 100;
            Datos.ta = ta.ToString();
           
            largo = Convert.ToInt32(float.Parse(TLargo.Text) * 1000);
            ancho = Convert.ToInt32(float.Parse(TAncho.Text) * 1000);
            alto = Convert.ToInt32(float.Parse(TAlto.Text) * 1000);
            epx = Convert.ToInt16(float.Parse(Cnoff6.Text));//Espesor de Parede
            esx = Convert.ToInt16(float.Parse(CBps.Text));//Espesor de suelo
            DinP = largo / 1000 - (epx * 2 / 1000);
            DinA = ancho / 1000 - (epx * 2 / 1000);
            DinH = alto / 1000 - ((esx + epx) / 1000);

            // Volumen
            //decimal V;//Volumen
            volu = DinP * DinA * DinH;
            if (volu < 100)
            {
                volu = Math.Round(volu, 1, MidpointRounding.ToEven);
            }

            if (volu > 100)
            {
                volu = Math.Round(volu, 0, MidpointRounding.ToEven);
            }

            TVolu.Text = Convert.ToDecimal(volu.ToString()).ToString();

            // 3. Peredida por transmisión de paredes en watio

            decimal PtP;//Perdidas por paredes Calculo del perimetro de la cámara
            PtP = Math.Round((DinP + DinA) * 2 * DinH, 4, MidpointRounding.ToEven);// 3.1 Perdida por paredes perimetro volumentrico

            //3.2 Calculo del diferencial de temperatura
            //decimal Tip;// Temperatura temperatura del producto


            //3.3 Calculo de perdidas horarias por transmisión de paredes
            decimal CdP;//Coeficiente de distribusión de paredes
            decimal PtPH;//Perdidas por paredes diaria
            decimal K1;// Coeficiente de valor PU
            decimal K2;// Coeficente de Valor PEP
            decimal K3;//Coefociente de valor PES
            decimal K4;//Coeficiente de valor HOR
            decimal PtPD;//Perdida por paredes po día
            decimal PtT;//Perdidas por transmisión de techo
            decimal CdT;//Coeficviente de distribucióm de techos
            decimal PtTH;//Perdidas por transmisión de techo por horas
            decimal PtTD;//Perdidas por transmisión de techo por día
            decimal DtP;//Diferencia de temperatura en psredes.
            decimal DtT;//Diferencial de temperatura de techo
            decimal DtS;//Diferencia de temperatura por suelo
            decimal CdS;//Coeficiente distribución de suelo
            decimal PtS;//Perdidas por transmisión de suelo
            decimal PtSH;//Perdidas por transmision de suelo por hora
            decimal PtSD;//Perdidas por transmisión de suelo pr día   
            K1 = 0;
            K2 = 0M;
            K3 = 0M;
            K4 = 0M;
            CdP = 0;
            CdT = 0;
            CdS = 0;

            //3.3 Calculo de perdidas horarias por transmisión de paredes
            PU = 1;
            if (epx == 0) { K1 = 1.630233M; K2 = 1.630233M; K3 = 1.630233M; K4 = 1.630233M; }
            if (epx == 5) { K1 = 1.511628M; K2 = 1.511628M; K3 = 1.511628M; K4 = 1.52093M; }
            if (epx == 10) { K1 = 1.395349M; K2 = 1.453488M; K3 = 1.453488M; K4 = 1.47093M; }
            if (epx == 15) { K1 = 1.27907M; K2 = 1.395349M; K3 = 1.395349M; K4 = 1.42093M; }
            if (epx == 20) { K1 = 1.133721M; K2 = 1.337209M; K3 = 1.337209M; K4 = 1.37093M; }
            if (epx == 25) { K1 = 0.906977M; K2 = 1.27907M; K3 = 1.27907M; K4 = 1.32093M; }
            if (epx == 30) { K1 = 0.755814M; K2 = 1.24031M; K3 = 1.162791M; K4 = 1.27093M; }
            if (epx == 35) { K1 = 0.647841M; K2 = 1.063123M; K3 = 0.996678M; K4 = 1.22093M; }
            if (epx == 40) { K1 = 0.56686M; K2 = 0.930233M; K3 = 0.872093M; K4 = 1.162791M; }
            if (epx == 45) { K1 = 0.503876M; K2 = 0.826873M; K3 = 0.775194M; K4 = 1.033592M; }
            if (epx == 50) { K1 = 0.453488M; K2 = 0.744186M; K3 = 0.697674M; K4 = 0.930233M; }
            if (epx == 55) { K1 = 0.412262M; K2 = 0.676533M; K3 = 0.634249M; K4 = 0.845666M; }
            if (epx == 60) { K1 = 0.377907M; K2 = 0.620155M; K3 = 0.581395M; K4 = 0.775194M; }
            if (epx == 65) { K1 = 0.348837M; K2 = 0.572451M; K3 = 0.536673M; K4 = 0.715564M; }
            if (epx == 70) { K1 = 0.32392M; K2 = 0.531561M; K3 = 0.498339M; K4 = 0.664452M; }
            if (epx == 75) { K1 = 0.302326M; K2 = 0.496124M; K3 = 0.465116M; K4 = 0.620155M; }
            if (epx == 80) { K1 = 0.28343M; K2 = 0.465116M; K3 = 0.436047M; K4 = 0.581395M; }
            if (epx == 85) { K1 = 0.266758M; K2 = 0.437756M; K3 = 0.410397M; K4 = 0.547196M; }
            if (epx == 90) { K1 = 0.251938M; K2 = 0.413437M; K3 = 0.387597M; K4 = 0.516796M; }
            if (epx == 95) { K1 = 0.238678M; K2 = 0.391677M; K3 = 0.367197M; K4 = 0.489596M; }
            if (epx == 100) { K1 = 0.226744M; K2 = 0.372093M; K3 = 0.348837M; K4 = 0.465116M; }
            if (epx == 105) { K1 = 0.215947M; K2 = 0.354374M; K3 = 0.332226M; K4 = 0.442968M; }
            if (epx == 110) { K1 = 0.206131M; K2 = 0.338266M; K3 = 0.317125M; K4 = 0.422833M; }
            if (epx == 115) { K1 = 0.197169M; K2 = 0.323559M; K3 = 0.303337M; K4 = 0.404449M; }
            if (epx == 120) { K1 = 0.188953M; K2 = 0.310078M; K3 = 0.290698M; K4 = 0.387597M; }
            if (epx == 125) { K1 = 0.181395M; K2 = 0.297674M; K3 = 0.27907M; K4 = 0.372093M; }
            if (epx == 130) { K1 = 0.174419M; K2 = 0.286225M; K3 = 0.268336M; K4 = 0.357782M; }
            if (epx == 135) { K1 = 0.167959M; K2 = 0.275624M; K3 = 0.258398M; K4 = 0.344531M; }
            if (epx == 140) { K1 = 0.16196M; K2 = 0.265781M; K3 = 0.249169M; K4 = 0.332226M; }
            if (epx == 145) { K1 = 0.156375M; K2 = 0.256616M; K3 = 0.240577M; K4 = 0.32077M; }
            if (epx == 150) { K1 = 0.151163M; K2 = 0.248062M; K3 = 0.232558M; K4 = 0.310078M; }
            if (epx == 155) { K1 = 0.146287M; K2 = 0.24006M; K3 = 0.225056M; K4 = 0.300075M; }
            if (epx == 160) { K1 = 0.141715M; K2 = 0.232558M; K3 = 0.218023M; K4 = 0.290698M; }
            if (epx == 165) { K1 = 0.137421M; K2 = 0.225511M; K3 = 0.211416M; K4 = 0.281889M; }
            if (epx == 170) { K1 = 0.133379M; K2 = 0.218878M; K3 = 0.205198M; K4 = 0.273598M; }
            if (epx == 175) { K1 = 0.129568M; K2 = 0.212625M; K3 = 0.199336M; K4 = 0.265781M; }
            if (epx == 180) { K1 = 0.125969M; K2 = 0.206718M; K3 = 0.193798M; K4 = 0.258398M; }
            if (epx == 185) { K1 = 0.122564M; K2 = 0.201131M; K3 = 0.188561M; K4 = 0.251414M; }
            if (epx == 190) { K1 = 0.119339M; K2 = 0.195838M; K3 = 0.183599M; K4 = 0.244798M; }
            if (epx == 195) { K1 = 0.116279M; K2 = 0.190817M; K3 = 0.178891M; K4 = 0.238521M; }
            if (epx == 200) { K1 = 0.113372M; K2 = 0.186047M; K3 = 0.174419M; K4 = 0.232558M; }
            if (epx == 205) { K1 = 0.110607M; K2 = 0.181509M; K3 = 0.170164M; K4 = 0.226886M; }
            if (epx == 210) { K1 = 0.107973M; K2 = 0.177187M; K3 = 0.166113M; K4 = 0.221484M; }
            if (epx == 215) { K1 = 0.105462M; K2 = 0.173067M; K3 = 0.16225M; K4 = 0.216333M; }
            if (epx == 220) { K1 = 0.103066M; K2 = 0.169133M; K3 = 0.158562M; K4 = 0.211416M; }
            if (epx == 225) { K1 = 0.100775M; K2 = 0.165375M; K3 = 0.155039M; K4 = 0.206718M; }
            if (epx == 230) { K1 = 0.098584M; K2 = 0.16178M; K3 = 0.151668M; K4 = 0.202224M; }
            if (epx == 235) { K1 = 0.096487M; K2 = 0.158337M; K3 = 0.148441M; K4 = 0.197922M; }
            if (epx == 240) { K1 = 0.094477M; K2 = 0.155039M; K3 = 0.145349M; K4 = 0.193798M; }
            if (epx == 245) { K1 = 0.092549M; K2 = 0.151875M; K3 = 0.142383M; K4 = 0.189843M; }
            if (epx == 250) { K1 = 0.090698M; K2 = 0.148837M; K3 = 0.139535M; K4 = 0.186047M; }
            if (epx == 255) { K1 = 0.088919M; K2 = 0.145919M; K3 = 0.136799M; K4 = 0.182399M; }
            if (epx == 260) { K1 = 0.087209M; K2 = 0.143113M; K3 = 0.134168M; K4 = 0.178891M; }
            if (epx == 265) { K1 = 0.085564M; K2 = 0.140412M; K3 = 0.131637M; K4 = 0.175516M; }
            if (epx == 270) { K1 = 0.083979M; K2 = 0.137812M; K3 = 0.129199M; K4 = 0.172265M; }
            if (epx == 275) { K1 = 0.082452M; K2 = 0.135307M; K3 = 0.12685M; K4 = 0.169133M; }
            if (epx == 280) { K1 = 0.08098M; K2 = 0.13289M; K3 = 0.124585M; K4 = 0.166113M; }
            if (epx == 285) { K1 = 0.079559M; K2 = 0.130559M; K3 = 0.122399M; K4 = 0.163199M; }
            if (epx == 290) { K1 = 0.078188M; K2 = 0.128308M; K3 = 0.120289M; K4 = 0.160385M; }
            if (epx == 295) { K1 = 0.076862M; K2 = 0.126133M; K3 = 0.11825M; K4 = 0.157667M; }
            if (epx == 300) { K1 = 0.075581M; K2 = 0.124031M; K3 = 0.116279M; K4 = 0.155039M; }
            if (epx == 305) { K1 = 0.074342M; K2 = 0.121998M; K3 = 0.114373M; K4 = 0.152497M; }
            if (epx == 310) { K1 = 0.073143M; K2 = 0.12003M; K3 = 0.112528M; K4 = 0.150038M; }
            if (epx == 315) { K1 = 0.071982M; K2 = 0.118125M; K3 = 0.110742M; K4 = 0.147656M; }
            if (epx == 320) { K1 = 0.070858M; K2 = 0.116279M; K3 = 0.109012M; K4 = 0.145349M; }
            if (epx == 325) { K1 = 0.069767M; K2 = 0.11449M; K3 = 0.107335M; K4 = 0.143113M; }
            if (epx == 330) { K1 = 0.06871M; K2 = 0.112755M; K3 = 0.105708M; K4 = 0.140944M; }
            if (epx == 335) { K1 = 0.067685M; K2 = 0.111073M; K3 = 0.104131M; K4 = 0.138841M; }
            if (epx == 340) { K1 = 0.066689M; K2 = 0.109439M; K3 = 0.102599M; K4 = 0.136799M; }
            if (epx == 345) { K1 = 0.065723M; K2 = 0.107853M; K3 = 0.101112M; K4 = 0.134816M; }
            if (epx == 350) { K1 = 0.064784M; K2 = 0.106312M; K3 = 0.099668M; K4 = 0.13289M; }
            if (epx == 355) { K1 = 0.063872M; K2 = 0.104815M; K3 = 0.098264M; K4 = 0.131019M; }
            if (epx == 360) { K1 = 0.062984M; K2 = 0.103359M; K3 = 0.096899M; K4 = 0.129199M; }
            if (epx == 365) { K1 = 0.062122M; K2 = 0.101943M; K3 = 0.095572M; K4 = 0.127429M; }
            if (epx == 370) { K1 = 0.061282M; K2 = 0.100566M; K3 = 0.09428M; K4 = 0.125707M; }
            if (epx == 375) { K1 = 0.060465M; K2 = 0.099225M; K3 = 0.093023M; K4 = 0.124031M; }
            if (epx == 380) { K1 = 0.05967M; K2 = 0.097919M; K3 = 0.091799M; K4 = 0.122399M; }
            if (epx == 385) { K1 = 0.058895M; K2 = 0.096648M; K3 = 0.090607M; K4 = 0.120809M; }
            if (epx == 390) { K1 = 0.05814M; K2 = 0.095408M; K3 = 0.089445M; K4 = 0.119261M; }
            if (epx == 395) { K1 = 0.057404M; K2 = 0.094201M; K3 = 0.088313M; K4 = 0.117751M; }



            if (CTPup.Text == "Poliuretano") { CdP = K1; CdT = K1; CdS = K1; }
            if (CTPup.Text == "Poliestireno") { CdP = K2; CdT = K2; CdS = K2; }
            if (CTPup.Text == "Poliuretano Suelo") { CdP = K3; CdT = K3; CdS = K3; }
            if (CTPup.Text == "Obra") { CdP = K4; CdT = K4; CdS = K4; }


            if (CTPus.Text == "Poliuretano") { CdP = K1; CdT = K1; CdS = K1; }
            if (CTPus.Text == "Poliestireno") { CdP = K2; CdT = K2; CdS = K2; }
            if (CTPus.Text == "Poliuretano Suelo") { CdP = K3; CdT = K3; CdS = K3; }
            if (CTPus.Text == "Obra") { CdP = K4; CdT = K4; CdS = K4; }
            //Perdidas por transmisión de paredes
            DtP = ta - Tip;
            PtPH = Math.Round((PtP * CdP * DtP), 4, MidpointRounding.ToEven);
            PtPD = Math.Round((PtPH * 24), 4, MidpointRounding.ToEven);
            //Perdidas por transmisión de techo
            DtT = (ta + 5) - Tip;
            PtT = DinP * DinA;
            PtTH = Math.Round((PtT * CdT * DtT), 4, MidpointRounding.ToEven);
            PtTD = PtTH * 24;
            // Perdidas de temperatura por SUELO
            DtS = (ta - 15) - Tip;
            PtS = DinP * DinA;
            PtSH = Math.Round((PtS * CdS * DtS), 4, MidpointRounding.ToEven);
            PtSD = PtSH * 24;

            //6.1 Perdidas de infirtración y aperturas de puertas
            decimal PtID;//Perdidas por infiltración y apertura de puerta diaria (W/dia)
            decimal Rn;//Renovaciones
            decimal PtIH;//Renovaciones por días
            decimal Rvolu;// Raiz cuadrada del volumen
            Rvolu = 0;
            decimal num = volu;
            double result = Math.Sqrt(Convert.ToDouble(num));
            Rvolu = Convert.ToDecimal(result);


            //Rvolu = Math.Sqrt(volu);

            if (Tip > -5) { Rn = 85 / Rvolu; PtID = Math.Round((volu * DtP * 0.66M / 0.86M * Rn), 4, MidpointRounding.ToEven); }
            else { Rn = 70 / Rvolu; PtID = Math.Round((volu * DtP * 0.66M / 0.86M * Rn), 4, MidpointRounding.ToEven); }
            //6.2 Calculo de perdidas por infiltración y apertura de puerta horas
            PtIH = Math.Round((PtID / 24), 2, MidpointRounding.ToEven);
            //6.3 Calculo de perdidas por Alumbrado y Personas días
            decimal PtAD;
            PtAD = Math.Round((PtPD + PtTD + PtID) * 0.12M, 4, MidpointRounding.ToEven);
            //6.4 Calculo de perdidas por Alumbrado y Personas horas
            decimal PtAH;
            PtAH = Math.Round((PtPH + PtTH + PtIH) * 0.12M, 4, MidpointRounding.ToEven);
            //7.1 Calculo de perdidas por Potencia Ventiladores en W diaria
            decimal PpE= 0;//Por Evaporador, Potencia Ventiladores, valor numérico entrado
            decimal PpD = 0;//Por Evaporador, potencia desescasrche, valor numerico de entreda
            decimal PpN = 0;//Número de evaporadores dentro de la cámara, sala de trabajo o tunel de congelacion
            try
            {
                PpE = Convert.ToInt16(float.Parse(TPrec.Text));
                PpN = Convert.ToInt16(float.Parse(TBcont.Text));
                PpD = Convert.ToInt32(float.Parse(TDesc.Text));
            }
            catch (FormatException ex)
            {
                // Manejo del error de formato
                Console.WriteLine("Error de formato: " + ex.Message);
            }
            catch (OverflowException ex)
            {
                // Manejo del error de desbordamiento
                Console.WriteLine("Error de desbordamiento: " + ex.Message);
            }

            int PpEL;//Condicion logica de existencia de valoe en PpE
            int PpNL;//Condición logica de exiastencia de valor en PpN
            int PpDL;//Condición logica de existewncia de valor en PpD
            int PpEN;//Condicion logica de existencia de valor en PpE Negado
            int PpNN;//Condición logica de exiastencia de valor en PpN Negado
            int PpDN;//Condición logica de existewncia de valor en PpD Negado

            decimal PpED;//Perdida por potencia en W/días
            decimal PpEH;//Perdida por potencia en W/horas
            decimal Suma;//Sumatoria de valores carga calculador: Perdodas días de Paredes + techo + suelo + infiltración y aperturas puertas + Alumbrado y personas
                         //+ Enfriar > 0ºC. +  Latente + Enfriar < 0ºC. + Respiración carga +  Respiración almacén
            decimal CdED;//Factor de Carga de refrigeración con Evaporador ausente
            PpNN = 0;
            PpEN = 0;
            PpDN = 0;
            PpED = 0;
            PpEH = 0;
            if (PpE > 0) { PpEN = 1; } else { PpEN = 0; }
            if (PpN > 0) { PpNN = 1; } else { PpNN = 0; }
            if (PpD > 0) { PpDN = 1; } else { PpDN = 0; }
            if (PpEN == 1)
                if (PpDN == 1) { PpED = Math.Round((PpEN * PpDN * PpE * Ht * PpN), 4, MidpointRounding.ToEven); PpEH = Math.Round((PpEN * PpDN * PpE * PpN), 4, MidpointRounding.ToEven); }
                else
                {
                    Suma = PtPD + PtTD + PtSD + PtID + PtAD;//Suma de cargas diarias
                    if (Tip < -5) { CdED = Math.Round((Suma * FcGC), 4, MidpointRounding.ToEven); PpED = CdED; }//Factor de carga congelación,Potencia de ventilación en W/diaria
                    if (Tip > -5) { CdED = Math.Round((Suma * FcGR), 4, MidpointRounding.ToEven); PpED = CdED; }//Factor de carga refrigeración,Potencia de ventilación en W/diaria
                }

            //7.2 Calculo de perdidas por Potencia Ventiladores en W horas
            decimal CdEH;//Potencia de ventiladores en W/horas
            CdEH = Math.Round((PpEN * PpDN * FcG), 4, MidpointRounding.ToEven);
            PpEH = Math.Round((PpE * PpN), 2, MidpointRounding.ToEven);
            //8 Perdidas por Desescarche Resit. o Gas Kw
            //8.1 Calculo de perdidas por Desescarche Resit. o Gas Kw diaria
            decimal PpDD;//Perdidas por Desescarche Resit. o Gas Kw W diaria
            decimal PpDH;//Perdidas por Desescarche Resit. o Gas Kw W horas
            PpDD = Math.Round((PpD * 1000 * 0.9M * PpN), 4, MidpointRounding.ToEven);
            PpDH = Math.Round((PpDD / Ht), 4, MidpointRounding.ToEven);
            //9 Perdidas por enfriamiento motores en cámara
            //9.1 Calculo de perdidas por enfriamiento motores en cámara diaria
            decimal PpMD;//Potencia de motores dentro de la camara diario en Kw
            decimal PpMH;//Potencia de motores dentro de la camara horas en Kw
            decimal PpM;//Potencia de motres dentro de la cámara en Kw
            PpMD = 0;
            PpMH = 0;
            PpM = Convert.ToInt16(float.Parse(TCvta.Text) * 1000);//Valor entrado potencia de motores en Kw
            PpMD = Math.Round((PpM * 0.9M * Ht), 4, MidpointRounding.ToEven);
            //9.2 Calculo de perdidas por enfriamiento motores en cámara horas
            PpMH = Math.Round((PpM * 0.9M), 4, MidpointRounding.ToEven);
            //10 Perdidas Carga de Género día enfriar > 0ºC.
            //Coeficientes de densidad de carga Según volumen de cámara
            decimal CdCV;//Coeficiente dencidad de carga
            CdCV = Convert.ToInt16(float.Parse(TCxp.Text));//% Densidad de carga
            decimal Q1;//Coeficiente dencidad
            decimal Q2;//Valor porcentual de carga
            decimal FdC;//Dencidad de carga Kg/m³
            Q1 = 0;
            Q2 = 0;
            FdC = 0;
            if (volu >= 0) { Q1 = 0.6M; Q2 = 20M; }
            if (volu >= 5) { Q1 = 0.61M; Q2 = 20.3M; }
            if (volu >= 10) { Q1 = 0.62M; Q2 = 20.7M; }
            if (volu >= 15) { Q1 = 0.63M; Q2 = 21M; }
            if (volu >= 20) { Q1 = 0.64M; Q2 = 21.3M; }
            if (volu >= 25) { Q1 = 0.65M; Q2 = 21.7M; }
            if (volu >= 30) { Q1 = 0.66M; Q2 = 22M; }
            if (volu >= 35) { Q1 = 0.67M; Q2 = 22.3M; }
            if (volu >= 40) { Q1 = 0.68M; Q2 = 22.7M; }
            if (volu >= 45) { Q1 = 0.69M; Q2 = 23M; }
            if (volu >= 50) { Q1 = 0.7M; Q2 = 23.3M; }
            if (volu >= 55) { Q1 = 0.71M; Q2 = 23.7M; }
            if (volu >= 60) { Q1 = 0.72M; Q2 = 24M; }
            if (volu >= 65) { Q1 = 0.73M; Q2 = 24.3M; }
            if (volu >= 70) { Q1 = 0.74M; Q2 = 24.7M; }
            if (volu >= 75) { Q1 = 0.75M; Q2 = 25M; }
            if (volu >= 80) { Q1 = 0.76M; Q2 = 25.3M; }
            if (volu >= 85) { Q1 = 0.77M; Q2 = 25.7M; }
            if (volu >= 90) { Q1 = 0.78M; Q2 = 26M; }
            if (volu >= 95) { Q1 = 0.79M; Q2 = 26.3M; }
            if (volu >= 100) { Q1 = 0.8M; Q2 = 26.7M; }
            if (volu >= 105) { Q1 = 0.81M; Q2 = 27M; }
            if (volu >= 110) { Q1 = 0.82M; Q2 = 27.3M; }
            if (volu >= 115) { Q1 = 0.83M; Q2 = 27.7M; }
            if (volu >= 120) { Q1 = 0.84M; Q2 = 28M; }
            if (volu >= 125) { Q1 = 0.85M; Q2 = 28.3M; }
            if (volu >= 130) { Q1 = 0.86M; Q2 = 28.7M; }
            if (volu >= 135) { Q1 = 0.87M; Q2 = 29M; }
            if (volu >= 140) { Q1 = 0.88M; Q2 = 29.3M; }
            if (volu >= 145) { Q1 = 0.89M; Q2 = 29.7M; }
            if (volu >= 150) { Q1 = 0.9M; Q2 = 30M; }
            if (volu >= 155) { Q1 = 0.91M; Q2 = 30.3M; }
            if (volu >= 160) { Q1 = 0.92M; Q2 = 30.7M; }
            if (volu >= 165) { Q1 = 0.93M; Q2 = 31M; }
            if (volu >= 170) { Q1 = 0.94M; Q2 = 31.3M; }
            if (volu >= 175) { Q1 = 0.95M; Q2 = 31.7M; }
            if (volu >= 180) { Q1 = 0.96M; Q2 = 32M; }
            if (volu >= 185) { Q1 = 0.97M; Q2 = 32.3M; }
            if (volu >= 190) { Q1 = 0.98M; Q2 = 32.7M; }
            if (volu >= 195) { Q1 = 0.99M; Q2 = 33M; }
            if (volu >= 200) { Q1 = 1M; Q2 = 33.3M; }
            if (volu >= 210) { Q1 = 1.004M; Q2 = 33.5M; }
            if (volu >= 220) { Q1 = 1.008M; Q2 = 33.6M; }
            if (volu >= 230) { Q1 = 1.012M; Q2 = 33.7M; }
            if (volu >= 240) { Q1 = 1.016M; Q2 = 33.9M; }
            if (volu >= 250) { Q1 = 1.02M; Q2 = 34M; }
            if (volu >= 260) { Q1 = 1.024M; Q2 = 34.1M; }
            if (volu >= 270) { Q1 = 1.028M; Q2 = 34.3M; }
            if (volu >= 280) { Q1 = 1.032M; Q2 = 34.4M; }
            if (volu >= 290) { Q1 = 1.036M; Q2 = 34.5M; }
            if (volu >= 300) { Q1 = 1.04M; Q2 = 34.7M; }
            if (volu >= 310) { Q1 = 1.044M; Q2 = 34.8M; }
            if (volu >= 320) { Q1 = 1.048M; Q2 = 34.9M; }
            if (volu >= 330) { Q1 = 1.052M; Q2 = 35.1M; }
            if (volu >= 340) { Q1 = 1.056M; Q2 = 35.2M; }
            if (volu >= 350) { Q1 = 1.06M; Q2 = 35.3M; }
            if (volu >= 360) { Q1 = 1.064M; Q2 = 35.5M; }
            if (volu >= 370) { Q1 = 1.068M; Q2 = 35.6M; }
            if (volu >= 380) { Q1 = 1.072M; Q2 = 35.7M; }
            if (volu >= 390) { Q1 = 1.076M; Q2 = 35.9M; }
            if (volu >= 400) { Q1 = 1.08M; Q2 = 36M; }
            if (volu >= 410) { Q1 = 1.084M; Q2 = 36.1M; }
            if (volu >= 420) { Q1 = 1.088M; Q2 = 36.3M; }
            if (volu >= 430) { Q1 = 1.092M; Q2 = 36.4M; }
            if (volu >= 440) { Q1 = 1.096M; Q2 = 36.5M; }
            if (volu >= 450) { Q1 = 1.1M; Q2 = 36.7M; }
            if (volu >= 460) { Q1 = 1.104M; Q2 = 36.8M; }
            if (volu >= 470) { Q1 = 1.108M; Q2 = 36.9M; }
            if (volu >= 480) { Q1 = 1.112M; Q2 = 37.1M; }
            if (volu >= 490) { Q1 = 1.116M; Q2 = 37.2M; }
            if (volu >= 500) { Q1 = 1.12M; Q2 = 37.3M; }
            if (volu >= 550) { Q1 = 1.124M; Q2 = 37.5M; }
            if (volu >= 600) { Q1 = 1.128M; Q2 = 37.6M; }
            if (volu >= 650) { Q1 = 1.132M; Q2 = 37.7M; }
            if (volu >= 700) { Q1 = 1.136M; Q2 = 37.9M; }
            if (volu >= 750) { Q1 = 1.14M; Q2 = 38M; }
            if (volu >= 800) { Q1 = 1.144M; Q2 = 38.1M; }
            if (volu >= 850) { Q1 = 1.148M; Q2 = 38.3M; }
            if (volu >= 900) { Q1 = 1.152M; Q2 = 38.4M; }
            if (volu >= 950) { Q1 = 1.156M; Q2 = 38.5M; }
            if (volu >= 1000) { Q1 = 1.16M; Q2 = 38.7M; }
            if (volu >= 1100) { Q1 = 1.164M; Q2 = 38.8M; }
            if (volu >= 1200) { Q1 = 1.168M; Q2 = 38.9M; }
            if (volu >= 1300) { Q1 = 1.172M; Q2 = 39.1M; }
            if (volu >= 1400) { Q1 = 1.176M; Q2 = 39.2M; }
            if (volu >= 1500) { Q1 = 1.18M; Q2 = 39.3M; }
            if (volu >= 1600) { Q1 = 1.184M; Q2 = 39.5M; }
            if (volu >= 1700) { Q1 = 1.188M; Q2 = 39.6M; }
            if (volu >= 1800) { Q1 = 1.192M; Q2 = 39.7M; }
            if (volu >= 1900) { Q1 = 1.196M; Q2 = 39.9M; }
            if (volu >= 2000) { Q1 = 1.2M; Q2 = 40M; }
            if (volu >= 2100) { Q1 = 1.204M; Q2 = 40.1M; }
            if (volu >= 2200) { Q1 = 1.208M; Q2 = 40.3M; }
            if (volu >= 2300) { Q1 = 1.212M; Q2 = 40.4M; }
            if (volu >= 2400) { Q1 = 1.216M; Q2 = 40.5M; }
            if (volu >= 2500) { Q1 = 1.22M; Q2 = 40.7M; }
            if (volu >= 2600) { Q1 = 1.224M; Q2 = 40.8M; }
            if (volu >= 2700) { Q1 = 1.228M; Q2 = 40.9M; }
            if (volu >= 2800) { Q1 = 1.232M; Q2 = 41.1M; }
            if (volu >= 2900) { Q1 = 1.236M; Q2 = 41.2M; }
            if (volu >= 3000) { Q1 = 1.24M; Q2 = 41.3M; }
            if (volu >= 3100) { Q1 = 1.244M; Q2 = 41.5M; }
            if (volu >= 3200) { Q1 = 1.248M; Q2 = 41.6M; }
            if (volu >= 3300) { Q1 = 1.252M; Q2 = 41.7M; }
            if (volu >= 3400) { Q1 = 1.256M; Q2 = 41.9M; }
            if (volu >= 3500) { Q1 = 1.26M; Q2 = 42M; }
            if (volu >= 3600) { Q1 = 1.264M; Q2 = 42.1M; }
            if (volu >= 3700) { Q1 = 1.268M; Q2 = 42.3M; }
            if (volu >= 3800) { Q1 = 1.272M; Q2 = 42.4M; }
            if (volu >= 3900) { Q1 = 1.276M; Q2 = 42.5M; }
            if (volu >= 4000) { Q1 = 1.28M; Q2 = 42.7M; }
            if (volu >= 4100) { Q1 = 1.284M; Q2 = 42.8M; }
            if (volu >= 4200) { Q1 = 1.288M; Q2 = 42.9M; }
            if (volu >= 4300) { Q1 = 1.292M; Q2 = 43.1M; }
            if (volu >= 4400) { Q1 = 1.296M; Q2 = 43.2M; }
            if (volu >= 4500) { Q1 = 1.3M; Q2 = 43.3M; }
            if (volu >= 4600) { Q1 = 1.304M; Q2 = 43.5M; }
            if (volu >= 4700) { Q1 = 1.308M; Q2 = 43.6M; }
            if (volu >= 4800) { Q1 = 1.312M; Q2 = 43.7M; }
            if (volu >= 4900) { Q1 = 1.316M; Q2 = 43.9M; }
            if (volu >= 5000) { Q1 = 1.32M; Q2 = 44M; }
            if (volu >= 500000) { Q1 = 3M; Q2 = 100M; }


            FdC = Q1 * DC;
            //if (Tip <= 0) ; if (ta <= 0)
            //{ CcL = 0; }

            //decimal CcLD;// Valor calculado del Calor espesifico antes de congelación del producto en Kc/Kg/ºC

            //if (Tip > 0) ; if (ta > 0)
            //{ CcLD = CcL/0.86M; }
            //if (Npr == 1) { FdC = 0; } else { FdC = Q1 * DC; }
            decimal Cg;//Carga de genero en Kg/día
            decimal PcGD1;// Valor calculado nivel 1, Perdidas Carga de Género > 0 w / días
            decimal PcGD2;// Valor calculado nivel 2, Perdidas Carga de Género > 0 w / días
            decimal PcGD3;// Valor calculado nivel 2, Perdidas Carga de Género > 0 w / días
            decimal PcGD;// Perdidas Carga de Género > 0 w / días
            decimal PcGH;//Perdidas Carga de Género > 0 w / horas
            decimal FdX;//Densidad de carga desplazada valor
            decimal T5;//
            FdX = 0;
            PcGH = 0;
            PcGD1 = 0;
            PcGD2 = 0;
            PcGD3 = 0;
            PcGD = 0;
            Cg = 0;
            T5 = 0;
            //=SI(M5*V5<R5;"Error Carga";SI(BU19>0;R5;0))
            //=SI(J21=1;0;G11*J20*(J21-1)/100) = Cg
            if (CdCV + 1 > 1) { Cg = volu * FdC * (CdCV) / 100; }

            if (volu * FdC < Cg)
            {
                MessageBox.Show("Error de carga, valor de carga incorrecto:");

            }


            else
            {
                if (volu * FdC > Cg)
                {
                    //SI(T5 > -2; X5; 0)
                    if (TeTR > -2) { PcGD1 = Cg; PcGD2 = CcL / 0.86M; }
                    //if (ta <= -2) { PcGD1 = 0; }
                    if (TeTR <= 0) { PcGD3 = 0; } if (Tip < -5) { T5 = TeTR - (-2); } if (T5 - Tip < 0) { PcGD3 = 0; } else { PcGD3 = TeTR - TeRF; }
                    //SI(T5<=0;0;SI(P5<-5;T5--2;SI(T5-P5<0;0;T5-U5))) = PcGD3
                    SPtp150.Text = Math.Round(FdC, 2, MidpointRounding.ToEven).ToString();//Capacidad de Carga Kg/m³
                    TTcmc9.Text = Math.Round(Cg / 24, 2, MidpointRounding.ToEven).ToString();//Carga por Genero Kg-días.

                    PcGD = Math.Round((PcGD1 * PcGD2 * PcGD3), 4, MidpointRounding.ToEven);
                    PcGH = Math.Round(PcGD / Ht, 4, MidpointRounding.ToEven);


                }



            }
            

            CeLD = 0;
            if (Tip < -2)
                if (TeTR > -2.1M) { CeLD = CeL / 0.86M; } // Calor latente KCal/Kg
            decimal PcLD;// Perdidas calor latente w / días
            decimal PcLD1;// Valor calculador de calor latente nivel 1
            decimal PcLD2;//Valor calculador de calor latente nivel 2
            decimal PcLH;// Perdidas calor latente w / hora
            PcLD1 = 0;
            PcLD2 = 0;

            //=SI(P5>-5;0;SI(T5>-2;Z5;0))
            //=SI(Tip>-5;0;SI(TeTR>-2;CeL/0,86;0))

            if (Tip > -5) { PcLD2 = 0; } else { if (TeTR > -2) { PcLD2 = CeLD; } else { PcLD2 = 0; } }

            if (volu * FdC < Cg)
            {
                MessageBox.Show("Error de carga, valor de carga incorrecto:");

            }
            else { if (PcLD2 > 0) { PcLD1 = Cg; } }
            //=SI(volu*FdC<Cg;"Error Carga";SI(BU20>0;Cg;0))
            //= BS20 * BU20

            PcLD = Math.Round(PcLD1 * PcLD2, 4, MidpointRounding.ToEven);
            PcLH = Math.Round(PcLD / Ht, 4, MidpointRounding.ToEven);

            //12 Perdidas Carga de Género día enfriar < 0ºC.

            decimal PcCD1;//Perdidas Carga de Género día enfriar < 0ºC w/días
            decimal PcCD;//Perdidas Carga de Género día enfriar < 0ºC w/ días
            decimal PcCH;//Perdidas Carga de Género día enfriar < 0ºC w/ horas

            PcCD1 = 0;
            PcCH = 0;

            //=SI(M5*V5<R5;"Error Carga";SI(BU21>0;R5;0))
            //=SI(P5>-5;0;Y5)

            //= SI(N5 = 1; ""; SI(Y(P5 < 0; T5 < 0); ""; BUSCARV(N5; Base; 11; 0)/ 0,86))
            if (Npr == 1) { FdX = 0; }
            if (Tip < 0)
                if (TeTR < 0) { FdX = CcF / 0.86M; }

            if (Tip > -5) { FdX = 0; }
            if (Tip < -5)
                if (Tip >= -2) { FdX = 0; }
            if (Tip < -2)
                if (volu * FdC < Cg)
                {
                    MessageBox.Show("Error de carga, valor de carga incorrecto:");

                }
                else { if (volu * FdC > Cg) { FdX = CcF / 0.86M; } }
            //=SI(P5>=-5;0;SI(T5-P5<=0;0;SI(T5<-2;T5-P5;-2-U5)))
            if (Tip >= -5) { PcCD1 = 0; }
            if (ta - Tip <= 0) { PcCD1 = 0; }
            if (TeTR < -2) { PcCD1 = TeTR - Tip; }
            else { PcCD1 = -2 - TeRF; }
            PcCD = Math.Round(FdX * PcCD1 * Cg, 4, MidpointRounding.ToEven);
            PcCH = Math.Round(PcCD / Ht, 4, MidpointRounding.ToEven);

            //13 Perdidas Respiración carga.
            decimal CrCD;//Valor de calor de respiración del producto Kcal/Tm/24h
            decimal CgKD;//Carga de genero kilogramo/dia
            decimal CrC1;//Valor de calor de respiración del producto Kcal/Tm/24h calculado
            decimal PcRD;//Perdidas Respiración carga w/días
            decimal PcRH;//Perdidas Respiración carga w/horas
            CgKD = 0;
            CrCD = 0;
            Crc = 0;
            CrC1 = 0;
            if (ta < -2) { CrCD = 0; }
            if (ta > 0 - 2) { CrCD = Crc / 0.86M; }
            if (volu * FdC < Cg)
            {
                MessageBox.Show("Error de carga, valor de carga incorrecto:");

            }
            else { if (Crc == 0) { CrCD = 0; } }
            if (Crc > 0) { CrCD = Cg / 1000; }
            PcRD = Math.Round(CrC1 * Crc, 4, MidpointRounding.ToEven);
            PcRH = Math.Round(PcRD / Ht, 4, MidpointRounding.ToEven);

            //14 Perdidas Respiración Almacén.
            decimal CrA1;//Calor de respiración de almacenamiento 
            decimal FrA;//Funcionamiento en respiración de alamacen
            decimal PrAD;//Perdidas Respiración Almacén w/días
            decimal PrAH;//Perdidas Respiración Almacén w/horas
            FrA = 0;
            CrA = 0;
            DC = 0;
            CrA1 = CrA / 0.86M;//Calor 1 Respiración Kcal/Tm/24h
            if (DC == 0)
                if (FrA == 0)
                    if (DC > 0) { FrA = volu * FdC / 1000; }
            PrAD = Math.Round(DC * FrA, 4, MidpointRounding.ToEven);
            PrAH = Math.Round(PrAD / Ht, 4, MidpointRounding.ToEven);

            //15 Corrección por evaporador ausente en W
            decimal suma2;//Valor calculado de suma valores diarios
            suma2 = 0;
            if (PpE == 0)
                //15.1 Calculo de perdidas por Potencia Ventiladores en W diaria
                if (PpD == 0) { suma2 = PtPD + PtTD + PtSD + PtID + PtAD + PcGD + PcLD + PcCD + PcRD + PrAD; }
            if (Tip < -5) { CdED = Math.Round(suma2 * FcGC, 4, MidpointRounding.ToEven); } else { CdED = Math.Round(suma2 * FcGR, 2, MidpointRounding.ToEven); }
            if (PpED == 0) { PpED = CdED; }
            //15.2 Calculo de perdidas por Potencia Ventiladores en W horas
            if (PpE == 0)
                if (PpD == 0) { CdEH = Math.Round((PcGH + PcLH + PcCH + PcRH + PrAH) * 0.2M, 4, MidpointRounding.ToEven); }
            if (PpEH == 0) { PpEH = CdEH; }

            //16 Calculo Subtotal base.
            decimal CsBD;// subtotal base por días
            decimal CsBH;// Subtotal bade por horas
            CsBD = Math.Round(PtPD + PtTD + PtSD + PtID + PtAD + PpED + PpDD + PpMD + PcGD + PcLD + PcCD + PcRD + PrAD, 4, MidpointRounding.ToEven);
            CsBH = Math.Round(PtPH + PtTH + PtSH + PtIH + PtAH + PpEH + PpMH + PcGH + PcLH + PcCH + PcRH + PrAH, 4, MidpointRounding.ToEven);

            //17 Calculo total con margen seguridad.
            decimal CtMD;//Calculo total con margen seguridad w/días
            decimal CtMH;//Calculo total con margen seguridad w / hora
            decimal MsG;//Margen de seguridad horario;
            MsG = Convert.ToInt16(float.Parse(TMuc.Text));
            CtMD = (MsG + 100) / 100 * CsBD;
            CtMH = 1.1M * CsBH;

            //18 Calculo total demanda promedio por horas.
            decimal CtDP;//Calculo total demanda promedio diaria watt
            decimal CtDH;//Calculo total demanda por horas watt
            decimal CtCD;//Calculo total en conservación demanda
            CtDP = 0;
            CtCD = 0;

           
            if (Tip < -5)
                if (Ht >= Hp14) { CtDP = CtMD / Ht; }
            if (Tip >= -4.9M); if (Ht >= Hp12) { CtDP = CtMD / Ht; } else { CtDP = CtMH; }



            //19 Calculo total en conservación demanda

            if (Tip <= -5)
            { CtCD = Math.Round((PtPD + PtTD + PtSD + PtID + PtAD + PpED + PpDD + PpMD) / Hp14 * (MsG + 100) / 100, 2, MidpointRounding.ToEven); }//Conservaciôn demanda
            else
            { CtCD = Math.Round((PtPD + PtTD + PtSD + PtID + PtAD + PpED + PpDD + PpMD) / Hp12 * (MsG + 100) / 100, 2, MidpointRounding.ToEven); }//Conservaciôn demanda




            //20 Potencia frigorífica a instalar, PfIn Kw en (DpOB= “Salas de elaboración”), de lo contrario
            
            decimal PfIn;//Potencia frigorífica a instalar
            decimal DpOB;//Datos de potencia “Salas de elaboración”
            decimal DpCC;// Datos de potencia de cámaras “Conservación”
            decimal DpCR;//
            decimal DpCG;//
            decimal DpTG;//
            PfIn = 0;
           

            if (Ctpd.Text == "") { PfIn = 0; MessageBox.Show("Producto no seleccionado:"); } else { if (CtDP > CtCD) { PfIn = CtDP; } else { PfIn = CtCD; } }


            //Mensaje de error de entradas
            if (epx == 0) { MessageBox.Show("Aislamiento pared no seleccionado:"); }
            if (Tip < -5)
                if (esx == 0) { MessageBox.Show("Aislamiento pared no seleccionado:"); }
            if (ta > 50) { MessageBox.Show("Temperatura de entrada no es correcta:"); }
            if (Ht > 22) { MessageBox.Show("Horas de trabajo no apropiadas:"); }

           
            TCmod.Text = Math.Round(PfIn, 2, MidpointRounding.ToEven).ToString();
            decimal PfIQ;// Valor en Kw de la potencia calculada PfIn
            PfIQ = 0;
            PfIQ = Math.Round(PfIn / 1000, 2, MidpointRounding.ToEven);
            TFWe.Text = Math.Round(PfIQ, 2, MidpointRounding.ToEven).ToString();
            decimal m2;
            m2 = Math.Round(DinP * DinA, 2, MidpointRounding.ToEven);
            TBdes.Text = m2.ToString();//Superficie de cámara
            TBfec.Text = ta.ToString();//%Humedad de la Càmara
            TDmce.Text = Hur.ToString();//%Humedad relativa ambinte
            TCmodp.Text = Hurp.ToString();//%Humedad en la cámara.
            
            decimal TrPD;//Toneladas de renovación de producto diario
            TrPD = 0;

            if (TeRF > -18) { TrPD = Math.Round((FdC * volu * FcGR * CdCV) / 1000, 2, MidpointRounding.ToEven); }
            if (TeRF <= -18) { TrPD = Math.Round((FdC * volu * FcGC * CdCV) / 1000, 2 ,MidpointRounding.ToEven); }
            if (-32 < TeRF)
                if (TeRF <= -18) { TrPD = Math.Round((FdC* volu *FcGC * CdCV)/ 1000, 2, MidpointRounding.ToEven); }
            decimal CeMP;//Calor especifico medio del producto
            CeMP = 0;
            if (CeMP == 0) { CeMP = 1; } else { CeMP = (CcL / 0.86M) * (CcF / 0.86M); }
            decimal TipC;//Temperatura del producto por temperatura final
            TipC = 0;
            TipC = TeRF + 2;
            decimal CtPT;//Sumatoria de perdidas por Paredes, Techo y suelos
            CtPT = 0;
            CtPT = PtPH + PtTH + PtSH;
            decimal PpCP;//Perdidas por carga de productos
            PpCP = 0;
            //= "Carga de Género día " & DECIMAL(BS19 / (1 - BS11 * V5) * 100; 0)*-1 & " %"
            TCmodp.Text = Hur1.ToString() + " / " + Hur2.ToString();

            //Base de transferencia de datos Form1 a Form5 utilizando Clase Datos
            
            Datos.TFWe = Convert.ToDecimal(PfIQ);
            Datos.TCentd = Convert.ToInt16(TCentx.Text);
            Datos.Hur11 = Convert.ToUInt16(Hur1);//Emisor de valores
            Datos.Hur21 = Convert.ToUInt16(Hur2);//Emisor de valores
            

            //BASE DE DATOS TRANSFERENCIA DE FROM5 A FROM1 EN CLASE DE DATOS


            if (RBun13.Checked == true)
            {

                StreamWriter sw = new StreamWriter(@"C:\01 OFERTAS 2023\00 OFERTAS\" + TNP.Text + "_C" + CClit1.Text + "_" + String.Format("{0:yyyy.MM.dd.hhmm}", DateTime.Now) + ".Txt", true);

                sw.WriteLine("CALCULO DE CARGA TERMICA PROYECTO:               " + TNP.Text + "");
                sw.WriteLine("01 Dimensiones interiores:                       " + DinP + " m x " + DinA + "  m x " + DinH + " m");
                sw.WriteLine("02 Superficie:                                   " + TBdes.Text + " m²");
                sw.WriteLine("03 Volumen:                                      " + volu + " m³");
                sw.WriteLine("04 Lugar de instalación sistema frigorífico:     " + CTLup.Text + " /Ubicación");
                sw.WriteLine("05 Temperatura ambiente:                         " + ta + " °C");
                sw.WriteLine("06 Humedad relativa en localidad de instalación: " + TDmce.Text + " %Hr");
                sw.WriteLine("07 Humedad en el recinto:                        " + Hur1.ToString() + "/" + Hur2.ToString() + " %Hr");
                sw.WriteLine("08 Temperatura mínima de genero:                 " + Tip + " °C");
                sw.WriteLine("09 Temperatura de entrada del producto:          " + TeTR + " °C");
                sw.WriteLine("10 Temperatura final del producto:               " + TCentx.Text + " °C");
                sw.WriteLine("11 Carga estimada diaria por producto:           " + TCxp.Text + " %");
                sw.WriteLine("12 Toneladas renovación de producto diario:      " + TrPD + " Ton-días");
                sw.WriteLine("13 Flujo transmisión térmica paneles aislantes:  " + Math.Round((0.00697333M * PtIH), 2, MidpointRounding.ToEven) + " W/h-m");
                sw.WriteLine("14 Relación de Carga por volumen de cámara:      " + Math.Round(FdC, 2, MidpointRounding.ToEven) + " Kg/m³");
                sw.WriteLine("15 Calor específico medio del producto:          " + Math.Round(CeMP, 2, MidpointRounding.ToEven) + " W/kg");
                sw.WriteLine("16 Densidad " + CdCV + " % Carga x Prod:                   " + Math.Round(((Cg / (1 - volu * FdC) * 100) * -1) * Cg * CcL / 0.86M * (TeTR - TeRF) / 100, 2, MidpointRounding.ToEven) + " W/días");
                sw.WriteLine("17 Enfriamiento de los productos:                " + Math.Round(Cg, 2, MidpointRounding.ToEven) + " Kg/día");
                sw.WriteLine("18 Enfriamiento de los embalajes:                " + Math.Round(CsBH - (CsBH * 0.8M), 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("19 Enfriamiento de los pallets:                  " + Math.Round((CsBH * 0.1M), 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("20 Calor por transmisión (paredes-techos-suelo): " + Math.Round(CtPT, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("21 N° ren/día Infiltración aperturas puertas:    " + Math.Round(Rn, 2, MidpointRounding.ToEven) + " Ren/día");
                sw.WriteLine("22 Calor por Infiltración y apertura puerta:     " + Math.Round(PtIH, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("23 Calor aportado por ventiladores evaporador:   " + Math.Round(PpEH, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("24 Calor aportado por la iluminación:            " + Math.Round(PtAH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("25 Carga térmica mantenimiento total:            " + Math.Round(CsBH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("26 Margen de seguridad:                          " + Math.Round(MsG, 2, MidpointRounding.ToEven) + " %");
                sw.WriteLine("27 Carga demandada por el sistema:               " + Math.Round(CtMH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("28 Carga total con margen de seguridad:          " + Math.Round(CtDP, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("29 Funcionamiento                                " + Ht + " h/días");
                sw.WriteLine("30 Potencia frigorífica a instalar:              " + TCmod.Text + " W; (" + PfIQ + " Kw)");
                sw.WriteLine("31 Capacidad de Carga en producto:               " + Math.Round((Cg / 0.3225M) / CeMP * Ht / 24, 1, MidpointRounding.ToEven) + " Kg-C");
                sw.WriteLine("");
                sw.WriteLine("DATOS GENERALES DE EVAPORACION CALCULADOS");
                sw.WriteLine("32 Temperatura de Evaporación:                   " + Datos.Tcmc8 + "°C");
                sw.WriteLine("33 Modelo de Evaporador:                         " + Datos.TValv + " p/u");
                sw.WriteLine("34 Fabricante Evaporador:                        " + Datos.TCsist1);
                sw.WriteLine("35 Tipo de Evaporador:                           " + Datos.CSerie);
                sw.WriteLine("36 Cantidad Evaporadores                         " + Datos.TPem + " p/u");
                sw.WriteLine("37 Potencia de evaporación entregada:            " + Datos.CTEvap + " kW");
                sw.WriteLine("38 Rendimienro de evaporación:                   " + Datos.TRdinE);
                sw.WriteLine("39 Diferencial de temperatura calculado:         " + Datos.DTpd + " °C");
                sw.WriteLine("40 Diferencial de temperatura fijado:            " + Datos.DTsl + " °C");
                sw.WriteLine("41 Temperatura de Condensación fijada:           " + Datos.Tcmc6 + " °C");
                sw.WriteLine("");
                sw.WriteLine("DATOS GENERALES DE EXPANSION CALCULADOS");
                sw.WriteLine("42 Fabricante Válvula Expansión:                 " + Datos.TPrs);
                sw.WriteLine("43 Tipo de Valvula Expansión:                    " + Datos.TPex);
                sw.WriteLine("44 Rendimiento de Expansión:                     " + Datos.PValE + " kW");
                sw.WriteLine("45 Modelo de Vávula Expansión:                   " + Datos.Modex);
                sw.WriteLine("46 Potencia Máx de la Válvula Expansión:         " + Datos.CrgaM + " kW");
                sw.WriteLine("47 Refrigerante del Sistema:                     " + Datos.Rfcalx);
                sw.WriteLine("48 Pressión de Condensación:                     " + Datos.TCPbar + " bar");
                sw.WriteLine("49 Diferencial de Presión:                       " + Datos.DTpress + " bar");
                sw.WriteLine("50 Presión del Liquido                           " + Datos.TLPbar + " bar");
                sw.WriteLine("50 Temperatura del liquido                       " + Datos.TLtemp + " °C");
                sw.WriteLine("51 Modelo del Orificio:                          " + Datos.tev2);
                sw.WriteLine("52 Codigo del Orificio:                          " + Datos.tev3);
                sw.WriteLine("");

                sw.Close();
                TextReader Leer = new StreamReader(@"C:\01 OFERTAS 2023\00 OFERTAS\" + TNP.Text + "_C" + CClit1.Text + "_" + String.Format("{0:yyyy.MM.dd.hhmm}", DateTime.Now) + ".Txt");
                MessageBox.Show(Leer.ReadToEnd());
                Leer.Close();
            }
            else 
            {
                File.Delete(@"C:\01 OFERTAS 2023\00 OFERTAS\DATOS.Txt");
                StreamWriter sw = new StreamWriter(@"C:\01 OFERTAS 2023\00 OFERTAS\DATOS.Txt", true);

                sw.WriteLine("CALCULO DE CARGA TERMICA PROYECTO:               " + TNP.Text + "");
                sw.WriteLine("");
                sw.WriteLine("01 Dimensiones interiores:                       " + DinP + " m x " + DinA + "  m x " + DinH + " m");
                sw.WriteLine("02 Superficie:                                   " + TBdes.Text + " m²");
                sw.WriteLine("03 Volumen:                                      " + volu + " m³");
                sw.WriteLine("04 Lugar de instalación sistema frigorífico:     " + CTLup.Text + " /Ubicación");
                sw.WriteLine("05 Temperatura ambiente:                         " + ta + " °C");
                sw.WriteLine("06 Humedad relativa en localidad de instalación: " + TDmce.Text + " %Hr");
                sw.WriteLine("07 Humedad en el recinto:                        " + Hur1.ToString() + "/" + Hur2.ToString() + " %Hr");
                sw.WriteLine("08 Temperatura mínima de genero:                 " + Tip + " °C");
                sw.WriteLine("09 Temperatura de entrada del producto:          " + TeTR + " °C");
                sw.WriteLine("10 Temperatura final del producto:               " + TCentx.Text + " °C");
                sw.WriteLine("11 Carga estimada diaria por producto:           " + TCxp.Text + " %");
                sw.WriteLine("12 Toneladas renovación de producto diario:      " + TrPD + " Ton-días");
                sw.WriteLine("13 Flujo transmisión térmica paneles aislantes:  " + Math.Round((0.00697333M * PtIH), 2, MidpointRounding.ToEven) + " W/h-m");
                sw.WriteLine("14 Relación de Carga por volumen de cámara:      " + Math.Round(FdC, 2, MidpointRounding.ToEven) + " Kg/m³");
                sw.WriteLine("15 Calor específico medio del producto:          " + Math.Round(CeMP, 2, MidpointRounding.ToEven) + " W/kg");
                sw.WriteLine("16 Densidad " + CdCV + " % Carga x Prod:                   " + Math.Round(((Cg / (1 - volu * FdC) * 100) * -1) * Cg * CcL / 0.86M * (TeTR - TeRF) / 100, 2, MidpointRounding.ToEven) + " W/días");
                sw.WriteLine("17 Enfriamiento de los productos:                " + Math.Round(Cg, 2, MidpointRounding.ToEven) + " Kg/día");
                sw.WriteLine("18 Enfriamiento de los embalajes:                " + Math.Round(CsBH - (CsBH * 0.8M), 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("19 Enfriamiento de los pallets:                  " + Math.Round((CsBH * 0.1M), 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("20 Calor por transmisión (paredes-techos-suelo): " + Math.Round(CtPT, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("21 N° ren/día Infiltración aperturas puertas:    " + Math.Round(Rn, 2, MidpointRounding.ToEven) + " Ren/día");
                sw.WriteLine("22 Calor por Infiltración y apertura puerta:     " + Math.Round(PtIH, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("23 Calor aportado por ventiladores evaporador:   " + Math.Round(PpEH, 2, MidpointRounding.ToEven) + " W/h");
                sw.WriteLine("24 Calor aportado por la iluminación:            " + Math.Round(PtAH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("25 Carga térmica mantenimiento total:            " + Math.Round(CsBH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("26 Margen de seguridad:                          " + Math.Round(MsG, 2, MidpointRounding.ToEven) + " %");
                sw.WriteLine("27 Carga demandada por el sistema:               " + Math.Round(CtMH, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("28 Carga total con margen de seguridad:          " + Math.Round(CtDP, 2, MidpointRounding.ToEven) + " W");
                sw.WriteLine("29 Funcionamiento                                " + Ht + " h/días");
                sw.WriteLine("30 Potencia frigorífica a instalar:              " + TCmod.Text + " W; (" + PfIQ + " Kw)");
                sw.WriteLine("31 Capacidad de Carga en producto:               " + Math.Round((Cg / 0.3225M) / CeMP * Ht/24 , 1, MidpointRounding.ToEven) + " Kg-C");

                sw.Close();

                TextReader Leer = new StreamReader(@"C:\01 OFERTAS 2023\00 OFERTAS\DATOS.Txt");
                MessageBox.Show(Leer.ReadToEnd());
                Leer.Close();
            }



            Datos.TNP1 = TNP.Text;
            Datos.CClit2 = CClit1.Text;
           
           
        }

        internal void GuardarValores(double tConD, double tLPbar)
        {
            throw new NotImplementedException();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            Form7 myPassForm = new Form7(this);
            myPassForm.ShowDialog();
        }

        private void button9_Click(object sender, EventArgs e)
        {
            var reportManager = new ReportManager();
            reportManager.GenerateThermalLoadReport(
                Datos.TNP1, // Nombre del Proyecto
                Datos.DinP + " mm" + Datos.DinA + " mm" + Datos.DinH + " mm", // Dimensiones
                    Datos.TBdes, // Superficie
                    Datos.volu, // Volumen
                    Datos.CTLup, // Lugar de instalación
                    Datos.ta // Temperatura ambiente
                );
        }
    }



    public class CExcelWork
    {
        //Variables de Excel
        Microsoft.Office.Interop.Excel.Application ExcelApp = null;
        public _Workbook myWorkBook = null;
        public _Worksheet myWorkSheet = null;

        Range myRange = null;
        public Array myArray = null;
        Array myValues = null;

        string local;
        COferta Oferta;
        Thread myThread;
        Form1 thisForm;

        //---------------Methods declarations

        public CExcelWork(COferta OFERTA, string LOCAL, Form1 FORM)
        {
            local = LOCAL;
            Oferta = OFERTA;
            thisForm = FORM;
        }

        public void SetThread(Thread thisThread)
        {
            myThread = thisThread;
        }

        // Crea la aplicacion de Excel
        /// <summary>
        /// Crea la aplicaion de Excel
        /// </summary>
        public void CreateApplication()
        {
            ExcelApp = new Microsoft.Office.Interop.Excel.Application();
            ExcelApp.Visible = false;
            ExcelApp.UserControl = false;
            ExcelApp.DisplayAlerts = false;
        }

        public void Disconnect()
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            try
            {
                Marshal.FinalReleaseComObject(myWorkSheet);
            }
            catch { }
            myWorkBook.Close(false, Type.Missing, Type.Missing);
            Marshal.FinalReleaseComObject(myWorkBook);
            ExcelApp.Quit();
            Marshal.FinalReleaseComObject(ExcelApp);
        }

        // FORMATO DE IMPRESION
        void PrintOut(Object From, Object to, Object Copies, Object Preview, Object ActivePrinter, Object PrintToFile, Object Collate, Object PrToFileName)
        {
            From = true;
            to = true;
            PrintToFile = true;
            PrToFileName = true;
        }
        

        //Abre el archivo
        /// <summary> 
        /// Abre el archivo de excel
        /// </summary>
        /// <returns>Devuelve false si no es posible abrirlo</returns>
        public bool OpenFile(string local, bool cuc)
        {
            try
            {
                if (cuc)
                {
                    myWorkBook = ExcelApp.Workbooks.Open(local + @"\fnstptecuc.tpt",
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing,
                        Type.Missing);
                }
                   
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                 "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false; 
            }
            //-----------------------------------------------------------------
            return true;

        }

        public void ChangeSheet(int num)
        {
            try
            {
                myWorkSheet = (_Worksheet)myWorkBook.Worksheets[num];
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }        

        /// <summary>
        /// Genera la oferta incertando los datos en la plantilla
        /// </summary>
        /// <param name="datos">Es un un objeto de la estructura datos</param>
        /// <returns>devuelve false si no se ha podido procesar</returns>
        public void GenerarOferta()
        {
           
            
            
        }
        
        

    }

    [Serializable]
    public class CCam
    {
        string NC;
        string Temp;
        string Largo;
        string Ancho;
        string Alto;
        string Volu;
        string TP;
        string DT;
        string CE;
        string CF;
        string FW;
        string Qfw;
        string Cmod;
        string Cmodd;
        string Cmodp;
        string Desc;
        string Prec;
        string Qfep;
        string Scdro;
        string Spsi;
        string Stemp;
        string Apsi;
        string Emevp;
        string Sup;
        string IT;
        string DEC;
        string DECE;
        string DECH;
        string DECF;           
        string Centx;
        string Cdin;
        string Cxp;
        string Digt;
        string Ccion;
        string Castre;
        string Cpcion;
        string Cfrio;
        string Ceq1;
        string Ceq2;
        string Ceq3;
        string Refrig;
        string TMuc;
        string TMevp;
        string TSol;
        string TValv;
        string TCvta;
        string TCuadro;
        string CBexpo;
       
        string CSumi;
        string CTCond;
        string CTamb;
        string Ctpd;
        string Cnoff1;
        string Coff1;
        string Cnoff2;
        string Cnoff3;
        string Coff3;
        string Coff4;
        string Coff5;
        string Cnoff6;
        string Coff6;
        string CTEvap;
       
        string CodValv;
        string Tmos;
        string TInc;
        string TInev;
        string TVnev;
        string TIned;
        string TIncd;
        string TIpv;
        string TIcc;
        string TQevp;
        string TQevpd;
        string TTint;
        string TEquip;
        string TCint1;
        string TCint2;
        string TCint3;
        string TMcc;
        string TCmce;
        string TPmce;
        string TDmce;
        string TDlq1;
        string TDlq2;
        string TDlq3;
        string TDlq11;
        string TDlq21;
        string TDlq31;
        string TLcc;
        string TLss;
        string TPcmc;
        string TPcond;
        string TModex;
        string TPvq;
        string TPls;
        string TPosc;
        string TPsq;
        string TPcq;
        string TPcy;
        string TPcp;
        string TPex;
        string TPrs;
        string TCsist;
        string TPem;
        string TPnt;
        string TPml;
        string TCt150;
        string TCt04;
        string TCt150m;
        string TCt04m;
        string TTcmc1;
        string TTcmc2;
        string TTcmc3;
        string FWe;
        string Sptp75;
        string Sptp74;
        string TTcmc6;
        string TTcmc8;
        string TTcmc9;
        string CBps;
        string CTPup;
        string CTPus;
        string CTLup;
        string TLtemp;  
        string TLPbar;
        string DTpress;
        string TCPbar;
        string TEPbar;

        /// <summary>
        /// Contador que me lleva la constancia de las imagenes
        /// </summary>
        int pk;

        public CCam(string nc, string temp, string largo, string ancho, string alto, string volu, string cf, string fw, string qfw, string cmod, string cmodd, 
            string cmodp, string desc, string prec, string qfep, string scdro, string spsi, string stemp, string apsi, string emevp, string sup,
            string centx, string cxp, string caster, string cpcion, string cfrio, string ceq1, string ceq2, string ceq3, 
            string tmuc, string tmevp, string tsol, string tvalv, string tcvta, string tcuadro, string cbexpo, string csumi, string ctpd, string coff3, string cnoff6, 
            string ctevap, string codvalv, string tinc, string tinev, string tvnev, string tined, string tincd, string tipv, string ticc, string tqevp, string tqevpd, string ttint, string tequip, 
            string tcint1, string tcint2, string tcint3, string tmcc, string tcmce, string tpmce, string tdmce, string tlcc, string tlss, string tpcmc, string tpcond,
            string tmodex, string tpvq, string tpls, string tposc, string tpsq, string tpcq, string tpcy, string tpcp, string tpex, string tprs, string tcsist, string tpem, 
            string tpnt, string tpml, string tct150, string tct04, string tct150m, string tct04m, string sptp75, string sptp74, string ttcmc1, string ttcmc2, string ttcmc3, 
            string fwe,  string ttcmc6, string ttcmc8, string ttcmc9, string cbps,string ctpup, string ctpus, string ctlup, string tltemp, string tlpbar, string dtpress, string tcpbar, string tepbar, int PK) 
        {
            NC = nc;
            Temp = temp;
            Largo = largo;
            Ancho = ancho;
            Alto = alto;
            Volu = volu;
            CF = cf;
            FW = fw;
            Qfw = qfw;
            Cmod = cmod;
            Cmodd = cmodd;
            Cmodp = cmodp;
            Desc = desc;
            Prec = prec;
            Qfep = qfep;
            Scdro = scdro;
            Spsi = spsi;
            Stemp = stemp;
            Apsi = apsi;
            Emevp = emevp;
            Sup = sup;
            Centx = centx;
            Cxp = cxp;
            Centx = centx;
            Castre = caster;
            Cpcion = cpcion;
            Cfrio = cfrio;
            Ceq1 = ceq1;
            Ceq2 = ceq2;
            Ceq3 = ceq3;
            TMuc = tmuc;
            TMevp = tmevp;
            TSol = tsol;
            TValv = tvalv;
            TCvta = tcvta;
            TCuadro = tcuadro;
            CBexpo = cbexpo;
            CSumi = csumi;
            Ctpd = ctpd;
            Coff3 = coff3;
            Cnoff6 = cnoff6;
            CTEvap = ctevap;
            CodValv = codvalv;
            TInc = tinc;
            TInev = tinev;
            TVnev = tvnev;
            TIned = tined;
            TIncd = tincd;
            TIpv = tipv;
            TIcc = ticc;
            TQevp = tqevp;
            TQevpd = tqevpd;
            TTint = ttint;
            TEquip = tequip;
            TCint1 = tcint1;
            TCint2 = tcint2;
            TCint3 = tcint3;
            TMcc = tmcc;
            TCmce = tcmce;
            TPmce = tpmce;
            TDmce = tdmce;
            TLcc = tlcc;
            TLss = tlss;
            TPcmc = tpcmc;
            TPcond = tpcond;
           TModex = tmodex;
            TPvq = tpvq;
            TPls = tpls;
            TPosc = tposc;
            TPsq = tpsq;
            TPcq = tpcq;
            TPcy = tpcy;
            TPcp = tpcp;
            TPex = tpex;
            TPrs = tprs;
            TCsist = tcsist;
            TPem = tpem;
            TPnt = tpnt;
            TPml = tpml;
            TCt150 = tct150;
            TCt04 = tct04;
            TCt04m = tct04m;
            TCt150m = tct150m;
            Sptp75 = sptp75;
            Sptp74 = sptp74;
            TTcmc1 = ttcmc1;
            TTcmc2 = ttcmc2;
            TTcmc3 = ttcmc3;
            FWe = fwe;
            TTcmc6 = ttcmc6;
            TTcmc8 = ttcmc8;
            TTcmc9 = ttcmc9;
            CBps = cbps;
            CTPup = ctpup;
            CTPus = ctpus;
            CTLup = ctlup;
            TLtemp = tltemp;
            TLPbar = tlpbar;
            DTpress = dtpress;
            TCPbar = tcpbar;
            TEPbar = tepbar;
            pk = PK;
        }

        //Aqui vienen los gets!
        public string GetNC()
        {
            return NC;
        }

        public string GetTemp()
        {
            return Temp;
        }

        public string GetLargo()
        {
            return Largo;
        }

        public string GetAncho()
        {
            return Ancho;
        }

        public string GetAlto()
        {
            return Alto;
        }
        public string GetVolu()
        {
            return Volu;
        }

        public string GetTP()
        {
            return TP;
        }

        public string GetDT()
        {
            return DT;
        }

        public string GetCE()
        {
            return CE;
        }

        public string GetCF()
        {
            return CF;
        }
        public string GetFW()
        {
            return FW;
        }
        public string GetQfw()
        {
            return Qfw;
        }

        public string GetCmod()
        {
            return Cmod;
        }
        public string GetCmodd()
        {
            return Cmodd;
        }
        public string GetCmodp()
        {
            return Cmodp;
        }
        public string GetDesc()
        {
            return Desc;
        }
        public string GetPrec()
        {
            return Prec;
        }
        public string GetQfep()
        {
            return Qfep;
        }
        public string GetScdro()
        {
            return Scdro;
        }

        public string GetSpsi()
        {
            return Spsi;
        }
        public string GetStemp()
        {
            return Stemp;
        }
        public string GetApsi()
        {
            return Apsi;
        }
        public string GetEmevp()
        {
            return Emevp;
        }

        public string GetSUP()
        {
            return Sup;
        }

        public string GetIT()
        {
            return IT;
        }

        public string GetDEC()
        {
            return DEC;
        }
        public string GetDECE()
        {
            return DECE;
        }
        public string GetDECH()
        {
            return DECH;
        }
        public string GetDECF()
        {
            return DECF;
        }

              

        public string GetCentx()
        {
            return Centx;
        }
        public string GetCdin()
        {
            return Cdin;
        }
        public string GetCxp()
        {
            return Cxp;
        }
       

        public string GetRefrig()
        {
            return Refrig;
        }
        public string GetDigt()
        {
            return Digt;
        }

        public string GetCcion()
        {
            return Ccion;
        }
        public string GetCastre()
        { 
            return Castre;
        }
        public string GetCpcion()
        {
            return Cpcion;
        }
        public string GetCfrio()
        {
            return Cfrio;
        }
        public string GetCeq1()
        {
            return Ceq1;
        }
        public string GetCeq2()
        {
            return Ceq2;
        }
        public string GetCeq3()
        {
            return Ceq3;
        }
        

       

        public string GetTMuc()
        {
            return TMuc;
        }

        public string GetTMevp()
        {
            return TMevp;
        }
        
        
        public string GetTSol()
        {
            return TSol;
        }

        public string GetTValv()
        {
            return TValv;
        }
        

        public string GetTCvta()
        {
            return TCvta;
        }

       
        public string GetTCuadro()
        {
            return TCuadro;
        }
        
        
       

        public string GetCBexpo()
        {
            return CBexpo;
        }
 
        
        public string GetCSumi()
        {
            return CSumi;
        }
        
        public string GetCTCond()
        {
            return CTCond;
        }
        public string GetCTamb()
        {
            return CTamb;
        }
        
        public string GetCtpd()
        {
            return Ctpd;
        }

        public string GetCnoff1()
        {
            return Cnoff1;
        }
        public string GetCoff1()
        {
            return Coff1;
        }
        public string GetCnoff2()
        {
            return Cnoff2;
        }
       
        public string GetCnoff3()
        {
            return Cnoff3;
        }
        public string GetCoff3()
        {
            return Coff3;
        }
        
        public string GetCoff4()
        {
            return Coff4;
        }
       
        public string GetCoff5()
        {
            return Coff5;
        }
        public string GetCnoff6()
        {
            return Cnoff6;
        }
        public string GetCoff6()
        {
            return Coff6;
        }

        
       

        public string GetCTEvap()
        {
            return CTEvap;
        }
        public string GetTInc()
        {
            return TInc;
        }

        public string GetTInev()
        {
            return TInev;
        }

        public string GetTVnev()
        {
            return TVnev;
        }
        public string GetTIned()
        {
            return TIned;
        }
        public string GetTIncd()
        {
            return TIncd;
        }
        public string GetTIpv()
        {
            return TIpv;
        }
        public string GetTIcc()
        {
            return TIcc;
        }

        public string GetTQevp()
        {
            return TQevp;
        }
        public string GetTQevpd()
        {
            return TQevpd;
        }
        
        public string GetTTint()
        {
            return TTint;
        }

        public string GetTEquip()
        {
            return TEquip;
        }

        public string GetTCint1()
        {
            return TCint1;
        }

        public string GetTCint2()
        {
            return TCint2;
        }

        public string GetTCint3()
        {
            return TCint3;
        }
        public string GetTMcc()
        {
            return TMcc;
        }
        public string GetTCmce()
        {
            return TCmce;
        }
        public string GetTPmce()
        {
            return TPmce;
        }
        public string GetTDmce()
        {
            return TDmce;
        }
       
       
        public string GetTDlq1()
        {
            return TDlq1;
        }
        public string GetTDlq2()
        {
            return TDlq2;
        }
        public string GetTDlq3()
        {
            return TDlq3;
        }
        public string GetTDlq11()
        {
            return TDlq11;
        }
        public string GetTDlq21()
        {
            return TDlq21;
        }
        public string GetTDlq31()
        {
            return TDlq31;
        }
       
       
      
        public string GetTLcc()
        {
            return TLcc;
        }
        public string GetTLss()
        {
            return TLss;
        }
        public string GetTPcmc()
        {
            return TPcmc;
        }
       
        
        
       
      
        public string GetTTcmc1()
        {
            return TTcmc1;
        }
        public string GetTTcmc2()
        {
            return TTcmc2;
        }
        public string GetTTcmc3()
        {
            return TTcmc3;
        }
       
       
        public string GetTTcmc6()
        {
            return TTcmc6;
        }
       
        public string GetTTcmc8()
        {
            return TTcmc8;
        }
        public string GetTTcmc9()
        {
            return TTcmc9;
        }
        
        public string GetCBps()
        {
            return CBps;
        }
        public string GetTPcond()
        {
            return TPcond;
        }
        public string GetTModex()
        {
            return TModex;
        }
        public string GetTPvq()
        {
            return TPvq;
        }
        public string GetTPls()
        {
            return TPls;
        }
        public string GetTPosc()
        {
            return TPosc;
        }
        public string GetTPsq()
        {
            return TPsq;
        }
        public string GetTPcq()
        {
            return TPcq;
        }
       
        public string GetTPcy()
        {
            return TPcy;
        }
        public string GetTPcp()
        {
            return TPcp;
        }
        public string GetTPex()
        {
            return TPex;
        }
        public string GetTPrs()
        {
            return TPrs;
        }
        public string GetTCsist()
        {
            return TCsist;
        }
        public string GetTPem()
        {
            return TPem;
        }
        public string GetTPnt()
        {
            return TPnt;
        }
        public string GetTPml()
        {
            return TPml;
        }
        
     
        public string GetTCt150()
        {
            return TCt150;
        }
        public string GetTCt04()
        {
            return TCt04;
        }
       
       
        public string GetTCt150m()
        {
            return TCt150m;
        }
        public string GetTCt04m()
        {
            return TCt04m;
        }
       
        public string GetFWe()
        {
            return FWe;
        }
        public string GetCTPup()
        {
            return CTPup;
        }
        public string GetCTPus()
        {
            return CTPus;
        }
        public string GetCTLup()
        {
            return CTLup;
        }
        public string GetTLtemp() 
        {
            return TLtemp;
        }
        public string GetTLPbar()
        {
            return TLPbar;
        }
        public string GetDTpress()
        {
            return DTpress;
        }
        public string GetTCPbar()
        {
            return TCPbar;
        }
        public string GetTEPbar()
        {
            return TEPbar;
        }
        public int GetPK()
        {
            return pk;
        }

        //*********************************
        public void setVolu(string volu)
        {
            Volu = volu;
        }
        public void setCF(string cf)
        {
            CF = cf;
        }
        public void setFW(string fw)
        {
            FW = fw;
        }
        public void setQfw(string qfw)
        {
            Qfw = qfw;
        }

        public void setCmod(string cmod)
        {
            Cmod = cmod;
        }
        public void setCmodd(string cmodd)
        {
            Cmodd = cmodd;
        }
        public void setCmodp(string cmodp)
        {
            Cmodp = cmodp;
        }
        public void setDesc(string desc)
        {
            Desc = desc;
        }
        public void setPrec(string prec)
        {
            Prec = prec;
        }
       
        public void setQfep(string qfep)
        {
            Qfep = qfep;
        }
        public void setScdro(string scdro)
        {
            Scdro = scdro;
        }
        public void setSpsi(string spsi)
        {
            Spsi = spsi;
        }
        public void setStemp(string stemp)
        {
            Stemp = stemp;
        }
        public void setApsi(string apsi)
        {
            Apsi = apsi;
        }
        public void setEmevp(string emevp)
        {
            Emevp = emevp;
        }

     

        public void SetMuc(string muc)
        {
            TMuc = muc;
        }

        public void SetMevp(string mevp)
        {
            TMevp = mevp;
        }
                
        
        public void SetSol(string sol)
        {
            TSol = sol;
        }

        public void SetValv(string valv)
        {
            TValv = valv;
        }
        
        public void SetCvta(string cvta)
        {
            TCvta = cvta;
        }
       
        public void SetCuadro(string cuadro)
        {
            TCuadro = cuadro;
        }
       
        
       

        public void Setcbexpo(string cbexpo)
        {
            CBexpo = cbexpo;
        }
  
        
        
        public void SetCSumi(string csumi)
        {
            CSumi = csumi;
        }
        
       
        public void SetCTamb(string ctamb)
        {
            CTamb = ctamb;
        }
        public void SetCtpd(string ctpd)
        {
            Ctpd = ctpd;
        }
        public void SetCnoff1(string cnoff1)
        {
            Cnoff1 = cnoff1;
        }

        public void SetCoff1(string coff1)
        {
            Coff1 = coff1;
        }
        public void SetCnoff2(string cnoff2)
        {
            Cnoff2 = cnoff2;
        }

       
        public void SetCnoff3(string cnoff3)
        {
            Cnoff3 = cnoff3;
        }

        public void SetCoff3(string coff3)
        {
            Coff3 = coff3;
        }
       
        public void SetCoff4(string coff4)
        {
            Coff4 = coff4;
        }
       
        public void SetCoff5(string coff5)
        {
            Coff5 = coff5;
        }
        public void SetCnoff6(string cnoff6)
        {
            Cnoff6 = cnoff6;
        }
        public void SetCoff6(string coff6)
        {
            Coff6 = coff6;
        }

       
       

        public void SetCTEvap(string ctevap)
        {
            CTEvap = ctevap;
        }
        
        public void SetCodValv(string codvalv)
        {
            CodValv = codvalv;
        }
       
        
        public string GetCodValv()
        {
            return CodValv;
        }
        public string GetTmos()
        {
            return Tmos;
        }

        public void SetTInc(string tinc)
        {
            TInc = tinc;
        }

        public void SetTInev(string tinev)
        {
            TInev = tinev;
        }

        public void SetTVnev(string tvnev)
        {
            TVnev = tvnev;
        }
        public void SetTIned(string tined)
        {
            TIned = tined;
        }
        public void SetTIncd(string tincd)
        {
            TIncd = tincd;
        }
        public void SetTIpv(string tipv)
        {
            TIpv = tipv;
        }
        public void SetTIcc(string ticc)
        {
            TIcc = ticc;
        }
        public void SetTQevp(string tqevp)
        {
            TQevp = tqevp;
        }
        public void SetTQevpd(string tqevpd)
        {
            TQevpd = tqevpd;
        }
      
        public void setTTint(string ttint)
        {
            TTint = ttint;
        }
        
        public void SetTEquip(string tequip)
        {
            TEquip = tequip;
        }

        public void SetTCint1(string tcint1)
        {
            TCint1 = tcint1;
        }

        public void SetTCint2(string tcint2)
        {
            TCint2 = tcint2;
        }

        public void SetTCint3(string tcint3)
        {
            TCint3 = tcint3;
        }
        public void SetTMcc(string tmcc)
        {
            TMcc = tmcc;
        }
        public void SetTCmce(string tcmce)
        {
            TCmce = tcmce;
        }
        public void SetTPmce(string tpmce)
        {
            TPmce = tpmce;
        }
        public void SetTDmce(string tdmce)
        {
            TDmce = tdmce;
        }
        
       
        public void SetTDlq1(string tdlq1)
        {
            TDlq1 = tdlq1;
        }
        public void SetTDlq2(string tdlq2)
        {
            TDlq2 = tdlq2;
        }
        public void SetTDlq3(string tdlq3)
        {
            TDlq3 = tdlq3;
        }
        public void SetTDlq11(string tdlq11)
        {
            TDlq11 = tdlq11;
        }
        public void SetTDlq21(string tdlq21)
        {
            TDlq21 = tdlq21;
        }
        public void SetTDlq31(string tdlq31)
        {
            TDlq31 = tdlq31;
        }
        
        public void SetTLcc(string tlcc)
        {
            TLcc = tlcc;
        }
        public void SetTLss(string tlss)
        {
            TLss = tlss;
        }
        public void SetTPcmc(string tpcmc)
        {
            TPcmc = tpcmc;
        }
        
       
     
        
        public void SetTPcond(string tpcond)
        {
            TPcond = tpcond;
        }
        public void SetTModex(string tmodex)
        {
            TModex = tmodex;
        }
        public void SetTPvq(string tpvq)
        {
            TPvq = tpvq;
        }
        public void SetTPls(string tpls)
        {
            TPls = tpls;
        }
        public void SetTPosc(string tposc)
        {
            TPosc = tposc;
        }
        public void SetTPsq(string tpsq)
        {
            TPsq = tpsq;
        }
        public void SetTPcq(string tpcq)
        {
            TPcq = tpcq;
        }
      
        public void SetTPcy(string tpcy)
        {
            TPcy = tpcy;
        }
        public void SetTPcp(string tpcp)
        {
            TPcp = tpcp;
        }
        public void SetTPex(string tpex)
        {
            TPex = tpex;
        }
        public void SetTPrs(string tprs)
        {
            TPrs = tprs;
        }
        public void SetTCsist(string tcsist)
        {
            TCsist = tcsist;
        }
        public void SetTPem(string tpem)
        {
            TPem = tpem;
        }
        public void SetTPnt(string tpnt)
        {
            TPnt = tpnt;
        }
        public void SetTPml(string tpml)
        {
            TPml = tpml;
        }
       
       
       
        public void SetTCt150(string tct150)
        {
            TCt150 = tct150;
        }
        public void SetTCt04(string tct04)
        {
            TCt04 = tct04;
        }
       
        
        public void SetTCt150m(string tct150m)
        {
            TCt150m = tct150m;
        }
        public void SetTCt04m(string tct04m)
        {
            TCt04m = tct04m;
        }
        
        public void SetTTcmc1(string ttcmc1)
        {
            TTcmc1 = ttcmc1;
        }
        public void SetTTcmc2(string ttcmc2)
        {
            TTcmc2 = ttcmc2;
        }
        public void SetTTcmc3(string ttcmc3)
        {
            TTcmc3 = ttcmc3;
        }
       
       
        public void SetTTcmc6(string ttcmc6)
        {
            TTcmc6 = ttcmc6;
        }
        
        public void SetTTcmc8(string ttcmc8)
        {
            TTcmc8 = ttcmc8;
        }
        public void SetTTcmc9(string ttcmc9)
        {
            TTcmc9 = ttcmc9;
        }
        public void SetFWe(string fwe)
        {
            FWe = fwe;
        }
       
        public void SetCBps(string cbps)
        {
            CBps = cbps;
        }
        public void SetCTPup(string ctpup)
        {
            CTPup = ctpup;
        }
        public void SetCTPus(string ctpus)
        {
            CTPus = ctpus;
        }
        public void SetCTLup(string ctlup)
        {
            CTLup = ctlup;
        }
        public void SetTLtemp(string tltemp)
        {
            TLtemp = tltemp;
        }
        public void SetTLPbar(string tlpbar)
        {
            TLPbar = tlpbar;
        }
        public void SetDTpress(string dtpress)
        {
            DTpress = dtpress;
        }
        public void SetTCPbar(string tcpbar)
        {
            TCPbar = tcpbar;
        }
        public void SetTEPbar(string tepbar)
        {
            TEPbar = tepbar;
        }
        public void SetPK(int PK)
        {
            pk = PK;
        }
    }

    [Serializable]
    public class COferta
    {
        string NP;
        string Ref;
        string NO;
        string Cmat;
        string Fecha;
        string Dsu130;
        string Dsu230;
        string Dsu330;
        string Dsu110;
        string Dsu210;
        string Dsu310;
       
       
        
        string Dmc2;
        string Dmc3;
        string Dmc4;
        string Dmc5;
        string Dmc6;
       
        string Dmc8;
        string Lcc;
        string Lss;
        string Pcmc;
        string Pcond;
        string Modex;
        string Pvq;
        string Pls;
        string Posc;
        string Psq;
        string Pcq;
        string Pcy;
        string Pcp;
        string Pex;
        string Prs;
        string Csist;
        string Pem;
        string Pnt;
        string Pml;
        string Ct150;
        string Ct04;
        string Ct150m;
        string Ct04m;
        string Inc;
        string Inev;
        string Vnev;
        string Ined;
        string Incd;
        string Ipv;
        string Icc;
        string Qevp;
        string Qevpd;
        string Tint;
        string Equip;
        string Sumi;
        string DT;
        CCam[] Camaras;
        int CantCam;
        int cont;
        string Digt;
        string Ccion;
        string Castre;
        string Cpcion;
        string Cfrio;
        string Ceq1;
        string Ceq2;
        string Ceq3;
        string inc;
        string Lugar;
        string Clit;
        string Clit1;
        string Clitm;
        string Bscu;
        string Bcont;
        string Bcos;
        string Bdir;
        string Benv;
        string Bpo;
        string Bfec;
        string Bdes;
        string Cdc;
        string Flet;
        string Cgr;
        string Intr;
        string Desct;
        string Ncont;
        string Fwc;
        string Dcmc1;
        string Dcmc2;
        string Dcmc3;
        string Tcmc1;
        string Tcmc2;
        string Tcmc3;
        string Tcmc4;
        string Tcmc5;
        string Tcmc6;
        string Tcmc7;
        string Tcmc8;
        string Tcmc9;
        string FWe;
        string Bps;
        string TPup;
        string TPus;
        string TLup;
        string Ltemp; 
        string LPbar;
        string Tpress;
        string CPbar;
        string EPbar;


        public COferta(string np, string REF, string no, string cmat, string fecha, string lcc, string lss, string pcmc, string pcond, string modex, string pvq, string pls,
            string posc, string psq, string pcq, string pcy, string pcp, string pex, string prs, string csist, string pem, string pnt, string pml, string ct150,
            string ct04, string ct150m, string ct04m, string spt75, string spt74, string inc, string inev, string vnev, string ined, string incd, string ipv, string icc, 
            string qevp, string qevpd, string equip, string sumi, string text,int cantcam, string castre, string cpcion, string cfrio, string ceq1, string ceq2, string ceq3,
            string lugar, string clit, string clit1, string bscu, string bcont, string bcos, string bdir, string benv, string bpo, string bfec, string bdes, string cdc,
            string flet, string cgr, string intr, string desct, string ncont, string tcmc1, string tcmc2, string tcmc3, string fwe, string tcmc6, string tcmc8, string tcmc9, 
            string bps, string tpup, string tpus, string tlup, string ltemp, string lpbar, string tpress, string cpbar, string epbar)
        {
            NP = np;
            Ref = REF;
            NO = no;
            Cmat = cmat;
            Fecha = fecha;
          
            Lcc = lcc;
            Lss = lss;
            Pcmc = pcmc;
           
            Pcond = pcond;
            Modex = modex;
            Pvq = pvq;
            Pls = pls;
            Posc = posc;
            Psq = psq;
            Pcq = pcq;
           
            Pcy = pcy;
            Pcp = pcp;
            Pex = pex;
            Prs = prs;
            Csist = csist;
            Pem = pem;
            Pnt = pnt;
           
           
            Ct150 = ct150;
            Ct04 = ct04;
            
            Ct04m = ct04m;
            
            Ct150m = ct150m;
           
            
            Inc = inc;
            Inev = inev;
            Vnev = vnev;
            Ined = ined;
            Incd = incd;
            Ipv = ipv;
            Icc = icc;
            Qevp = qevp;
            Qevpd = qevpd;
           
          
            Equip = equip;
            Sumi = sumi;
           
            Camaras = new CCam[100];
            CantCam = cantcam;
            cont = 0;
            
            
           
            Castre = castre;
            Cpcion = cpcion;
            Cfrio = cfrio;
            Ceq1 = ceq1;
            Ceq2 = ceq2;
            Ceq3 = ceq3;
           
            inc = Inc;
           
          
            Lugar = lugar;
            Clit = clit;
            Clit1 = clit1;
         
            
           
           
            Bscu = bscu;
            Bcont = bcont;
            Bcos = bcos;
            Bdir = bdir;
            Benv = benv;
            Bpo = bpo;
            Bfec = bfec;
            Bdes = bdes;
            Clit = clit;
            Cdc = cdc;
            Flet = flet;
            Cgr = cgr;
            Intr = intr;
            Desct = desct;
            Ncont = ncont;
           
            Tcmc1 = tcmc1;
            Tcmc2 = tcmc2;
            Tcmc3 = tcmc3;
           
           
            FWe = fwe;
            Tcmc6 = tcmc6;
            
            Tcmc8 = tcmc8;
            Tcmc9 = tcmc9;
           
            Bps = bps;
            TPup = tpup;
            TPus = tpus;
            TLup = tlup;
            Ltemp = ltemp;
            LPbar = lpbar;
            Tpress = tpress;
            CPbar = cpbar;
            EPbar = epbar;
        }
        /// <summary>
        /// Anade una c'amara nueva
        /// </summary>
        /// <param name="camara">camara que se va a anadir</param>
        /// <returns>devuelve falso si no se puede anadir mas</returns>
        public bool AddCam(CCam camara)
        {
            if (cont == CantCam)
                return false;
            else
            {
                Camaras[cont] = camara;
                cont++;
                return true;
            }
        }
        /// <summary>
        /// Borra la camara dado la posicion y organiza el arreglo
        /// </summary>
        /// <param name="num">Posicion en el arreglo</param>
        public void BorrarCam(int num)
        {
            for (int i = (num - 1); i < cont; i++)
            {
                if (i < (cont - 1))
                    Camaras[i] = Camaras[i + 1];
                else
                {
                    Camaras[i] = null;
                    cont--;
                }
            }
        }
        /// <summary>
        /// Actualiza la camara actual
        /// </summary>
        /// <param name="camara">camara a actualizar</param>
        /// <param name="num">numero de la c'amara</param>
        public void actualizar(CCam camara, int num)
        {
            Camaras[num - 1] = camara;
        }
        /// <summary>
        /// Devuelve la camara almaceneda en el indice indicado
        /// </summary>
        /// <param name="num">Indice de la camara deseada</param>
        /// <returns>Devuelve la camara deseada (CCam)</returns>
        public CCam GetCam(int num)
        {
            return Camaras[num];
        }

        public void SetCam(CCam cam, int pos)
        {
            this.Camaras[pos] = cam;
        }
        /// <summary>
        /// Devuelve el contador de camaras
        /// </summary>
        /// <returns>int cont</returns>
        public int GetCont()
        {
            return cont;
        }
        /// <summary>
        /// Devuelve un valor referente a la cantidad de camaras actual 
        /// mas exacta.
        /// </summary>
        /// <returns></returns>
        public int GetNumCam()
        {
            int i = 0;
            while (true)
            {
                CCam exCam;
                
                exCam = this.GetCam(i);
                if (exCam == null)
                    return (i);
                i++;
            }
        }

        public string GetNP()
        {
            return NP;
        }

        public string GetREF()
        {
            return Ref;
        }

        public string GetNO()
        {
            return NO;
        }
        public string GetCmat()
        {
            return Cmat;
        }
       
      
        public string GetDsu130()
        {
            return Dsu130;
        }
        public string GetDsu230()
        {
            return Dsu230;
        }
        public string GetDsu330()
        {
            return Dsu330;
        }
        public string GetDsu110()
        {
            return Dsu110;
        }
        public string GetDsu210()
        {
            return Dsu210;
        }
        public string GetDsu310()
        {
            return Dsu310;
        }
       
     
        public string GetDmc2()
        {
            return Dmc2;
        }
        public string GetDmc3()
        {
            return Dmc3;
        }

        public string GetDmc4()
        {
            return Dmc4;
        }

        public string GetDmc5()
        {
            return Dmc5;
        }

        public string GetDmc6()
        {
            return Dmc6;
        }
       
        public string GetDmc8()
        {
            return Dmc8;
        }
        public string GetLcc()
        {
            return Lcc;
        }
        public string GetLss()
        {
            return Lss;
        }
        public string GetPcmc()
        {
            return Pcmc;
        }
       
       
       
        public string GetPcond()
        {
            return Pcond;
        }
        public string GetModex()
        {
            return Modex;
        }
        public string GetPvq()
        {
            return Pvq;
        }
        public string GetPls()
        {
            return Pls;
        }
        public string GetPosc()
        {
            return Posc;
        }
        public string GetPsq()
        {
            return Psq;
        }
        public string GetPcq()
        {
            return Pcq;
        }
       
        public string GetPcy()
        {
            return Pcy;
        }
        public string GetPcp()
        {
            return Pcp;
        }
        public string GetPex()
        {
            return Pex;
        }
        public string GetPrs()
        {
            return Prs;
        }
        public string GetCsist()
        {
            return Csist;
        }

        public string GetPem()
        {
            return Pem;
        }
        public string GetPnt()
        {
            return Pnt;
        }
        public string GetPml()
        {
            return Pml;
        }
        
       
      
       
        
        
       
        public string GetCt150()
        {
            return Ct150;
        }
        public string GetCt04()
        {
            return Ct04;
        }
       
        
        public string GetCt04m()
        {
            return Ct04m;
        }
        public string GetCt150m()
        {
            return Ct150m;
        }

        public string GetInc()
        {
            return Inc;
        }

        public string GetInev()
        {
            return Inev;
        }
        public string GetVnev()
        {
            return Vnev;
        }

        public string GetIned()
        {
            return Ined;
        }
        public string GetIncd()
        {
            return Incd;
        }
        public string GetIpev()
        {
            return Ipv;
        }
        public string GetIcc()
        {
            return Icc;
        }

        public string GetQevp()
        {
            return Qevp;
        }
        public string GetQevpd()
        {
            return Qevpd;
        }
       
        
        public string GetTint()
        {
            return Tint;
        }

        public string GetEquip()
        {
            return Equip;
        }
        public string GetDT()
        {
            return DT;
        }
        
        public int GetCantCam()
        {
            return CantCam;
        }

        public string GetFecha()
        {
            return Fecha;
        }

        public string GetBscu()
        {
            return Bscu;
        }
        public string GetBcont()
        {
            return Bcont;
        }
        public string GetBcos()
        {
            return Bcos;
        }
        public string GetBdir()
        {
            return Bdir;
        }
        public string GetBenv()
        {
            return Benv;
        }
        public string GetBpo()
        {
            return Bpo;
        }
        public string GetBfec()
        {
            return Bfec;
        }
        public string GetBdes()
        {
            return Bdes;
        }
        public string GetClit()
        {
            return Clit;

        }
        public string GetClit1()
        {
            return Clit1;
        }
        public string GetClitm()
        {
            return Clitm;
        }
        public string GetCdc()
        {
            return Cdc;
        }
        public string GetFlet()
        {
            return Flet;
        }
        public string GetCgr()
        {
            return Cgr;
        }
        public string GetIntr()
        {
            return Intr;
        }
        public string GetDesct()
        {
            return Desct;
        }
        public string GetNcont()
        {
            return Ncont;
        }
      

       
     
        public string GetFwc()
        {
            return Fwc;
        }
        public string GetDcmc1()
        {
            return Dcmc1;
        }
        public string GetDcmc2()
        {
            return Dcmc2;
        }
        public string GetDcmc3()
        {
            return Dcmc3;
        }
        public string GetTcmc1()
        {
            return Tcmc1;
        }
        public string GetTcmc2()
        {
            return Tcmc2;
        }
        public string GetTcmc3()
        {
            return Tcmc3;
        }
        public string GetTcmc4()
        {
            return Tcmc4;
        }
        public string GetTcmc5()
        {
            return Tcmc5;
        }
        public string GetFWe()
        {
            return FWe;
        }
        public string GetTcmc6()
        {
            return Tcmc6;
        }
        public string GetTcmc7()
        {
            return Tcmc7;
        }
        public string GetTcmc8()
        {
            return Tcmc8;
        }
        public string GetTcmc9()
        {
            return Tcmc9;
        }
       
        public string GetBps()
        {
            return Bps;
        }
        public string GetTPup()
        {
            return TPup;
        }
        public string GetTPus()
        {
            return TPus;
        }
        public string GetTLup()
        {
            return TLup;
        }

        public string GetLtemp() 
        {
            return Ltemp;
        }
        public string GetLPbar()
        {
            return LPbar;
        }
        public string GetTpress()
        {
            return Tpress;
        }
        public string GetCPbar()
        {
            return CPbar;
        }
        public string GetEPbar()
        {
            return EPbar;
        }

        //-----------------------------------Los Sets

        public void SetNP(string np)
        {
            NP = np;
        }

        public void SetREF(string REF)
        {
            Ref = REF;
        }

        public void SetNO(string no)
        {
            NO = no;
        }
        public void SetCmat(string cmat)
        {
            Cmat = cmat;
        }
       
       
       
       
        
        public void SetLcc(string lcc)
        {
            Lcc = lcc;
        }
        public void SetLss(string lss)
        {
            Lss = lss;
        }
        public void SetPcmc(string pcmc)
        {
            Pcmc = pcmc;
        }
       
        public void SetPcond(string pcond)
        {
            Pcond = pcond;
        }
        public void SetModex(string modex)
        {
            Modex = modex;
        }
        public void SetPvq(string pvq)
        {
            Pvq = pvq;
        }
        public void SetPls(string pls)
        {
            Pls = pls;
        }
        public void SetPosc(string posc)
        {
            Posc = posc;
        }
        public void SetPsq(string psq)
        {
            Psq = psq;
        }
        public void SetPcq(string pcq)
        {
            Pcq = pcq;
        }
      
        public void SetPcy(string pcy)
        {
            Pcy = pcy;
        }
        public void SetPcp(string pcp)
        {
            Pcp = pcp;
        }
        public void SetPex(string pex)
        {
            Pex = pex;
        }
        public void SetPrs(string prs)
        {
            Prs = prs;
        }
        public void SetCsist(string csist)
        {
            Csist = csist;
        }
        public void SetPem(string pem)
        {
            Pem = pem;
        }
        public void SetPnt(string pnt)
        {
            Pnt = pnt;
        }
        public void SetPml(string pml)
        {
            Pml = pml;
        }
       
       
       
       
        public void SetCt150(string ct150)
        {
            Ct150 = ct150;
        }
        public void SetCt04(string ct04)
        {
            Ct04 = ct04;
        }
      
        public void SetCt150m(string ct150m)
        {
            Ct150m = ct150m;
        }
        public void SetCt04m(string ct04m)
        {
            Ct04m = ct04m;
        }
       
      
        public void SetInc(string inc)
        {
            Inc = inc;
        }
        public void SetInev(string inev)
        {
            Inev = inev;
        }
        public void SetVnev(string vnev)
        {
            Vnev = vnev;
        }
        public void SetIned(string ined)
        {
            Ined = ined;
        }
        public void SetIncd(string incd)
        {
            Incd = incd;
        }
        public void SetIpv(string ipv)
        {
            Ipv = ipv;
        }
        public void SetIcc(string icc)
        {
            Icc = icc;
        }

        public void SetQevp(string qevp)
        {
            Qevp = qevp;
        }
        public void SetQevpd(string qevpd)
        {
            Qevpd = qevpd;
        }
       

        public void SetTint(string tint)
        {
            Tint = tint;
        }
        public void SetEquip(string equip)
        {
            Equip = equip;
        }

       

        public void SetCantCam(int cc)
        {
            CantCam = cc;
        }

        public void SetFecha(string fecha)
        {
            Fecha = fecha;
        }

        public void SetBscu(string bscu)
        {
            Bscu = bscu;
        }
        public void SetBcont(string bcont)
        {
            Bcont = bcont;
        }
        public void SetBcos(string bcos)
        {
            Bcos = bcos;
        }
        public void SetBdir(string bdir)
        {
            Bdir = bdir;
        }
        public void SetBenv(string benv)
        {
            Benv = benv;
        }
        public void SetBpo(string bpo)
        {
            Bpo = bpo;
        }
        public void SetBfec(string bfec)
        {
            Bfec = bfec;
        }
        public void SetBdes(string bdes)
        {
            Bdes = bdes;
        }
        public void SetClit(string clit)
        {
            Clit = clit;
        }

        public void SetClit1(string clit1)
        {
            Clit1 = clit1;
        }
        public void SetClitm(string clitm)
        {
            Clitm = clitm;
        }
        public void SetCdc(string cdc)
        {
            Cdc = cdc;
        }
        public void SetFlet(string flet)
        {
            Flet = flet;
        }
        public void SetCgr(string cgr)
        {
            Cgr = cgr;
        }
        public void SetIntr(string intr)
        {
            Intr = intr;
        }
        public void SetDesct(string desct)
        {
            Desct = desct;
        }
        public void SetNcont(string ncont)
        {
            Ncont = ncont;
        }
       
        public void SetTcmc1(string tcmc1)
        {
            Tcmc1 = tcmc1;
        }
        public void SetTcmc2(string tcmc2)
        {
            Tcmc2 = tcmc2;
        }
        public void SetTcmc3(string tcmc3)
        {
            Tcmc3 = tcmc3;
        }
       
       

        public void SetFWe(string fwe)
        {
            FWe = fwe;
        }
        public void SetTcmc6(string tcmc6)
        {
            Tcmc6 = tcmc6;
        }
       
        public void SetTcmc8(string tcmc8)
        {
            Tcmc8 = tcmc8;
        }
        public void SetTcmc9(string tcmc9)
        {
            Tcmc9 = tcmc9;
        }

        

        public void SetBps(string bps)
        {
            Bps = bps;
        }

        public void SetTPup(string tpup)
        {
            TPup = tpup;
        }
        public void SetTPus(string tpus)
        {
            TPus = tpus;
        }
        public void SetTLup(string tlup)
        {
            TLup = tlup;
        }
        public void SetLtemp(string ltemp)  ////  TLPbar.ToString(), DTpress.ToString(), TCPbar.ToString(), TEPbar.ToString()
        {
            Ltemp = ltemp;
        }
        public void SetLPbar(string lpbar)
        {
            LPbar = lpbar;
        }
        public void SetTpress(string tpress)
        {
            Tpress = tpress;
        }
        public void SetCPbar(string cpbar)
        {
            CPbar = cpbar;
        }
        public void SetEPbar(string epbar)
        {
            EPbar = epbar;
        }

        //-------------------------------------------------------------------------
        public string GetDigt()
        {
            return Digt;
        }
        public string GetCcion()
        {
            return Ccion;
        }
        public string GetCastre()
        {
            return Castre;
        }
       
        public string GetCpcion()
        {
            return Cpcion;
        }
        public string GetCfrio()
        {
            return Cfrio;
        }
        public string GetCeq1()
        {
            return Ceq1;
        }
        public string GetCeq2()
        {
            return Ceq2;
        }
        public string GetCeq3()
        {
            return Ceq3;
        }
        public string Getinc()
        {
            return inc;
        }

      
 
        //------------------------------------------------------------------
        
        
      
        public void SetCastre(string castre)
        {
            Castre = castre;
        }
        public void SetCpcion(string cpcion)
        {
            Cpcion = cpcion;
        }
        public void SetCfrio(string cfrio)
        {
            Cfrio = cfrio;
        }
        public void SetCeq1(string ceq1)
        {
            Ceq1 = ceq1;
        }
        public void SetCeq2(string ceq2)
        {
            Ceq2 = ceq2;
        }
        public void SetCeq3(string ceq3)
        {
            Ceq3 = ceq3;
        }
      

        public void Setinc(string Inc)
        {
            inc = Inc;
        }
        


        public void SetLugar(string lugar)
        {
            Lugar = lugar;
        }
       
       
        
        public string GetLugar()
        {
           return Lugar;
        }
         
                
      }
}
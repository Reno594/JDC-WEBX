using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Interop.Excel;
//using olib;
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
        //----------------------------------------------------

        public Form1()
        {
            InitializeComponent();
            myUserControl = new UserControl1();
            this.DateActualizer();
            local = Directory.GetCurrentDirectory();
            TNP.Focus();
            myProcesses = Process.GetProcesses();
                        
        }
        
        /// <summary>
        /// Add a camera
        /// </summary>
        /// <returns>returns true or false</returns>
        bool add()
        {
            if (this.Validateit())
            {
                CCam myCam = new CCam(TNC.Text, TTem.Text, TLargo.Text,
                    TAncho.Text, TAlto.Text, TTP.Text, CDT.Text, CCE.Text,
                    TCF.Text, TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text, TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text, CITPuerta.Text, TDEC.Text, TDECE.Text, TDECH.Text, TDECF.Text, TCantEv.Text, TCentx.Text, TCdin.Text, TCxp.Text, CRefrig.Text, CDigt.Text, CCcion.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text, RBInteriores.Checked, CKPanel.Checked, CKNanauf.Checked, CKMonile.Checked,
                    CKCable.Checked, CKUnid.Checked, CKPuerta.Checked, CKexpo.Checked, CKepiso.Checked, CKDrenaje.Checked, CKSellaje.Checked, CKEmerg.Checked, CKBrida.Checked, CKCort.Checked, CKTor.Checked, CKTcobre.Checked, CKRefrig.Checked, 
                    CKSopr.Checked, CKValv.Checked, CKCelect.Checked, CKPerf.Checked, CKAUni.Checked, CKAlum.Checked, CKMobra.Checked, CKSD.Checked, CKSMin.Checked, CKPAI.Checked, CKpmtal.Checked, CKlux.Checked, CKvsol.Checked, CKp10.Checked, CKp12.Checked, CKp15.Checked, CKp15t.Checked,
                    CKdt.Checked, CKat.Checked, CKbt.Checked, CKmt.Checked, CKmod.Checked, CRvent.Checked, CKantc.Checked, CKppc.Checked, CKepc.Checked, CKcion.Checked, CKastre.Checked, CKpcion.Checked, CKfrio.Checked, CKeq1.Checked, CKeq2.Checked, CKeq3.Checked, CKsu1.Checked, CKsu2.Checked, CKsu3.Checked, CKpps.Checked, CVolt.Text,
                    CFase.Text, TConD.Text, TMuc.Text, TMevp.Text, TSol.Text, TValv.Text, TCvta.Text, Txm2.Text, TCuadro.Text, CBint.Text, CBcm.Text, CCmci.Text, CBexpo.Text,
                    Cevap.Text, Cmodex.Text, Ctxv.Text, CSumi.Text, CTCond.Text, CTamb.Text,
                    Ctpd.Text, Cnoff1.Text, Coff1.Text, Cnoff2.Text, Coff2.Text, Cnoff3.Text, Coff3.Text, Cnoff4.Text, Coff4.Text, Cnoff5.Text, Coff5.Text, Cnoff6.Text, Coff6.Text, Cnoff7.Text, Coff7.Text, Cnoff8.Text, Coff8.Text, CTEvap.Text, TCodValv.Text, TTmos.Text, TInc.Text, TInev.Text, TIned.Text, TIncd.Text,
                    TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TQevpc.Text, TTint.Text, TEquip.Text, TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text,
                    TDmce.Text, TDmc1.Text, TDmc11.Text, TDmc12.Text, TDlq1.Text, TDlq2.Text, TDlq3.Text, TDlq11.Text, TDlq21.Text, TDlq31.Text, TDsu130.Text, TDsu230.Text, TDsu330.Text, TDsu110.Text, TDsu210.Text, TDsu310.Text, TDsu105.Text, TDsu205.Text, TDsu305.Text, TDmc2.Text, TDmc3.Text, TDmc4.Text, TDmc5.Text, TDmc6.Text, TDmc7.Text, TDmc8.Text, TLcc.Text, TLss.Text, TPcmc.Text, TPcmc1.Text, TPcmc2.Text, TPcmc3.Text,
                    TPcond.Text, TPlq.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcv.Text, TPcx.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, TPnt.Text, TPml.Text, TPdtr.Text, TPdtc.Text, TPdt2.Text, TPdt1.Text, TCp80.Text, TCp100.Text, TCp120.Text, TCp150.Text, TCt80.Text, TCt100.Text, TCt120.Text, TCt150.Text, TCp80m.Text, TCp100m.Text, TCp120m.Text, TCp150m.Text, TCt80m.Text, TCt100m.Text, TCt120m.Text, TCt150m.Text,
                    SPtp84.Text, SPtp83.Text, SPtp78.Text, SPtp76.Text, SPtp75.Text, SPtp74.Text, SPtp73.Text, SPtp72.Text, SPtp85.Text, SPtp86.Text, SPtp87.Text, SPtp88.Text, SPtp89.Text, SPtp90.Text, SPtp91.Text, SPtp92.Text, SPtp93.Text, SPtp94.Text, SPtp95.Text, SPtp96.Text, SPtp97.Text, SPtp98.Text, SPtp99.Text, SPtp100.Text, SPtp101.Text, SPtp102.Text, SPtp103.Text, SPtp104.Text, SPtp105.Text, SPtp106.Text, SPtp107.Text,
                    SPtp108.Text, SPtp109.Text, SPtp110.Text, SPtp111.Text, SPtp136.Text, SPtp137.Text, SPtp141.Text, SPtp142.Text, SPtp143.Text, SPtp144.Text, SPtp145.Text, SPtp146.Text, SPtp147.Text, SPtp148.Text, SPtp149.Text, SPtp150.Text, SPtp151.Text, TIn1.Text, TIn2.Text, TIn3.Text, TIn4.Text, TIn5.Text, TIn6.Text, TIn7.Text, TIn8.Text, TIn9.Text, TIn10.Text, TIn11.Text, TIn12.Text, TIn13.Text, TIn14.Text, TIn15.Text, TIn16.Text, TIn17.Text, TIn18.Text, TIn19.Text, TIn20.Text, TIn21.Text, TIn22.Text, TIn23.Text,
                     TIp1.Text, TIp2.Text, TIp3.Text, TIp4.Text, TIp5.Text, TIp6.Text, TIp7.Text, TIp8.Text, TIp9.Text, TIp10.Text, TIp11.Text, TIp12.Text, TIp13.Text, TIp14.Text, TIp15.Text, TIp16.Text, TIp17.Text, TIp18.Text, TIp19.Text, TIp20.Text, TIp21.Text, TIp23.Text, pk);
                //myCam.SetCodValv(codValv);
                if (myOferta == null)
                {
                    LEstado.Text = "Estado: Creando oferta...";
                    try
                    {
                        myOferta = new COferta(TNP.Text, TREF.Text, TNO.Text, TCmat.Text, TDmc1.Text, TDmc11.Text, TDmc12.Text, TDlq1.Text, TDlq2.Text, TDlq3.Text, TDlq11.Text, TDlq21.Text, TDlq31.Text, TDsu130.Text, TDsu230.Text, TDsu330.Text, TDsu110.Text, TDsu210.Text, TDsu310.Text, TDsu105.Text, TDsu205.Text, TDsu305.Text, TDmc2.Text, TDmc3.Text, TDmc4.Text, TDmc5.Text, TDmc6.Text, TDmc7.Text, TDmc8.Text, TLcc.Text, TLss.Text, TPcmc.Text, TPcmc1.Text, TPcmc2.Text, TPcmc3.Text, TPcond.Text,
                            TPlq.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcv.Text, TPcx.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, TPnt.Text, TPml.Text, TPdtr.Text, TPdtc.Text, TPdt2.Text, TPdt1.Text, TCp80.Text, TCp100.Text, TCp120.Text, TCp150.Text, TCt80.Text, TCt100.Text, TCt120.Text, TCt150.Text, TCp80m.Text, TCp100m.Text, TCp120m.Text, TCp150m.Text, TCt80m.Text, TCt100m.Text, TCt120m.Text, TCt150m.Text,
                            SPtp84.Text, SPtp83.Text, SPtp78.Text, SPtp76.Text, SPtp75.Text, SPtp74.Text, SPtp73.Text, SPtp72.Text, SPtp85.Text, SPtp86.Text, SPtp87.Text, SPtp88.Text, SPtp89.Text, SPtp90.Text, SPtp91.Text, SPtp92.Text, SPtp93.Text, SPtp94.Text, SPtp95.Text, SPtp96.Text, SPtp97.Text, SPtp98.Text, SPtp99.Text, SPtp100.Text, SPtp101.Text, SPtp102.Text, SPtp103.Text, SPtp104.Text, SPtp105.Text, SPtp106.Text, SPtp107.Text,
                            SPtp108.Text, SPtp109.Text, SPtp110.Text, SPtp111.Text, SPtp136.Text, SPtp137.Text, SPtp141.Text, SPtp142.Text, SPtp143.Text, SPtp144.Text, SPtp145.Text, SPtp146.Text, SPtp147.Text, SPtp148.Text, SPtp149.Text, SPtp150.Text, SPtp151.Text, TInc.Text, TInev.Text, TIned.Text, TIncd.Text, TIpv.Text, TIcc.Text,
                            TQevp.Text, TQevpd.Text, TQevpc.Text, TTint.Text, TEquip.Text, CSumi.Text, CDT.Text, int.Parse(TCC.Text), CDigt.Text, CCcion.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text, TFecha.Text, TRMol.Text, TConstinc.Text,
                            TContCivP.Text, TEquipFrig.Text, TPuertasFrig.Text, TDesE.Text, TTasa.Text, TDsc.Text, TGAdmObras.Text, TGIndObras.Text, TGIndObrascuc.Text,
                            Credito.Text, Creditocup.Text, CLugar.Text, CClit.Text, CClit1.Text, CClitm.Text, CRcivil.Text, CRpiso.Text, RBmoni.Checked, RB60H.Checked, RBun.Checked, RBun2.Checked, RBun3.Checked, RBun4.Checked, RBun5.Checked, RBun6.Checked, RBun7.Checked, RBun8.Checked, RBun9.Checked, RBinvert.Checked, RB360.Checked, CKeur.Checked, CKmod.Checked, RBsup.Checked, TBscu.Text, TBcont.Text, TBcos.Text, TBdir.Text, TBenv.Text, TBpo.Text, TBfec.Text, TBdes.Text,
                            CCdc.Text, CFlet.Text, CCgr.Text, CIntr.Text, CDesct.Text, CNcont.Text, CMon.Text, TIn1.Text, TIn2.Text, TIn3.Text, TIn4.Text, TIn5.Text, TIn6.Text, TIn7.Text, TIn8.Text, TIn9.Text, TIn10.Text, TIn11.Text, TIn12.Text, TIn13.Text, TIn14.Text, TIn15.Text, TIn16.Text, TIn17.Text, TIn18.Text, TIn19.Text, TIn20.Text, TIn21.Text, TIn22.Text, TIn23.Text,
                            TIp1.Text, TIp2.Text, TIp3.Text, TIp4.Text, TIp5.Text, TIp6.Text, TIp7.Text, TIp8.Text, TIp9.Text, TIp10.Text, TIp11.Text, TIp12.Text, TIp13.Text, TIp14.Text, TIp15.Text, TIp16.Text, TIp17.Text, TIp18.Text, TIp19.Text, TIp20.Text, TIp21.Text, TIp23.Text);
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
            if (TTem.Text == "")
                return false;
            if (CDT.Text == "")
                return false;
            /*
            if (TCF.Text == "")
                return false;
            if (TFW.Text == "")
                return false;
            */
            if (TDEC.Text == "")
                return false;
            if (TDECE.Text == "")
                return false;
            if (TDECH.Text == "")
                return false;
            
            //if (TDECF.Text == "")
                //return false;

            //-----------dimension
            if (TLargo.Text == "")
                return false;
            if (TAncho.Text == "")
                return false;
            if (TAlto.Text == "")
                return false;
            if (CCE.Text == "")
                return false;
            if (CITPuerta.Text == "")
                return false;
            if (CSup.Text == "")
                return false;
            if (TCantEv.Text == "")
                return false;                       
            if (TCentx.Text == "")
                return false;
            if (Ctxv.Text == "")
                return false;
            
            if (TCdin.Text == "")
                return false;
            if (TCxp.Text == "")
                return false;
            if (CRefrig.Text == "")
                return false;
            if (TConD.Text == "")
                return false;
            if (CClit.Text == "")
                return false;
            if (CClit1.Text == "")
                return false;
            if (CClitm.Text == "")
                return false;
            if (CCdc.Text == "")
                return false;
            
            return true;
        }
        
        /// <summary>
        /// asigna a los textbox los valores de la camara pasada como parametro
        /// </summary>
        /// <param name="numcam">camara a extraer valores</param>
        void asignarcam(int numcam)
        {            
            CCam myCam = myOferta.GetCam(numcam - 1);
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
            TCantEv.Text = myCam.GetCantEv();
            TQfep.Text = myCam.GetQfep();
            TScdro.Text = myCam.GetScdro();
            TSpsi.Text = myCam.GetSpsi();
            TStemp.Text = myCam.GetStemp();
            TApsi.Text = myCam.GetApsi();
            TEmevp.Text = myCam.GetEmevp();
            TTP.Text = myCam.GetTP();
            CDT.Text = myCam.GetDT();
            TDEC.Text = myCam.GetDEC();
            TDECE.Text = myCam.GetDECE();
            TDECH.Text = myCam.GetDECH();
            TDECF.Text = myCam.GetDECF();
            TLargo.Text = myCam.GetLargo();
            TAncho.Text = myCam.GetAncho();
            CCE.Text = myCam.GetCE();
            TAlto.Text = myCam.GetAlto();
            CITPuerta.Text = myCam.GetIT();
            CSup.Text = myCam.GetSUP();
            actualcam = CCamara.Text;
            TCantEv.Text = myCam.GetCantEv();                      
            TCentx.Text = myCam.GetCentx();
            Ctxv.Text = myCam.GetCtxv();
            TCdin.Text = myCam.GetCdin();
            TCxp.Text = myCam.GetCxp();
            CRefrig.Text = myCam.GetRefrig();
            CDigt.Text = myCam.GetDigt();
            CCcion.Text = myCam.GetCcion();
            CCastre.Text = myCam.GetCastre();
            CCpcion.Text = myCam.GetCpcion();
            CCfrio.Text = myCam.GetCfrio();
            CCeq1.Text = myCam.GetCeq1();
            CCeq2.Text = myCam.GetCeq2();
            CCeq3.Text = myCam.GetCeq3();
            TConD.Text = myCam.GetConD();
            TMuc.Text = myCam.GetTMuc();
            TMevp.Text = myCam.GetTMevp();
            TValv.Text = myCam.GetTValv();
            TCvta.Text = myCam.GetTCvta();
            Txm2.Text = myCam.GetTxm2();
            TSol.Text = myCam.GetTSol();
            
            TCuadro.Text = myCam.GetTCuadro();
            CBint.Text = myCam.GetCBint();
            CBcm.Text = myCam.GetCBcm();
            CCmci.Text = myCam.GetCCmci();
            CBexpo.Text = myCam.GetCBexpo();
            Cevap.Text = myCam.GetCevap();
            Cmodex.Text = myCam.GetCmodex();
            CSumi.Text = myCam.GetCSumi();
            CTCond.Text = myCam.GetCTCond();
            CTamb.Text = myCam.GetCTamb();
            Ctpd.Text = myCam.GetCtpd();
            Cnoff1.Text = myCam.GetCnoff1();
            Coff1.Text = myCam.GetCoff1();
            Cnoff2.Text = myCam.GetCnoff2();
            Coff2.Text = myCam.GetCoff2();
            Cnoff3.Text = myCam.GetCnoff3();
            Coff3.Text = myCam.GetCoff3();
            Cnoff4.Text = myCam.GetCnoff4();
            Coff4.Text = myCam.GetCoff4();
            Cnoff5.Text = myCam.GetCnoff5();
            Coff5.Text = myCam.GetCoff5();
            Cnoff6.Text = myCam.GetCnoff6();
            Coff6.Text = myCam.GetCoff6();
            Cnoff7.Text = myCam.GetCnoff7();
            Coff7.Text = myCam.GetCoff7();
            if (myCam.GetKmod() == false)
            {
                Cnoff8.Text = myCam.GetCnoff8();
                Coff8.Text = myCam.GetCoff8();             
            }
            /*
            if (myCam.GetKmod() == true)
            {
                Cnoff8.Text = myCam.GetCmod();
                Coff8.Text = "1";
            }
            */
            
            

            CTEvap.Text = myCam.GetCTEvap();
            TCodValv.Text = myCam.GetCodValv();
            TTmos.Text = myCam.GetTmos();
            TInc.Text = myCam.GetTInc();
            TInev.Text = myCam.GetTInev();
            TIned.Text = myCam.GetTIned();
            TIncd.Text = myCam.GetTIncd();
            TIpv.Text = myCam.GetTIpv();
            TIcc.Text = myCam.GetTIcc();
            TQevp.Text = myCam.GetTQevp();
            TQevpd.Text = myCam.GetTQevpd();
            TQevpc.Text = myCam.GetTQevpc();
            TTint.Text = myCam.GetTTint();
            TEquip.Text = myCam.GetTEquip();
            TCint1.Text = myCam.GetTCint1();
            TCint2.Text = myCam.GetTCint2();
            TCint3.Text = myCam.GetTCint3();
            TMcc.Text = myCam.GetTMcc();
            TCmce.Text = myCam.GetTCmce();
            TPmce.Text = myCam.GetTPmce();
            TDmce.Text = myCam.GetTDmce();
            TDmc1.Text = myCam.GetTDmc1();
            TDmc11.Text = myCam.GetTDmc11();
            TDmc12.Text = myCam.GetTDmc12();
            TDlq1.Text = myCam.GetTDlq1();
            TDlq2.Text = myCam.GetTDlq2();
            TDlq3.Text = myCam.GetTDlq3();
            TDlq11.Text = myCam.GetTDlq11();
            TDlq21.Text = myCam.GetTDlq21();
            TDlq31.Text = myCam.GetTDlq31();
            TDsu130.Text = myCam.GetTDsu130();
            TDsu230.Text = myCam.GetTDsu230();
            TDsu330.Text = myCam.GetTDsu330();
            TDsu110.Text = myCam.GetTDsu110();
            TDsu210.Text = myCam.GetTDsu210();
            TDsu310.Text = myCam.GetTDsu310();
            TDsu105.Text = myCam.GetTDsu105();
            TDsu205.Text = myCam.GetTDsu205();
            TDsu305.Text = myCam.GetTDsu305();
            TDmc2.Text = myCam.GetTDmc2();
            TDmc3.Text = myCam.GetTDmc3();
            TDmc4.Text = myCam.GetTDmc4();
            TDmc5.Text = myCam.GetTDmc5();
            TDmc6.Text = myCam.GetTDmc6();
            TDmc7.Text = myCam.GetTDmc7();
            TDmc8.Text = myCam.GetTDmc8();
            TLcc.Text = myCam.GetTLcc();
            TLss.Text = myCam.GetTLss();
            TPcmc.Text = myCam.GetTPcmc();
            TPcmc1.Text = myCam.GetTPcmc1();
            TPcmc2.Text = myCam.GetTPcmc2();
            TPcmc3.Text = myCam.GetTPcmc3();
            TPcond.Text = myCam.GetTPcond();
            TPlq.Text = myCam.GetTPlq();
            TPvq.Text = myCam.GetTPvq();
            TPls.Text = myCam.GetTPls();
            TPosc.Text = myCam.GetTPosc();
            TPsq.Text = myCam.GetTPsq();
            TPcq.Text = myCam.GetTPcq();
            TPcv.Text = myCam.GetTPcv();
            TPcx.Text = myCam.GetTPcx();
            TPcy.Text = myCam.GetTPcy();
            TPcp.Text = myCam.GetTPcp();
            TPex.Text = myCam.GetTPex();
            TPrs.Text = myCam.GetTPrs();
            TCsist.Text = myCam.GetTCsist();
            TPem.Text = myCam.GetTPem();
            TPnt.Text = myCam.GetTPnt();
            TPml.Text = myCam.GetTPml();
            TPdtr.Text = myCam.GetTPdtr();
            TPdtc.Text = myCam.GetTPdtc();
            TPdt2.Text = myCam.GetTPdt2();
            TPdt1.Text = myCam.GetTPdt1();
            TCp80.Text = myCam.GetTCp80();
            TCp100.Text = myCam.GetTCp100();
            TCp120.Text = myCam.GetTCp120();
            TCp150.Text = myCam.GetTCp150();
            TCt80.Text = myCam.GetTCt80();
            TCt100.Text = myCam.GetTCt100();
            TCt120.Text = myCam.GetTCt120();
            TCt150.Text = myCam.GetTCt150();
            TCp80m.Text = myCam.GetTCp80m();
            TCp100m.Text = myCam.GetTCp100m();
            TCp120m.Text = myCam.GetTCp120m();
            TCp150m.Text = myCam.GetTCp150m();
            TCt80m.Text = myCam.GetTCt80m();
            TCt100m.Text = myCam.GetTCt100m();
            TCt120m.Text = myCam.GetTCt120m();
            TCt150m.Text = myCam.GetTCt150m();
            SPtp84.Text = myCam.GetSPtp84();
            SPtp83.Text = myCam.GetSPtp83();
            SPtp78.Text = myCam.GetSPtp78();
            SPtp76.Text = myCam.GetSPtp76();
            SPtp75.Text = myCam.GetSPtp75();
            SPtp74.Text = myCam.GetSPtp74();
            SPtp73.Text = myCam.GetSPtp73();
            SPtp72.Text = myCam.GetSPtp72();
            SPtp85.Text = myCam.GetSPtp85();
            SPtp86.Text = myCam.GetSPtp86();
            SPtp87.Text = myCam.GetSPtp87();
            SPtp88.Text = myCam.GetSPtp88();
            SPtp89.Text = myCam.GetSPtp89();
            SPtp90.Text = myCam.GetSPtp90();
            SPtp91.Text = myCam.GetSPtp91();
            SPtp92.Text = myCam.GetSPtp92();
            SPtp93.Text = myCam.GetSPtp93();
            SPtp94.Text = myCam.GetSPtp94();
            SPtp95.Text = myCam.GetSPtp95();
            SPtp96.Text = myCam.GetSPtp96();
            SPtp97.Text = myCam.GetSPtp97();
            SPtp98.Text = myCam.GetSPtp98();
            SPtp99.Text = myCam.GetSPtp99();
            SPtp100.Text = myCam.GetSPtp100();
            SPtp101.Text = myCam.GetSPtp101();
            SPtp102.Text = myCam.GetSPtp102();
            SPtp103.Text = myCam.GetSPtp103();
            SPtp104.Text = myCam.GetSPtp104();
            SPtp105.Text = myCam.GetSPtp105();
            SPtp106.Text = myCam.GetSPtp106();
            SPtp107.Text = myCam.GetSPtp107();
            SPtp108.Text = myCam.GetSPtp108();
            SPtp109.Text = myCam.GetSPtp109();
            SPtp110.Text = myCam.GetSPtp110();
            SPtp111.Text = myCam.GetSPtp111();
            SPtp136.Text = myCam.GetSPtp136();
            SPtp137.Text = myCam.GetSPtp137();
            SPtp141.Text = myCam.GetSPtp141();
            SPtp142.Text = myCam.GetSPtp142();
            SPtp143.Text = myCam.GetSPtp143();
            SPtp144.Text = myCam.GetSPtp144();
            SPtp145.Text = myCam.GetSPtp145();
            SPtp146.Text = myCam.GetSPtp146();
            SPtp147.Text = myCam.GetSPtp147();
            SPtp148.Text = myCam.GetSPtp148();
            SPtp149.Text = myCam.GetSPtp149();
            SPtp150.Text = myCam.GetSPtp150();
            SPtp151.Text = myCam.GetSPtp151();
            TIn1.Text = myCam.GetTIn1();
            TIn2.Text = myCam.GetTIn2();
            TIn3.Text = myCam.GetTIn3();
            TIn4.Text = myCam.GetTIn4();
            TIn5.Text = myCam.GetTIn5();
            TIn6.Text = myCam.GetTIn6();
            TIn7.Text = myCam.GetTIn7();
            TIn8.Text = myCam.GetTIn8();
            TIn9.Text = myCam.GetTIn9();
            TIn10.Text = myCam.GetTIn10();
            TIn11.Text = myCam.GetTIn11();
            TIn12.Text = myCam.GetTIn12();
            TIn13.Text = myCam.GetTIn13();
            TIn14.Text = myCam.GetTIn14();
            TIn15.Text = myCam.GetTIn15();
            TIn16.Text = myCam.GetTIn16();
            TIn17.Text = myCam.GetTIn17();
            TIn18.Text = myCam.GetTIn18();
            TIn19.Text = myCam.GetTIn19();
            TIn20.Text = myCam.GetTIn20();
            TIn21.Text = myCam.GetTIn21();
            TIn22.Text = myCam.GetTIn22();
            TIn23.Text = myCam.GetTIn23();
            TIp1.Text = myCam.GetTIp1();
            TIp2.Text = myCam.GetTIp2();
            TIp3.Text = myCam.GetTIp3();
            TIp4.Text = myCam.GetTIp4();
            TIp5.Text = myCam.GetTIp5();
            TIp6.Text = myCam.GetTIp6();
            TIp7.Text = myCam.GetTIp7();
            TIp8.Text = myCam.GetTIp8();
            TIp9.Text = myCam.GetTIp9();
            TIp10.Text = myCam.GetTIp10();
            TIp11.Text = myCam.GetTIp11();
            TIp12.Text = myCam.GetTIp12();
            TIp13.Text = myCam.GetTIp13();
            TIp14.Text = myCam.GetTIp14();
            TIp15.Text = myCam.GetTIp15();
            TIp16.Text = myCam.GetTIp16();
            TIp17.Text = myCam.GetTIp17();
            TIp18.Text = myCam.GetTIp18();
            TIp19.Text = myCam.GetTIp19();
            TIp20.Text = myCam.GetTIp20();
            TIp21.Text = myCam.GetTIp21();
            TIp23.Text = myCam.GetTIp23();

            pictureBox2.Image = myUserControl.IMAGES.Images[myCam.GetPK()];
            try
            {
                CTEvap.Text = (int.Parse(TTem.Text) - int.Parse(CDT.Text)).ToString();
            }
            catch { }

            if (myCam.GetIE())
                RBInteriores.Checked = true;
            else
                RBInteriores.Checked = false;

            if (myCam.GetKP())
                CKPanel.Checked = true;
            else
                CKPanel.Checked = false;

            if (myCam.GetKN())
                CKNanauf.Checked = true;
            else
                CKNanauf.Checked = false;

            if (myCam.GetKM())
                CKMonile.Checked = true;
            else
                CKMonile.Checked = false;
            
            if (myCam.GetKC())
                CKCable.Checked = true;
            else
                CKCable.Checked = false;
            
            if (myCam.GetKU())
                CKUnid.Checked = true;
            else
                CKUnid.Checked = false;

            if (myCam.GetKPR())
                CKPuerta.Checked = true;
            else
                CKPuerta.Checked = false;

            if (myCam.GetKexpo())
                CKexpo.Checked = true;
            else
                CKexpo.Checked = false;

            if (myCam.GetKepiso())
                CKepiso.Checked = true;
            else
                CKepiso.Checked=false;

            if (myCam.GetKD())
                CKDrenaje.Checked = true;
            else
                CKDrenaje.Checked = false;

            if (myCam.GetKS())
                CKSellaje.Checked = true;
            else
                CKSellaje.Checked = false;

            if (myCam.GetKCE())
                CKEmerg.Checked = true;
            else
                CKEmerg.Checked = false;

            if (myCam.GetKB())
                CKBrida.Checked = true;
            else
                CKBrida.Checked = false;

            if (myCam.GetKCO())
                CKCort.Checked = true;
            else
                CKCort.Checked = false;

            if (myCam.GetKTO())
                CKTor.Checked = true;
            else
                CKTor.Checked = false;

            if (myCam.GetKTC())
                CKTcobre.Checked = true;
            else
                CKTcobre.Checked = false;

            if (myCam.GetKRE())
                CKRefrig.Checked = true;
            else
                CKRefrig.Checked = false;

            if (myCam.GetKSO())
                CKSopr.Checked = true;
            else
                CKSopr.Checked = false;

            if (myCam.GetKVA())
                CKValv.Checked = true;
            else
                CKValv.Checked = false;

            if (myCam.GetKCL())
                CKCelect.Checked = true;
            else
                CKCelect.Checked = false;

            if (myCam.GetKPE())
                CKPerf.Checked = true;
            else
                CKPerf.Checked = false;

            if (myCam.GetKUA())
                CKAUni.Checked = true;
            else
                CKAUni.Checked = false;

            if (myCam.GetKAL())
                CKAlum.Checked = true;
            else
                CKAlum.Checked = false;

            if (myCam.GetKMO())
                CKMobra.Checked = true;
            else
                CKMobra.Checked = false;

            if (myCam.GetKSD())
                CKSD.Checked = true;
            else
                CKSD.Checked = false;

            if (myCam.GetKSMin())
                CKSMin.Checked = true;
            else
                CKSMin.Checked = false;

            if (myCam.GetKPAI())
                CKPAI.Checked = true;
            else
                CKPAI.Checked = false;

            if (myCam.GetKpmtal())
                CKpmtal.Checked = true;
            else
                CKpmtal.Checked = false;

            if (myCam.GetKlux())
                CKlux.Checked = true;
            else
                CKlux.Checked = false;

            if (myCam.GetKvsol())
                CKvsol.Checked = true;
            else
                CKvsol.Checked = false;

            if (myCam.GetKp10())
                CKp10.Checked = true;
            else
                CKp10.Checked = false;

            if (myCam.GetKp12())
                CKp12.Checked = true;
            else
                CKp12.Checked = false;

            if (myCam.GetKp15())
                CKp15.Checked = true;
            else
                CKp15.Checked = false;
            
            if (myCam.GetKp15t())
                CKp15t.Checked = true;
            else
                CKp15t.Checked = false;

            if (myCam.GetKdt())
                CKdt.Checked = true;
            else
                CKdt.Checked = false;

            if (myCam.GetKat())
                CKat.Checked = true;
            else
                CKat.Checked = false;

            if (myCam.GetKbt())
                CKbt.Checked = true;
            else
                CKbt.Checked = false;

            if (myCam.GetKmt())
                CKmt.Checked = true;
            else
                CKmt.Checked = false;

            if (myCam.GetKmod())
                CKmod.Checked = true;
            else
                CKmod.Checked = false;

            if (myCam.GetRvent())
                CRvent.Checked = true;
            else
                CRvent.Checked = false;

            if (myCam.GetKantc())
                CKantc.Checked = true;
            else
                CKantc.Checked = false;

            if (myCam.GetKppc())
                CKppc.Checked = true;
            else
                CKppc.Checked = false;

            if (myCam.GetKepc())
                CKepc.Checked = true;
            else
                CKepc.Checked = false;

            if (myCam.GetKcion())
                CKcion.Checked = true;
            else
                CKcion.Checked = false;

            if (myCam.GetKpcion())
                CKpcion.Checked = true;
            else
                CKpcion.Checked = false;

            if (myCam.GetKastre())
                CKastre.Checked = true;
            else
                CKastre.Checked = false;

            if (myCam.GetKfrio())
                CKfrio.Checked = true;
            else
                CKfrio.Checked = false;

            if (myCam.GetKeq1())
                CKeq1.Checked = true;
            else
                CKeq1.Checked = false;

            if (myCam.GetKeq2())
                CKeq2.Checked = true;
            else
                CKeq2.Checked = false;

            if (myCam.GetKeq3())
                CKeq3.Checked = true;
            else
                CKeq3.Checked = false;

            if (myCam.GetKsu1())
                CKsu1.Checked = true;
            else
                CKsu1.Checked = false;

            if (myCam.GetKsu2())
                CKsu2.Checked = true;
            else
                CKsu2.Checked = false;

            if (myCam.GetKsu3())
                CKsu3.Checked = true;
            else
                CKsu3.Checked = false;
            
            if (myCam.GetKpps())
                CKpps.Checked = true;
            else
                CKpps.Checked = false;

            CVolt.Text = myCam.GetVol();
            CFase.Text = myCam.GetFase();
            if (int.Parse(myCam.GetTemp()) < 0)
            {
                TCapC.Text = ((float.Parse(myCam.GetAncho()) * (float.Parse(myCam.GetAlto())) * (float.Parse(myCam.GetLargo()))) * 71 * float.Parse("0,49744081")).ToString();
                string myvar = TCapC.Text;
                TCapC.Text = "";
                for (int y = 0; y < myvar.Length; y++)
                    if (myvar[y] != ',')
                        TCapC.Text += myvar[y].ToString();
                    else
                        break;
            }
            else
            {
                TCapC.Text = ((float.Parse(myCam.GetAncho()) * (float.Parse(myCam.GetAlto())) * (float.Parse(myCam.GetLargo()))) * 71 * float.Parse("0,99488169")).ToString();
                string myvar = TCapC.Text;
                TCapC.Text = "";
                for (int y = 0; y < myvar.Length; y++)
                    if (myvar[y] != ',')
                        TCapC.Text += myvar[y].ToString();
                    else
                        break;
            }

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
            TTP.Text = "Derecha";
            CDT.Text = "6";
            CCE.Text = "CC";
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
            CSup.Text = "CON";
            CITPuerta.Text = "SIN";
            TDEC.Text = "18";
            TDECE.Text = "12";
            TDECH.Text = "10";
            TDECF.Text = "";
            CCamara.Items.Clear();
            TNP.Clear();
            TNO.Clear();
            TCmat.Clear();
            TDmc1.Clear();
            TDmc11.Clear();
            TDmc12.Clear();
            TDlq1.Clear();
            TDlq2.Clear();
            TDlq3.Clear();
            TDlq11.Clear();
            TDlq21.Clear();
            TDlq31.Clear();
            TDsu130.Clear();
            TDsu230.Clear();
            TDsu330.Clear();
            TDsu110.Clear();
            TDsu210.Clear();
            TDsu310.Clear();
            TDsu105.Clear();
            TDsu205.Clear();
            TDsu305.Clear();
            TDmc2.Clear();
            TDmc3.Clear();
            TDmc4.Clear();
            TDmc5.Clear();
            TDmc6.Clear();
            TDmc7.Clear();
            TDmc8.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcmc1.Clear();
            TPcmc2.Clear();
            TPcmc3.Clear();
            TPcond.Clear();
            TPlq.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcv.Clear();
            TPcx.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TPdtr.Clear();
            TPdtc.Clear();
            TPdt2.Clear();
            TPdt1.Clear();
            TCp80.Clear();
            TCp100.Clear();
            TCp120.Clear();
            TCp150.Clear();
            TCt80.Clear();
            TCt100.Clear();
            TCt120.Clear();
            TCt150.Clear();
            TCp80m.Clear();
            TCp100m.Clear();
            TCp120m.Clear();
            TCp150m.Clear();
            TCt80m.Clear();
            TCt100m.Clear();
            TCt120m.Clear();
            TCt150m.Clear();
            SPtp84.Clear();
            SPtp83.Clear();
            SPtp78.Clear();
            SPtp76.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp73.Clear();
            SPtp72.Clear();
            SPtp85.Clear();
            SPtp86.Clear();
            SPtp87.Clear();
            SPtp88.Clear();
            SPtp89.Clear();
            SPtp90.Clear();
            SPtp91.Clear();
            SPtp92.Clear();
            SPtp93.Clear();
            SPtp94.Clear();
            SPtp95.Clear();
            SPtp96.Clear();
            SPtp97.Clear();
            SPtp98.Clear();
            SPtp99.Clear();
            SPtp100.Clear();
            SPtp101.Clear();
            SPtp102.Clear();
            SPtp103.Clear();
            SPtp104.Clear();
            SPtp105.Clear();
            SPtp106.Clear();
            SPtp107.Clear();
            SPtp108.Clear();
            SPtp109.Clear();
            SPtp110.Clear();
            SPtp111.Clear();
            SPtp136.Clear();
            SPtp137.Clear();
            SPtp141.Clear();
            SPtp142.Clear();
            SPtp143.Clear();
            SPtp144.Clear();
            SPtp145.Clear();
            SPtp146.Clear();
            SPtp147.Clear();
            SPtp148.Clear();
            SPtp149.Clear();
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
            TDmc1.Clear();
            TDmc11.Clear();
            TDmc12.Clear();
            TDlq1.Clear();
            TDlq2.Clear();
            TDlq3.Clear();
            TDlq11.Clear();
            TDlq21.Clear();
            TDlq31.Clear();
            TDsu130.Clear();
            TDsu230.Clear();
            TDsu330.Clear();
            TDsu110.Clear();
            TDsu210.Clear();
            TDsu310.Clear();
            TDsu105.Clear();
            TDsu205.Clear();
            TDsu305.Clear();
            TDmc2.Clear();
            TDmc3.Clear();
            TDmc4.Clear();
            TDmc5.Clear();
            TDmc6.Clear();
            TDmc7.Clear();
            TDmc8.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcmc1.Clear();
            TPcmc2.Clear();
            TPcmc3.Clear();
            TPcond.Clear();
            TPcond.Clear();
            TPlq.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcv.Clear();
            TPcx.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TPdtr.Clear();
            TPdtc.Clear();
            TPdt2.Clear();
            TPdt1.Clear();
            TCp80.Clear();
            TCp100.Clear();
            TCp120.Clear();
            TCp150.Clear();
            TCt80.Clear();
            TCt100.Clear();
            TCt120.Clear();
            TCt150.Clear();
            TCp80m.Clear();
            TCp100m.Clear();
            TCp120m.Clear();
            TCp150m.Clear();
            TCt80m.Clear();
            TCt100m.Clear();
            TCt120m.Clear();
            TCt150m.Clear();
            SPtp84.Clear();
            SPtp83.Clear();
            SPtp78.Clear();
            SPtp76.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp73.Clear();
            SPtp72.Clear();
            SPtp85.Clear();
            SPtp86.Clear();
            SPtp87.Clear();
            SPtp88.Clear();
            SPtp89.Clear();
            SPtp90.Clear();
            SPtp91.Clear();
            SPtp92.Clear();
            SPtp93.Clear();
            SPtp94.Clear();
            SPtp95.Clear();
            SPtp96.Clear();
            SPtp97.Clear();
            SPtp98.Clear();
            SPtp99.Clear();
            SPtp100.Clear();
            SPtp101.Clear();
            SPtp102.Clear();
            SPtp103.Clear();
            SPtp104.Clear();
            SPtp105.Clear();
            SPtp106.Clear();
            SPtp107.Clear();
            SPtp108.Clear();
            SPtp109.Clear();
            SPtp110.Clear();
            SPtp111.Clear();
            SPtp136.Clear();
            SPtp137.Clear();
            SPtp141.Clear();
            SPtp142.Clear();
            SPtp143.Clear();
            SPtp144.Clear();
            SPtp145.Clear();
            SPtp146.Clear();
            SPtp147.Clear();
            SPtp148.Clear();
            SPtp149.Clear();
            SPtp150.Clear();
            SPtp151.Clear();
            TQevp.Clear();
            TQevpd.Clear();
            TQevpc.Clear();
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
            TCantEv.Text = "1";                       
            TCentx.Text = "25";
            TCdin.Text = "0";
            TCxp.Text = "0";
            TConD.Text = "1";
            CRefrig.Text = "R449A";
            CDigt.Text = "DIG";
            CCcion.Text = "";
            CCastre.Text = "";
            CCpcion.Text = "";
            CCfrio.Text = "";
            CCeq1.Text = "";
            CCeq2.Text = "";
            CCeq3.Text = "";
            this.DateActualizer();
            CLugar.Text = "Habana";
            CClit.Text = "Sonia Aleida";
            CClit1.Text = "";
            CClitm.Text = "";
            RBmoni.Checked = true;
            RB60H.Checked = true;
            RBun.Checked = true;
            RBun2.Checked = true;
            RBun3.Checked = true;
            RBun4.Checked = true;
            RBun5.Checked = true;
            RBun6.Checked = true;
            RBun7.Checked = true;
            RBun8.Checked = true;
            RBun9.Checked = true;
            RBinvert.Checked = true;
            RB360.Checked = true;
            CKeur.Checked = true;
            RBsup.Checked = true;
            // MARGEN POR PRODUCTO FIJADO
            TRMol.Text = "1,35";
            TConstinc.Text = "1,3";
            TContCivP.Text = "1,2";
            TEquipFrig.Text = "1,3";
            TPuertasFrig.Text = "1,3";
            TDesE.Text = "1,3";
            // COSTO DEL PRODUCTO
            CVolt.Text = "380";
            CFase.Text = "3";
            Ctxv.Text = "TX3";
            Ctpd.Text = "";
            Cnoff1.Text = "";
            Coff1.Text = "";
            Cnoff2.Text = "";
            Coff2.Text = "";
            Cnoff3.Text = "";
            Coff3.Text = "";
            Cnoff4.Text = "";
            Coff4.Text = "";
            Cnoff5.Text = "";
            Coff5.Text = "";
            Cnoff6.Text = "";
            Coff6.Text = "";
            Cnoff7.Text = "";
            Coff7.Text = "";
            Cnoff8.Text = "";
            Coff8.Text = "";
            CCdc.Text = "";
            CFlet.Text = "";
            CCgr.Text = "";
            CIntr.Text = "";
            CDesct.Text = "";
            CNcont.Text = "";
            CMon.Text = "";
            TIn1.Clear();
            TIn2.Clear();
            TIn3.Clear();
            TIn4.Clear();
            TIn5.Clear();
            TIn6.Clear();
            TIn7.Clear();
            TIn8.Clear();
            TIn9.Clear();
            TIn10.Clear();
            TIn11.Clear();
            TIn12.Clear();
            TIn13.Clear();
            TIn14.Clear();
            TIn15.Clear();
            TIn16.Clear();
            TIn17.Clear();
            TIn18.Clear();
            TIn19.Clear();
            TIn20.Clear();
            TIn21.Clear();
            TIn22.Clear();
            TIn23.Clear();
            TIp1.Clear();
            TIp2.Clear();
            TIp3.Clear();
            TIp4.Clear();
            TIp5.Clear();
            TIp6.Clear();
            TIp7.Clear();
            TIp8.Clear();
            TIp9.Clear();
            TIp10.Clear();
            TIp11.Clear();
            TIp12.Clear();
            TIp13.Clear();
            TIp14.Clear();
            TIp15.Clear();
            TIp16.Clear();
            TIp17.Clear();
            TIp18.Clear();
            TIp19.Clear();
            TIp20.Clear();
            TIp21.Clear();
            TIp23.Clear();
            
            //TPcCUC.Text = "";
            //TMpiso.Text = "";
            //TPanel.Text = "";
            //Tcuc.Text = "";
            //TConsC.Text = "";
            //TCup.Text = "";

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
                    TDmc1.Text = myOferta.GetDmc1();
                    TDmc11.Text = myOferta.GetDmc11();
                    TDmc12.Text = myOferta.GetDmc12();
                    TDlq1.Text = myOferta.GetDlq1();
                    TDlq2.Text = myOferta.GetDlq2();
                    TDlq3.Text = myOferta.GetDlq3();
                    TDlq11.Text = myOferta.GetDlq1();
                    TDlq21.Text = myOferta.GetDlq2();
                    TDlq31.Text = myOferta.GetDlq3();
                    TDsu130.Text = myOferta.GetDsu130();
                    TDsu230.Text = myOferta.GetDsu230();
                    TDsu330.Text = myOferta.GetDsu330();
                    TDsu110.Text = myOferta.GetDsu110();
                    TDsu210.Text = myOferta.GetDsu210();
                    TDsu310.Text = myOferta.GetDsu310();
                    TDsu105.Text = myOferta.GetDsu105();
                    TDsu205.Text = myOferta.GetDsu205();
                    TDsu305.Text = myOferta.GetDsu305();
                    TDmc2.Text = myOferta.GetDmc2();
                    TDmc3.Text = myOferta.GetDmc3();
                    TDmc4.Text = myOferta.GetDmc4();
                    TDmc5.Text = myOferta.GetDmc5();
                    TDmc6.Text = myOferta.GetDmc6();
                    TDmc7.Text = myOferta.GetDmc7();
                    TDmc8.Text = myOferta.GetDmc8();
                    TLcc.Text = myOferta.GetLcc();
                    TLss.Text = myOferta.GetLss();
                    TPcmc.Text = myOferta.GetPcmc();
                    TPcmc1.Text = myOferta.GetPcmc1();
                    TPcmc2.Text = myOferta.GetPcmc2();
                    TPcmc3.Text = myOferta.GetPcmc3();
                    TPcond.Text = myOferta.GetPcond();
                    TPlq.Text = myOferta.GetPlq();
                    TPvq.Text = myOferta.GetPvq();
                    TPls.Text = myOferta.GetPls();
                    TPosc.Text = myOferta.GetPosc();
                    TPsq.Text = myOferta.GetPsq();
                    TPcq.Text = myOferta.GetPcq();
                    TPcv.Text = myOferta.GetPcv();
                    TPcx.Text = myOferta.GetPcx();
                    TPcy.Text = myOferta.GetPcy();
                    TPcp.Text = myOferta.GetPcp();
                    TPex.Text = myOferta.GetPex();
                    TPrs.Text = myOferta.GetPrs();
                    TCsist.Text = myOferta.GetCsist();
                    TPem.Text = myOferta.GetPem();
                    TPnt.Text = myOferta.GetPnt();
                    TPml.Text = myOferta.GetPml();
                    TPdtr.Text = myOferta.GetPdtr();
                    TPdtc.Text = myOferta.GetPdtc();
                    TPdt2.Text = myOferta.GetPdt2();
                    TPdt1.Text = myOferta.GetPdt1();
                    TCp80.Text = myOferta.GetCp80();
                    TCp100.Text = myOferta.GetCp100();
                    TCp120.Text = myOferta.GetCp120();
                    TCp150.Text = myOferta.GetCp150();
                    TCt80.Text = myOferta.GetCt80();
                    TCt100.Text = myOferta.GetCt100();
                    TCt120.Text = myOferta.GetCt120();
                    TCt150.Text = myOferta.GetCt150();
                    TCp80m.Text = myOferta.GetCp80m();
                    TCp100m.Text = myOferta.GetCp100m();
                    TCp120m.Text = myOferta.GetCp120m();
                    TCp150m.Text = myOferta.GetCp150m();
                    TCt80m.Text = myOferta.GetCt80m();
                    TCt100m.Text = myOferta.GetCt100m();
                    TCt120m.Text = myOferta.GetCt120m();
                    TCt150m.Text = myOferta.GetCt150m();
                    SPtp84.Text = myOferta.GetPtp84();
                    SPtp83.Text = myOferta.GetPtp83();
                    SPtp78.Text = myOferta.GetPtp78();
                    SPtp76.Text = myOferta.GetPtp76();
                    SPtp75.Text = myOferta.GetPtp75();
                    SPtp74.Text = myOferta.GetPtp74();
                    SPtp73.Text = myOferta.GetPtp73();
                    SPtp72.Text = myOferta.GetPtp72();
                    SPtp85.Text = myOferta.GetPtp85();
                    SPtp86.Text = myOferta.GetPtp86();
                    SPtp87.Text = myOferta.GetPtp87();
                    SPtp88.Text = myOferta.GetPtp88();
                    SPtp89.Text = myOferta.GetPtp89();
                    SPtp90.Text = myOferta.GetPtp90();
                    SPtp91.Text = myOferta.GetPtp91();
                    SPtp92.Text = myOferta.GetPtp92();
                    SPtp93.Text = myOferta.GetPtp93();
                    SPtp94.Text = myOferta.GetPtp94();
                    SPtp95.Text = myOferta.GetPtp95();
                    SPtp96.Text = myOferta.GetPtp96();
                    SPtp97.Text = myOferta.GetPtp97();
                    SPtp98.Text = myOferta.GetPtp98();
                    SPtp99.Text = myOferta.GetPtp99();
                    SPtp100.Text = myOferta.GetPtp100();
                    SPtp101.Text = myOferta.GetPtp101();
                    SPtp102.Text = myOferta.GetPtp102();
                    SPtp103.Text = myOferta.GetPtp103();
                    SPtp104.Text = myOferta.GetPtp104();
                    SPtp105.Text = myOferta.GetPtp105();
                    SPtp106.Text = myOferta.GetPtp106();
                    SPtp107.Text = myOferta.GetPtp107();
                    SPtp108.Text = myOferta.GetPtp108();
                    SPtp109.Text = myOferta.GetPtp109();
                    SPtp110.Text = myOferta.GetPtp110();
                    SPtp111.Text = myOferta.GetPtp111();
                    SPtp136.Text = myOferta.GetPtp136();
                    SPtp137.Text = myOferta.GetPtp137();
                    SPtp141.Text = myOferta.GetPtp141();
                    SPtp142.Text = myOferta.GetPtp142();
                    SPtp143.Text = myOferta.GetPtp143();
                    SPtp144.Text = myOferta.GetPtp144();
                    SPtp145.Text = myOferta.GetPtp145();
                    SPtp146.Text = myOferta.GetPtp146();
                    SPtp147.Text = myOferta.GetPtp147();
                    SPtp148.Text = myOferta.GetPtp148();
                    SPtp149.Text = myOferta.GetPtp149();
                    SPtp150.Text = myOferta.GetPtp150();
                    SPtp151.Text = myOferta.GetPtp151();
                    TInc.Text = myOferta.GetInc();
                    TInev.Text = myOferta.GetInev();
                    TIned.Text = myOferta.GetIned();
                    TIncd.Text = myOferta.GetIncd();
                    TIpv.Text = myOferta.GetIpev();
                    TIcc.Text = myOferta.GetIcc();
                    TQevp.Text = myOferta.GetQevp();
                    TQevpd.Text = myOferta.GetQevpd();
                    TQevpc.Text = myOferta.GetQevpc();
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
                    CClitm.Text = myOferta.GetClitm();
                    CCdc.Text = myOferta.GetCdc();
                    CFlet.Text = myOferta.GetFlet();
                    CCgr.Text = myOferta.GetCgr();
                    CIntr.Text = myOferta.GetIntr();
                    CDesct.Text = myOferta.GetDesct();
                    CNcont.Text = myOferta.GetNcont();
                    CMon.Text = myOferta.GetMon();
                    

                    if (myOferta.GetConstCivPan() == "")
                    {
                        TContCivP.Text = "1,2";
                        myOferta.SetConstCivPan("1,2");
                    }
                    else
                        TContCivP.Text = myOferta.GetConstCivPan();

                    if (myOferta.GetEquipFrig() == "")
                    {
                        TEquipFrig.Text = "1,3";
                        myOferta.SetEquipFrig("1,3");
                    }
                    else
                        TEquipFrig.Text = myOferta.GetEquipFrig();

                    TGAdmObras.Text = myOferta.GetGastosAdmObra();
                    TGIndObras.Text = myOferta.GetGastosIndObra();
                    TGIndObrascuc.Text = myOferta.GetGastosIndObracuc();
                    Credito.Text = myOferta.GetCredito();
                    Creditocup.Text = myOferta.GetCreditocup();
                    CRcivil.Text = myOferta.GetCRcivil();
                    CRpiso.Text = myOferta.GetCRpiso();
                    TTasa.Text = myOferta.GetTasa();
                    TDsc.Text = myOferta.GetDsc();

                    // MARGEN POR PRODUCTO FIJADO
                    if (myOferta.GetBun9() == false)
                    {
                        if (myOferta.Getinc() == "")
                        {
                            TConstinc.Text = "1,3";
                            myOferta.Setinc("1,3");
                        }
                        else
                            TConstinc.Text = myOferta.Getinc();

                        if (myOferta.GetPuertasFrig() == "")
                        {
                            TPuertasFrig.Text = "1,3";
                            myOferta.SetPuertasFrig("1,3");
                        }
                        else
                            TPuertasFrig.Text = myOferta.GetPuertasFrig();

                        if (myOferta.GetResinaM() == "")
                        {
                            TRMol.Text = "1,35";
                            myOferta.SetResinaM("1,35");
                        }
                        else
                            TRMol.Text = myOferta.GetResinaM();

                        if (myOferta.GetDesE() == "")
                        {
                            TDesE.Text = "1,3";
                            myOferta.SetDesE("1,3");
                        }
                        else
                            TDesE.Text = myOferta.GetDesE();
                    }
                     /*
                    if (myOferta.Getinc() == "")
                    {
                        TConstinc.Text = "1,3";
                        myOferta.Setinc("1,3");
                    }
                    else
                        TConstinc.Text = myOferta.Getinc();

                    if (myOferta.GetPuertasFrig() == "")
                    {
                        TPuertasFrig.Text = "1,3";
                        myOferta.SetPuertasFrig("1,3");
                    }
                    else
                        TPuertasFrig.Text = myOferta.GetPuertasFrig();

                    if (myOferta.GetResinaM() == "")
                    {
                        TRMol.Text = "1,35";
                        myOferta.SetResinaM("1,35");
                    }
                    else
                        TRMol.Text = myOferta.GetResinaM();

                    if (myOferta.GetDesE() == "")
                    {
                        TDesE.Text = "1,3";
                        myOferta.SetDesE("1,3");
                    }
                    else
                        TDesE.Text = myOferta.GetDesE();
                    */
                    // COSTO DEL PRODUCTO
                    else
                    {
                        TConstinc.Text = myOferta.Getinc();
                        TPuertasFrig.Text = myOferta.GetPuertasFrig();
                        TRMol.Text = myOferta.GetResinaM();
                        TDesE.Text = myOferta.GetDesE();
                    }
                    
                    CLugar.Text = myOferta.GetLugar();
                    CClit.Text = myOferta.GetClit();
                    CClit1.Text = myOferta.GetClit1();
                    CClitm.Text = myOferta.GetClitm();
                    RBmoni.Checked = myOferta.GetBmoni();
                    RB60H.Checked = myOferta.GetB60H();
                    RBun.Checked = myOferta.GetBun();
                    RBun2.Checked = myOferta.GetBun2();
                    RBun3.Checked = myOferta.GetBun3();
                    RBun4.Checked = myOferta.GetBun4();
                    RBun5.Checked = myOferta.GetBun5();
                    RBun6.Checked = myOferta.GetBun6();
                    RBun7.Checked = myOferta.GetBun7();
                    RBun8.Checked = myOferta.GetBun8();
                    RBun9.Checked = myOferta.GetBun9();
                    RBinvert.Checked = myOferta.GetBinvert();
                    RB360.Checked = myOferta.GetB360();
                    CKeur.Checked = myOferta.GetKeur();
                    RBsup.Checked = myOferta.GetBsup();
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
                    LEstado.Text = "Abriendo archivo xlsx...";
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

                        myOferta = new COferta("---", "---", "---", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 40, "---", "1,25", "1,2",
                            "1,2", "1,3", "1,25", "1,3", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "");
                        Range myRange = myWorksheet.get_Range("A1", "I45");
                        Array myValues = (Array)myRange.Value2;                        

                        while (true)
                        {
                            try
                            {
                                //CONSTRUCTOR**
                                CCam myCam = new CCam(myValues.GetValue(i, 4).ToString(), myValues.GetValue(i, 5).ToString(), myValues.GetValue(i, 6).ToString(),
                                    myValues.GetValue(i, 7).ToString(), myValues.GetValue(i, 8).ToString(), "Derecha", "7", "CC", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "CON", "SIN",
                                    "18", "12", "10", "10", "1", "25", "0", "0", "R404A", "", "", "", "", "", "", "", "", true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true,  true, true, true, true, true, true, true, true, true, true,
                                    true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, true, "380", "3", "1", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 
                                    "ECO", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "",
                                    "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", 0);
                                myOferta.AddCam(myCam);
                                CCam newCalcCam = myOferta.GetCam(i - 6);
                                newCalcCam.setCF(Calcular(i - 5));
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
                        this.SetText(TNC, myCam1.GetNC());
                        //TNC.Text = myCam1.GetNC();
                        this.SetText(TTem, myCam1.GetTemp());
                        //TTem.Text = myCam1.GetTemp();
                        this.SetText(TCF, myCam1.GetCF());
                        //TCF.Text = myCam1.GetCF();
                        this.SetText(TFW, myCam1.GetFW());
                        this.SetText(TQfw, myCam1.GetQfw());
                        this.SetText(TCmod, myCam1.GetCmod());
                        this.SetText(TCmodd, myCam1.GetCmodd());
                        this.SetText(TCmodp, myCam1.GetCmodp());
                        this.SetText(TDesc, myCam1.GetDesc());
                        this.SetText(TPrec, myCam1.GetPrec());
                        this.SetText(TCantEv, myCam1.GetCantEv());
                        this.SetText(TQfep, myCam1.GetQfep());
                        this.SetText(TScdro, myCam1.GetScdro());
                        this.SetText(TSpsi, myCam1.GetSpsi());
                        this.SetText(TStemp, myCam1.GetStemp());
                        this.SetText(TApsi, myCam1.GetApsi());
                        this.SetText(TEmevp, myCam1.GetEmevp());
                        this.SetText(TTP, myCam1.GetTP());
                        //TTP.Text = myCam1.GetTP();
                        this.SetText(CDT, myCam1.GetDT());
                        //CDT.Text = myCam1.GetDT();
                        this.SetText(TDEC, myCam1.GetDEC());
                        //TDEC.Text = myCam1.GetDEC();
                        this.SetText(TDECE, myCam1.GetDECE());
                        this.SetText(TDECH, myCam1.GetDECH());
                        this.SetText(TDECF, myCam1.GetDECF());
                        this.SetText(TLargo, myCam1.GetLargo());
                        //TLargo.Text = myCam1.GetLargo();
                        this.SetText(TAncho, myCam1.GetAncho());
                        //TAncho.Text = myCam1.GetAncho();
                        this.SetText(CCE, myCam1.GetCE());
                        //CCE.Text = myCam1.GetCE();
                        this.SetText(TAlto, myCam1.GetAlto());
                        //TAlto.Text = myCam1.GetAlto();
                        this.SetText(CITPuerta, myCam1.GetIT());
                        //CITPuerta.Text = myCam1.GetIT();
                        this.SetText(CSup, myCam1.GetSUP());
                        //CSup.Text = myCam1.GetSUP();
                        this.SetText(TCantEv, myCam1.GetCantEv());
                        //TCantEv.Text = myCam1.GetCantEv();                                             
                        this.SetText(TCentx, myCam1.GetCentx());
                        this.SetText(TCdin, myCam1.GetCdin());
                        this.SetText(TCxp, myCam1.GetCxp());
                        this.SetText(TConD, myCam1.GetConD());
                        //CD.Text = myCam1.GetCD();
                        this.SetText(CRefrig, myCam1.GetRefrig());
                        this.SetText(CDigt, myCam1.GetDigt());
                        this.SetText(CCcion, myCam1.GetCcion());
                        this.SetText(CCastre, myCam1.GetCastre());
                        this.SetText(CCpcion, myCam1.GetCpcion());
                        this.SetText(CCfrio, myCam1.GetCfrio());
                        this.SetText(CCeq1, myCam1.GetCeq1());
                        this.SetText(CCeq2, myCam1.GetCeq2());
                        this.SetText(CCeq3, myCam1.GetCeq3());
                        //CRefrig.Text = myCam1.GetRefrig();
                        
                        
                        this.SetText(CSumi, myCam1.GetCSumi());
                        
                        this.SetText(CTCond, myCam1.GetCTCond());
                        this.SetText(CTamb, myCam1.GetCTamb());
                        this.SetText(Ctxv, myCam1.GetCtxv());
                        
                        this.SetText(Ctpd, myCam1.GetCtpd());
                        this.SetText(Cnoff1, myCam1.GetCnoff1());
                        this.SetText(Coff1, myCam1.GetCoff1());
                        this.SetText(Cnoff2, myCam1.GetCnoff2());
                        this.SetText(Coff2, myCam1.GetCoff2());
                        this.SetText(Cnoff3, myCam1.GetCnoff3());
                        this.SetText(Coff3, myCam1.GetCoff3());
                        this.SetText(Cnoff4, myCam1.GetCnoff4());
                        this.SetText(Coff4, myCam1.GetCoff4());
                        this.SetText(Cnoff5, myCam1.GetCnoff5());
                        this.SetText(Coff5, myCam1.GetCoff5());
                        this.SetText(Cnoff6, myCam1.GetCnoff6());
                        this.SetText(Coff6, myCam1.GetCoff6());
                        this.SetText(Cnoff7, myCam1.GetCnoff7());
                        this.SetText(Coff7, myCam1.GetCoff7());
                        this.SetText(Cnoff8, myCam1.GetCnoff8());
                        this.SetText(Coff8, myCam1.GetCoff8());
       
                        this.SetText(CTEvap, myCam1.GetCTEvap());

                        if (myOferta.GetBun9() == false)
                        {
                            this.SetText(TContCivP, "1,2");
                            this.SetText(TEquipFrig, "1,3");
                            this.SetText(TGAdmObras, "7");
                            this.SetText(TGIndObras, "0,0476011231");
                            this.SetText(TGIndObrascuc, "0,012514743");
                            this.SetText(TConstinc, "1,3");
                            this.SetText(TPuertasFrig, "1,3");
                            this.SetText(TDesE, "1,3");
                            this.SetText(TRMol, "1,35");
                            this.SetText(TTasa, "1");
                            this.SetText(TDsc, "1,3");
                        }
                        else
                        {
                            this.SetText(TContCivP, "1");
                            this.SetText(TEquipFrig, "1");
                            this.SetText(TGAdmObras, "7");
                            this.SetText(TGIndObras, "0,0476011231");
                            this.SetText(TGIndObrascuc, "0,012514743");
                            this.SetText(TConstinc, "1");
                            this.SetText(TPuertasFrig, "1");
                            this.SetText(TDesE, "1");
                            this.SetText(TRMol, "1");
                            //this.SetText(TTasa, "1");
                            this.SetText(TDsc, "1");
                        }
                       
                        
                        myOferta.SetCantCam(myOferta.GetCont());
                        this.SetText(TCC, (myOferta.GetCont().ToString()));
                        myExcel.Quit();
                        //TCC.Text = myOferta.GetCont().ToString();
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
                CKepiso.Visible = true;
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
            DateTime fechaF = Convert.ToDateTime("06/30/2021");
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
            DateTime fechaX = Convert.ToDateTime("06/30/2021");
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
                    myOferta.SetDmc1(TDmc1.Text);
                    myOferta.SetDmc11(TDmc11.Text);
                    myOferta.SetDmc12(TDmc12.Text);
                    myOferta.SetDlq1(TDlq1.Text);
                    myOferta.SetDlq2(TDlq2.Text);
                    myOferta.SetDlq3(TDlq3.Text);
                    myOferta.SetDlq11(TDlq11.Text);
                    myOferta.SetDlq21(TDlq21.Text);
                    myOferta.SetDlq31(TDlq31.Text);
                    myOferta.SetDsu130(TDsu130.Text);
                    myOferta.SetDsu230(TDsu230.Text);
                    myOferta.SetDsu330(TDsu330.Text);
                    myOferta.SetDsu110(TDsu110.Text);
                    myOferta.SetDsu210(TDsu210.Text);
                    myOferta.SetDsu310(TDsu310.Text);
                    myOferta.SetDsu105(TDsu105.Text);
                    myOferta.SetDsu205(TDsu205.Text);
                    myOferta.SetDsu305(TDsu305.Text);
                    myOferta.SetDmc2(TDmc2.Text);
                    myOferta.SetDmc3(TDmc3.Text);
                    myOferta.SetDmc4(TDmc4.Text);
                    myOferta.SetDmc5(TDmc5.Text);
                    myOferta.SetDmc6(TDmc6.Text);
                    myOferta.SetDmc7(TDmc7.Text);
                    myOferta.SetDmc8(TDmc8.Text);
                    myOferta.SetLcc(TLcc.Text);
                    myOferta.SetLss(TLss.Text);
                    myOferta.SetPcmc(TPcmc.Text);
                    myOferta.SetPcmc1(TPcmc1.Text);
                    myOferta.SetPcmc2(TPcmc2.Text);
                    myOferta.SetPcmc3(TPcmc3.Text);
                    myOferta.SetPcond(TPcond.Text);
                    myOferta.SetPlq(TPlq.Text);
                    myOferta.SetPvq(TPvq.Text);
                    myOferta.SetPls(TPls.Text);
                    myOferta.SetPosc(TPosc.Text);
                    myOferta.SetPsq(TPsq.Text);
                    myOferta.SetPcq(TPcq.Text);
                    myOferta.SetPcv(TPcv.Text);
                    myOferta.SetPcx(TPcx.Text);
                    myOferta.SetPcy(TPcy.Text);
                    myOferta.SetPcp(TPcp.Text);
                    myOferta.SetPex(TPex.Text);
                    myOferta.SetPrs(TPrs.Text);
                    myOferta.SetCsist(TCsist.Text);
                    myOferta.SetPem(TPem.Text);
                    myOferta.SetPnt(TPnt.Text);
                    myOferta.SetPml(TPml.Text);
                    myOferta.SetPdtr(TPdtr.Text);
                    myOferta.SetPdtc(TPdtc.Text);
                    myOferta.SetPdt2(TPdt2.Text);
                    myOferta.SetPdt1(TPdt1.Text);
                    myOferta.SetCp80(TCp80.Text);
                    myOferta.SetCp100(TCp100.Text);
                    myOferta.SetCp120(TCp120.Text);
                    myOferta.SetCp150(TCp150.Text);
                    myOferta.SetCt80(TCt80.Text);
                    myOferta.SetCt100(TCt100.Text);
                    myOferta.SetCt120(TCt120.Text);
                    myOferta.SetCt150(TCt150.Text);
                    myOferta.SetCp80m(TCp80.Text);
                    myOferta.SetCp100m(TCp100m.Text);
                    myOferta.SetCp120m(TCp120m.Text);
                    myOferta.SetCp150m(TCp150m.Text);
                    myOferta.SetCt80m(TCt80m.Text);
                    myOferta.SetCt100m(TCt100m.Text);
                    myOferta.SetCt120m(TCt120m.Text);
                    myOferta.SetCt150m(TCt150m.Text);
                    myOferta.SetPtp84(SPtp84.Text);
                    myOferta.SetPtp83(SPtp83.Text);
                    myOferta.SetPtp78(SPtp78.Text);
                    myOferta.SetPtp76(SPtp76.Text);
                    myOferta.SetPtp75(SPtp75.Text);
                    myOferta.SetPtp74(SPtp74.Text);
                    myOferta.SetPtp73(SPtp73.Text);
                    myOferta.SetPtp72(SPtp72.Text);
                    myOferta.SetPtp85(SPtp85.Text);
                    myOferta.SetPtp86(SPtp86.Text);
                    myOferta.SetPtp87(SPtp87.Text);
                    myOferta.SetPtp88(SPtp88.Text);
                    myOferta.SetPtp89(SPtp89.Text);
                    myOferta.SetPtp90(SPtp90.Text);
                    myOferta.SetPtp91(SPtp91.Text);
                    myOferta.SetPtp92(SPtp92.Text);
                    myOferta.SetPtp93(SPtp93.Text);
                    myOferta.SetPtp94(SPtp94.Text);
                    myOferta.SetPtp95(SPtp95.Text);
                    myOferta.SetPtp96(SPtp96.Text);
                    myOferta.SetPtp97(SPtp97.Text);
                    myOferta.SetPtp98(SPtp98.Text);
                    myOferta.SetPtp99(SPtp99.Text);
                    myOferta.SetPtp100(SPtp100.Text);
                    myOferta.SetPtp101(SPtp101.Text);
                    myOferta.SetPtp102(SPtp102.Text);
                    myOferta.SetPtp103(SPtp103.Text);
                    myOferta.SetPtp104(SPtp104.Text);
                    myOferta.SetPtp105(SPtp105.Text);
                    myOferta.SetPtp106(SPtp106.Text);
                    myOferta.SetPtp107(SPtp107.Text);
                    myOferta.SetPtp108(SPtp108.Text);
                    myOferta.SetPtp109(SPtp109.Text);
                    myOferta.SetPtp110(SPtp110.Text);
                    myOferta.SetPtp111(SPtp111.Text);
                    myOferta.SetPtp136(SPtp136.Text);
                    myOferta.SetPtp137(SPtp137.Text);
                    myOferta.SetPtp141(SPtp141.Text);
                    myOferta.SetPtp142(SPtp142.Text);
                    myOferta.SetPtp143(SPtp143.Text);
                    myOferta.SetPtp144(SPtp144.Text);
                    myOferta.SetPtp145(SPtp145.Text);
                    myOferta.SetPtp146(SPtp146.Text);
                    myOferta.SetPtp147(SPtp147.Text);
                    myOferta.SetPtp148(SPtp148.Text);
                    myOferta.SetPtp149(SPtp149.Text);
                    myOferta.SetPtp150(SPtp150.Text);
                    myOferta.SetPtp151(SPtp151.Text);
                    myOferta.SetInc(TInc.Text);
                    myOferta.SetInev(TInev.Text);
                    myOferta.SetIned(TIned.Text);
                    myOferta.SetIncd(TIncd.Text);
                    myOferta.SetIpv(TIpv.Text);
                    myOferta.SetIcc(TIcc.Text);
                    myOferta.SetQevp(TQevp.Text);
                    myOferta.SetQevpd(TQevpd.Text);
                    myOferta.SetQevpc(TQevpc.Text);
                    myOferta.SetTint(TTint.Text);
                    myOferta.SetEquip(TEquip.Text);
                    myOferta.SetCantCam(int.Parse(TCC.Text));
                    myOferta.SetREF(TREF.Text);
                    myOferta.SetConstCivPan(TContCivP.Text);
                    myOferta.SetEquipFrig(TEquipFrig.Text);
                    myOferta.SetGastosAdmObra(TGAdmObras.Text);
                    myOferta.SetGastosIndObra(TGIndObras.Text);
                    myOferta.SetGastosIndObracuc(TGIndObrascuc.Text);
                    myOferta.SetCredito(Credito.Text);
                    myOferta.SetCreditocup(Creditocup.Text);
                    myOferta.SetCRcivil(CRcivil.Text);
                    myOferta.SetCRpiso(CRpiso.Text);
                    myOferta.Setinc(TConstinc.Text);
                    myOferta.SetPuertasFrig(TPuertasFrig.Text);
                    myOferta.SetDesE(TDesE.Text);
                    myOferta.SetDigt(CDigt.Text);
                    myOferta.SetCcion(CCcion.Text);
                    myOferta.SetCastre(CCastre.Text);
                    myOferta.SetCpcion(CCpcion.Text);
                    myOferta.SetCfrio(CCfrio.Text);
                    myOferta.SetCeq1(CCeq1.Text);
                    myOferta.SetCeq2(CCeq2.Text);
                    myOferta.SetCeq3(CCeq3.Text);
                    myOferta.SetResinaM(TRMol.Text);
                    myOferta.SetTasa(TTasa.Text);
                    myOferta.SetDsc(TDsc.Text);
                    myOferta.SetLugar(CLugar.Text);
                    myOferta.SetClit(CClit.Text);
                    myOferta.SetClit1(CClit1.Text);
                    myOferta.SetClitm(CClitm.Text);
                    myOferta.SetBmoni(RBmoni.Checked);
                    myOferta.SetB60H(RB60H.Checked);
                    myOferta.SetBun(RBun.Checked);
                    myOferta.SetBun2(RBun2.Checked);
                    myOferta.SetBun3(RBun3.Checked);
                    myOferta.SetBun4(RBun4.Checked);
                    myOferta.SetBun5(RBun5.Checked);
                    myOferta.SetBun6(RBun6.Checked);
                    myOferta.SetBun7(RBun7.Checked);
                    myOferta.SetBun8(RBun8.Checked);
                    myOferta.SetBun9(RBun9.Checked);
                    myOferta.SetBinvert(RBinvert.Checked);
                    myOferta.SetB360(RB360.Checked);
                    myOferta.SetKeur(CKeur.Checked);
                    myOferta.SetBsup(RBsup.Checked);
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
                    myOferta.SetMon(CMon.Text);
                    
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

        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            
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
                        if (myOferta.GetClit() == "Ibrahin López")
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
                        if (myOferta.GetClit() == "Suzzete Díaz")
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
                        if (myOferta.GetClit() == "Laura Olazabal")
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
                        if (myOferta.GetClit() == "Yunikleyvis Hernández")
                        {
                            if (RCX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RGX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RFTX.Checked)
                                myXlsxSaveDialog.FileName = myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                            if (RODC.Checked)
                                myXlsxSaveDialog.FileName = " « " + "SCU - " + myOferta.GetNO() + " - " + String.Format("{0:yyyy}", DateTime.Now) + " " + myOferta.GetNP() + " " + myOferta.GetClit1() + " " + myOferta.GetREF() + ".xlsx";
                        }

                    
                    else
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
                CalcWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(CalcWorkBook);

                GC.Collect();
                GC.WaitForPendingFinalizers();
                try
                {
                    Marshal.FinalReleaseComObject(SelecWorkSheet);
                }
                catch { }
                CalcWorkBook.Close(false, Type.Missing, Type.Missing);
                Marshal.FinalReleaseComObject(SelecWorkBook);
                NewExcelApp.Quit();
                Marshal.FinalReleaseComObject(NewExcelApp);   
                TrabExcel.Disconnect();
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
            TTP.Text = "Derecha";           
            CDT.Text = "6";
            CCE.Text = "CC";
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
            CSup.Text = "CON";
            CITPuerta.Text = "SIN";
            TDEC.Text = "18";
            TDECE.Text = "12";
            TDECH.Text = "10";
            TDECF.Text = "";
            TCantEv.Text = "1";                      
            TCentx.Text = "25";
            Ctxv.Text = "TX3";
            
            TCdin.Text = "0";
            TCxp.Text = "0";
            TConD.Text = "1";
            CRefrig.Text = "R404A";
            CDigt.Text = "DIG";
            CCcion.Text = "";
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
            Txm2.Clear();
            TMuc.Clear();
            TSol.Clear();
            
            TValv.Clear();
            TCvta.Clear();
            TCodValv.Clear();
            TTmos.Clear();
            TInc.Clear();
            TInev.Clear();
            TIned.Clear();
            TIncd.Clear();
            TIpv.Clear();
            TDmc1.Clear();
            TDmc11.Clear();
            TDmc12.Clear();
            TDlq1.Clear();
            TDlq2.Clear();
            TDlq3.Clear();
            TDlq11.Clear();
            TDlq21.Clear();
            TDlq31.Clear();
            TDsu130.Clear();
            TDsu230.Clear();
            TDsu330.Clear();
            TDsu110.Clear();
            TDsu210.Clear();
            TDsu310.Clear();
            TDsu105.Clear();
            TDsu205.Clear();
            TDsu305.Clear();
            TDmc2.Clear();
            TDmc3.Clear();
            TDmc4.Clear();
            TDmc5.Clear();
            TDmc6.Clear();
            TDmc7.Clear();
            TDmc8.Clear();
            TLcc.Clear();
            TLss.Clear();
            TPcmc.Clear();
            TPcmc1.Clear();
            TPcmc2.Clear();
            TPcmc3.Clear();
            TPcond.Clear();
            TPlq.Clear();
            TPvq.Clear();
            TPls.Clear();
            TPosc.Clear();
            TPsq.Clear();
            TPcq.Clear();
            TPcv.Clear();
            TPcx.Clear();
            TPcy.Clear();
            TPcp.Clear();
            TPex.Clear();
            TPrs.Clear();
            TCsist.Clear();
            TPem.Clear();
            TPnt.Clear();
            TPml.Clear();
            TPdtr.Clear();
            TPdtc.Clear();
            TPdt2.Clear();
            TPdt1.Clear();
            TCp80.Clear();
            TCp100.Clear();
            TCp120.Clear();
            TCp150.Clear();
            TCt80.Clear();
            TCt100.Clear();
            TCt120.Clear();
            TCt150.Clear();
            TCp80m.Clear();
            TCp100m.Clear();
            TCp120m.Clear();
            TCp150m.Clear();
            TCt80m.Clear();
            TCt100m.Clear();
            TCt120m.Clear();
            TCt150m.Clear();
            SPtp84.Clear();
            SPtp83.Clear();
            SPtp78.Clear();
            SPtp76.Clear();
            SPtp75.Clear();
            SPtp74.Clear();
            SPtp73.Clear();
            SPtp72.Clear();
            SPtp85.Clear();
            SPtp86.Clear();
            SPtp87.Clear();
            SPtp88.Clear();
            SPtp89.Clear();
            SPtp90.Clear();
            SPtp91.Clear();
            SPtp92.Clear();
            SPtp93.Clear();
            SPtp94.Clear();
            SPtp95.Clear();
            SPtp96.Clear();
            SPtp97.Clear();
            SPtp98.Clear();
            SPtp99.Clear();
            SPtp100.Clear();
            SPtp101.Clear();
            SPtp102.Clear();
            SPtp103.Clear();
            SPtp104.Clear();
            SPtp105.Clear();
            SPtp106.Clear();
            SPtp107.Clear();
            SPtp108.Clear();
            SPtp109.Clear();
            SPtp110.Clear();
            SPtp111.Clear();
            SPtp136.Clear();
            SPtp137.Clear();
            SPtp141.Clear();
            SPtp142.Clear();
            SPtp143.Clear();
            SPtp144.Clear();
            SPtp145.Clear();
            SPtp146.Clear();
            SPtp147.Clear();
            SPtp148.Clear();
            SPtp149.Clear();
            SPtp150.Clear();
            SPtp151.Clear();
            TIcc.Clear();
            TQevp.Clear();
            TQevpd.Clear();
            TQevpc.Clear();
            TTint.Clear();
            TEquip.Clear();
            TCint1.Clear();
            TCint2.Clear();
            TCint3.Clear();
            TMcc.Clear();
            TCmce.Clear();
            TPmce.Clear();
            TDmce.Clear();
            Cnoff1.Text = "";
            Coff1.Text = "";
            Cnoff2.Text = "";
            Coff2.Text = "";
            Cnoff3.Text = "";
            Coff3.Text = "";
            Cnoff4.Text = "";
            Coff4.Text = "";
            Cnoff5.Text = "";
            Coff5.Text = "";
            Cnoff6.Text = "";
            Coff6.Text = "";
            Cnoff7.Text = "";
            Coff7.Text = "";
            Cnoff8.Text = "";
            Coff8.Text = "";
            
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
                    //this.SetText(CCamara, "1");
                    CCamara.Text = "1";
                    RBInteriores.Checked = myOferta.GetCam(0).GetIE();
                    CKPanel.Checked = myOferta.GetCam(0).GetKP();
                    CKNanauf.Checked = myOferta.GetCam(0).GetKN();
                    CKMonile.Checked = myOferta.GetCam(0).GetKM();
                    CKCable.Checked = myOferta.GetCam(0).GetKC();
                    CKUnid.Checked = myOferta.GetCam(0).GetKU();
                    CKPuerta.Checked = myOferta.GetCam(0).GetKPR();
                    CKexpo.Checked = myOferta.GetCam(0).GetKexpo();
                    CKepiso.Checked = myOferta.GetCam(0).GetKepiso();
                    CKDrenaje.Checked = myOferta.GetCam(0).GetKD();
                    CKSellaje.Checked = myOferta.GetCam(0).GetKS();
                    CKEmerg.Checked = myOferta.GetCam(0).GetKCE();
                    CKBrida.Checked = myOferta.GetCam(0).GetKB();
                    CKCort.Checked = myOferta.GetCam(0).GetKCO();
                    CKTor.Checked = myOferta.GetCam(0).GetKTO();
                    CKTcobre.Checked = myOferta.GetCam(0).GetKTC();
                    CKRefrig.Checked = myOferta.GetCam(0).GetKRE();
                    CKSopr.Checked = myOferta.GetCam(0).GetKSO();
                    CKValv.Checked = myOferta.GetCam(0).GetKVA();
                    CKCelect.Checked = myOferta.GetCam(0).GetKCL();
                    CKPerf.Checked = myOferta.GetCam(0).GetKPE();
                    CKAUni.Checked = myOferta.GetCam(0).GetKUA();
                    CKAlum.Checked = myOferta.GetCam(0).GetKAL();
                    CKMobra.Checked = myOferta.GetCam(0).GetKMO();
                    CKSD.Checked = myOferta.GetCam(0).GetKSD();
                    CKSMin.Checked = myOferta.GetCam(0).GetKSMin();
                    CKPAI.Checked = myOferta.GetCam(0).GetKPAI();
                    CKpmtal.Checked = myOferta.GetCam(0).GetKpmtal();
                    CKlux.Checked = myOferta.GetCam(0).GetKlux();
                    CKvsol.Checked = myOferta.GetCam(0).GetKvsol();
                    CKp10.Checked = myOferta.GetCam(0).GetKp10();
                    CKp12.Checked = myOferta.GetCam(0).GetKp12();
                    CKp15.Checked = myOferta.GetCam(0).GetKp15();
                    CKp15t.Checked = myOferta.GetCam(0).GetKp15t();
                    CKdt.Checked = myOferta.GetCam(0).GetKdt();
                    CKat.Checked = myOferta.GetCam(0).GetKat();
                    CKbt.Checked = myOferta.GetCam(0).GetKbt();
                    CKmt.Checked = myOferta.GetCam(0).GetKmt();
                    CKmod.Checked = myOferta.GetCam(0).GetKmod();
                    CRvent.Checked = myOferta.GetCam(0).GetRvent();
                    CKantc.Checked = myOferta.GetCam(0).GetKantc();
                    CKppc.Checked = myOferta.GetCam(0).GetKppc();
                    CKepc.Checked = myOferta.GetCam(0).GetKepc();
                    CKcion.Checked = myOferta.GetCam(0).GetKcion();
                    CKastre.Checked = myOferta.GetCam(0).GetKastre();
                    CKpcion.Checked = myOferta.GetCam(0).GetKpcion();
                    CKfrio.Checked = myOferta.GetCam(0).GetKfrio();
                    CKeq1.Checked = myOferta.GetCam(0).GetKeq1();
                    CKeq2.Checked = myOferta.GetCam(0).GetKeq2();
                    CKeq3.Checked = myOferta.GetCam(0).GetKeq3();
                    CKsu1.Checked = myOferta.GetCam(0).GetKsu1();
                    CKsu2.Checked = myOferta.GetCam(0).GetKsu2();
                    CKsu3.Checked = myOferta.GetCam(0).GetKsu3();
                    CKpps.Checked = myOferta.GetCam(0).GetKpps();

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
                catch {}/*(Exception ex)
                {
                    MessageBox.Show(ex.Message,
                                "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }*/
            }
            else
                MessageBox.Show("No hay ninguna camara seleccionada",
                            "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
            if (int.Parse(CCamara.Text) <= 5)
            {
                CBexpo.Visible = true;
                CKepiso.Visible = true;
                CKexpo.Visible = true;
            }
            if (int.Parse(CCamara.Text) <= 6)
            {
                Coff1.Visible = true;
                Cnoff1.Visible = true;
                Coff2.Visible = true;
                Cnoff2.Visible = true;
                Coff3.Visible = true;
                Cnoff3.Visible = true;
                Coff4.Visible = true;
                Cnoff4.Visible = true;
                Coff5.Visible = true;
                Cnoff5.Visible = true;
                Coff6.Visible = true;
                Cnoff6.Visible = true;
                Coff7.Visible = true;
                Cnoff7.Visible = true;
                Coff8.Visible = true;
                Cnoff8.Visible = true;
                label71.Visible = true;
            }
                
            
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


            if (int.Parse(CCamara.Text) > 5)
            {
                CBexpo.Visible = false;
                CKepiso.Visible = false;
                CKexpo.Visible = false;
            }
            if (int.Parse(CCamara.Text) > 6)
            {
                Coff1.Visible = false;
                Cnoff1.Visible = false;
                Coff2.Visible = false;
                Cnoff2.Visible = false;
                Coff3.Visible = false;
                Cnoff3.Visible = false;
                Coff4.Visible = false;
                Cnoff4.Visible = false;
                Coff5.Visible = false;
                Cnoff5.Visible = false;
                Coff6.Visible = false;
                Cnoff6.Visible = false;
                Coff7.Visible = false;
                Cnoff7.Visible = false;
                Coff8.Visible = false;
                Cnoff8.Visible = false;
                label71.Visible = false;
            }
                    
            
        }
        
        private void groupBox1_Enter(object sender, EventArgs e)
        {

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

        private void RUSD_CheckedChanged(object sender, EventArgs e)
        {

        }
        private void RGX_CheckedChanged(object sender, EventArgs e)
        { 
        
        }
        private void button1_Click_1(object sender, EventArgs e)
        {

            Form4 myPassForm = new Form4(this);           
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

        private void button2_Click(object sender, EventArgs e)
        {
            this.calcbeging();
            this.Calcular2();
        }
        
       
        public string Calcular(int cam)
        {
            
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
            CalcWorkSheet.Cells[14, 13] = myOferta.GetCam(cam - 1).GetLargo();
            CalcWorkSheet.Cells[15, 13] = myOferta.GetCam(cam - 1).GetAncho();
            CalcWorkSheet.Cells[16, 13] = myOferta.GetCam(cam - 1).GetAlto();
            CalcWorkSheet.Cells[18, 13] = myOferta.GetCam(cam - 1).GetTemp();
            CalcWorkSheet.Cells[31, 18] = myOferta.GetCam(cam - 1).GetCtpd();
            CalcWorkSheet.Cells[27, 13] = myOferta.GetCam(cam - 1).GetCantEv();
            CalcWorkSheet.Cells[31, 27] = myOferta.GetCam(cam - 1).GetCdin();
            CalcWorkSheet.Cells[31, 25] = myOferta.GetCam(cam - 1).GetCxp();   
            CalcWorkSheet.Cells[31, 26] = myOferta.GetCam(cam - 1).GetCentx();
            CalcWorkSheet.Cells[31, 20] = myOferta.GetCam(cam - 1).GetTMevp();
            CalcWorkSheet.Cells[26, 13] = myOferta.GetCam(cam - 1).GetCTamb();
            CalcWorkSheet.Cells[31, 28] = myOferta.GetCam(cam - 1).GetCtxv();
            CalcWorkSheet.Cells[27, 10] = myOferta.GetCam(cam - 1).GetCBint();
            CalcWorkSheet.Cells[31, 4] = myOferta.GetCam(cam - 1).GetCBcm();
            CalcWorkSheet.Cells[32, 4] = myOferta.GetCam(cam - 1).GetCCmci();
            CalcWorkSheet.Cells[24, 9] = ("=REDONDEAR.MAS(CONVERTIR(F24;\"Wh\";\"BTU\");0)");
            CalcWorkSheet.Cells[29, 15] = ("=SI.ERROR(U31;1)");
            CalcWorkSheet.Cells[29, 13] = ("=SI(O29<>1;INDICE(W31:W336;U31);SI(O29=1;75))");
            CalcWorkSheet.Cells[30, 13] = ("=SI(O29<>1;INDICE(X31:X336;U31);SI(O29=1;0))");
            CalcWorkSheet.Cells[22, 16] = myOferta.GetCam(cam - 1).GetCSumi();
            //CalcWorkSheet.Cells[23, 16] = CCamara.Text;
            for (int i = 1; i <= cam; i++)
            {
                CCam myCam = myOferta.GetCam(i - 1);
                CalcWorkSheet.Cells[i + 33, 5] = myOferta.GetCam(i - 1).GetCantEv();
                CalcWorkSheet.Cells[i + 33, 6] = myOferta.GetCam(i - 1).GetFW();
                CalcWorkSheet.Cells[i + 33, 7] = myOferta.GetCam(i - 1).GetQfep();
                CalcWorkSheet.Cells[i + 33, 8] = myOferta.GetCam(i - 1).GetCTEvap();
                CalcWorkSheet.Cells[i + 33, 9] = myOferta.GetCam(i - 1).GetScdro();
                CalcWorkSheet.Cells[i + 33, 10] = myOferta.GetCam(i - 1).GetCSumi();
                CalcWorkSheet.Cells[i + 33, 11] = myOferta.GetCam(i - 1).GetRefrig();
                CalcWorkSheet.Cells[i + 33, 38] = myOferta.GetCam(i - 1).GetCBint();
                CalcWorkSheet.Cells[i + 33, 35] = myOferta.GetCam(i - 1).GetTInc();
                CalcWorkSheet.Cells[i + 33, 36] = myOferta.GetCam(i - 1).GetTInev();
                CalcWorkSheet.Cells[i + 33, 37] = myOferta.GetCam(i - 1).GetTIned();
                CalcWorkSheet.Cells[i + 33, 4] = myOferta.GetCam(i - 1).GetCBcm();
                CalcWorkSheet.Cells[i + 33, 3] = myOferta.GetCam(i - 1).GetCCmci();
                CalcWorkSheet.Cells[i + 33, 12] = myOferta.GetCam(i - 1).GetTDmc1();
                CalcWorkSheet.Cells[i + 33, 90] = myOferta.GetCam(i - 1).GetCtxv();
                CalcWorkSheet.Cells[i + 33, 93] = myOferta.GetCam(i - 1).GetKsu1().ToString();// Tuberia extención tramo 1
                CalcWorkSheet.Cells[i + 33, 94] = myOferta.GetCam(i - 1).GetKsu2().ToString();// Tuberia extención tramo 2
                CalcWorkSheet.Cells[i + 33, 95] = myOferta.GetCam(i - 1).GetKsu3().ToString();// Tuberia extención tramo 3

            }
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[22];
            for (int a = 1; a <= cam; a++)
            {
                CCam myCam = myOferta.GetCam(a - 1);
                CalcWorkSheet.Cells[a + 14, 2] = "";
                CalcWorkSheet.Cells[a + 14, 2] = "=SI.ERROR(SI(B8=" + a.ToString() + ";SI(F10=\"INTEG\";INDICE(P1:Y14;1;B8);SI(F10=\"400QS\";INDICE(P21:P58;B8);0)));0)";
            }
            /*
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[22];
            fo0r (int b = 1; b <= cam; b++)
            {
                CCam myCam = myOferta.GetCam(b - 1);
                CalcWorkSheet.Cells[b + 14, 1] = "=SI.ERROR(SI(B8=" + b.ToString() + ";SI(F10=\"INTEG\";INDICE(P15:Y15;COINCIDIR(C10;P18:Y18;0));INDICE(L15:L52;B8));0)";

            }
            
           */
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[22];
            CalcWorkSheet.Cells[11, 8] = "=CONTAR.SI.CONJUNTO(I15:I52;\"=INTEGRADO1\";M15:M52;\">0\")";
            CalcWorkSheet.Cells[12, 8] = "=SI(H11>1;2,65;2,1)";
            //CalcWorkSheet.Cells[6, 3] = "=SI(F10=\"INTEG\";INDICE(P1:Y14;1;C11);INDICE(P21:P58;B8))";
            //CalcWorkSheet.Cells[9, 3] = "=SI(F10=\"INTEG\";INDICE(P15:Y15;COINCIDIR(C10;P18:Y18;0));INDICE(L15:L52;B8))";
            CalcWorkSheet.Cells[15, 16] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO1\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO1\")>=1;1)*2400*H12";
            CalcWorkSheet.Cells[15, 17] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO2\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO2\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 18] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO3\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO3\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 19] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO4\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO4\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 20] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO5\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO5\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 21] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO6\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO6\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 22] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO7\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO7\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 23] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO8\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO8\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 24] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO9\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO9\")>=1;1)*2400";
            CalcWorkSheet.Cells[15, 25] = "=SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO10\")+SI(SUMAR.SI.CONJUNTO($L$15:$L$52;$I$15:$I$52;\"INTEGRADO10\")>=1;1)*2400";
            CalcWorkSheet.Cells[16, 16] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=1\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 17] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=2\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 18] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=3\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 19] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=4\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 20] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=5\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 21] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=6\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 22] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=7\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 23] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=8\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 24] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=9\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[16, 25] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=10\";$D$15:$D$52;\">=-5\")";
            CalcWorkSheet.Cells[17, 16] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=1\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 17] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=2\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 18] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=3\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 19] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=4\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 20] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=5\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 21] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=6\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 22] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=7\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 23] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=8\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 24] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=9\";$D$15:$D$52;\"<-5\")";
            CalcWorkSheet.Cells[17, 25] = "=CONTAR.SI.CONJUNTO($M$15:$M$52;\"=10\";$D$15:$D$52;\"<-5\")";
            //**************************************************************************** DATOS CUADRO ELECTRICO**
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[22];
            CalcWorkSheet.Cells[8, 27] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\">=5\";Datos!CL34:CL70;\"=TX3\")*114";
            CalcWorkSheet.Cells[8, 28] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<5\";Datos!H34:H70;\">-6\";Datos!CL34:CL70;\"=TX3\")*300";
            CalcWorkSheet.Cells[8, 29] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<=-6\";Datos!CL34:CL70;\"=TX3\")*1751";
            CalcWorkSheet.Cells[9, 27] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\">=5\";Datos!CL34:CL70;\"=EX2-M00\")*355";
            CalcWorkSheet.Cells[9, 28] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<5\";Datos!H34:H70;\">-6\";Datos!CL34:CL70;\"=EX2-M00\")*355";
            CalcWorkSheet.Cells[9, 29] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<=-6\";Datos!CL34:CL70;\"=EX2-M00\")*1751";
            CalcWorkSheet.Cells[10, 27] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\">=5\";Datos!CL34:CL70;\"=EX3-2000\")*455";
            CalcWorkSheet.Cells[10, 28] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<5\";Datos!H34:H70;\">-6\";Datos!CL34:CL70;\"=EX3-2000\")*455";
            CalcWorkSheet.Cells[10, 29] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<=-6\";Datos!CL34:CL70;\"=EX3-2000\")*1851";
            CalcWorkSheet.Cells[11, 27] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\">=5\";Datos!CL34:CL70;\"=TCLE\")*114";
            CalcWorkSheet.Cells[11, 28] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<5\";Datos!H34:H70;\">-6\";Datos!CL34:CL70;\"=TCLE\")*300";
            CalcWorkSheet.Cells[11, 29] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<=-6\";Datos!CL34:CL70;\"=TCLE\")*1651";
            CalcWorkSheet.Cells[12, 27] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\">=5\";Datos!CL34:CL70;\"=DANF\")*114";
            CalcWorkSheet.Cells[12, 28] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<5\";Datos!H34:H70;\">-6\";Datos!CL34:CL70;\"=DANF\")*300";
            CalcWorkSheet.Cells[12, 29] = "=SUMAR.SI.CONJUNTO(Datos!E34:E70;Datos!H34:H70;\"<=-6\";Datos!CL34:CL70;\"=DANF\")*1651";

            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
            for (int b = 1; b <= cam; b++)
            {
                CCam myCam = myOferta.GetCam(b - 1);
                CalcWorkSheet.Cells[9, 46] = "=SUMAR.SI.CONJUNTO(AS12:AS46;AU12:AU46;AT8)";
            }
            CalcWorkSheet.Cells[6, 11] = "=SI.ERROR(SI(K3=0;0;AP26);0)";
            //*****************************************************************************
            Range NewRangeCalc = CalcWorkSheet.get_Range("R9", "R22");
            Array myNewArr = (Array)NewRangeCalc.Value2;
            String[,] myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
            for (int i = 0; i < myNewArr.GetLength(0); i++)
            {
                for (int j = 0; j < myNewArr.GetLength(1); j++)
                {
                    long[] indices = new long[] { i + 1, j + 1 };
                    myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                }
            }
           
            //CalcWorkBook.SaveCopyAs(local + "\\prueba.xls");
            return myNewArrCalc.GetValue(12, 1).ToString();
            return myNewArrCalc.GetValue(13,1).ToString();
            return myNewArrCalc.GetValue(14,1).ToString();
            CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
            //CalcWorkSheet.Cells[16, 26] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<=-5\")");
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
                //TPcCUC.Text = TrabExcel.myArray.GetValue(1, 1).ToString();
                //TPcCUP.Text = TrabExcel.myArray.GetValue(2, 1).ToString();
                //Tcuc.Text = TrabExcel.myArray.GetValue(3, 1).ToString();
                //TCup.Text = TrabExcel.myArray.GetValue(4, 1).ToString();
                //TConsC.Text = TrabExcel.myArray.GetValue(6, 1).ToString();
                //TEfriG.Text = TrabExcel.myArray.GetValue(7, 1).ToString();
                //TPanel.Text = TrabExcel.myArray.GetValue(8, 1).ToString();
                
                //TCvta.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTCvta();
                //TValv.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTValv();
                //TSol.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTSol();
                
                //TMuc.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTMuc();
                //TMevp.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTMevp();
                //Txm2.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTxm2();
                //TCuadro.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetTCuadro();
                //CBint.Text = myOferta.GetCam(int.Parse(CCamara.Text)-1).GetCBint();
                //CBexpo.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCBexpo();
                //Cevap.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCevap();
                //Cmodex.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCmodex();
                //Ctxv.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCtxv();
                
                
                //CSumi.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCSumi();
                //TFW.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetFW();
                //CTCond.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCTCond();
                //CTamb.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCTamb();
                //Ctpd.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCtpd();
                //Cnoff3.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCnoff3();
                //Coff3.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCoff3();
                //Cnoff2.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCnoff2();
                //Coff2.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCoff2();
                //Cnoff.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCnoff();
                //Coff.Text= myOferta.GetCam(int.Parse(CCamara.Text)-1).GetCoff();
                //CTEvap.Text = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCTEvap();
                
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
            //this.Calcular2();
            this.GuardarDatos();
            ProCamara = true;
            this.Processar();
        }

        private void TCantEv_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (int.Parse(TCantEv.Text) > 1)
                label49.Visible = true;
            else
                label49.Visible = false;
        }
        private void CKmod_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (CKmod.Checked==true)
                Cnoff8.Visible = false;
        }

        private void CKppc_Click(object sender, EventArgs e)
        {
            if (CKppc.Checked == true)
                label118.Visible = false;
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
                            //string value = myOferta.GetCam(int.Parse(CCamara.Text) - 1).GetCodValv();
                            CCam myCam = new CCam(TNC.Text, TTem.Text, TLargo.Text,
                                TAncho.Text, TAlto.Text, TTP.Text, CDT.Text, CCE.Text,
                                TCF.Text,  TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text, TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text, CITPuerta.Text, TDEC.Text, TDECE.Text, TDECH.Text, TDECF.Text, TCantEv.Text,
                                TCentx.Text, TCdin.Text, TCxp.Text, CRefrig.Text, CDigt.Text, CCcion.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text, RBInteriores.Checked, CKPanel.Checked,
                                CKNanauf.Checked, CKMonile.Checked, CKCable.Checked, CKUnid.Checked, CKPuerta.Checked, CKexpo.Checked, CKepiso.Checked,
                                CKDrenaje.Checked, CKSellaje.Checked, CKEmerg.Checked, CKBrida.Checked, CKCort.Checked,
                                CKTor.Checked, CKTcobre.Checked, CKRefrig.Checked, CKSopr.Checked, CKValv.Checked, 
                                CKCelect.Checked, CKPerf.Checked, CKAUni.Checked, CKAlum.Checked, CKMobra.Checked, CKSD.Checked,
                                CKSMin.Checked, CKPAI.Checked, CKpmtal.Checked, CKlux.Checked, CKvsol.Checked, CKp10.Checked, 
                                CKp12.Checked, CKp15.Checked, CKp15t.Checked, CKdt.Checked, CKat.Checked, CKbt.Checked, CKmt.Checked, CKmod.Checked, CRvent.Checked, CKantc.Checked, CKppc.Checked, CKepc.Checked, CKcion.Checked, CKastre.Checked, CKpcion.Checked, CKfrio.Checked, CKeq1.Checked, CKeq2.Checked, CKeq3.Checked, CKsu1.Checked, CKsu2.Checked, CKsu3.Checked, CKpps.Checked,  CVolt.Text, CFase.Text, TConD.Text, TMuc.Text, TMevp.Text,
                                TSol.Text, TValv.Text, 
                                TCvta.Text, Txm2.Text, TCuadro.Text, CBint.Text, CBcm.Text, CCmci.Text, CBexpo.Text,
                                Cevap.Text, Cmodex.Text, Ctxv.Text, CSumi.Text, CTCond.Text,
                                CTamb.Text, Ctpd.Text, Cnoff1.Text, Coff1.Text, Cnoff2.Text, Coff2.Text, Cnoff3.Text, Coff3.Text, Cnoff4.Text, Coff4.Text, Cnoff5.Text, Coff5.Text, Cnoff6.Text, Coff6.Text, Cnoff7.Text, Coff7.Text, Cnoff8.Text, Coff8.Text, CTEvap.Text, TCodValv.Text, TTmos.Text, TInc.Text, TInev.Text, TIned.Text, TIncd.Text,
                                TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TQevpc.Text, TTint.Text, TEquip.Text, TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text,
                                TDmce.Text, TDmc1.Text, TDmc11.Text, TDmc12.Text, TDlq1.Text, TDlq2.Text, TDlq3.Text, TDlq11.Text, TDlq21.Text, TDlq31.Text, TDsu130.Text, TDsu230.Text, TDsu330.Text, TDsu110.Text, TDsu210.Text, TDsu310.Text, TDsu105.Text, TDsu205.Text, TDsu305.Text, TDmc2.Text, TDmc3.Text, TDmc4.Text, TDmc5.Text, TDmc6.Text, TDmc7.Text, TDmc8.Text, TLcc.Text, TLss.Text, TPcmc.Text, TPcmc1.Text, TPcmc2.Text, TPcmc3.Text, TPcond.Text, TPlq.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcv.Text, TPcx.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, TPnt.Text, TPml.Text, TPdtr.Text, TPdtc.Text, TPdt2.Text, TPdt1.Text, TCp80.Text, TCp100.Text, TCp120.Text, TCp150.Text, TCt80.Text, TCt100.Text, TCt120.Text, TCt150.Text, TCp80m.Text, TCp100m.Text, TCp120m.Text, TCp150m.Text, TCt80m.Text, TCt100m.Text, TCt120m.Text, TCt150m.Text,
                                SPtp84.Text, SPtp83.Text, SPtp78.Text, SPtp76.Text, SPtp75.Text, SPtp74.Text, SPtp73.Text, SPtp72.Text, SPtp85.Text, SPtp86.Text, SPtp87.Text, SPtp88.Text, SPtp89.Text, SPtp90.Text, SPtp91.Text, SPtp92.Text, SPtp93.Text, SPtp94.Text, SPtp95.Text, SPtp96.Text, SPtp97.Text, SPtp98.Text, SPtp99.Text, SPtp100.Text, SPtp101.Text, SPtp102.Text, SPtp103.Text, SPtp104.Text, SPtp105.Text, SPtp106.Text, SPtp107.Text,
                                SPtp108.Text, SPtp109.Text, SPtp110.Text, SPtp111.Text, SPtp136.Text, SPtp137.Text, SPtp141.Text, SPtp142.Text, SPtp143.Text, SPtp144.Text, SPtp145.Text, SPtp146.Text, SPtp147.Text, SPtp148.Text, SPtp149.Text, SPtp150.Text, SPtp151.Text, TIn1.Text, TIn2.Text, TIn3.Text, TIn4.Text, TIn5.Text, TIn6.Text, TIn7.Text, TIn8.Text, TIn9.Text, TIn10.Text, TIn11.Text, TIn12.Text, TIn13.Text, TIn14.Text, TIn15.Text, TIn16.Text, TIn17.Text, TIn18.Text, TIn19.Text, TIn20.Text, TIn21.Text, TIn22.Text, TIn23.Text, TIp1.Text, TIp2.Text, TIp3.Text, TIp4.Text, TIp5.Text, TIp6.Text, TIp7.Text, TIp8.Text, TIp9.Text, TIp10.Text, TIp11.Text, TIp12.Text, TIp13.Text, TIp14.Text, TIp15.Text, TIp16.Text, TIp17.Text, TIp18.Text, TIp19.Text, TIp20.Text, TIp21.Text, TIp23.Text, pk);
                            //myCam.SetCodValv(value);
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

        private void CDT_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                CTEvap.Text = (int.Parse(TTem.Text) - int.Parse(CDT.Text)).ToString();
            }
            catch { }
        }

        void Calcular2()
        {
            try
            {
                TCF.Text = Calcular(int.Parse(CCamara.Text));
                TFW.Text = Calcular(int.Parse(CCamara.Text));
                TCmod.Text = Calcular(int.Parse(CCamara.Text));
                TSpsi.Text = Calcular(int.Parse(CCamara.Text));
                TStemp.Text = Calcular(int.Parse(CCamara.Text));
                TApsi.Text = Calcular(int.Parse(CCamara.Text));
                TEmevp.Text = Calcular(int.Parse(CCamara.Text));
                TDesc.Text = Calcular(int.Parse(CCamara.Text));
                TPrec.Text = Calcular(int.Parse(CCamara.Text));
                TQfep.Text = Calcular(int.Parse(CCamara.Text));
                TScdro.Text = Calcular(int.Parse(CCamara.Text));
                CCam newCam = new CCam(TNC.Text, TTem.Text, TLargo.Text,
                                    TAncho.Text, TAlto.Text, TTP.Text, CDT.Text, CCE.Text,
                                    TCF.Text, TFW.Text, TQfw.Text, TCmod.Text, TCmodd.Text, TCmodp.Text, TDesc.Text, TPrec.Text, TQfep.Text, TScdro.Text, TSpsi.Text, TStemp.Text, TApsi.Text, TEmevp.Text, CSup.Text, CITPuerta.Text, TDEC.Text, TDECE.Text, TDECH.Text, TDECF.Text, TCantEv.Text, TCentx.Text, TCdin.Text, TCxp.Text, CRefrig.Text, CDigt.Text, CCcion.Text, CCastre.Text, CCpcion.Text, CCfrio.Text, CCeq1.Text, CCeq2.Text, CCeq3.Text,
                                    RBInteriores.Checked, CKPanel.Checked, CKNanauf.Checked, CKMonile.Checked, CKCable.Checked,
                                    CKUnid.Checked, CKPuerta.Checked, CKexpo.Checked, CKepiso.Checked, CKDrenaje.Checked, CKSellaje.Checked, CKEmerg.Checked, CKBrida.Checked,
                                    CKCort.Checked, CKTor.Checked, CKTcobre.Checked, CKRefrig.Checked, CKSopr.Checked, CKValv.Checked, CKCelect.Checked, CKPerf.Checked, CKAUni.Checked, CKAlum.Checked, CKMobra.Checked, CKSD.Checked,
                                    CKSMin.Checked, CKPAI.Checked, CKpmtal.Checked, CKlux.Checked, CKvsol.Checked, CKp10.Checked, CKsu1.Checked, CKsu2.Checked, CKsu3.Checked, CKpps.Checked,
                                    CKp12.Checked, CKp15.Checked, CKp15t.Checked, CKdt.Checked, CKat.Checked, CKbt.Checked, CKmt.Checked, CKmod.Checked, CKantc.Checked, CKppc.Checked, CKcion.Checked, CKastre.Checked, CKpcion.Checked, CKfrio.Checked, CKeq1.Checked, CKeq2.Checked, CKeq3.Checked, CRvent.Checked, CKepc.Checked, CVolt.Text, CFase.Text, TConD.Text, TMuc.Text, TMevp.Text,
                                    TSol.Text, TValv.Text, 
                                    TCvta.Text, Txm2.Text, TCuadro.Text, CBint.Text, CBcm.Text, CCmci.Text, CBexpo.Text,
                                    Cevap.Text, Cmodex.Text, Ctxv.Text, CSumi.Text, CTCond.Text, CTamb.Text, Ctpd.Text, Cnoff1.Text, Coff1.Text, Cnoff2.Text, Coff2.Text, Cnoff3.Text, Coff3.Text, Cnoff4.Text, Coff4.Text,
                                    Cnoff5.Text, Coff5.Text, Cnoff6.Text, Coff6.Text, Cnoff7.Text, Coff7.Text, Cnoff8.Text, Coff8.Text,
                                    CTEvap.Text, TCodValv.Text, TTmos.Text, TInc.Text, TInev.Text, TIned.Text, TIncd.Text, TIpv.Text, TIcc.Text, TQevp.Text, TQevpd.Text, TQevpc.Text, TTint.Text, TEquip.Text, TCint1.Text, TCint2.Text, TCint3.Text, TMcc.Text, TCmce.Text, TPmce.Text,
                                    TDmce.Text, TDmc1.Text, TDmc11.Text, TDmc12.Text, TDlq1.Text, TDlq2.Text, TDlq3.Text, TDlq1.Text, TDlq2.Text, TDlq3.Text, TDsu130.Text, TDsu230.Text, TDsu330.Text, TDsu110.Text, TDsu210.Text, TDsu310.Text, TDsu105.Text, TDsu205.Text, TDsu305.Text, TDmc2.Text, TDmc3.Text, TDmc4.Text, TDmc5.Text, TDmc6.Text, TDmc7.Text, TDmc8.Text, TLcc.Text, TLss.Text, TPcmc.Text, TPcmc1.Text, TPcmc2.Text, TPcmc3.Text, TPcond.Text, TPlq.Text, TPvq.Text, TPls.Text, TPosc.Text, TPsq.Text, TPcq.Text, TPcv.Text, TPcx.Text, TPcy.Text, TPcp.Text, TPex.Text, TPrs.Text, TCsist.Text, TPem.Text, TPnt.Text, TPml.Text, TPdtr.Text, TPdtc.Text, TPdt2.Text, TPdt1.Text, TCp80.Text, TCp100.Text, TCp120.Text, TCp150.Text, TCt80.Text, TCt100.Text, TCt120.Text, TCt150.Text, TCp80m.Text, TCp100m.Text, TCp120m.Text, TCp150m.Text, TCt80m.Text, TCt100m.Text, TCt120m.Text, TCt150m.Text,
                                    SPtp84.Text, SPtp83.Text, SPtp78.Text, SPtp76.Text, SPtp75.Text, SPtp74.Text, SPtp73.Text, SPtp72.Text, SPtp85.Text, SPtp86.Text, SPtp87.Text, SPtp88.Text, SPtp89.Text, SPtp90.Text, SPtp91.Text, SPtp92.Text, SPtp93.Text, SPtp94.Text, SPtp95.Text, SPtp96.Text, SPtp97.Text, SPtp98.Text, SPtp99.Text, SPtp100.Text, SPtp101.Text, SPtp102.Text, SPtp103.Text, SPtp104.Text, SPtp105.Text, SPtp106.Text, SPtp107.Text,
                                    SPtp108.Text, SPtp109.Text, SPtp110.Text, SPtp111.Text, SPtp136.Text, SPtp137.Text, SPtp141.Text, SPtp142.Text, SPtp143.Text, SPtp144.Text, SPtp145.Text, SPtp146.Text, SPtp147.Text, SPtp148.Text, SPtp149.Text, SPtp150.Text, SPtp151.Text, TIn1.Text, TIn2.Text, TIn3.Text, TIn4.Text, TIn5.Text, TIn6.Text, TIn7.Text, TIn8.Text, TIn9.Text, TIn10.Text, TIn11.Text, TIn12.Text, TIn13.Text, TIn14.Text, TIn15.Text, TIn16.Text, TIn17.Text, TIn18.Text, TIn19.Text, TIn20.Text, TIn21.Text, TIn22.Text, TIn23.Text, TIp1.Text, TIp2.Text, TIp3.Text, TIp4.Text, TIp5.Text, TIp6.Text, TIp7.Text, TIp8.Text, TIp9.Text, TIp10.Text, TIp11.Text, TIp12.Text, TIp13.Text, TIp14.Text, TIp15.Text, TIp16.Text, TIp17.Text, TIp18.Text, TIp19.Text, TIp20.Text, TIp21.Text, TIp23.Text, pk);
                myOferta.actualizar(newCam, int.Parse(CCamara.Text));

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                CalcWorkSheet.Cells[30, 18] = CVolt.Text;
                CalcWorkSheet.Cells[34, 39] = CKAUni.Checked.ToString();//39
                CalcWorkSheet.Cells[34, 40] = CKValv.Checked.ToString();//40
                CalcWorkSheet.Cells[34, 41] = CKMonile.Checked.ToString();//41
                CalcWorkSheet.Cells[34, 42] = CKUnid.Checked.ToString();//42
                CalcWorkSheet.Cells[34, 43] = CKTcobre.Checked.ToString();//43
                CalcWorkSheet.Cells[34, 44] = CKCelect.Checked.ToString();//44
                CalcWorkSheet.Cells[34, 45] = CKAlum.Checked.ToString();//45
                CalcWorkSheet.Cells[34, 46] = CKMobra.Checked.ToString();//46
                CalcWorkSheet.Cells[34, 47] = CKSD.Checked.ToString();//47
                CalcWorkSheet.Cells[34, 48] = CKlux.Checked.ToString();//48
                CalcWorkSheet.Cells[34, 49] = CKvsol.Checked.ToString();//49
                CalcWorkSheet.Cells[34, 50] = CKCort.Checked.ToString();//50
                CalcWorkSheet.Cells[34, 51] = CKTor.Checked.ToString();//51
                CalcWorkSheet.Cells[34, 52] = CKPanel.Checked.ToString();//52
                CalcWorkSheet.Cells[34, 53] = CKSopr.Checked.ToString();//53
                CalcWorkSheet.Cells[34, 54] = CKCable.Checked.ToString();//54
                CalcWorkSheet.Cells[34, 55] = CKPuerta.Checked.ToString();//55
                CalcWorkSheet.Cells[34, 56] = CKDrenaje.Checked.ToString();//56
                CalcWorkSheet.Cells[34, 57] = CKSellaje.Checked.ToString();//57
                CalcWorkSheet.Cells[34, 58] = CKEmerg.Checked.ToString();//58
                CalcWorkSheet.Cells[34, 59] = CKBrida.Checked.ToString();//59
                CalcWorkSheet.Cells[34, 60] = CKNanauf.Checked.ToString();//60
                CalcWorkSheet.Cells[34, 61] = CKRefrig.Checked.ToString();//61
                CalcWorkSheet.Cells[34, 62] = CKPerf.Checked.ToString();//62
                CalcWorkSheet.Cells[34, 63] = CKpmtal.Checked.ToString();//62
                CalcWorkSheet.Cells[34, 64] = CRvent.Checked.ToString();//64
                CalcWorkSheet.Cells[34, 65] = CKppc.Checked.ToString();//65
                CalcWorkSheet.Cells[34, 66] = CKepc.Checked.ToString();//66
                CalcWorkSheet.Cells[9, 13] =  CKp10.Checked.ToString();
                CalcWorkSheet.Cells[10, 13] = CKp12.Checked.ToString();
                CalcWorkSheet.Cells[11, 13] = CKp15.Checked.ToString();
                CalcWorkSheet.Cells[12, 13] = CKp15t.Checked.ToString();
                CalcWorkSheet.Cells[14, 17] = CFase.Text;
                CalcWorkSheet.Cells[5, 13] = RBun.Checked.ToString();
                CalcWorkSheet.Cells[4, 13] = RBun2.Checked.ToString();
                CalcWorkSheet.Cells[16, 29] = RBun3.Checked.ToString();
                CalcWorkSheet.Cells[16, 27] = RBun.Checked.ToString();// BT+ MT
                CalcWorkSheet.Cells[34, 67] = CKsu1.Checked.ToString();//67
                CalcWorkSheet.Cells[34, 68] = CKsu2.Checked.ToString();//68
                CalcWorkSheet.Cells[34, 69] = CKsu3.Checked.ToString();//69

                CalcWorkSheet.Cells[34,  99] = "=SI(CU26=1;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($BY$35:$BY$71;$CR$34:$CR$70;\"=1\")";//  99 SxCMC1 SUCCION  +5°C T1
                CalcWorkSheet.Cells[34, 100] = "=SI(CU26=1;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($BY$35:$BY$71;$CS$34:$CS$70;\"=1\")";// 100 SxCMC1 SUCCION  +5°C T2
                CalcWorkSheet.Cells[34, 101] = "=SI(CU26=1;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($BY$35:$BY$71;$CT$34:$CT$70;\"=1\")";// 101 SxCMC1 SUCCION  +5°C T3
                CalcWorkSheet.Cells[34, 102] = "=SI(CU26=1;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($BZ$35:$BZ$71;$CR$34:$CR$70;\"=1\")";// 102 SxCMC1 SUCCION -10°C T1
                CalcWorkSheet.Cells[34, 103] = "=SI(CU26=1;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($BZ$35:$BZ$71;$CR$34:$CR$70;\"=1\")";// 103 SxCMC1 SUCCION -10°C T2
                CalcWorkSheet.Cells[34, 104] = "=SI(CU26=1;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($BZ$35:$BZ$71;$CT$34:$CT$70;\"=1\")";// 104 SxCMC1 SUCCION -10°C T3
                CalcWorkSheet.Cells[34, 105] = "=SI(CU26=1;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CA$35:$CA$71;$CR$34:$CR$70;\"=1\")";// 105 SxCMC1 SUCCION -30°C T1
                CalcWorkSheet.Cells[34, 106] = "=SI(CU26=1;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CA$35:$CA$71;$CS$34:$CS$70;\"=1\")";// 106 SxCMC1 SUCCION -30°C T2
                CalcWorkSheet.Cells[34, 107] = "=SI(CU26=1;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CA$35:$CA$71;$CT$34:$CT$70;\"=1\")";// 107 SxCMC1 SUCCION -30°C T3

                CalcWorkSheet.Cells[39,  99] = "=SI(CU26=2;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CB$35:$CB$71;$CR$34:$CR$70;\"=1\")";//  99 SxCMC2 SUCCION  +5°C T1
                CalcWorkSheet.Cells[39, 100] = "=SI(CU26=2;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CB$35:$CB$71;$CS$34:$CS$70;\"=1\")";// 100 SxCMC2 SUCCION  +5°C T2
                CalcWorkSheet.Cells[39, 101] = "=SI(CU26=2;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CB$35:$CB$71;$CT$34:$CT$70;\"=1\")";// 101 SxCMC2 SUCCION  +5°C T3
                CalcWorkSheet.Cells[39, 102] = "=SI(CU26=2;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CC$35:$CC$71;$CR$34:$CR$70;\"=1\")";// 102 SxCMC2 SUCCION -10°C T1
                CalcWorkSheet.Cells[39, 103] = "=SI(CU26=2;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CC$35:$CC$71;$CS$34:$CS$70;\"=1\")";// 103 SxCMC2 SUCCION -10°C T2
                CalcWorkSheet.Cells[39, 104] = "=SI(CU26=2;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CC$35:$CC$71;$CT$34:$CT$70;\"=1\")";// 104 SxCMC2 SUCCION -10°C T3
                CalcWorkSheet.Cells[39, 105] = "=SI(CU26=2;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CD$35:$CD$71;$CR$34:$CR$70;\"=1\")";// 105 SxCMC2 SUCCION -30°C T1
                CalcWorkSheet.Cells[39, 106] = "=SI(CU26=2;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CD$35:$CD$71;$CS$34:$CS$70;\"=1\")";// 106 SxCMC2 SUCCION -30°C T2
                CalcWorkSheet.Cells[39, 107] = "=SI(CU26=2;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CD$35:$CD$71;$CT$34:$CT$70;\"=1\")";// 107 SxCMC2 SUCCION -30°C T3

                CalcWorkSheet.Cells[44,  99] = "=SI(CU26=3;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CE$35:$CE$71;$CR$34:$CR$70;\"=1\")";//  99 SxCMC3 SUCCION  +5°C T1
                CalcWorkSheet.Cells[44, 100] = "=SI(CU26=3;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CE$35:$CE$71;$CS$34:$CS$70;\"=1\")";// 100 SxCMC3 SUCCION  +5°C T2
                CalcWorkSheet.Cells[44, 101] = "=SI(CU26=3;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CE$35:$CE$71;$CT$34:$CT$70;\"=1\")";// 101 SxCMC3 SUCCION  +5°C T3
                CalcWorkSheet.Cells[44, 102] = "=SI(CU26=3;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CF$35:$CF$71;$CR$34:$CR$70;\"=1\")";// 102 SxCMC3 SUCCION -10°C T1
                CalcWorkSheet.Cells[44, 103] = "=SI(CU26=3;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CF$35:$CF$71;$CS$34:$CS$70;\"=1\")";// 103 SxCMC3 SUCCION -10°C T2
                CalcWorkSheet.Cells[44, 104] = "=SI(CU26=3;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CF$35:$CF$71;$CT$34:$CT$70;\"=1\")";// 104 SxCMC3 SUCCION -10°C T3
                CalcWorkSheet.Cells[44, 105] = "=SI(CU26=3;1)*SI(CU27=1;1)*SUMAR.SI.CONJUNTO($CG$35:$CG$71;$CR$34:$CR$70;\"=1\")";// 105 SxCMC3 SUCCION -30°C T1
                CalcWorkSheet.Cells[44, 106] = "=SI(CU26=3;1)*SI(CU28=1;1)*SUMAR.SI.CONJUNTO($CG$35:$CG$71;$CS$34:$CS$70;\"=2\")";// 106 SxCMC3 SUCCION -30°C T2
                CalcWorkSheet.Cells[44, 107] = "=SI(CU26=3;1)*SI(CU29=1;1)*SUMAR.SI.CONJUNTO($CG$35:$CG$71;$CT$34:$CT$70;\"=3\")";// 107 SxCMC3 SUCCION -30°C T3


                if(RBun3.Checked == true)
                {
                    CalcWorkSheet.Cells[151, 18] = "=SI.ERROR(AC19;\"\")";
                    CalcWorkSheet.Cells[152, 18] = "=SI.ERROR(AC20;\"\")";
                }
                else
                {
                    CalcWorkSheet.Cells[151, 18] = "";
                    CalcWorkSheet.Cells[152, 18] = "";
                }
                CalcWorkSheet.Cells[16, 28] = RBun4.Checked.ToString();
                CalcWorkSheet.Cells[16, 31] = RBun5.Checked.ToString();
                CalcWorkSheet.Cells[16, 32] = RBun6.Checked.ToString(); // EQUIPADO COMPRESOR
                CalcWorkSheet.Cells[16, 33] = RBun7.Checked.ToString(); // EQUIPADO EVAPORADOR
                CalcWorkSheet.Cells[16, 34] = RBun8.Checked.ToString(); // EQUIPAMIENTO BITZER/COPELAND

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[2];
                CalcWorkSheet.Cells[5, 34] = "=SI(N5=1;\"\";SI(IZQUIERDA(AI5;5)=\"Error\";AI5;SI(CA27>CA28;CA27;CA28)))";
                CalcWorkSheet.Cells[27, 79] = "=SI(Y(P5<-5;BU27>=14);CA25/BU27;SI(Y(P5>-4,9;BU27>=12);CA25/BU27;CC25))";
                CalcWorkSheet.Cells[28, 79] = "=SI(P5<-5;(SUMA(CA7:CA16)/14)*BY25;(SUMA(CA7:CA16)/12)*BY25)";
                CalcWorkSheet.Cells[1, 7] = "Software ProJDC v.6.4";
                CalcWorkSheet.Cells[5, 22]="=SI(N5=1;\"\";BUSCARV(VALOR(M5);Coef_den_carga;2;1)*BUSCARV(N5;Base;17;0))";
                CalcWorkSheet.Cells[5, 27] = "=SI(T5<-2;0;REDONDEAR(BUSCARV(N5;Base;16;0)/0,86;-1))";
                CalcWorkSheet.Cells[5, 28] = "=REDONDEAR(BUSCARV(N5;Base;15;0)/0,86;-1)";
               

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[7];
                CalcWorkSheet.Cells[833, 5] = "=SI(Datos!AB31=\"EX2-M00\";SI.ERROR(SI(C22<=12;H841;0);0);0)";
                CalcWorkSheet.Cells[834, 5] = "=SI(Datos!AB31=\"EX2-M00\";SI.ERROR(SI(C22>12;H842;0);0);0)";

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                CalcWorkSheet.Cells[23, 16] = CCamara.Text;
                CalcWorkSheet.Cells[14, 13] = TLargo.Text;
                CalcWorkSheet.Cells[15, 13] = TAncho.Text;
                CalcWorkSheet.Cells[16, 13] = TAlto.Text;
                CalcWorkSheet.Cells[18, 13] = TTem.Text;
                CalcWorkSheet.Cells[13, 13] = pk.ToString();
                CalcWorkSheet.Cells[7, 13] = CKPanel.Checked.ToString();
                CalcWorkSheet.Cells[8, 13] = CKexpo.Checked.ToString();
                CalcWorkSheet.Cells[31, 18] = Ctpd.Text;
                CalcWorkSheet.Cells[27, 13] = TCantEv.Text;
                CalcWorkSheet.Cells[31, 26] = TCentx.Text;
                CalcWorkSheet.Cells[31, 27] = TCdin.Text;
                CalcWorkSheet.Cells[31, 25] = TCxp.Text;
                CalcWorkSheet.Cells[31, 20] = TMevp.Text;
                CalcWorkSheet.Cells[26, 13] = CTamb.Text;
                CalcWorkSheet.Cells[31, 28] = Ctxv.Text;
                CalcWorkSheet.Cells[2, 18] = TDECF.Text;
                CalcWorkSheet.Cells[3, 18] = TDEC.Text;
                CalcWorkSheet.Cells[4, 18] = TDECE.Text;
                CalcWorkSheet.Cells[5, 18] = TDECH.Text;
                CalcWorkSheet.Cells[6, 18] = TDmc1.Text;
                CalcWorkSheet.Cells[6, 19] = TDmc11.Text;
                CalcWorkSheet.Cells[6, 20] = TDmc12.Text;
                CalcWorkSheet.Cells[161, 18] = "=SI.ERROR(SI(TUBERIA!O5=0;\"\";TUBERIA!O5);\"\")";
                CalcWorkSheet.Cells[162, 18] = "=SI.ERROR(SI(TUBERIA!O6=0;\"\";TUBERIA!O6);\"\")";
                CalcWorkSheet.Cells[163, 18] = "=SI.ERROR(SI(TUBERIA!O9=0;\"\";TUBERIA!O9);\"\")";
                CalcWorkSheet.Cells[164, 18] = "=SI.ERROR(SI(TUBERIA!O10=0;\"\";TUBERIA!O10);\"\")";
                CalcWorkSheet.Cells[165, 18] = "=SI.ERROR(SI(TUBERIA!O12=0;\"\";TUBERIA!O12);\"\")";
                CalcWorkSheet.Cells[166, 18] = "=SI.ERROR(SI(TUBERIA!O13=0;\"\";TUBERIA!O13);\"\")";
                CalcWorkSheet.Cells[167, 18] = "=SI.ERROR(SI(TUBERIA!O16=0;\"\";TUBERIA!O16);\"\")";
                CalcWorkSheet.Cells[168, 18] = "=SI.ERROR(SI(TUBERIA!O17=0;\"\";TUBERIA!O17);\"\")";
                CalcWorkSheet.Cells[169, 18] = "=SI.ERROR(SI(TUBERIA!O19=0;\"\";TUBERIA!O19);\"\")";
                CalcWorkSheet.Cells[170, 18] = "=SI.ERROR(SI(TUBERIA!O20=0;\"\";TUBERIA!O20);\"\")";
                CalcWorkSheet.Cells[171, 18] = "=SI.ERROR(SI(TUBERIA!O23=0;\"\";TUBERIA!O23);\"\")";
                CalcWorkSheet.Cells[172, 18] = "=SI.ERROR(SI(TUBERIA!O24=0;\"\";TUBERIA!O24);\"\")";
                CalcWorkSheet.Cells[6, 21] = TDlq1.Text;
                CalcWorkSheet.Cells[6, 22] = TDlq2.Text;
                CalcWorkSheet.Cells[6, 23] = TDlq3.Text;
                CalcWorkSheet.Cells[6, 24] = TDsu130.Text;
                CalcWorkSheet.Cells[6, 25] = TDsu230.Text;
                CalcWorkSheet.Cells[6, 26] = TDsu330.Text;
                CalcWorkSheet.Cells[6, 27] = TDsu110.Text;
                CalcWorkSheet.Cells[6, 28] = TDsu210.Text;
                CalcWorkSheet.Cells[6, 29] = TDsu310.Text;
                CalcWorkSheet.Cells[6, 30] = TDsu105.Text;
                CalcWorkSheet.Cells[6, 31] = TDsu205.Text;
                CalcWorkSheet.Cells[6, 32] = TDsu305.Text;
                CalcWorkSheet.Cells[2, 16] = TDECF.Text;
                CalcWorkSheet.Cells[3, 16] = CTEvap.Text;
                CalcWorkSheet.Cells[4, 16] = CTCond.Text;
                CalcWorkSheet.Cells[5, 16] = CRefrig.Text;
                CalcWorkSheet.Cells[32, 13] = CDigt.Text;
                CalcWorkSheet.Cells[24, 9] = ("=REDONDEAR.MAS(CONVERTIR(F24;\"Wh\";\"BTU\");0)");
                CalcWorkSheet.Cells[29, 15] = ("=SI.ERROR(U31;1)");
                CalcWorkSheet.Cells[29, 13] = ("=SI(O29<>1;INDICE(W31:W336;U31);SI(O29=1;75))");
                CalcWorkSheet.Cells[30, 13] = ("=SI(O29<>1;INDICE(X31:X336;U31);SI(O29=1;0))");
                CalcWorkSheet.Cells[8, 16] = RB60H.Checked;
                CalcWorkSheet.Cells[11, 16] = CDT.Text;
                CalcWorkSheet.Cells[14, 16] = CVolt.Text;
                CalcWorkSheet.Cells[17, 16] = CKdt.Checked;
                CalcWorkSheet.Cells[21, 16] = TTP.Text;
                CalcWorkSheet.Cells[22, 16] = CSumi.Text;
                CalcWorkSheet.Cells[2, 23] = CBint.Text;
                CalcWorkSheet.Cells[3, 23] = TCC.Text;
                CalcWorkSheet.Cells[14, 17] = CFase.Text;
                CalcWorkSheet.Cells[28, 15] = "=PANELES!P6";
                CalcWorkSheet.Cells[28, 16] = "=PANELES!Q6";
                CalcWorkSheet.Cells[28, 17] = "=PANELES!R6";
                CalcWorkSheet.Cells[114, 18] = "=SI.ERROR(SI(PSTP!D38=0;\"\";PSTP!D38);REDONDEAR(P6/0,48;2))";
                CalcWorkSheet.Cells[115, 18] = "=SI.ERROR(SI(PSTP!D39=0;\"\";PSTP!D39);REDONDEAR(P6/7,0;2))";
                CalcWorkSheet.Cells[147, 18] = "=SI.ERROR(SI(SCMC!AQ26=0;\"\";SCMC!AQ26);0)";
                CalcWorkSheet.Cells[34, 13] = "=SUMAR.SI.CONJUNTO(E34:E70;H34:H70;\">=-5\";D34:D70;N34)";
                CalcWorkSheet.Cells[35, 13] = "=SUMAR.SI.CONJUNTO(E34:E70;H34:H70;\"<-5\";D34:D70;N34)";

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[30];// PAGINA COND-INT
                CalcWorkSheet.Cells[121, 8] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"NT1\";J38:J107;0));\"\")";
                CalcWorkSheet.Cells[122, 8] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"NT2\";J38:J107;0));\"\")";
                CalcWorkSheet.Cells[123, 8] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"NT3\";J38:J107;0));\"\")";
                CalcWorkSheet.Cells[121, 9] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"MT1\";K38:K107;0));\"\")";
                CalcWorkSheet.Cells[122, 9] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"MT2\";K38:K107;0));\"\")";
                CalcWorkSheet.Cells[123, 9] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"MT3\";K38:K107;0));\"\")";
                CalcWorkSheet.Cells[121, 10] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"BT1\";L38:L107;0));\"\")";
                CalcWorkSheet.Cells[122, 10] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"BT2\";L38:L107;0));\"\")";
                CalcWorkSheet.Cells[123, 10] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"BT3\";L38:L107;0));\"\")";
                CalcWorkSheet.Cells[122, 11] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"MT4\";K38:K107;0));\"\")";
                CalcWorkSheet.Cells[123, 11] = "=SI.ERROR(INDICE(B38:B107;COINCIDIR(\"BT4\";L38:L107;0));\"\")";
                CalcWorkSheet.Cells[124, 8] = "=SUMA(J38:J107)";
                CalcWorkSheet.Cells[124, 9] = "=SUMA(K38:K107)";
                CalcWorkSheet.Cells[124, 10] = "=SUMA(L38:L107)*SI(BUSCAR(\"BT4\";L38:L107)=\"BT4\";0;1)";
                CalcWorkSheet.Cells[124, 11] = "=SUMA(L38:L107)*SI(BUSCAR(\"BT4\";L38:L107)=\"BT4\";1;0)";
                //CalcWorkSheet.Cells[115, 6] = "=SI.ERROR(INDICE($H$38:$H$101;COINCIDIR(G115;$B$38:$B$101;0));0)";
                //CalcWorkSheet.Cells[116, 6] = "=SI.ERROR(INDICE($H$38:$H$101;COINCIDIR(G116;$B$38:$B$101;0));0)";
                //CalcWorkSheet.Cells[117, 6] = "=SI.ERROR(INDICE($H$38:$H$101;COINCIDIR(G117;$B$38:$B$101;0));0)";
                //CalcWorkSheet.Cells[124, 8] = "=SI.ERROR(INDICE($H$38:$H$107;COINCIDIR(SI(H121<>\"\";H121;SI(H122<>\"\";H122;SI(H123<>\"\";H123)));$B$38:$B$107;0));0)";
                //CalcWorkSheet.Cells[124, 9] = "=SI.ERROR(INDICE($H$38:$H$107;COINCIDIR(SI(I121<>\"\";I121;SI(I122<>\"\";I122;SI(I123<>\"\";I123)));$B$38:$B$107;0));0)";
                //CalcWorkSheet.Cells[124, 10] = "=SI.ERROR(INDICE($H$38:$H$107;COINCIDIR(SI(J121<>\"\";J121;SI(J122<>\"\";J122;SI(J123<>\"\";J123)));$B$38:$B$107;0));0)";
                //CalcWorkSheet.Cells[124, 11] = "=SI.ERROR(INDICE($H$38:$H$107;COINCIDIR(SI(K121<>\"\";K121;SI(K122<>\"\";K122;SI(K123<>\"\";K123)));$B$38:$B$107;0));0)";
                CalcWorkSheet.Cells[127, 11] = "=SI.ERROR(SI(K124=0;\"\";CONCATENAR(\" CAPACIDAD EN BAJA TEMPERATURA \";C115;\" Kw, CONFORMADO POR \";K124;\" COMPRESORES SEMIHERMETICOS MODELO \";K125));\"\")";
                CalcWorkSheet.Cells[127, 10] = "=SI.ERROR(SI(J124=0;\"\";CONCATENAR(\" CAPACIDAD EN BAJA TEMPERATURA \";C115;\" Kw, CONFORMADO POR \";J124;\" COMPRESORES SEMIHERMETICOS MODELO \";J125));\"\")";
                CalcWorkSheet.Cells[127, 9] = "=SI.ERROR(SI(I124=0;\"\";CONCATENAR(\" CAPACIDAD EN MEDIA TEMPERATURA DE \";C113;\" Kw CONFORMADO POR \";I124;\" COMPRESORES SEMIHERMETICOS MODELO \";I125));\"\")";
                CalcWorkSheet.Cells[127, 8] = "=SI.ERROR(SI(H124=0;\"\";CONCATENAR(\" CAPACIDAD EN ALTA TEMPERATURA DE \";C111;\" Kw CONFORMADO POR \";H124;\" COMPRESORES SEMIHERMETICOS MODELO \";H125));\"\")";

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                CalcWorkSheet.Cells[30, 3] = CKpps.Checked.ToString();

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                Range NewRangeCalc = CalcWorkSheet.get_Range("M28", "Q28");
                Array myNewArr = (Array)NewRangeCalc.Value2;
                String[,] myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                for (int i = 0; i < myNewArr.GetLength(0); i++)
                {
                    for (int j = 0; j < myNewArr.GetLength(1); j++)
                    {
                        long[] indices = new long[] { i + 1, j + 1 };
                        myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                    }
                }
                //CalcWorkBook.SaveCopyAs(@"C:\prueba.xlsx");
                TCF.Text = myNewArrCalc.GetValue(1, 1).ToString();
                TFW.Text = myNewArrCalc.GetValue(1, 2).ToString();
                TCmod.Text = myNewArrCalc.GetValue(1, 3).ToString();
                TCmodd.Text = myNewArrCalc.GetValue(1, 4).ToString();
                TCmodp.Text = myNewArrCalc.GetValue(1, 5).ToString();
                //TInc.Text = myNewArrCalc.GetValue(16, 1).ToString();
                //TCint1.Text = myNewArrCalc.GetValue(17, 1).ToString();
                //TCint2.Text = myNewArrCalc.GetValue(18, 1).ToString();
                //TCint3.Text = myNewArrCalc.GetValue(19, 1).ToString();
                if (newCam.GetKmod() == true)
                {
                    CalcWorkSheet.Cells[218, 16] = CKmod.ToString();
                    TCmod.Text = myNewArrCalc.GetValue(1, 3).ToString();
                    TCmodd.Text = myNewArrCalc.GetValue(1, 4).ToString();
                    TCmodp.Text = myNewArrCalc.GetValue(1, 5).ToString();
                }
                if (newCam.GetKmod() == false)
                {
                    TCmod.Clear();
                    TCmod.Text = "";
                    TCmodd.Clear();
                    TCmodd.Text = "";
                    TCmodp.Clear();
                    TCmodp.Text = "";
                }
                
                CalcWorkSheet.Cells[6, 16] = TFW.Text;
                CalcWorkSheet.Cells[7, 16] = TConD.Text;
                CalcWorkSheet.Cells[12, 16] = Cmodex.Text;
                CalcWorkSheet.Cells[13, 16] = Cevap.Text;

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[25];
                CalcWorkSheet.Cells[8, 7] = "=REDOND.MULT(REDONDEAR.MAS((G6*2);0);9)";
                CalcWorkSheet.Cells[8, 8] = "=REDOND.MULT(REDONDEAR.MAS((G6*2);0);16)";
                
                // ******************************************************************************************* 25/11/2018
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[27];
                CalcWorkSheet.Cells[6, 4] = "=SI.ERROR(INDICE(A106:A120;COINCIDIR(CONCATENAR(H9;E7);B106:B120;0));INDICE(A106:A120;COINCIDIR(CONCATENAR(H9;E8);B106:B120;0)))";
                /*
                NewRangeCalc = CalcWorkSheet.get_Range("D2", "D12");
                myNewArr = (Array)NewRangeCalc.Value2;
                
                SPtp149.Text = myNewArrCalc.GetValue(1, 5).ToString();
                CalcWorkSheet.Cells[6, 4] = SPtp149.Text;
                // ******************************************************************************************* 25/11/2018
                */
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[28];
                CalcWorkSheet.Cells[4, 21] = pk.ToString();
                CalcWorkSheet.Cells[9, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q9;$C$5:$C$68;1));0)";
                CalcWorkSheet.Cells[10, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q10;$C$5:$C$68;1));0)";
                CalcWorkSheet.Cells[11, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q11;$C$5:$C$68;1));0)";
                CalcWorkSheet.Cells[12, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q12;$C$5:$C$68;1));0)";
                CalcWorkSheet.Cells[13, 19] = "=SI.ERROR(SI(W9<Q9;INDICE($H$5:$H$68;COINCIDIR(Q9-INDICE($C$5:$C$68;COINCIDIR(Q9;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                CalcWorkSheet.Cells[14, 19] = "=SI.ERROR(SI(W10<Q10;INDICE($H$5:$H$68;COINCIDIR(Q10-INDICE($C$5:$C$68;COINCIDIR(Q10;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                CalcWorkSheet.Cells[15, 19] = "=SI.ERROR(SI(W11<Q11;INDICE($H$5:$H$68;COINCIDIR(Q11-INDICE($C$5:$C$68;COINCIDIR(Q11;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                CalcWorkSheet.Cells[16, 19] = "=SI.ERROR(SI(W12<Q12;INDICE($H$5:$H$68;COINCIDIR(Q12-INDICE($C$5:$C$68;COINCIDIR(Q12;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                // LEER VALORES CALCULADOS.
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                NewRangeCalc = CalcWorkSheet.get_Range("R9", "R300");
                myNewArr = (Array)NewRangeCalc.Value2;
                myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                for (int i = 0; i < myNewArr.GetLength(0); i++)
                {
                    for (int j = 0; j < myNewArr.GetLength(1); j++)
                    {
                        long[] indices = new long[] { i + 1, j + 1 };
                        myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                    }
                }
                TQevp.Text = myNewArrCalc.GetValue(1, 1).ToString();
                TTint.Text = myNewArrCalc.GetValue(20, 1).ToString();
                TEquip.Text = myNewArrCalc.GetValue(2, 1).ToString();
                TValv.Text = myNewArrCalc.GetValue(7, 1).ToString();
                TCodValv.Text = myNewArrCalc.GetValue(8, 1).ToString();
                TTmos.Text = myNewArrCalc.GetValue(25, 1).ToString();
                TInc.Text = myNewArrCalc.GetValue(16, 1).ToString();
                TInev.Text = myNewArrCalc.GetValue(15, 1).ToString();
                TIned.Text = myNewArrCalc.GetValue(21, 1).ToString();
                TIncd.Text = myNewArrCalc.GetValue(132, 1).ToString();
                TIpv.Text = myNewArrCalc.GetValue(3, 1).ToString();
                TIcc.Text = myNewArrCalc.GetValue(4, 1).ToString();
                TCint1.Text = myNewArrCalc.GetValue(17, 1).ToString();
                TCint2.Text = myNewArrCalc.GetValue(18, 1).ToString();
                TCint3.Text = myNewArrCalc.GetValue(19, 1).ToString();
                TCmce.Text = myNewArrCalc.GetValue(5, 1).ToString();
                TPmce.Text = myNewArrCalc.GetValue(6, 1).ToString();
                TDmce.Text = myNewArrCalc.GetValue(24, 1).ToString();
                TDmc4.Text = myNewArrCalc.GetValue(26, 1).ToString();
                TDmc5.Text = myNewArrCalc.GetValue(27, 1).ToString();
                TDmc6.Text = myNewArrCalc.GetValue(28, 1).ToString();
                TDmc3.Text = myNewArrCalc.GetValue(29, 1).ToString();
                TDmc7.Text = myNewArrCalc.GetValue(30, 1).ToString();
                TDmc8.Text = myNewArrCalc.GetValue(31, 1).ToString();
                TDmc2.Text = myNewArrCalc.GetValue(32, 1).ToString();
                TLcc.Text = myNewArrCalc.GetValue(33, 1).ToString();
                TLss.Text = myNewArrCalc.GetValue(34, 1).ToString();
                TPcmc.Text = myNewArrCalc.GetValue(35, 1).ToString();
                TPcond.Text = myNewArrCalc.GetValue(36, 1).ToString();
                TPlq.Text = myNewArrCalc.GetValue(37, 1).ToString();
                TPvq.Text = myNewArrCalc.GetValue(38, 1).ToString();
                TPls.Text = myNewArrCalc.GetValue(39, 1).ToString();
                TPosc.Text = myNewArrCalc.GetValue(40, 1).ToString();
                TPsq.Text = myNewArrCalc.GetValue(41, 1).ToString();
                TPcq.Text = myNewArrCalc.GetValue(42, 1).ToString();
                TPcv.Text = myNewArrCalc.GetValue(43, 1).ToString();
                TPcx.Text = myNewArrCalc.GetValue(44, 1).ToString();
                TPcy.Text = myNewArrCalc.GetValue(45, 1).ToString();
                TPcp.Text = myNewArrCalc.GetValue(46, 1).ToString();
                TPex.Text = myNewArrCalc.GetValue(47, 1).ToString();
                TPrs.Text = myNewArrCalc.GetValue(48, 1).ToString();
                TPem.Text = myNewArrCalc.GetValue(49, 1).ToString();
                TPnt.Text = myNewArrCalc.GetValue(50, 1).ToString();
                TPml.Text = myNewArrCalc.GetValue(51, 1).ToString();
                TPdtr.Text = myNewArrCalc.GetValue(52, 1).ToString();
                TPdtc.Text = myNewArrCalc.GetValue(53, 1).ToString();
                TPdt2.Text = myNewArrCalc.GetValue(54, 1).ToString();
                TPdt1.Text = myNewArrCalc.GetValue(55, 1).ToString();
                TCp80.Text = myNewArrCalc.GetValue(56, 1).ToString();// Cantidad de paneles de pared 80mm
                TCp100.Text = myNewArrCalc.GetValue(57, 1).ToString();// Cantidad de paneles de pared 100mm
                TCp120.Text = myNewArrCalc.GetValue(58, 1).ToString();// Cantidad de paneles de pared 120mm
                TCp150.Text = myNewArrCalc.GetValue(109, 1).ToString();// Cantidad de paneles de pared 150mm
                TCt80.Text = myNewArrCalc.GetValue(59, 1).ToString();// Cantidad de paneles de techo 80mm
                TCt100.Text = myNewArrCalc.GetValue(60, 1).ToString();// Cantidad de paneles de techo 100mm
                TCt120.Text = myNewArrCalc.GetValue(61, 1).ToString();// Cantidad de paneles de techo 120mm
                TCt150.Text = myNewArrCalc.GetValue(62, 1).ToString();// Cantidad de paneles de techo 150mm
                TCp80m.Text = myNewArrCalc.GetValue(63, 1).ToString();// Largo de paneles de pared 80mm
                TCp100m.Text = myNewArrCalc.GetValue(64, 1).ToString();// Largo de paneles de pared 100mm
                TCp120m.Text = myNewArrCalc.GetValue(65, 1).ToString();// Largo de paneles de pared 120mm
                TCp150m.Text = myNewArrCalc.GetValue(110, 1).ToString();// Largo de paneles de pared 150mm
                TCt80m.Text = myNewArrCalc.GetValue(66, 1).ToString();// Largo de paneles de techo 80mm
                TCt100m.Text = myNewArrCalc.GetValue(67, 1).ToString();// Largo de paneles de techo 100mm
                TCt120m.Text = myNewArrCalc.GetValue(68, 1).ToString();// Largo de paneles de techo 120mm
                TCt150m.Text = myNewArrCalc.GetValue(69, 1).ToString();// Largo de paneles de techo 150mm
                SPtp84.Text = myNewArrCalc.GetValue(70, 1).ToString();
                SPtp83.Text = myNewArrCalc.GetValue(71, 1).ToString();
                SPtp78.Text = myNewArrCalc.GetValue(72, 1).ToString();
                SPtp76.Text = myNewArrCalc.GetValue(73, 1).ToString();
                SPtp75.Text = myNewArrCalc.GetValue(74, 1).ToString();
                SPtp74.Text = myNewArrCalc.GetValue(75, 1).ToString();
                SPtp73.Text = myNewArrCalc.GetValue(76, 1).ToString();
                SPtp72.Text = myNewArrCalc.GetValue(77, 1).ToString();
                SPtp85.Text = myNewArrCalc.GetValue(79, 1).ToString();
                SPtp86.Text = myNewArrCalc.GetValue(80, 1).ToString();
                SPtp87.Text = myNewArrCalc.GetValue(81, 1).ToString();
                SPtp88.Text = myNewArrCalc.GetValue(82, 1).ToString();
                SPtp89.Text = myNewArrCalc.GetValue(83, 1).ToString();
                SPtp90.Text = myNewArrCalc.GetValue(84, 1).ToString();
                SPtp91.Text = myNewArrCalc.GetValue(85, 1).ToString();
                SPtp92.Text = myNewArrCalc.GetValue(86, 1).ToString();
                SPtp93.Text = myNewArrCalc.GetValue(87, 1).ToString();
                SPtp94.Text = myNewArrCalc.GetValue(88, 1).ToString();
                SPtp95.Text = myNewArrCalc.GetValue(89, 1).ToString();
                SPtp96.Text = myNewArrCalc.GetValue(90, 1).ToString();
                SPtp97.Text = myNewArrCalc.GetValue(91, 1).ToString();
                SPtp98.Text = myNewArrCalc.GetValue(92, 1).ToString();
                SPtp99.Text = myNewArrCalc.GetValue(93, 1).ToString();
                SPtp100.Text = myNewArrCalc.GetValue(94, 1).ToString();
                SPtp101.Text = myNewArrCalc.GetValue(95, 1).ToString();
                SPtp102.Text = myNewArrCalc.GetValue(96, 1).ToString();
                SPtp103.Text = myNewArrCalc.GetValue(97, 1).ToString();
                SPtp104.Text = myNewArrCalc.GetValue(98, 1).ToString();
                SPtp105.Text = myNewArrCalc.GetValue(99, 1).ToString();
                SPtp106.Text = myNewArrCalc.GetValue(100, 1).ToString();
                SPtp107.Text = myNewArrCalc.GetValue(101, 1).ToString();
                SPtp108.Text = myNewArrCalc.GetValue(102, 1).ToString();
                SPtp109.Text = myNewArrCalc.GetValue(103, 1).ToString();
                SPtp110.Text = myNewArrCalc.GetValue(104, 1).ToString();
                SPtp111.Text = myNewArrCalc.GetValue(105, 1).ToString();
                SPtp136.Text = myNewArrCalc.GetValue(106, 1).ToString();
                SPtp137.Text = myNewArrCalc.GetValue(107, 1).ToString();
                SPtp141.Text = myNewArrCalc.GetValue(133, 1).ToString();
                SPtp142.Text = myNewArrCalc.GetValue(134, 1).ToString();
                SPtp143.Text = myNewArrCalc.GetValue(135, 1).ToString();
                SPtp144.Text = myNewArrCalc.GetValue(136, 1).ToString();
                SPtp145.Text = myNewArrCalc.GetValue(137, 1).ToString();
                SPtp146.Text = myNewArrCalc.GetValue(138, 1).ToString();
                SPtp147.Text = myNewArrCalc.GetValue(139, 1).ToString();
                SPtp148.Text = myNewArrCalc.GetValue(140, 1).ToString();
                SPtp150.Text = myNewArrCalc.GetValue(143, 1).ToString();
                SPtp151.Text = myNewArrCalc.GetValue(144, 1).ToString();
                TPcmc1.Text = myNewArrCalc.GetValue(145, 1).ToString();
                TPcmc2.Text = myNewArrCalc.GetValue(146, 1).ToString();
                TPcmc3.Text = myNewArrCalc.GetValue(147, 1).ToString();
                TQevpd.Text = myNewArrCalc.GetValue(151, 1).ToString();
                TQevpc.Text = myNewArrCalc.GetValue(152, 1).ToString();
                TCsist.Text = myNewArrCalc.GetValue(108, 1).ToString();
                TQfw.Text = myNewArrCalc.GetValue(11, 1).ToString();

                TDlq11.Text = myNewArrCalc.GetValue(153, 1).ToString();//Liquido -30°C T1:
                TDsu130.Text = myNewArrCalc.GetValue(154, 1).ToString();//Succion -30°C T1:
                TDsu105.Text = myNewArrCalc.GetValue(155, 1).ToString();//Succiön +5°C T1:
                TDsu110.Text = myNewArrCalc.GetValue(156, 1).ToString();//Succión -10°C T1:

                TDlq21.Text = myNewArrCalc.GetValue(157, 1).ToString();//Liquido -30°C T2:
                TDsu230.Text = myNewArrCalc.GetValue(158, 1).ToString();//Succion -30°C T2:
                TDsu205.Text = myNewArrCalc.GetValue(159, 1).ToString();//Succiön +5°C T2:
                TDsu210.Text = myNewArrCalc.GetValue(160, 1).ToString();//Succión -10°C T2:

                TDlq31.Text = myNewArrCalc.GetValue(161, 1).ToString();//Liquido -30°C T3:
                TDsu330.Text = myNewArrCalc.GetValue(162, 1).ToString();//Succion -30°C T3:
                TDsu305.Text = myNewArrCalc.GetValue(163, 1).ToString();//Succiön +5°C T3:
                TDsu310.Text = myNewArrCalc.GetValue(164, 1).ToString();//Succión -10°C T3:

                TIn1.Text = myNewArrCalc.GetValue(165, 1).ToString();// Abrasaderas sifonicas LIQ LIN- CTRAL.
                TIn2.Text = myNewArrCalc.GetValue(167, 1).ToString();//Abrasaderas sifonicas SUCC LIN- CTRAL
                TIn3.Text = myNewArrCalc.GetValue(168, 1).ToString();//Abrasaderas sifonicas SUCC 1er EVAP.
                TIn4.Text = myNewArrCalc.GetValue(169, 1).ToString();//Abrasaderas sifonicas SUCC 2do EVAP
                TIn5.Text = myNewArrCalc.GetValue(170, 1).ToString();//Abrasaderas sifonicas SUCC 3er EVAP.
                TIn6.Text = myNewArrCalc.GetValue(171, 1).ToString();//Abrasaderas sifonicas SUCC LIN T1(-30°C).
                TIn7.Text = myNewArrCalc.GetValue(172, 1).ToString();//Abrasaderas sifonicas SUCC LIN T2 (-30°C).
                TIn8.Text = myNewArrCalc.GetValue(173, 1).ToString();// Abrasaderas sifonicas SUCC LIN T3(-30°C)..
                TIn9.Text = myNewArrCalc.GetValue(174, 1).ToString();// Abrasaderas sifonicas SUCC LIN T1(-10°C).
                TIn10.Text = myNewArrCalc.GetValue(175, 1).ToString();// Abrasaderas sifonicas SUCC LIN T2(-10°C).
                TIn11.Text = myNewArrCalc.GetValue(176, 1).ToString();// Abrasaderas sifonicas SUCC LIN T3(-10°C).
                TIn12.Text = myNewArrCalc.GetValue(177, 1).ToString();// Abrasaderas sifonicas SUCC LIN T1( 5°C).
                TIn13.Text = myNewArrCalc.GetValue(178, 1).ToString();// Abrasaderas sifonicas SUCC LIN T2( 5°C).
                TIn14.Text = myNewArrCalc.GetValue(179, 1).ToString();// Abrasaderas sifonicas SUCC LIN T3( 5°C).
                TIn15.Text = myNewArrCalc.GetValue(193, 1).ToString();// BARILLA ROSCADA M10-M8
                TIn16.Text = myNewArrCalc.GetValue(194, 1).ToString();// TUERCAS M10-M8
                TIn17.Text = myNewArrCalc.GetValue(195, 1).ToString();// EXPANCIONES M10-M8
                TIn18.Text = myNewArrCalc.GetValue(196, 1).ToString();// ARANDELAS CUAD M10-M8
                TIn19.Text = myNewArrCalc.GetValue(197, 1).ToString();// ARANDELAS M10-M8
                TIn20.Text = myNewArrCalc.GetValue(198, 1).ToString();// PERFIL DE CARGA
                TIn21.Text = myNewArrCalc.GetValue(205, 1).ToString();// BANDEJA DE 300MM
                TIn22.Text = myNewArrCalc.GetValue(206, 1).ToString();// TAPA BANDEJA 300MM
                TIn23.Text = myNewArrCalc.GetValue(208, 1).ToString();// Abrasadera LIQ LIN EVP
                
                TIp1.Text = myNewArrCalc.GetValue(166, 1).ToString();// CANTIDAD Abrasaderas sifonicas LIQ LIN- CTRAL.
                TIp2.Text = myNewArrCalc.GetValue(180, 1).ToString();//Cantidad de abrasaderas sifonicas SUCC LIN- CTRAL.
                TIp3.Text = myNewArrCalc.GetValue(181, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC 1er EVAP.
                TIp4.Text = myNewArrCalc.GetValue(182, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC 2do EVAP
                TIp5.Text = myNewArrCalc.GetValue(183, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC 3er EVAP.
                TIp6.Text = myNewArrCalc.GetValue(184, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T1(-30°C).
                TIp7.Text = myNewArrCalc.GetValue(185, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T2 (-30°C).
                TIp8.Text = myNewArrCalc.GetValue(186, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T3(-30°C).
                TIp9.Text = myNewArrCalc.GetValue(187, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T1(-10°C).
                TIp10.Text = myNewArrCalc.GetValue(188, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T3(-10°C).
                TIp11.Text = myNewArrCalc.GetValue(189, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T3(-10°C).
                TIp12.Text = myNewArrCalc.GetValue(190, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T1( 5°C).
                TIp13.Text = myNewArrCalc.GetValue(191, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T2( 5°C).
                TIp14.Text = myNewArrCalc.GetValue(192, 1).ToString();// CANTIDAD Abrasaderas sifonicas SUCC LIN T3( 5°C).
                TIp15.Text = myNewArrCalc.GetValue(199, 1).ToString();// CANTIDAD BARILLA ROSCADA M10-M8
                TIp16.Text = myNewArrCalc.GetValue(200, 1).ToString();// CANTIDAD TUERCAS M10-M8
                TIp17.Text = myNewArrCalc.GetValue(201, 1).ToString();// CANTIDAD EXPANCIONES M10-M8
                TIp18.Text = myNewArrCalc.GetValue(202, 1).ToString();// CANTIDAD ARANDELAS CUAD M10-M8
                TIp19.Text = myNewArrCalc.GetValue(203, 1).ToString();// CANTIDAD ARANDELAS M10-M8
                TIp20.Text = myNewArrCalc.GetValue(204, 1).ToString();// CANTIDAD PERFIL DE CARGA
                TIp21.Text = myNewArrCalc.GetValue(207, 1).ToString();// CANTIDAD BANDEJA DE 300MM + TAPA BANDEJA 300MM
                TIp23.Text = myNewArrCalc.GetValue(209, 1).ToString();// CANTIDAD Abrasadera LIQ LIN EVP

                //TIn23.Text = myNewArrCalc.GetValue(167, 1).ToString();//
                //TIn24.Text = myNewArrCalc.GetValue(168, 1).ToString();//
                //TIn25.Text = myNewArrCalc.GetValue(169, 1).ToString();//
                //TIn26.Text = myNewArrCalc.GetValue(170, 1).ToString();//
                //TIn27.Text = myNewArrCalc.GetValue(171, 1).ToString();//
                //TIn28.Text = myNewArrCalc.GetValue(172, 1).ToString();//
                //TIn29.Text = myNewArrCalc.GetValue(173, 1).ToString();//
                //TIn30.Text = myNewArrCalc.GetValue(174, 1).ToString();//
                //TIn31.Text = myNewArrCalc.GetValue(175, 1).ToString();//
                //TIn32.Text = myNewArrCalc.GetValue(176, 1).ToString();//
                //TIn33.Text = myNewArrCalc.GetValue(177, 1).ToString();//
                //TIn34.Text = myNewArrCalc.GetValue(178, 1).ToString();//
                //TIn35.Text = myNewArrCalc.GetValue(179, 1).ToString();//
                //TIn36.Text = myNewArrCalc.GetValue(180, 1).ToString();//
                //TIn37.Text = myNewArrCalc.GetValue(181, 1).ToString();//        
                //TIn38.Text = myNewArrCalc.GetValue(182, 1).ToString();//
                //TIn39.Text = myNewArrCalc.GetValue(183, 1).ToString();//
                //TIn40.Text = myNewArrCalc.GetValue(184, 1).ToString();//
                //TIn41.Text = myNewArrCalc.GetValue(165, 1).ToString();//
                //TIn42.Text = myNewArrCalc.GetValue(166, 1).ToString();//
                //TIp22.Text = myNewArrCalc.GetValue(166, 1).ToString();//
                //TIp23.Text = myNewArrCalc.GetValue(167, 1).ToString();//
                //TIp24.Text = myNewArrCalc.GetValue(168, 1).ToString();//
                //TIp25.Text = myNewArrCalc.GetValue(169, 1).ToString();//
                //TIp26.Text = myNewArrCalc.GetValue(170, 1).ToString();//
                //TIp27.Text = myNewArrCalc.GetValue(171, 1).ToString();//
                //TIp28.Text = myNewArrCalc.GetValue(172, 1).ToString();//
                //TIp29.Text = myNewArrCalc.GetValue(173, 1).ToString();//
                //TIp30.Text = myNewArrCalc.GetValue(174, 1).ToString();//
                //TIp31.Text = myNewArrCalc.GetValue(175, 1).ToString();//
                //TIp32.Text = myNewArrCalc.GetValue(176, 1).ToString();//
                //TIp33.Text = myNewArrCalc.GetValue(177, 1).ToString();//
                //TIp34.Text = myNewArrCalc.GetValue(178, 1).ToString();//
                //TIp35.Text = myNewArrCalc.GetValue(179, 1).ToString();//
                //TIp36.Text = myNewArrCalc.GetValue(180, 1).ToString();//
                //TIp37.Text = myNewArrCalc.GetValue(181, 1).ToString();//
                //TIp38.Text = myNewArrCalc.GetValue(182, 1).ToString();//
                //TIp39.Text = myNewArrCalc.GetValue(183, 1).ToString();//
                //TIp40.Text = myNewArrCalc.GetValue(184, 1).ToString();// 
                //TIp41.Text = myNewArrCalc.GetValue(165, 1).ToString();// 
                //TIp42.Text = myNewArrCalc.GetValue(166, 1).ToString();//
                TCmod.Text = myNewArrCalc.GetValue(210, 1).ToString(); // CODIGO CMM
                TCmodd.Text = myNewArrCalc.GetValue(211, 1).ToString();// DESCRIPCION CMM
                TCmodp.Text = myNewArrCalc.GetValue(212, 1).ToString();// PRECIO CMM

                //codValv = myNewArrCalc.GetValue(8, 1).ToString();
                /*
                try
                {

                    myOferta.GetCam(int.Parse(CCamara.Text) - 1).SetCodValv(myNewArrCalc.GetValue(8, 1).ToString());

                }
                catch { }
                */
                TDesc.Text = myNewArrCalc.GetValue(9, 1).ToString();
                TPrec.Text = myNewArrCalc.GetValue(10, 1).ToString();
                TQfep.Text = myNewArrCalc.GetValue(11, 1).ToString();
                TScdro.Text = myNewArrCalc.GetValue(12, 1).ToString();

                //-----------------------------------------------------SIN PROBLEMA "PRINCIPAL"

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[3];               
                NewRangeCalc = CalcWorkSheet.get_Range("T361", "T372");
                myNewArr = (Array)NewRangeCalc.Value2;
                myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                for (int i = 0; i < myNewArr.GetLength(0); i++)
                {
                    for (int j = 0; j < myNewArr.GetLength(1); j++)
                    {
                        long[] indices = new long[] { i + 1, j + 1 };
                        myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                    }
                }
                TEmevp.Text = myNewArrCalc.GetValue(2, 1).ToString();
                TSpsi.Text = myNewArrCalc.GetValue(10, 1).ToString();
                TStemp.Text = myNewArrCalc.GetValue(11, 1).ToString();
                TApsi.Text = myNewArrCalc.GetValue(12, 1).ToString();

                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
                if(RBun.Checked == false)
                {
                    CalcWorkSheet.Cells[12, 26] = "=SI(AQ12=0;AQ13;AQ12)";
                    CalcWorkSheet.Cells[12, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-5\")");
                    CalcWorkSheet.Cells[13, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-5\")");
                    CalcWorkSheet.Cells[108, 69] = "=SI(BE120=9;SI.ERROR(INDICE(BI151:BI165;BI122+1);INDICE(BI151:BI165;BI122));SI(BE120=1;SI.ERROR(INDICE(BA151:BA165;BI122+1);INDICE(BA151:BA165;BI122))))";
                    CalcWorkSheet.Cells[109, 69] = "=SI(BE120=9;INDICE(BI151:BI165;BI122);SI(BE120=1;INDICE(BA151:BA165;BI122)))";
                    CalcWorkSheet.Cells[116, 69] = "=SI.ERROR(INDICE(BQ132:BQ143;BQ121+1);INDICE(BQ132:BQ143;BQ121))";

                    CalcWorkSheet.Cells[12, 27] = "=SI(AO12=0;AO13;AO12)";
                    if (int.Parse(CTEvap.Text) < -5)
                    {
                        CalcWorkSheet.Cells[12, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-5\")*1,04+D37;2)))";// POTENCIA CMC BT //
                        CalcWorkSheet.Cells[13, 41] = "0";
                    }
                    if (int.Parse(CTEvap.Text) >= -5)
                    {
                        CalcWorkSheet.Cells[13, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-5\")*1,02+D37;2)))";// POTENCIA CMC NT //
                        CalcWorkSheet.Cells[12, 41] = "0";
                    }

                    CalcWorkSheet.Cells[12, 28] = "=SI(AP12=0;AP13;AP12)";
                    CalcWorkSheet.Cells[12, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-10\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$10;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$I$23;COINCIDIR(D4;$E$22:$I$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,18);2);0))";
                    CalcWorkSheet.Cells[13, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-10\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$9;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$I$23;COINCIDIR(D4;$E$22:$I$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,1);2);0))";
                
                }
                else
                {
                    CalcWorkSheet.Cells[118, 46] = "1";
                    CalcWorkSheet.Cells[12, 26] = "=SI(AQ12=0;AQ13;AQ12)";
                    CalcWorkSheet.Cells[12, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\")");
                    CalcWorkSheet.Cells[13, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\")");
                    CalcWorkSheet.Cells[108, 69] = "=SI(BE120=9;SI.ERROR(INDICE(BI151:BI165;BI122+1);INDICE(BI151:BI165;BI122));SI(BE120=1;SI.ERROR(INDICE(BA151:BA165;BI122+1);INDICE(BA151:BA165;BI122))))";
                    CalcWorkSheet.Cells[109, 69] = "=SI(BE120=9;INDICE(BI151:BI165;BI122);SI(BE120=1;INDICE(BA151:BA165;BI122)))";
                    CalcWorkSheet.Cells[116, 69] = "=SI.ERROR(INDICE(BQ132:BQ143;BQ121+1);INDICE(BQ132:BQ143;BQ121))";
                    CalcWorkSheet.Cells[8, 42] = RBun.Checked.ToString();

                    CalcWorkSheet.Cells[12, 27] = "=SI(AO12=0;AO13;AO12)";
                    if (int.Parse(CTEvap.Text) < -5)
                    {
                        CalcWorkSheet.Cells[12, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\")*1,11+D37;2)))";
                        CalcWorkSheet.Cells[13, 41] = "0";
                    }
                    if (int.Parse(CTEvap.Text) >= -5)
                    {
                        CalcWorkSheet.Cells[13, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\")*1,11+D37;2)))";
                        CalcWorkSheet.Cells[12, 41] = "0";
                    }

                    CalcWorkSheet.Cells[12, 28] = "=SI(AP12=0;AP13;AP12)";
                    CalcWorkSheet.Cells[12, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$10;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$I$23;COINCIDIR(D4;$E$22:$I$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,24);2);0))";
                    CalcWorkSheet.Cells[13, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$9;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$I$23;COINCIDIR(D4;$E$22:$I$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,34);2);0))";
                
                }
                CalcWorkSheet.Cells[123, 61] = "=SI(BE120=9;SI.ERROR(SI(BI122>=1;COINCIDIR(BQ110;BI151:BI165;0);0);0);SI(BE120=1;SI.ERROR(SI(BI122>=1;COINCIDIR(BQ110;BA151:BA165;0);0);0)))";
                CalcWorkSheet.Cells[123, 63] = "=SI.ERROR(SI(BK122>=1;COINCIDIR(BS110;BK151:BK165;0);0);0)";
                CalcWorkSheet.Cells[123, 69] = "=SI.ERROR(SI(BQ121>=1;COINCIDIR(BQ118;BQ132:BQ143;0);0);0)";
                CalcWorkSheet.Cells[123, 71] = "=SI.ERROR(SI(BS121>=1;COINCIDIR(BS118;BS132:BS143;0);0);0)";
                
                // LECTURA DEL NUMERO DE MULTICOMPRESORA ACTIVA//
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
                NewRangeCalc = CalcWorkSheet.get_Range("Y11", "Y31");
                myNewArr = (Array)NewRangeCalc.Value2;
                myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                for (int i = 0; i < myNewArr.GetLength(0); i++)
                {
                    for (int j = 0; j < myNewArr.GetLength(1); j++)
                    {
                        long[] indices = new long[] { i + 1, j + 1 };
                        myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                    }
                }
                TMcc.Text = myNewArrCalc.GetValue(1, 1).ToString();

                //********************************************************* CONTROL DE LINEA CENTRAL CMC ****
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[24];// PAGINA ACCESS
                CalcWorkSheet.Cells[7, 11] = CBcm.Text;
                if (int.Parse(CCmci.Text) != 0)
                {
                    if (int.Parse(CCmci.Text) == 1)
                    {
                        CalcWorkSheet.Cells[8, 11] = TDmc1.Text;
                        //CalcWorkSheet.Cells[8, 9] = TDmc11.Text;
                        //CalcWorkSheet.Cells[9, 9] = TDmc12.Text;
                       
                    }
                    else
                    { }
                    if (int.Parse(CCmci.Text) == 2)
                    {
                        CalcWorkSheet.Cells[9, 11] = TDmc1.Text;
                       
                    }
                    else
                    { }
                    if (int.Parse(CCmci.Text) == 3)
                    {
                        CalcWorkSheet.Cells[10, 11] = TDmc1.Text;
                        
                    }
                    else
                    { }
                    if (int.Parse(CCmci.Text) == 4)
                    {
                        CalcWorkSheet.Cells[11, 11] = TDmc1.Text;
                       
                    }
                    else
                    { }
                    if (int.Parse(CCmci.Text) == 5)
                    {
                        CalcWorkSheet.Cells[12, 11] = TDmc1.Text;
                        
                    }
                    else
                    { }
                }

                //********************************************************* CONTROL DE VALVULAS SANHUA ****
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[31];// PAGINA ACCESS
                CalcWorkSheet.Cells[13, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN13:AU13;-1));0)";
                CalcWorkSheet.Cells[14, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN14:AU14;-1));0)";
                CalcWorkSheet.Cells[15, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN15:AU15;-1));0)";
                CalcWorkSheet.Cells[16, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN16:AU16;-1));0)";
                CalcWorkSheet.Cells[17, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN17:AU17;-1));0)";
                CalcWorkSheet.Cells[18, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN18:AU18;-1));0)";
                CalcWorkSheet.Cells[19, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN19:AU19;-1));0)";
                //********************************************************* CONTROL DE CMC 2-3 ****
                CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[32];// PAGINA ACCESS
                CalcWorkSheet.Cells[59, 6] = "=SI.ERROR(INDICE(M44:M57;COINCIDIR(F61;F44:F57;-1));INDICE(M17:M26;COINCIDIR(F61;F17:F26;-1)))";
                /*
                NewRangeCalc = CalcWorkSheet.get_Range("K8", "K15");
                myNewArrCalc = (Array)NewRangeCalc.Value2;
                if (int.Parse(CCmci.Text) == 1)
                {
                    CalcWorkSheet.Cells[1, 1] = myNewArrCalc.GetValue(1, 1).ToString();
                }
                if (int.Parse(CCmci.Text) == 2)
                {
                    CalcWorkSheet.Cells[2, 1] = myNewArrCalc.GetValue(2, 1).ToString();
                }
                if (int.Parse(CCmci.Text) == 3)
                {
                    CalcWorkSheet.Cells[3, 1] = myNewArrCalc.GetValue(3, 1).ToString();
                }
                if (int.Parse(CCmci.Text) == 4)
                {
                    CalcWorkSheet.Cells[4, 1] = myNewArrCalc.GetValue(4, 1).ToString();
                }
                if (int.Parse(CCmci.Text) == 5)
                {
                    CalcWorkSheet.Cells[5, 1] = myNewArrCalc.GetValue(5, 1).ToString();
                }
                */
                //***********************************************
            }
            catch
            {
                try
                {
                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                    CalcWorkSheet.Cells[30, 18] = CVolt.Text;
                    CalcWorkSheet.Cells[34, 39] = CKAUni.Checked.ToString();//39
                    CalcWorkSheet.Cells[34, 40] = CKValv.Checked.ToString();//40
                    CalcWorkSheet.Cells[34, 41] = CKMonile.Checked.ToString();//41
                    CalcWorkSheet.Cells[34, 42] = CKUnid.Checked.ToString();//42
                    CalcWorkSheet.Cells[34, 43] = CKTcobre.Checked.ToString();//43
                    CalcWorkSheet.Cells[34, 44] = CKCelect.Checked.ToString();//44
                    CalcWorkSheet.Cells[34, 45] = CKAlum.Checked.ToString();//45
                    CalcWorkSheet.Cells[34, 46] = CKMobra.Checked.ToString();//46
                    CalcWorkSheet.Cells[34, 47] = CKSD.Checked.ToString();//47
                    CalcWorkSheet.Cells[34, 48] = CKlux.Checked.ToString();//48
                    CalcWorkSheet.Cells[34, 49] = CKvsol.Checked.ToString();//49
                    CalcWorkSheet.Cells[34, 50] = CKCort.Checked.ToString();//50
                    CalcWorkSheet.Cells[34, 51] = CKTor.Checked.ToString();//51
                    CalcWorkSheet.Cells[34, 52] = CKPanel.Checked.ToString();//52
                    CalcWorkSheet.Cells[34, 53] = CKSopr.Checked.ToString();//53
                    CalcWorkSheet.Cells[34, 54] = CKCable.Checked.ToString();//54
                    CalcWorkSheet.Cells[34, 55] = CKPuerta.Checked.ToString();//55
                    CalcWorkSheet.Cells[34, 56] = CKDrenaje.Checked.ToString();//56
                    CalcWorkSheet.Cells[34, 57] = CKSellaje.Checked.ToString();//57
                    CalcWorkSheet.Cells[34, 58] = CKEmerg.Checked.ToString();//58
                    CalcWorkSheet.Cells[34, 59] = CKBrida.Checked.ToString();//59
                    CalcWorkSheet.Cells[34, 60] = CKNanauf.Checked.ToString();//60
                    CalcWorkSheet.Cells[34, 61] = CKRefrig.Checked.ToString();//61
                    CalcWorkSheet.Cells[34, 62] = CKPerf.Checked.ToString();//62
                    CalcWorkSheet.Cells[34, 63] = CKpmtal.Checked.ToString();//62
                    CalcWorkSheet.Cells[34, 64] = CRvent.Checked.ToString();//64
                    CalcWorkSheet.Cells[34, 65] = CKppc.Checked.ToString();//65
                    CalcWorkSheet.Cells[34, 66] = CKepc.Checked.ToString();//66
                    CalcWorkSheet.Cells[9, 13] = CKp10.Checked.ToString();
                    CalcWorkSheet.Cells[10, 13] = CKp12.Checked.ToString();
                    CalcWorkSheet.Cells[11, 13] = CKp15.Checked.ToString();
                    CalcWorkSheet.Cells[11, 13] = CKp15t.Checked.ToString();
                    CalcWorkSheet.Cells[14, 17] = CFase.Text;
                    CalcWorkSheet.Cells[151, 18] = "=SI.ERROR(AC19;\"\")";
                    CalcWorkSheet.Cells[152, 18] = "=SI.ERROR(AC20;\"\")";
                    CalcWorkSheet.Cells[16, 27] = RBun.Checked.ToString();// BT+ MT
                    CalcWorkSheet.Cells[34, 67] = CKsu1.Checked.ToString();//67
                    CalcWorkSheet.Cells[34, 68] = CKsu2.Checked.ToString();//68
                    CalcWorkSheet.Cells[34, 69] = CKsu3.Checked.ToString();//69

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[2];
                    CalcWorkSheet.Cells[5, 34] = "=SI(N5=1;\"\";SI(IZQUIERDA(AI5;5)=\"Error\";AI5;SI(CA27>CA28;CA27;CA28)))";
                    CalcWorkSheet.Cells[27, 79] = "=SI(Y(P5<-5;BU27>=14);CA25/BU27;SI(Y(P5>-4,9;BU27>=12);CA25/BU27;CC25))";
                    CalcWorkSheet.Cells[28, 79] = "=SI(P5<-5;(SUMA(CA7:CA16)/14)*BY25;(SUMA(CA7:CA16)/12)*BY25)";
                    CalcWorkSheet.Cells[1, 7] = "Software ProJDC v.6.4";
                    CalcWorkSheet.Cells[5, 22] = "=SI(N5=1;\"\";BUSCARV(VALOR(M5);Coef_den_carga;2;1)*BUSCARV(N5;Base;17;0))";
                    CalcWorkSheet.Cells[5, 27] = "=SI(T5<-2;0;REDONDEAR(BUSCARV(N5;Base;16;0)/0,86;-1))";
                    CalcWorkSheet.Cells[5, 28] = "=REDONDEAR(BUSCARV(N5;Base;15;0)/0,86;-1)";
                    

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[7];
                    CalcWorkSheet.Cells[833, 5] = "=SI(Datos!AB31=\"EX2-M00\";SI.ERROR(SI(C22<=12;H841;0);0);0)";
                    CalcWorkSheet.Cells[834, 5] = "=SI(Datos!AB31=\"EX2-M00\";SI.ERROR(SI(C22>12;H842;0);0);0)";

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];
                    CalcWorkSheet.Cells[23, 16] = CCamara.Text;
                    CalcWorkSheet.Cells[14, 13] = TLargo.Text;
                    CalcWorkSheet.Cells[15, 13] = TAncho.Text;
                    CalcWorkSheet.Cells[16, 13] = TAlto.Text;
                    CalcWorkSheet.Cells[18, 13] = TTem.Text;
                    CalcWorkSheet.Cells[13, 13] = pk.ToString();
                    CalcWorkSheet.Cells[7, 13] = CKPanel.Checked.ToString();
                    CalcWorkSheet.Cells[8, 13] = CKexpo.Checked.ToString();
                    CalcWorkSheet.Cells[31, 18] = Ctpd.Text;
                    CalcWorkSheet.Cells[27, 13] = TCantEv.Text;
                    CalcWorkSheet.Cells[31, 26] = TCentx.Text;
                    CalcWorkSheet.Cells[31, 27] = TCdin.Text;
                    CalcWorkSheet.Cells[31, 25] = TCxp.Text;
                    CalcWorkSheet.Cells[31, 20] = TMevp.Text;
                    CalcWorkSheet.Cells[26, 13] = CTamb.Text;
                    CalcWorkSheet.Cells[31, 28] = Ctxv.Text;
                    CalcWorkSheet.Cells[2, 18] = TDECF.Text;
                    CalcWorkSheet.Cells[3, 18] = TDEC.Text;
                    CalcWorkSheet.Cells[4, 18] = TDECE.Text;
                    CalcWorkSheet.Cells[5, 18] = TDECH.Text;
                    CalcWorkSheet.Cells[6, 18] = TDmc1.Text;
                    CalcWorkSheet.Cells[2, 16] = TDECF.Text;
                    CalcWorkSheet.Cells[3, 16] = CTEvap.Text;
                    CalcWorkSheet.Cells[4, 16] = CTCond.Text;
                    CalcWorkSheet.Cells[5, 16] = CRefrig.Text;
                    CalcWorkSheet.Cells[32, 13] = CDigt.Text;
                    CalcWorkSheet.Cells[24, 9] = ("=REDONDEAR.MAS(CONVERTIR(F24;\"Wh\";\"BTU\");0)");
                    CalcWorkSheet.Cells[29, 15] = ("=SI.ERROR(U31;1)");
                    CalcWorkSheet.Cells[29, 13] = ("=SI(O29<>1;INDICE(W31:W336;U31);SI(O29=1;75))");
                    CalcWorkSheet.Cells[30, 13] = ("=SI(O29<>1;INDICE(X31:X336;U31);SI(O29=1;0))");
                    CalcWorkSheet.Cells[8, 16] = RB60H.Checked;
                    CalcWorkSheet.Cells[11, 16] = CDT.Text;
                    CalcWorkSheet.Cells[14, 16] = CVolt.Text;
                    CalcWorkSheet.Cells[17, 16] = CKdt.Checked;
                    CalcWorkSheet.Cells[21, 16] = TTP.Text;
                    CalcWorkSheet.Cells[22, 16] = CSumi.Text;
                    CalcWorkSheet.Cells[2, 23] = CBint.Text;
                    CalcWorkSheet.Cells[3, 23] = TCC.Text;
                    CalcWorkSheet.Cells[14, 17] = CFase.Text;
                    CalcWorkSheet.Cells[28, 15] = "=PANELES!P6";
                    CalcWorkSheet.Cells[28, 16] = "=PANELES!Q6";
                    CalcWorkSheet.Cells[28, 17] = "=PANELES!R6";
                    CalcWorkSheet.Cells[114, 18] = "=SI.ERROR(SI(PSTP!D38=0;\"\";PSTP!D38);REDONDEAR(P6/0,48;2))";
                    CalcWorkSheet.Cells[115, 18] = "=SI.ERROR(SI(PSTP!D39=0;\"\";PSTP!D39);REDONDEAR(P6/7,0;2))";
                    CalcWorkSheet.Cells[147, 18] = "=SI.ERROR(SI(SCMC!AQ26=0;\"\";SCMC!AQ26);0)";
                    CalcWorkSheet.Cells[34, 13] = "=SUMAR.SI.CONJUNTO(E34:E70;H34:H70;\">=-5\";D34:D70;N34)";
                    CalcWorkSheet.Cells[35, 13] = "=SUMAR.SI.CONJUNTO(E34:E70;H34:H70;\"<-5\";D34:D70;N34)";

                    Range NewRangeCalc = CalcWorkSheet.get_Range("M28", "Q28");
                    Array myNewArr = (Array)NewRangeCalc.Value2;
                    String[,] myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                    for (int i = 0; i < myNewArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < myNewArr.GetLength(1); j++)
                        {
                            long[] indices = new long[] { i + 1, j + 1 };
                            myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                        }
                    }
                    //CalcWorkBook.SaveCopyAs(@"C:\prueba.xlsx");
                    TCF.Text = myNewArrCalc.GetValue(1, 1).ToString();
                    TFW.Text = myNewArrCalc.GetValue(1, 2).ToString();
                    TCmod.Text = myNewArrCalc.GetValue(1, 3).ToString();
                    CalcWorkSheet.Cells[6, 16] = TFW.Text;
                    CalcWorkSheet.Cells[7, 16] = TConD.Text;
                    CalcWorkSheet.Cells[12, 16] = Cmodex.Text;
                    CalcWorkSheet.Cells[13, 16] = Cevap.Text;

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[25];
                    CalcWorkSheet.Cells[8, 7] = "=REDOND.MULT(REDONDEAR.MAS((G6*2);0);9)";
                    CalcWorkSheet.Cells[8, 8] = "=REDOND.MULT(REDONDEAR.MAS((G6*2);0);16)";
                    

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[28];
                    CalcWorkSheet.Cells[9, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q9;$C$5:$C$68;1));0)";
                    CalcWorkSheet.Cells[10, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q10;$C$5:$C$68;1));0)";
                    CalcWorkSheet.Cells[11, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q11;$C$5:$C$68;1));0)";
                    CalcWorkSheet.Cells[12, 19] = "=SI.ERROR(INDICE($H$5:$H$68;COINCIDIR(Q12;$C$5:$C$68;1));0)";
                    CalcWorkSheet.Cells[13, 19] = "=SI.ERROR(SI(W9<Q9;INDICE($H$5:$H$68;COINCIDIR(Q9-INDICE($C$5:$C$68;COINCIDIR(Q9;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                    CalcWorkSheet.Cells[14, 19] = "=SI.ERROR(SI(W10<Q10;INDICE($H$5:$H$68;COINCIDIR(Q10-INDICE($C$5:$C$68;COINCIDIR(Q10;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                    CalcWorkSheet.Cells[15, 19] = "=SI.ERROR(SI(W11<Q11;INDICE($H$5:$H$68;COINCIDIR(Q11-INDICE($C$5:$C$68;COINCIDIR(Q11;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";
                    CalcWorkSheet.Cells[16, 19] = "=SI.ERROR(SI(W12<Q12;INDICE($H$5:$H$68;COINCIDIR(Q12-INDICE($C$5:$C$68;COINCIDIR(Q12;$C$5:$C$68;1));$C$5:$C$68;1));0);0)";

                    //********************************************************* CONTROL DE VALVULAS SANHUA ****
                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[31];// PAGINA ACCESS
                    CalcWorkSheet.Cells[13, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN13:AU13;-1));0)";
                    CalcWorkSheet.Cells[14, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN14:AU14;-1));0)";
                    CalcWorkSheet.Cells[15, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN15:AU15;-1));0)";
                    CalcWorkSheet.Cells[16, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN16:AU16;-1));0)";
                    CalcWorkSheet.Cells[17, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN17:AU17;-1));0)";
                    CalcWorkSheet.Cells[18, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN18:AU18;-1));0)";
                    CalcWorkSheet.Cells[19, 38] = "=SI.ERROR(INDICE($AN$2:$AU$2;COINCIDIR($M$30;AN19:AU19;-1));0)";
                    //********************************************************* CONTROL DE CMC 2-3 ****
                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[32];// PAGINA ACCESS
                    CalcWorkSheet.Cells[59, 6] = "=SI.ERROR(INDICE(M44:M57;COINCIDIR(F61;F44:F57;-1));INDICE(M17:M26;COINCIDIR(F61;F17:F26;-1)))";

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[1];

                    NewRangeCalc = CalcWorkSheet.get_Range("R9", "R300");
                    myNewArr = (Array)NewRangeCalc.Value2;
                    myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                    for (int i = 0; i < myNewArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < myNewArr.GetLength(1); j++)
                        {
                            long[] indices = new long[] { i + 1, j + 1 };
                            myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                        }
                    }
                    TQevp.Text = myNewArrCalc.GetValue(1, 1).ToString();
                    TTint.Text = myNewArrCalc.GetValue(20, 1).ToString();
                    TEquip.Text = myNewArrCalc.GetValue(2, 1).ToString();
                    TValv.Text = myNewArrCalc.GetValue(7, 1).ToString();
                    TCodValv.Text = myNewArrCalc.GetValue(8, 1).ToString();
                    TTmos.Text = myNewArrCalc.GetValue(25, 1).ToString();
                    TInc.Text = myNewArrCalc.GetValue(16, 1).ToString();
                    TInev.Text = myNewArrCalc.GetValue(15, 1).ToString();
                    TIned.Text = myNewArrCalc.GetValue(21, 1).ToString();
                    TIncd.Text = myNewArrCalc.GetValue(132, 1).ToString();
                    TIpv.Text = myNewArrCalc.GetValue(3, 1).ToString();
                    TIcc.Text = myNewArrCalc.GetValue(4, 1).ToString();
                    TCint1.Text = myNewArrCalc.GetValue(17, 1).ToString();
                    TCint2.Text = myNewArrCalc.GetValue(18, 1).ToString();
                    TCint3.Text = myNewArrCalc.GetValue(19, 1).ToString();
                    TCmce.Text = myNewArrCalc.GetValue(5, 1).ToString();
                    TPmce.Text = myNewArrCalc.GetValue(6, 1).ToString();
                    TDmce.Text = myNewArrCalc.GetValue(24, 1).ToString();
                    TDmc4.Text = myNewArrCalc.GetValue(26, 1).ToString();
                    TDmc5.Text = myNewArrCalc.GetValue(27, 1).ToString();
                    TDmc6.Text = myNewArrCalc.GetValue(28, 1).ToString();
                    TDmc3.Text = myNewArrCalc.GetValue(29, 1).ToString();
                    TDmc7.Text = myNewArrCalc.GetValue(30, 1).ToString();
                    TDmc8.Text = myNewArrCalc.GetValue(31, 1).ToString();
                    TDmc2.Text = myNewArrCalc.GetValue(32, 1).ToString();
                    TLcc.Text = myNewArrCalc.GetValue(33, 1).ToString();
                    TLss.Text = myNewArrCalc.GetValue(34, 1).ToString();
                    TPcmc.Text = myNewArrCalc.GetValue(35, 1).ToString();
                    TPcond.Text = myNewArrCalc.GetValue(36, 1).ToString();
                    TPlq.Text = myNewArrCalc.GetValue(37, 1).ToString();
                    TPvq.Text = myNewArrCalc.GetValue(38, 1).ToString();
                    TPls.Text = myNewArrCalc.GetValue(39, 1).ToString();
                    TPosc.Text = myNewArrCalc.GetValue(40, 1).ToString();
                    TPsq.Text = myNewArrCalc.GetValue(41, 1).ToString();
                    TPcq.Text = myNewArrCalc.GetValue(42, 1).ToString();
                    TPcv.Text = myNewArrCalc.GetValue(43, 1).ToString();
                    TPcx.Text = myNewArrCalc.GetValue(44, 1).ToString();
                    TPcy.Text = myNewArrCalc.GetValue(45, 1).ToString();
                    TPcp.Text = myNewArrCalc.GetValue(46, 1).ToString();
                    TPex.Text = myNewArrCalc.GetValue(47, 1).ToString();
                    TPrs.Text = myNewArrCalc.GetValue(48, 1).ToString();
                    TPem.Text = myNewArrCalc.GetValue(49, 1).ToString();
                    TPnt.Text = myNewArrCalc.GetValue(50, 1).ToString();
                    TPml.Text = myNewArrCalc.GetValue(51, 1).ToString();
                    TPdtr.Text = myNewArrCalc.GetValue(52, 1).ToString();
                    TPdtc.Text = myNewArrCalc.GetValue(53, 1).ToString();
                    TPdt2.Text = myNewArrCalc.GetValue(54, 1).ToString();
                    TPdt1.Text = myNewArrCalc.GetValue(55, 1).ToString();
                    TCp80.Text = myNewArrCalc.GetValue(56, 1).ToString();
                    TCp100.Text = myNewArrCalc.GetValue(57, 1).ToString();
                    TCp120.Text = myNewArrCalc.GetValue(58, 1).ToString();
                    TCp150.Text = myNewArrCalc.GetValue(109, 1).ToString();
                    TCt80.Text = myNewArrCalc.GetValue(59, 1).ToString();
                    TCt100.Text = myNewArrCalc.GetValue(60, 1).ToString();
                    TCt120.Text = myNewArrCalc.GetValue(61, 1).ToString();
                    TCt150.Text = myNewArrCalc.GetValue(62, 1).ToString();
                    TCp80m.Text = myNewArrCalc.GetValue(63, 1).ToString();
                    TCp100m.Text = myNewArrCalc.GetValue(64, 1).ToString();
                    TCp120m.Text = myNewArrCalc.GetValue(65, 1).ToString();
                    TCp150m.Text = myNewArrCalc.GetValue(110, 1).ToString();
                    TCt80m.Text = myNewArrCalc.GetValue(66, 1).ToString();
                    TCt100m.Text = myNewArrCalc.GetValue(67, 1).ToString();
                    TCt120m.Text = myNewArrCalc.GetValue(68, 1).ToString();
                    TCt150m.Text = myNewArrCalc.GetValue(69, 1).ToString();
                    SPtp84.Text = myNewArrCalc.GetValue(70, 1).ToString();
                    SPtp83.Text = myNewArrCalc.GetValue(71, 1).ToString();
                    SPtp78.Text = myNewArrCalc.GetValue(72, 1).ToString();
                    SPtp76.Text = myNewArrCalc.GetValue(73, 1).ToString();
                    SPtp75.Text = myNewArrCalc.GetValue(74, 1).ToString();
                    SPtp74.Text = myNewArrCalc.GetValue(75, 1).ToString();
                    SPtp73.Text = myNewArrCalc.GetValue(76, 1).ToString();
                    SPtp72.Text = myNewArrCalc.GetValue(77, 1).ToString();
                    SPtp85.Text = myNewArrCalc.GetValue(79, 1).ToString();
                    SPtp86.Text = myNewArrCalc.GetValue(80, 1).ToString();
                    SPtp87.Text = myNewArrCalc.GetValue(81, 1).ToString();
                    SPtp88.Text = myNewArrCalc.GetValue(82, 1).ToString();
                    SPtp89.Text = myNewArrCalc.GetValue(83, 1).ToString();
                    SPtp90.Text = myNewArrCalc.GetValue(84, 1).ToString();
                    SPtp91.Text = myNewArrCalc.GetValue(85, 1).ToString();
                    SPtp92.Text = myNewArrCalc.GetValue(86, 1).ToString();
                    SPtp93.Text = myNewArrCalc.GetValue(87, 1).ToString();
                    SPtp94.Text = myNewArrCalc.GetValue(88, 1).ToString();
                    SPtp95.Text = myNewArrCalc.GetValue(89, 1).ToString();
                    SPtp96.Text = myNewArrCalc.GetValue(90, 1).ToString();
                    SPtp97.Text = myNewArrCalc.GetValue(91, 1).ToString();
                    SPtp98.Text = myNewArrCalc.GetValue(92, 1).ToString();
                    SPtp99.Text = myNewArrCalc.GetValue(93, 1).ToString();
                    SPtp100.Text = myNewArrCalc.GetValue(94, 1).ToString();
                    SPtp101.Text = myNewArrCalc.GetValue(95, 1).ToString();
                    SPtp102.Text = myNewArrCalc.GetValue(96, 1).ToString();
                    SPtp103.Text = myNewArrCalc.GetValue(97, 1).ToString();
                    SPtp104.Text = myNewArrCalc.GetValue(98, 1).ToString();
                    SPtp105.Text = myNewArrCalc.GetValue(99, 1).ToString();
                    SPtp106.Text = myNewArrCalc.GetValue(100, 1).ToString();
                    SPtp107.Text = myNewArrCalc.GetValue(101, 1).ToString();
                    SPtp108.Text = myNewArrCalc.GetValue(102, 1).ToString();
                    SPtp109.Text = myNewArrCalc.GetValue(103, 1).ToString();
                    SPtp110.Text = myNewArrCalc.GetValue(104, 1).ToString();
                    SPtp111.Text = myNewArrCalc.GetValue(105, 1).ToString();
                    SPtp136.Text = myNewArrCalc.GetValue(106, 1).ToString();
                    SPtp137.Text = myNewArrCalc.GetValue(107, 1).ToString();
                    SPtp141.Text = myNewArrCalc.GetValue(134, 1).ToString();
                    SPtp142.Text = myNewArrCalc.GetValue(134, 1).ToString();
                    SPtp143.Text = myNewArrCalc.GetValue(135, 1).ToString();
                    SPtp144.Text = myNewArrCalc.GetValue(136, 1).ToString();
                    SPtp145.Text = myNewArrCalc.GetValue(137, 1).ToString();
                    SPtp146.Text = myNewArrCalc.GetValue(138, 1).ToString();
                    SPtp147.Text = myNewArrCalc.GetValue(139, 1).ToString();
                    SPtp148.Text = myNewArrCalc.GetValue(140, 1).ToString();
                    SPtp150.Text = myNewArrCalc.GetValue(143, 1).ToString();
                    SPtp151.Text = myNewArrCalc.GetValue(144, 1).ToString();
                    TPcmc1.Text = myNewArrCalc.GetValue(145, 1).ToString();
                    TPcmc2.Text = myNewArrCalc.GetValue(146, 1).ToString();
                    TPcmc3.Text = myNewArrCalc.GetValue(147, 1).ToString();
                    TQevpd.Text = myNewArrCalc.GetValue(151, 1).ToString();
                    TQevpc.Text = myNewArrCalc.GetValue(152, 1).ToString();

                    TCsist.Text = myNewArrCalc.GetValue(108, 1).ToString();
                    TQfw.Text = myNewArrCalc.GetValue(11, 1).ToString();
                    //codValv = myNewArrCalc.GetValue(8, 1).ToString();
                    /*
                    try
                    {

                        myOferta.GetCam(int.Parse(CCamara.Text) - 1).SetCodValv(myNewArrCalc.GetValue(8, 1).ToString());

                    }
                    catch { }
                    */
                    TDesc.Text = myNewArrCalc.GetValue(9, 1).ToString();
                    TPrec.Text = myNewArrCalc.GetValue(10, 1).ToString();
                    TQfep.Text = myNewArrCalc.GetValue(11, 1).ToString();
                    TScdro.Text = myNewArrCalc.GetValue(12, 1).ToString();

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[3];
                    NewRangeCalc = CalcWorkSheet.get_Range("T361", "T372");
                    myNewArr = (Array)NewRangeCalc.Value2;
                    myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                    for (int i = 0; i < myNewArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < myNewArr.GetLength(1); j++)
                        {
                            long[] indices = new long[] { i + 1, j + 1 };
                            myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                        }
                    }
                    TEmevp.Text = myNewArrCalc.GetValue(2, 1).ToString();
                    TSpsi.Text = myNewArrCalc.GetValue(10, 1).ToString();
                    TStemp.Text = myNewArrCalc.GetValue(11, 1).ToString();
                    TApsi.Text = myNewArrCalc.GetValue(12, 1).ToString();


                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
                    if (RBun.Checked == false)
                    {
                        CalcWorkSheet.Cells[12, 26] = "=SI(AQ12=0;AQ13;AQ12)";
                        CalcWorkSheet.Cells[12, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-5\")");
                        CalcWorkSheet.Cells[13, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-5\")");
                        CalcWorkSheet.Cells[108, 69] = "=SI.ERROR(INDICE(BI151:BI165;BI122+1);INDICE(BI151:BI165;BI122))";
                        CalcWorkSheet.Cells[116, 69] = "=SI.ERROR(INDICE(BQ132:BQ143;BQ121+1);INDICE(BQ132:BQ143;BQ121))";

                        CalcWorkSheet.Cells[12, 27] = "=SI(AO12=0;AO13;AO12)";
                        if (int.Parse(CTEvap.Text) < -5)
                        {
                            CalcWorkSheet.Cells[12, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-5\")*0,72+D37;2);REDONDEAR(1,18*SUMAR.SI.CONJUNTO($S$4:$S$40;$V$4:$V$40;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-5\");2)))";
                            CalcWorkSheet.Cells[13, 41] = "0";
                        }
                        if (int.Parse(CTEvap.Text) >= -5)
                        {
                            CalcWorkSheet.Cells[13, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-5\")*0,76+D37;2);REDONDEAR(1,18*SUMAR.SI.CONJUNTO($S$4:$S$40;$V$4:$V$40;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-5\");2)))";
                            CalcWorkSheet.Cells[12, 41] = "0";
                        }

                        CalcWorkSheet.Cells[12, 28] = "=SI(AP12=0;AP13;AP12)";
                        CalcWorkSheet.Cells[12, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\"<-10\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$10;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$H$23;COINCIDIR(D4;$E$22:$H$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,24);2);0))";
                        CalcWorkSheet.Cells[13, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\";$U$4:$U$40;\">=-10\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$9;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$H$23;COINCIDIR(D4;$E$22:$H$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,34);2);0))";

                    }
                    else
                    {
                        CalcWorkSheet.Cells[12, 26] = "=SI(AQ12=0;AQ13;AQ12)";
                        CalcWorkSheet.Cells[12, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\")");
                        CalcWorkSheet.Cells[13, 43] = ("=SUMAR.SI.CONJUNTO($T$4:$T$40;$V$4:$V$40;1;$W$4:$W$40;\"MULT-COMP-INT\")");
                        CalcWorkSheet.Cells[8, 42] = RBun.Checked.ToString();

                        CalcWorkSheet.Cells[12, 27] = "=SI(AO12=0;AO13;AO12)";
                        if (int.Parse(CTEvap.Text) < -5)
                        {
                            CalcWorkSheet.Cells[12, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\")*0,72+D37;2);REDONDEAR(1,18*SUMAR.SI.CONJUNTO($S$4:$S$40;$V$4:$V$40;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)))";
                            CalcWorkSheet.Cells[13, 41] = "0";
                        }
                        if (int.Parse(CTEvap.Text) >= -5)
                        {
                            CalcWorkSheet.Cells[13, 41] = "=si(Y10=0;0;SI(D8=VERDADERO;REDONDEAR(SUMAR.SI.CONJUNTO($S$4:$S$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\")*0,76+D37;2);REDONDEAR(1,18*SUMAR.SI.CONJUNTO($S$4:$S$40;$V$4:$V$40;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)))";
                            CalcWorkSheet.Cells[12, 41] = "0";
                        }

                        CalcWorkSheet.Cells[12, 28] = "=SI(AP12=0;AP13;AP12)";
                        CalcWorkSheet.Cells[12, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$10;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$H$23;COINCIDIR(D4;$E$22:$H$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,24);2);0))";
                        CalcWorkSheet.Cells[13, 42] = "=si(Y10=0;0;SI.ERROR(REDONDEAR((REDONDEAR(SUMAR.SI.CONJUNTO($T$4:$T$40;$AT$12:$AT$48;Y10;$W$4:$W$40;\"MULT-COMP-INT\");2)*INDICE($F$12:$P$16;COINCIDIR(D7;$E$12:$E$16;-1);COINCIDIR($Z$9;$F$11:$P$11;-1)))/($Z$5*INDICE($E$23:$H$23;COINCIDIR(D4;$E$22:$H$22;0))*INDICE($F$27:$M$27;COINCIDIR($Z$8;$F$26:$M$26;-1))*2,34);2);0))";

                    }
                    // LECTURA DEL NUMERO DE MULTICOMPRESORA ACTIVA//
                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[18];
                    NewRangeCalc = CalcWorkSheet.get_Range("Y11", "Y31");
                    myNewArr = (Array)NewRangeCalc.Value2;
                    myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                    for (int i = 0; i < myNewArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < myNewArr.GetLength(1); j++)
                        {
                            long[] indices = new long[] { i + 1, j + 1 };
                            myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                        }
                    }
                    TMcc.Text = myNewArrCalc.GetValue(1, 1).ToString();

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[30];// PAGINA COND-INT
                    //CalcWorkSheet.Cells[1, 1] = CCmci.Text;
                    //CalcWorkSheet.Cells[1, 118] = CCmci.Text;

                    CalcWorkSheet = (_Worksheet)CalcWorkBook.Worksheets[24];// PAGINA ACCESS
                    CalcWorkSheet.Cells[7, 11] = CCmci.Text;
                    if (int.Parse(CCmci.Text) == 1)
                    {
                        CalcWorkSheet.Cells[8, 12] = "=SI($K$7=1;$K$3;0)";
                    }
                    if (int.Parse(CCmci.Text) == 2)
                    {
                        CalcWorkSheet.Cells[9, 12] = "=SI($K$7=2;$K$3;0)";
                    }
                    if (int.Parse(CCmci.Text) == 3)
                    {
                        CalcWorkSheet.Cells[10, 12] = "=SI($K$7=3;$K$3;0)";
                    }
                    if (int.Parse(CCmci.Text) == 4)
                    {
                        CalcWorkSheet.Cells[11, 12] = "=SI($K$7=4;$K$3;0)";
                    }
                    if (int.Parse(CCmci.Text) == 5)
                    {
                        CalcWorkSheet.Cells[12, 12] = "=SI($K$7=5;$K$3;0)";
                    }

                    NewRangeCalc = CalcWorkSheet.get_Range("K8", "K15");
                    myNewArr = (Array)NewRangeCalc.Value2;
                    myNewArrCalc = new String[myNewArr.GetLength(0) + 1, myNewArr.GetLength(1) + 1];
                    for (int i = 0; i < myNewArr.GetLength(0); i++)
                    {
                        for (int j = 0; j < myNewArr.GetLength(1); j++)
                        {
                            long[] indices = new long[] { i + 1, j + 1 };
                            myNewArrCalc.SetValue(Convert.ToString(myNewArr.GetValue(i + 1, j + 1), new CultureInfo("es-ES")), indices);
                        }
                    }
                    if (int.Parse(CCmci.Text) == 1)
                    {
                        CalcWorkSheet.Cells[1, 1] = myNewArrCalc.GetValue(1, 1).ToString();
                    }
                    if (int.Parse(CCmci.Text) == 2)
                    {
                        CalcWorkSheet.Cells[2, 1] = myNewArrCalc.GetValue(2, 1).ToString();
                    }
                    if (int.Parse(CCmci.Text) == 3)
                    {
                        CalcWorkSheet.Cells[3, 1] = myNewArrCalc.GetValue(3, 1).ToString();
                    }
                    if (int.Parse(CCmci.Text) == 4)
                    {
                        CalcWorkSheet.Cells[4, 1] = myNewArrCalc.GetValue(4, 1).ToString();
                    }
                    if (int.Parse(CCmci.Text) == 5)
                    {
                        CalcWorkSheet.Cells[5, 1] = myNewArrCalc.GetValue(5, 1).ToString();
                    }

                    //***********************************************
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.Message,
                                "Error de datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

            }
        }

        private void TTem_TextChanged(object sender, EventArgs e)
        {
            try
            {
                CTEvap.Text = (int.Parse(TTem.Text) - int.Parse(CDT.Text)).ToString();
            }
            catch { }
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
        
        private void CBint_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void toolStrip2_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void label54_Click(object sender, EventArgs e)
        {

        }

        private void Cmode_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Cmode_SelectedIndexChanged_1(object sender, EventArgs e)
        {

        }

        private void Cnof_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Coff_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void Txm2_TextChanged(object sender, EventArgs e)
        {

        }

        private void myXlsxSaveDialog_FileOk(object sender, CancelEventArgs e)
        {

        }

        private void TCmeapm_TextChanged(object sender, EventArgs e)
        {

        }

        private void TFecha_TextChanged(object sender, EventArgs e)
        {

        }

        private void TMevp_TextChanged(object sender, EventArgs e)
        {

        }

        private void label55_Click(object sender, EventArgs e)
        {

        }

        private void label71_Click(object sender, EventArgs e)
        {

        }

        private void label63_Click(object sender, EventArgs e)
        {

        }

        private void TFW_TextChanged(object sender, EventArgs e)
        {

        }
       

        private void TJ5_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ6_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ7_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ8_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ16_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ15_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ14_TextChanged(object sender, EventArgs e)
        {

        }

        private void TJ13_TextChanged(object sender, EventArgs e)
        {

        }
        
        private void button5_Click(object sender, EventArgs e)
        {
            
            try
            {
                BExportar.Enabled = false;
                System.Windows.Forms.Application.Exit();
            }
            catch { }

            Microsoft.Office.Interop.Excel.Application ExcelApp = null;
            _Workbook myWorkBook = null;
            _Worksheet myWorkSheet = null;
            //myWorkBook.SaveCopyAs(myXlsxSaveDialog.FileName);
            //ExcelApp.ActiveWorkbook.Close(false, @"fnstpte.tpt", Type.Missing);
            //ExcelApp.Quit(); 
        }

        private void btnGenerarPDF_Click(object sender, EventArgs e)
        {

            string pdfPath = Path.Combine(System.Windows.Forms.Application.StartupPath, "archivo.pdf");

            Process.Start(pdfPath);

        }

        private void label51_Click(object sender, EventArgs e)
        {

        }

        private void label14_Click(object sender, EventArgs e)
        {

        }
        
    }


    //------------------------------------------------------------------------------------------
    //-------------------------------------------------------------------------------
     //---------------------------------------------------------------------------------
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
            try
            {
                thisForm.LEstado.Text = "Generando oferta... 0%";
                this.CreateApplication();
                thisForm.LEstado.Text = "Generando oferta... 10%";
                //this.OpenFile(local, thisForm.RUSD.Checked);
                this.OpenFile(local, thisForm.RCUC.Checked);
                this.OpenFile(local, thisForm.RGX.Checked);
                this.OpenFile(local, thisForm.RFTX.Checked);
                this.OpenFile(local, thisForm.RODC.Checked);
                thisForm.LEstado.Text = "Generando oferta... 15%";
                this.ChangeSheet(1);
                myWorkSheet.Cells[7, 4] = Oferta.GetNP();
                myWorkSheet.Cells[12, 4] = Oferta.GetNO();
                myWorkSheet.Cells[10, 4] = Oferta.GetREF();
                thisForm.LEstado.Text = "Generando oferta... 20%";
                this.ChangeSheet(8);
                //myWorkSheet.Cells[4,  6] = Oferta.GetBscu();
                myWorkSheet.Cells[11, 6] = "=SI(PRESUPUESTO!H70=0;\"\";PRESUPUESTO!H70)";
                myWorkSheet.Cells[12, 6] = "=SI(SI.ERROR(PRESUPUESTO!I70;\"\")=0;\"\";SI.ERROR(PRESUPUESTO!I70;\"\"))";
                myWorkSheet.Cells[17, 6] = "=SI.ERROR(REDONDEAR.MAS(F15/F16;0);\"\")";
                //myWorkSheet.Cells[19, 6] = "=SI(PRESUPUESTO!F53+PRESUPUESTO!F54+PRESUPUESTO!F55=0;\"\";PRESUPUESTO!F53+PRESUPUESTO!F54+PRESUPUESTO!F55)";
                myWorkSheet.Cells[21, 6] = "=F16-1";
                myWorkSheet.Cells[6, 6] = "=SI(PRESUPUESTO!F50=0;\"\";PRESUPUESTO!F50)";
                myWorkSheet.Cells[23, 6] = "=SI.ERROR(F22*2;\"\")";
                thisForm.LEstado.Text = "Generando oferta... 25%";
                this.ChangeSheet(3);
                myWorkSheet.Cells[30, 3] = "Sr(a). " + Oferta.GetClit() + "";
                myWorkSheet.Cells[31, 3] = Oferta.GetClit1();

                this.ChangeSheet(6);
                myWorkSheet.Cells[31, 6] = Oferta.GetClitm();

                this.ChangeSheet(7);
                for (int i = 0; i < thisForm.myOferta.GetCantCam(); i++)
                myWorkSheet.Cells[i + 6, 3] = thisForm.myOferta.GetCam(i).GetPK() + 1;
                myWorkSheet.Cells[48, 4] = Oferta.GetNP();
                
                //myWorkSheet.Cells[79, 13] = Oferta.GetNO();
                myWorkSheet.Cells[107, 4]=Oferta.GetCmat();
                
                myWorkSheet.Cells[380, 738] = Oferta.GetDmc3();
                myWorkSheet.Cells[381, 738] = Oferta.GetDmc4();
                myWorkSheet.Cells[382, 738] = Oferta.GetDmc5();
                myWorkSheet.Cells[383, 738] = Oferta.GetDmc6();
                //myWorkSheet.Cells[51, 18] = "Software ProJDC v.5.2";
                myWorkSheet.Cells[49, 5] = Oferta.GetBmoni();
                myWorkSheet.Cells[49, 6] = Oferta.GetB60H();
                myWorkSheet.Cells[49, 7] = Oferta.GetBinvert();
                //myWorkSheet.Cells[80, 13] = Oferta.GetREF();
                //myWorkSheet.Cells[8, 2] = " » " + "CODIGO OFERTA: " + Oferta.GetNumCam() + "." + String.Format("{0:yyyyMMdd.hhmm}", DateTime.Now) + " » ";
                myWorkSheet.Cells[2, 3] = Oferta.GetREF() + " » " + "CODIGO OFERTA: " + Oferta.GetNO() + " » " + "Software ProJDC v.6.4";
                //myWorkSheet.Cells[78, 13] = " » " +  Oferta.GetNumCam() + "." + String.Format("{0:yyyyMMdd.hhmm}", DateTime.Now) + " » ";
                //myWorkSheet.Cells[89, 16] = " » " + "CODIGO OFERTA: " + Oferta.GetNumCam() + "." + String.Format("{0:yyyyMMdd.hhmm}", DateTime.Now) + " » ";
                thisForm.LEstado.Text = "Generando oferta... 30%";
                /*
                if (thisForm.RCUC.Checked)
                    myWorkSheet.Cells[48, 5] = "1";
                else
                    myWorkSheet.Cells[48, 5] = "0";
                */       
                int cc = Oferta.GetCantCam();

                myWorkSheet.Unprotect("midea");
                int perby;

                string myval = (70 / cc).ToString();                
                perby = int.Parse( myval );
                
                
                //=====================================================================================
                this.ChangeSheet(2);
                myWorkSheet.Cells[1,10] = Oferta.GetDesE();
                myWorkSheet.Cells[1, 1] = Oferta.GetResinaM();
                myWorkSheet.Cells[1, 2] = Oferta.Getinc();
                myWorkSheet.Cells[1, 3] = Oferta.GetConstCivPan();
                myWorkSheet.Cells[1, 4] = Oferta.GetEquipFrig();
                myWorkSheet.Cells[1, 14] = Oferta.GetDigt();
                myWorkSheet.Cells[1, 9] = Oferta.GetTasa();
                myWorkSheet.Cells[731, 5] = "EX01009";
                myWorkSheet.Cells[731, 6] = "=PRESUPUESTO!AHY45";
                //myWorkSheet.Cells[782, 5] = Oferta.GetB360();
                /*
                if (thisForm.RGX.Checked == true)
                {
                    myWorkSheet.Cells[1, 9] = Oferta.GetTasa();
                }
                else
                {
                    myWorkSheet.Cells[1, 9] = "1";
                }
                */
                myWorkSheet.Cells[1, 7] = Oferta.GetDsc();
                myWorkSheet.Cells[1, 6] = Oferta.GetPuertasFrig();
                
                if (thisForm.RUSD.Checked == true)
                {
                    myWorkSheet.Cells[1, 8] = "0";
                }
                else
                {
                    myWorkSheet.Cells[1, 8] = "1";
                }
                if (Oferta.GetBsup() == true)
                {
                    myWorkSheet.Cells[138, 6] = "=PRESUPUESTO!AFD42";
                }
                else
                {
                    myWorkSheet.Cells[138, 6] = "0";
                }
                this.ChangeSheet(7);
                myWorkSheet.Cells[464, 717] = Oferta.GetBun().ToString();
                myWorkSheet.Cells[259, 483] = Oferta.GetTasa();
                myWorkSheet.Cells[132, 488] = Oferta.GetResinaM();
                myWorkSheet.Cells[288, 483] = Oferta.GetResinaM();
                //myWorkSheet.Cells[120, 11] = Oferta.GetCredito();
                myWorkSheet.Cells[70, 8] = Oferta.GetCRcivil();
                myWorkSheet.Cells[70, 9] = Oferta.GetCRpiso();
                myWorkSheet.Cells[322, 522] = Oferta.GetDsc();
                /*
                this.ChangeSheet(11);
                myWorkSheet.Cells[25, 9] = Oferta.GetGastosIndObra();
                myWorkSheet.Cells[26, 9] = Oferta.GetGastosAdmObra();
                myWorkSheet.Cells[25, 10] = Oferta.GetGastosIndObracuc();
                myWorkSheet.Cells[26, 10] = Oferta.GetCredito();
                myWorkSheet.Cells[27, 9] = Oferta.GetCreditocup();
                */
                //***************************************** 22/02/2013
                this.ChangeSheet(7);
                myRange = myWorkSheet.get_Range("A1", "YK168");
                myValues = (Array)myRange.Value2;
                if (cc < 5)
                {
                    for (int i = 1; i <= cc; i++)
                    {
                        CCam myCam = Oferta.GetCam(i - 1);
                        myWorkSheet.Cells[i + 129, 559] = myCam.GetCBexpo();
                        if (myCam.GetKexpo())
                            myWorkSheet.Cells[i + 129, 558] = "0";
                        else
                            myWorkSheet.Cells[i + 129, 558] = "1";

                    }
                }
                else
                {
                    for (int i = 1; i <= 5; i++)
                    {
                        CCam myCam = Oferta.GetCam(i - 1);
                        myWorkSheet.Cells[i + 129, 559] = myCam.GetCBexpo();
                        if (myCam.GetKexpo())
                            myWorkSheet.Cells[i + 129, 558] = "0";
                        else
                            myWorkSheet.Cells[i + 129, 558] = "1";

                    }
                }
                this.ChangeSheet(7);
                myRange = myWorkSheet.get_Range("A1", "YK168");
                myValues = (Array)myRange.Value2;
                for (int i = 1; i <= cc; i++)
                {
                    CCam myCam = Oferta.GetCam(i - 1);
                    if (myCam.GetKepiso())
                        myWorkSheet.Cells[i + 129, 561] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 561] = "0";
                }
                this.ChangeSheet(2);
                myRange = myWorkSheet.get_Range("A1", "M998");
                myValues = (Array)myRange.Value2;
                if (cc < 5)
                {
                    for (int i = 1; i <= cc; i++)
                    {
                        CCam myCam = Oferta.GetCam(i - 1);

                        if (myCam.GetKexpo())
                            myWorkSheet.Cells[i + 82, 3] = "1";
                        else
                            myWorkSheet.Cells[i + 82, 3] = "0";

                    }

                }
                else
                {
                    for (int i = 1; i <= 5; i++)
                    {
                        CCam myCam = Oferta.GetCam(i - 1);

                        if (myCam.GetKexpo())
                            myWorkSheet.Cells[i + 82, 3] = "1";
                        else
                            myWorkSheet.Cells[i + 82, 3] = "0";

                    }
                }
                //====================================================================================                  
                this.ChangeSheet(7);
                myRange = myWorkSheet.get_Range("A1", "YK168");
                myValues = (Array)myRange.Value2;
               
                int percent = 30;
                int xfila = 6;
                int x2fila = 130;
                int x3fila = 390;
                int x4fila = 190;
                for (int i = 1; i <= cc; i++)
                {                    
                    CCam myCam = Oferta.GetCam(i - 1);
                    /*
                   //************************************** 03/02/2014 ******************************
                    myWorkSheet.Cells[i + 5, 13] = "=REDONDEAR.MAS(G" + xfila.ToString() + "*F"+ xfila.ToString() + ";2)";
                    myWorkSheet.Cells[i + 5, 14] = "=REDONDEAR.MAS(((F" + xfila.ToString() + "+G" + xfila.ToString() + ")*2)*H" + xfila.ToString() + "+(F" + xfila.ToString() + "*G" + xfila.ToString() + "*2);0)";
                    myWorkSheet.Cells[i + 5, 15] = "=E" + xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 16] ="=CONVERTIR(O" + xfila.ToString()+";\"C\";\"F\")";
                    myWorkSheet.Cells[i + 5, 17] = "=REDONDEAR.MAS(GB" + xfila.ToString() + " *71*SI(O" + xfila.ToString() + " <0;0,49744081)*1+GB" + xfila.ToString() + " *71*SI(O" + xfila.ToString() + ">=0;0,99488169)*1;0)";
                    myWorkSheet.Cells[i + 5, 18] = "=JB" + x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 20] = "=B"+ xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 21] = "=D"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 23] = "=I" + xfila.ToString() + "/V" + xfila.ToString();
                    myWorkSheet.Cells[i + 5, 24] = "=REDONDEAR.MAS(CONVERTIR(W" + xfila.ToString() + ";\"BTU\";\"Wh\");0)";
                    myWorkSheet.Cells[i + 5, 25] = "=REDONDEAR.MAS(CONVERTIR(W" + xfila.ToString() + ";\"BTU\";\"kcal\");0)";
                    myWorkSheet.Cells[i + 5, 26] = "=SI(CA" +xfila.ToString() +"<>\"MULT-COMP-INT\";PV" + x2fila.ToString() +";SI(JV"+ x2fila.ToString() +"=\"INTEGRADO_1 \";SI(E"+ xfila.ToString()+">-5;$AEB$375;$AEB$374);SI(JV"+x2fila.ToString()+"=\"INTEGRADO_2\";SI(E"+xfila.ToString()+ ">-5;$AEB$377;$AEB$376);SI(JV"+ x2fila.ToString() +"=\"INTEGRADO_3\";SI(E"+ xfila.ToString()+">-5;$AEB$379;$AEB$378)))))";
                    myWorkSheet.Cells[i + 5, 30] = "=HK"+ x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 32] = "=REDONDEAR.MAS(ABE" + x3fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 33] = "=AH"+ xfila.ToString()+"*AK"+ xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 34] = "=II"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 35] = "=ACH"+x4fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 41] = "=AK"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 43] = "=ABC"+x3fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 44] = "=ABD"+x3fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 45] = "=ABH"+x3fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 57] = "=B"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 60] = "=RAIZ(BN"+xfila.ToString()+")*(BM"+xfila.ToString()+"*BJ"+xfila.ToString()+"*0,7)/1000";
                    myWorkSheet.Cells[i + 5, 61] = "=BH"+xfila.ToString()+"/0,746";
                    myWorkSheet.Cells[i + 5, 62] = "=SI(CA" + xfila.ToString() + "<>\"MULT-COMP-INT\";PX" + x2fila.ToString() + ";REDONDEAR.MAS(SI(JV" + x2fila.ToString() + "=\"INTEGRADO 1\";SI(E" + xfila.ToString() + ">-5;$AEE$375;$AEE$374);SI(JV" + x2fila.ToString() + "=\"INTEGRADO 2\";SI(E" + xfila.ToString() + ">-5;$AEE$377;$AEE$376);SI(JV" + x2fila.ToString() + "=\"INTEGRADO 3\";SI(E" + xfila.ToString() + ">-5;$AEE$379;$AEE$378))))*X" + xfila.ToString() + "/100/SI(JV" + x2fila.ToString() + "=\"INTEGRADO 1\";SI(E" + xfila.ToString() + ">-5;$ADS$375;$ADS$374);SI(JV" + x2fila.ToString() + "=\"INTEGRADO 2\";SI(E" + xfila.ToString() + ">-5;$ADS$377;$ADS$376);SI(JV" + x2fila.ToString() + "=\"INTEGRADO 3\";SI(E" + xfila.ToString() + ">-5;$ADS$379;$ADS$378))))%;2))";
                    myWorkSheet.Cells[i + 5, 63] = "=PY"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 64] = "=BK"+xfila.ToString()+"+BJ"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 67] = "=SI(CA"+xfila.ToString()+"<>\"MULT-COMP-INT\";ABG244;CONCATENAR(REDONDEAR(X"+xfila.ToString()+"*100/SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 1\";SI(E"+xfila.ToString()+">-5;$ADS$375;$ADS$374);SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 2\";SI(E"+xfila.ToString()+">-5;$ADS$377;$ADS$376);SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 3\";SI(E"+xfila.ToString()+">-5;$ADS$379;$ADS$378))));0);\"%-\";EXTRAE(SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 1\";SI(E"+xfila.ToString()+">-5;$AED$375;$AED$374);SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 2\";SI(E"+xfila.ToString()+">-5;$AED$377;$AED$376);SI(JV"+ x2fila.ToString()+"=\"INTEGRADO 3\";SI(E"+xfila.ToString()+">-5;$AED$379;$AED$378))));1;8)))";
                    myWorkSheet.Cells[i + 5, 70] = "=SI.ERROR(JM"+x2fila.ToString()+";0)";
                    myWorkSheet.Cells[i + 5, 71] = "=REDONDEAR.MAS((BW"+xfila.ToString()+"*BR"+xfila.ToString()+"*0,587)/1000;2)";
                    myWorkSheet.Cells[i + 5, 72] = "=SI.ERROR(REDONDEAR.MAS(BU"+xfila.ToString()+"/BM"+xfila.ToString()+"/1,73;0);0)";
                    myWorkSheet.Cells[i + 5, 73] = "=SI.ERROR(JL"+x2fila.ToString()+"*SI(E"+xfila.ToString()+"<0;\"1\")+JL"+x2fila.ToString()+"*SI(E"+xfila.ToString()+">0;\"0\");0)";
                    myWorkSheet.Cells[i + 5, 74] = "=BR"+xfila.ToString()+"+BL"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 75] = "208";
                    myWorkSheet.Cells[i + 5, 76] = "=AI"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 77] = "=JV"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 81] = "=B"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 82] = "=D"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 83] = "=F"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 84] = "=G"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 85] = "=H"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 86] = "=SI(CE"+xfila.ToString()+">=CF"+xfila.ToString()+";\"1\")*CE"+xfila.ToString()+"+SI(CE"+xfila.ToString()+"<CF"+xfila.ToString()+";\"1\")*CF"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 87] = "=SI(CE"+xfila.ToString()+"<=CF"+xfila.ToString()+";\"1\")*CE"+xfila.ToString()+"+SI(CE"+xfila.ToString()+">CF"+xfila.ToString()+";\"1\")*CF"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 88] = "=REDONDEAR.MAS(CE"+xfila.ToString()+"*CF"+xfila.ToString()+";0)";
                    myWorkSheet.Cells[i + 5, 89] = "=O"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 90] = "MONILE";
                    myWorkSheet.Cells[i + 5, 91] = "=NO(CL"+xfila.ToString()+"<>\"CONCAVOS\")*1+NO(CL"+xfila.ToString()+"<>\"ALUMINIO\")*2+NO(CL"+xfila.ToString()+"<>\"MONILE\")*3+NO(CL"+xfila.ToString()+"<>\"MLE + CONC.\")*4";
                    myWorkSheet.Cells[i + 5, 92] = "1: Unica";
                    myWorkSheet.Cells[i + 5, 93] = "=NO(CN"+xfila.ToString()+"<>\"1: Unica\")*1+NO(CN"+xfila.ToString()+"<>\"2-A:MIXTO\")*2+NO(CN"+xfila.ToString()+"<>\"3-L:MIXTO\")*3+NO(CN6<>\"4-L+ A:MIXTO\")*4+NO(CN"+xfila.ToString()+"<>\"5-2A+ L:MIXTO\")*5+NO(CN6<>\"6-2A+ 2L:MIXTO\")*6+NO(CN"+xfila.ToString()+"<>\"7-A  ABIERTO\")*7+NO(CN"+xfila.ToString()+"<>\"8-L  ABIERTO\")*8+NO(CN"+xfila.ToString()+"<>\"9-2L+ A MIXTO\")*9";
                    myWorkSheet.Cells[i + 5, 94] = "0";
                    myWorkSheet.Cells[i + 5, 97] = "=(NO(CR"+xfila.ToString()+"<>\"Derecha\")*1+NO(CR"+xfila.ToString()+"<>\"Izquierda\")*2+NO(CR"+xfila.ToString()+"<>\"Derecha + Derecha\")*3+NO(CR"+xfila.ToString()+"<>\"Derecha + Izquierda\")*4+NO(CR"+xfila.ToString()+"<>\"Izquierda + Izquierda\")*5)";
                    myWorkSheet.Cells[i + 5, 98] = "=CC6";
                    myWorkSheet.Cells[i + 5, 99] = "=SI(CG"+xfila.ToString()+"<=4;SI(CW"+xfila.ToString()+">100;120;SI(CH"+xfila.ToString()+"*CI"+xfila.ToString()+"<>0;\"1\")*((SI(CK"+xfila.ToString()+">=1;\"80\")+SI(CK"+xfila.ToString()+"<1;\"100\"))*SI(CK"+xfila.ToString()+"<>0;\"1\")));120)";
                    myWorkSheet.Cells[i + 5, 100] = "=CG"+xfila.ToString()+"+0,1";
                    myWorkSheet.Cells[i + 5, 101] = "=SI(ADI"+x2fila.ToString()+"+ADJ"+x2fila.ToString()+"+ADK"+x2fila.ToString()+"=0;SI(CH"+xfila.ToString()+"*CI"+xfila.ToString()+"<>0;\"1\")*((SI(CK"+xfila.ToString()+">=1;\"80\")+SI(CK"+xfila.ToString()+"<1;\"100\"))*SI(CI"+xfila.ToString()+"<=4;\"1\")+SI(CI"+xfila.ToString()+">=7;\"150\";SI(CI"+xfila.ToString()+">=4;\"120\"))))+SI(ADI"+x2fila.ToString()+"=1;100;SI(ADJ"+x2fila.ToString()+"=2;120;SI(ADK"+x2fila.ToString()+"=3;150)))";
                    myWorkSheet.Cells[i + 5, 102] = "=((CU"+xfila.ToString()+"/1000)*2)+CI"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 103] = "=REDONDEAR.MAS(SI(C"+xfila.ToString()+"=1;DL"+xfila.ToString()+"*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=2;(DL"+xfila.ToString()+"-(CI"+xfila.ToString()+"/2))*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=3;(DL"+xfila.ToString()+"-(CH"+xfila.ToString()+"/2))*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=4;(DL"+xfila.ToString()+"-CH"+xfila.ToString()+")*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=5;(DL"+xfila.ToString()+"-CI"+xfila.ToString()+")*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=6;(DL"+xfila.ToString()+"-CH"+xfila.ToString()+"-CI"+xfila.ToString()+")*CG"+xfila.ToString()+";SI(C"+xfila.ToString()+"=7;(DL"+xfila.ToString()+"-CH"+xfila.ToString()+"-CI"+xfila.ToString()+"-CI"+xfila.ToString()+")*CG"+xfila.ToString()+")))))))*CZ"+xfila.ToString()+";0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 104] = "=SI(C"+xfila.ToString()+"=1;1,01;SI(C"+xfila.ToString()+"=2;1,02;SI(C"+xfila.ToString()+"=3;1,04;SI(C"+xfila.ToString()+"=4;1,05;SI(C"+xfila.ToString()+"=5;1,06;SI(C"+xfila.ToString()+"=6;1,06;SI(C"+xfila.ToString()+"=7;1,03)))))))*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 105] = "=REDONDEAR.MAS(SI(CW"+xfila.ToString()+"=150;(CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*1,05;SI(CW"+xfila.ToString()+"=120;(CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*1,05;SI(CW"+xfila.ToString()+"=80;(CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*1,05;SI(CW"+xfila.ToString()+"=100;(CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*1,05))))*1,035;0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 106] = "=REDONDEAR(SI(C"+xfila.ToString()+"=1;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12);SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=2;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=3;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=4;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=5;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=6;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(C"+xfila.ToString()+"=7;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=80;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=100;0));0);SI(CU"+xfila.ToString()+"=0;0))))))))*CZ"+xfila.ToString()+";0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 107] = "=REDONDEAR(SI(C"+xfila.ToString()+"=1;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12);SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=2;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=3;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=4;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=5;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=6;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=7;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=100;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(CU"+xfila.ToString()+"=0;0))))))))*CZ"+xfila.ToString()+";0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 108] = "=REDONDEAR(SI(C"+xfila.ToString()+"=1;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12);SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=2;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=3;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/2/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=4;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CH"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=5;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-CI"+xfila.ToString()+"/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=6;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(C"+xfila.ToString()+"=7;REDONDEAR.MAS(SI(CU"+xfila.ToString()+"=120;((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2/1,12)-(CH"+xfila.ToString()+"+CI"+xfila.ToString()+"+CI"+xfila.ToString()+")/1,12;SI(CU"+xfila.ToString()+"=80;0));0);SI(CU"+xfila.ToString()+"=0;0;SI(CU"+xfila.ToString()+"=80;0;SI(CU"+xfila.ToString()+"=100;0))))))))))*CZ"+xfila.ToString()+";0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 109] = "=REDONDEAR.MAS(SI(CW"+xfila.ToString()+"=80;(CH"+xfila.ToString()+"/1,12))*1,02;0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 110] = "=REDONDEAR.MAS(SI(CW"+xfila.ToString()+"=100;(CH"+xfila.ToString()+"/1,12))*1,02;0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 111] = "=REDONDEAR.MAS(SI(CW"+xfila.ToString()+"=120;(CH"+xfila.ToString()+"/1,12)*1,02;SI(CW"+xfila.ToString()+"=150;(CH"+xfila.ToString()+"/1,16)*1,02));0)*SU"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    //myWorkSheet.Cells[i + 5, 112] = "";
                    //myWorkSheet.Cells[i + 5, 113] = "";
                    //myWorkSheet.Cells[i + 5, 114] = "";
                    myWorkSheet.Cells[i + 5, 115] = "=SI(CH"+xfila.ToString()+"*CI"+xfila.ToString()+"<>0;\"1\")*CI"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 116] = "=((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2)";
                    myWorkSheet.Cells[i + 5, 117] = "=(SI(CH"+xfila.ToString()+"*CI"+xfila.ToString()+"<>0;\"1\")*(((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2*3)*NO(DX"+xfila.ToString()+"<>\"CANAL\")+((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2*6)*NO(DX"+xfila.ToString()+"<>\"ANGULAR\")))*TF"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 118] = "=(SI(CH"+xfila.ToString()+"*CI"+xfila.ToString()+"<>0;\"1\")*((CQ"+xfila.ToString()+"*6)+6*(ED"+xfila.ToString()+"+EE"+xfila.ToString()+")))*TF"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 119] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*((3*EB" + xfila.ToString() + ")+6*(EF" + xfila.ToString() + "+EG" + xfila.ToString() + "+EH" + xfila.ToString() + ")))*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 120] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*(2*(CH" + xfila.ToString() + "+CI" + xfila.ToString() + ")/1,13+0,5))*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 121] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*(((2,5*(900+(2*2000))/1000)*2)*(SI(CS" + xfila.ToString() + "=1;\"1\")+SI(CS" + xfila.ToString() + "=2;\"1\")+SI(CS" + xfila.ToString() + "=5;\"1\")+SI(CS" + xfila.ToString() + "=6;\"1\")+SI(CS" + xfila.ToString() + "=7;\"1\")+SI(CS" + xfila.ToString() + "=11;\"1\")+SI(CS" + xfila.ToString() + "=12;\"1\")+SI(CS" + xfila.ToString() + "=15;\"1\")+SI(CS" + xfila.ToString() + "=16;\"1\")+SI(CS" + xfila.ToString() + "=17;\"1\"))+((2,5*(1200+(2*2000))/1000)*2)*(SI(CS" + xfila.ToString() + "=3;\"1\")+SI(CS" + xfila.ToString() + "=4;\"1\")+SI(CS" + xfila.ToString() + "=8;\"1\")+SI(CS" + xfila.ToString() + "=9;\"1\")+SI(CS" + xfila.ToString() + "=13;\"1\")+SI(CS" + xfila.ToString() + "=14;\"1\")+SI(CS" + xfila.ToString() + "=19;\"1\")+SI(CS" + xfila.ToString() + "=20;\"1\"))))*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 122] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*(((2,5*(900+(2*2000))/1000)*2)*(SI(CS" + xfila.ToString() + "=1;\"1\")+SI(CS" + xfila.ToString() + "=2;\"1\")+SI(CS" + xfila.ToString() + "=5;\"1\")+SI(CS" + xfila.ToString() + "=6;\"1\")+SI(CS" + xfila.ToString() + "=7;\"1\")+SI(CS" + xfila.ToString() + "=11;\"1\")+SI(CS" + xfila.ToString() + "=12;\"1\")+SI(CS" + xfila.ToString() + "=15;\"1\")+SI(CS" + xfila.ToString() + "=16;\"1\")+SI(CS" + xfila.ToString() + "=17;\"1\"))+((2,5*(1200+(2*2000))/1000)*2)*(SI(CS" + xfila.ToString() + "=3;\"1\")+SI(CS" + xfila.ToString() + "=4;\"1\")+SI(CS" + xfila.ToString() + "=8;\"1\")+SI(CS" + xfila.ToString() + "=9;\"1\")+SI(CS" + xfila.ToString() + "=13;\"1\")+SI(CS" + xfila.ToString() + "=14;\"1\")+SI(CS" + xfila.ToString() + "=19;\"1\")+SI(CS" + xfila.ToString() + "=20;\"1\"))))*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 123] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*40*CQ" + xfila.ToString() + ")*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 124] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*50*CQ" + xfila.ToString() + ")*TF" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 125] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*(DM" + xfila.ToString() + "/50+0,5));0)*TF" + x2fila.ToString() + "";
                    myWorkSheet.Cells[i + 5, 126] = "=REDONDEAR.MAS((SI(CH" + xfila.ToString() + "*CI" + xfila.ToString() + "<>0;\"1\")*(((DS" + xfila.ToString() + "+DT" + xfila.ToString() + ")/300)));0)*TF" + x2fila.ToString() + "";
                    myWorkSheet.Cells[i + 5, 127] = "=CT"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 128] = "=SI(CP"+xfila.ToString()+"=0;\"ANGULAR\";SI(CP"+xfila.ToString()+"<>0;\"CANAL\"))";
                    myWorkSheet.Cells[i + 5, 129] = "=SI(CK"+xfila.ToString()+"<0;\"1\")*NO(DX"+xfila.ToString()+"<>\"CANAL\")*(CZ"+xfila.ToString()+"*1,13)*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 130] = "=SI(CK"+xfila.ToString()+">=0;\"1\")*NO(DX"+xfila.ToString()+"<>\"CANAL\")*(CY"+xfila.ToString()+"*1,13)*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 131] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*NO(DX"+xfila.ToString()+"<>\"ANGULAR\")*((DC"+xfila.ToString()+"*1,13)+(DB"+xfila.ToString()+"*1,13))*2*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 132] = "=SI.ERROR(REDONDEAR.MAS(CY" + xfila.ToString() + "*1,12/4;0);0)*TL" + x2fila.ToString() + "";
                    myWorkSheet.Cells[i + 5, 133] = "=4*4*AA"+xfila.ToString()+"*TL"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 134] = "=(F"+xfila.ToString()+"+G"+xfila.ToString()+")*2*TL"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 135] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*2*(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*TL"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 136] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*(SI(CO"+xfila.ToString()+"=1;\"1\")*((CH"+xfila.ToString()+"+CI"+xfila.ToString()+")*2)+SI(CO"+xfila.ToString()+"=2;\"1\")*(2*CH"+xfila.ToString()+"+CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=3;\"1\")*(CH"+xfila.ToString()+"+2*CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=4;\"1\")*(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=5;\"1\")*CH"+xfila.ToString()+"+SI(CO"+xfila.ToString()+"=6;\"1\")*(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=7;\"1\")*(2*CH"+xfila.ToString()+"+CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=8;\"1\")*(CH"+xfila.ToString()+"+CI"+xfila.ToString()+"*2))*ST"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 137] = "=((AA"+xfila.ToString()+"*H"+xfila.ToString()+")+(F"+xfila.ToString()+"+G"+xfila.ToString()+"))*2*ST"+x2fila.ToString()+"*UL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 138] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*(SI(CO"+xfila.ToString()+"=1;\"1\")*(0)+SI(CO"+xfila.ToString()+"=2;\"1\")*(CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=3;\"1\")*(CH"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=4;\"1\")*(CH"+xfila.ToString()+"+CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=5;\"1\")*(2*CI"+xfila.ToString()+"+CH"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=6;\"1\")*(2+CH"+xfila.ToString()+"+2*CI"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=7;\"1\")*(2*CH"+xfila.ToString()+")+SI(CO"+xfila.ToString()+"=8;\"1\")*(2*CI"+xfila.ToString()+"))*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 139] = "=(SI(CO"+xfila.ToString()+"=1;\"1\")*(2*CG"+xfila.ToString()+"/2)+SI(CO"+xfila.ToString()+"=2;\"1\")*(2*CG"+xfila.ToString()+"/2)+SI(CO"+xfila.ToString()+"=3;\"1\")*(2*CG"+xfila.ToString()+"/2)+SI(CO"+xfila.ToString()+"=4;\"1\")*(CG"+xfila.ToString()+"/2)+SI(CO"+xfila.ToString()+"=5;\"1\")*(0)+SI(CO"+xfila.ToString()+"=6;\"1\")*(0)+SI(CO"+xfila.ToString()+"=7;\"1\")*(0)+SI(CO"+xfila.ToString()+"=8;\"1\")*(0))*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 140] = "=REDONDEAR.MAS((EI" + xfila.ToString() + "/3)*TL" + x2fila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 141] = "=AA"+xfila.ToString()+"*8*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 142] = "=AA"+xfila.ToString()+"*4*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 143] = "=AA"+xfila.ToString()+"*4*TL"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 144] = "=AA6*4*TL130";
                    myWorkSheet.Cells[i + 5, 147] = "=ER"+xfila.ToString()+"/2*AA"+xfila.ToString()+"*TA"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 148] = "=1*AA"+xfila.ToString()+"*TA"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 149] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*(ER"+xfila.ToString()+"/50)*TA"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 150] = "=NO(CL"+xfila.ToString()+"=\"ENCABEZADO\")*(ER"+xfila.ToString()+"/100)*TA"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 151] = "=REDONDEAR.MAS(NO(CL" + xfila.ToString() + "=\"ENCABEZADO\")*(EQ" + xfila.ToString() + "/1,5);0)";
                    myWorkSheet.Cells[i + 5, 152] = "=SI(MW"+x2fila.ToString()+"=2;\"1\"*CQ"+xfila.ToString()+";SI(MW"+x2fila.ToString()+"=1;\"0\"))*AA"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 153] = "=2*AA"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 154] = "=SI((CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*10/90>1;REDONDEA.PAR((CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*10/90);SI((CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*10/90<=1;REDONDEAR.MAS((CH"+xfila.ToString()+"*CI"+xfila.ToString()+")*10/90;0)))*TN"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 155] = "UC:380V/3";
                    myWorkSheet.Cells[i + 5, 160] = "=EO"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 164] = "=REDONDEAR.MAS(M" + xfila.ToString() + "*SW" + x2fila.ToString() + ";1)";
                    myWorkSheet.Cells[i + 5, 165] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "/9;1)";
                    myWorkSheet.Cells[i + 5, 166] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "*6/9;1)";
                    myWorkSheet.Cells[i + 5, 167] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "*6/9;1)";
                    myWorkSheet.Cells[i + 5, 168] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "*2/9;1)";
                    myWorkSheet.Cells[i + 5, 169] = "=REDONDEAR.MAS(((F" + xfila.ToString() + "+G" + xfila.ToString() + ")*2*0,1)*SW" + x2fila.ToString() + ";1)";
                    myWorkSheet.Cells[i + 5, 170] = "=REDONDEAR.MAS(FM" + xfila.ToString() + "/9;1)";
                    myWorkSheet.Cells[i + 5, 171] = "=REDONDEAR.MAS(FM" + xfila.ToString() + "*6/9;1)";
                    myWorkSheet.Cells[i + 5, 172] = "=REDONDEAR.MAS(FM" + xfila.ToString() + "*6/9;1)";
                    myWorkSheet.Cells[i + 5, 173] = "=REDONDEAR.MAS(FM" + xfila.ToString() + "*2/9;1)";
                    myWorkSheet.Cells[i + 5, 174] = "=REDONDEAR.MAS((FH" + xfila.ToString() + "/80)*SW" + x2fila.ToString() + ";1)";
                    myWorkSheet.Cells[i + 5, 175] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "/90;1)";
                    myWorkSheet.Cells[i + 5, 176] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "*6/90;1)";
                    myWorkSheet.Cells[i + 5, 178] = "=AK"+xfila.ToString()+"*2*AA"+xfila.ToString()+"*TI"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 180] = "=REDONDEAR.MAS(((DB"+xfila.ToString()+"*(CV"+xfila.ToString()+"+1,12)*2)+(DC"+xfila.ToString()+"*(CV"+xfila.ToString()+"+1,12)*2)+(DG"+xfila.ToString()+"*(CX"+xfila.ToString()+"+1,12)*2))/15;0)";
                    myWorkSheet.Cells[i + 5, 181] = "=REDONDEAR.MAS(((DB"+xfila.ToString()+"*CV"+xfila.ToString()+")+(DC"+xfila.ToString()+"*CV"+xfila.ToString()+")+(DE"+xfila.ToString()+"*CX"+xfila.ToString()+")+(DF"+xfila.ToString()+"*CX"+xfila.ToString()+")+(DG"+xfila.ToString()+"*CX"+xfila.ToString()+"))/9;0)";
                    myWorkSheet.Cells[i + 5, 182] = "=FY"+xfila.ToString()+"/10";
                    myWorkSheet.Cells[i + 5, 183] = "=Y(O"+xfila.ToString()+"<=-15;O6>-40)*1*TC"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 185] = "=SI(E"+xfila.ToString()+"<=-1;SI((3,66*GB"+xfila.ToString()+"*SI(E"+xfila.ToString()+"<(-10);2;1))<250;1))*TJ"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 186] = "=SI(E"+xfila.ToString()+"<=-1;SI(3,66*GB"+xfila.ToString()+"*SI(E"+xfila.ToString()+"<(-10);2;1)>=650;0;SI((3,66*GB"+xfila.ToString()+"*SI(E"+xfila.ToString()+"<(-10);2;1))>250;1)))*TJ"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 187] = "=SI(E"+xfila.ToString()+"<=-1;SI(3,66*GB"+xfila.ToString()+"*SI(E"+xfila.ToString()+"<(-10);2;1)>4300;2;SI((3,66*GB"+xfila.ToString()+"*SI(E"+xfila.ToString()+"<(-10);2;1))>650;1)))*TJ"+x2fila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 188] = "=FF"+xfila.ToString()+"*FG"+xfila.ToString()+"+(CU"+xfila.ToString()+"/1000*FF"+xfila.ToString()+")+(CU"+xfila.ToString()+"/1000*FG"+xfila.ToString()+")";
                    myWorkSheet.Cells[i + 5, 189] = "=REDONDEAR.MAS(N" + xfila.ToString() + "*(CU" + xfila.ToString() + "/1000);0)";
                    myWorkSheet.Cells[i + 5, 190] = "=REDONDEAR.MAS(((0,000004*(SUMA(DM" + xfila.ToString() + ":DV" + xfila.ToString() + ")+EJ" + xfila.ToString() + "))+(DY" + xfila.ToString() + "*(0,16*2,5*0,04)+DZ" + xfila.ToString() + "*(0,16*2,5*0,04)+EA" + xfila.ToString() + "*(0,16*2,5*0,04)+EB" + xfila.ToString() + "*(0,16*2,5*0,04)+EC" + xfila.ToString() + "*(0,16*2,5*0,04)+ED" + xfila.ToString() + "*(0,16*2,5*0,04)+EE" + xfila.ToString() + "*(0,16*2,5*0,04)+EF" + xfila.ToString() + "*(0,16*2,5*0,04)+EG" + xfila.ToString() + "*(0,16*2,5*0,04)+EH" + xfila.ToString() + "*(0,16*2,5*0,04)))*AA" + xfila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 191] = "=REDONDEAR.MAS(FH" + xfila.ToString() + "/5;0)";
                    myWorkSheet.Cells[i + 5, 192] = "=REDONDEAR.MAS(((CQ" + xfila.ToString() + "*0,9*2*0,12)+(0,62*(AA" + xfila.ToString() + "+AK" + xfila.ToString() + "))*2)*AA" + xfila.ToString() + ";0)";
                    myWorkSheet.Cells[i + 5, 193] = "=REDONDEAR.MAS((GG6+GH6+GI6+GJ6)*AA6*1,3;0)";
                    myWorkSheet.Cells[i + 5, 194] = "=B"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 195] = "=D"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 196] = "=E"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 197] = "=X"+xfila.ToString()+"/1000";
                    myWorkSheet.Cells[i + 5, 198] = "=AG"+xfila.ToString()+"";
                    myWorkSheet.Cells[i + 5, 199] = "=SI(SI.ERROR(QZ"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(QZ"+x2fila.ToString()+";3);0);SI(SI.ERROR(QZ"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 200] = "=SI(SI.ERROR(RA"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RA"+x2fila.ToString()+";3);0);SI(SI.ERROR(RA"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 201] = "=SI(SI.ERROR(RB"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RB"+x2fila.ToString()+";3);0);SI(SI.ERROR(RB"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 202] = "=SI(SI.ERROR(RC"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RC"+x2fila.ToString()+";3);0);SI(SI.ERROR(RC"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 203] = "=SI(SI.ERROR(RD"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RD"+x2fila.ToString()+";3);0);SI(SI.ERROR(RD"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 204] = "=SI(SI.ERROR(RE"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RE"+x2fila.ToString()+";3);0);SI(SI.ERROR(RE"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 205] = "=SI(SI.ERROR(RG"+x2fila.ToString()+";0)<>0;SI.ERROR(REDONDEAR(RG"+x2fila.ToString()+";3);0);SI(SI.ERROR(RG"+x2fila.ToString()+";0)=0;0))";
                    myWorkSheet.Cells[i + 5, 206] = "=SI(SI.ERROR(SUMA(GQ"+xfila.ToString()+":GW"+xfila.ToString()+");0)<>0;SI.ERROR(SUMA(GQ"+xfila.ToString()+":GW"+xfila.ToString()+");0);SI(SI.ERROR(SUMA(GQ"+xfila.ToString()+":GW"+xfila.ToString()+");0)=0;0))";
                    */
                    xfila++;
                    x2fila++;
                    x3fila++;
                    x4fila++;
                    percent += perby;
                    thisForm.LEstado.Text = "Actualizando base de datos... " + percent.ToString() + "%";
                    
                   // ************************************* 03/02/2014 *****************************
                    
                    myWorkSheet.Cells[i + 5, 4] = myCam.GetNC();
                    myWorkSheet.Cells[i + 5, 5] = myCam.GetTemp();
                    myWorkSheet.Cells[i + 5, 6] = myCam.GetLargo();
                    myWorkSheet.Cells[i + 5, 7] = myCam.GetAncho();
                    myWorkSheet.Cells[i + 5, 8] = myCam.GetAlto();
                    myWorkSheet.Cells[i + 5, 9] = myCam.GetCF();
                    myWorkSheet.Cells[i + 5, 10] = myCam.GetSpsi();
                    myWorkSheet.Cells[i + 5, 11] = myCam.GetStemp();
                    myWorkSheet.Cells[i + 5, 12] = myCam.GetApsi();
                    myWorkSheet.Cells[i + 5, 15] = myCam.GetTIpv();
                    myWorkSheet.Cells[i + 5, 20] = myCam.GetTIcc();
                    myWorkSheet.Cells[i + 5, 19] = myCam.GetEmevp();
                    myWorkSheet.Cells[i + 5, 31] = myCam.GetDT();
                    myWorkSheet.Cells[i + 5, 54] = myCam.GetCE();
                    myWorkSheet.Cells[i + 129, 701] = myCam.GetCE();
                    myWorkSheet.Cells[i + 389, 737] = myCam.GetDEC();
                    myWorkSheet.Cells[i + 129, 365] = myCam.GetDECE();
                    myWorkSheet.Cells[i + 389, 738] = myCam.GetDECH();
                    myWorkSheet.Cells[i + 389, 743] = myCam.GetDECF();
                    myWorkSheet.Cells[i + 5, 55] = myCam.GetSUP();
                    myWorkSheet.Cells[i + 5, 56] = myCam.GetIT();
                    myWorkSheet.Cells[i + 5, 37] = myCam.GetCantEv();                                 
                    myWorkSheet.Cells[i + 129, 511] = myCam.GetCentx();
                    myWorkSheet.Cells[i + 129, 512] = myCam.GetCdin();
                    myWorkSheet.Cells[i + 129, 513] = myCam.GetCxp();
                    myWorkSheet.Cells[i + 129, 416] = myCam.GetCTCond();
                    myWorkSheet.Cells[i + 129, 430] = myCam.GetCTamb();
                    myWorkSheet.Cells[i + 5, 22] = myCam.GetConD();
                    myWorkSheet.Cells[i + 5, 96] = myCam.GetTP();
                    myWorkSheet.Cells[i + 5, 58] = myCam.GetRefrig();
                    myWorkSheet.Cells[i + 5, 80] = myCam.GetCBint();
                    myWorkSheet.Cells[i + 129, 559] = myCam.GetCBexpo();
                    myWorkSheet.Cells[i + 129, 499] = myCam.GetCevap();
                    myWorkSheet.Cells[i + 129, 501] = myCam.GetCmodex();
                    myWorkSheet.Cells[i + 129, 550] = myCam.GetCtxv();
                    myWorkSheet.Cells[i + 189, 336] = myCam.GetCTEvap();
                    myWorkSheet.Cells[i + 129, 506] = myCam.GetCtpd();
                    myWorkSheet.Cells[i + 129, 503] = myCam.GetCnoff1();
                    myWorkSheet.Cells[i + 129, 504] = myCam.GetCoff1();
                    myWorkSheet.Cells[i + 129, 780] = myCam.GetCnoff2();
                    myWorkSheet.Cells[i + 129, 779] = myCam.GetCoff2();
                    myWorkSheet.Cells[i + 129, 709] = myCam.GetCnoff3();
                    myWorkSheet.Cells[i + 129, 710] = myCam.GetCoff3();
                    myWorkSheet.Cells[i + 129, 703] = myCam.GetCnoff4();
                    myWorkSheet.Cells[i + 129, 704] = myCam.GetCoff4();
                    myWorkSheet.Cells[i + 129, 705] = myCam.GetCnoff5();
                    myWorkSheet.Cells[i + 129, 706] = myCam.GetCoff5();
                    myWorkSheet.Cells[i + 129, 707] = myCam.GetCnoff6();
                    myWorkSheet.Cells[i + 129, 708] = myCam.GetCoff6();
                    myWorkSheet.Cells[i + 129, 719] = myCam.GetCnoff7();
                    myWorkSheet.Cells[i + 129, 720] = myCam.GetCoff7();
                    myWorkSheet.Cells[i + 5, 794] = myCam.GetTCp80();
                    myWorkSheet.Cells[i + 5, 795] = myCam.GetTCp80m(); 
                    myWorkSheet.Cells[i + 5, 790] = myCam.GetTCp100();
                    myWorkSheet.Cells[i + 5, 791] = myCam.GetTCp100m();
                    myWorkSheet.Cells[i + 5, 786] = myCam.GetTCp120();
                    myWorkSheet.Cells[i + 5, 787] = myCam.GetTCp120m();
                    myWorkSheet.Cells[i + 5, 711] = myCam.GetTCp150();
                    myWorkSheet.Cells[i + 5, 712] = myCam.GetTCp150m();
                   
                    if(myCam.GetKppc() == false)
                    {
                        myWorkSheet.Cells[i + 5, 796] = myCam.GetTCt80();
                        myWorkSheet.Cells[i + 5, 796] = myCam.GetTCt80();
                    }
                    else
                    {
                        myWorkSheet.Cells[i + 5, 709] = myCam.GetTCt80();
                        myWorkSheet.Cells[i + 5, 710] = myCam.GetTCt80();
                    }
                    
                    
                    myWorkSheet.Cells[i + 5, 797] = myCam.GetTCt80m();
                    myWorkSheet.Cells[i + 5, 792] = myCam.GetTCt100();
                    myWorkSheet.Cells[i + 5, 793] = myCam.GetTCt100m();
                    myWorkSheet.Cells[i + 5, 788] = myCam.GetTCt120();
                    myWorkSheet.Cells[i + 5, 789] = myCam.GetTCt120m();
                    myWorkSheet.Cells[i + 5, 784] = myCam.GetTCt150();
                    myWorkSheet.Cells[i + 5, 785] = myCam.GetTCt150m();
                    myWorkSheet.Cells[i + 5, 780] = myCam.GetSPtp84();
                    myWorkSheet.Cells[i + 5, 779] = myCam.GetSPtp83();
                    myWorkSheet.Cells[i + 5, 778] = myCam.GetSPtp78();
                    myWorkSheet.Cells[i + 5, 777] = myCam.GetSPtp76();
                    myWorkSheet.Cells[i + 5, 776] = myCam.GetSPtp75();
                    myWorkSheet.Cells[i + 5, 775] = myCam.GetSPtp74();
                    myWorkSheet.Cells[i + 5, 774] = myCam.GetSPtp73();
                    myWorkSheet.Cells[i + 5, 773] = myCam.GetSPtp72();
                    myWorkSheet.Cells[i + 5, 772] = myCam.GetSPtp85();
                    myWorkSheet.Cells[i + 5, 771] = myCam.GetSPtp86();
                    myWorkSheet.Cells[i + 5, 770] = myCam.GetSPtp87();
                    myWorkSheet.Cells[i + 5, 769] = myCam.GetSPtp88();
                    myWorkSheet.Cells[i + 5, 768] = myCam.GetSPtp89();
                    myWorkSheet.Cells[i + 5, 767] = myCam.GetSPtp90();
                    myWorkSheet.Cells[i + 5, 766] = myCam.GetSPtp91();
                    myWorkSheet.Cells[i + 5, 765] = myCam.GetSPtp92();
                    myWorkSheet.Cells[i + 5, 764] = myCam.GetSPtp93();
                    myWorkSheet.Cells[i + 5, 763] = myCam.GetSPtp94();
                    myWorkSheet.Cells[i + 5, 762] = myCam.GetSPtp95();
                    myWorkSheet.Cells[i + 5, 761] = myCam.GetSPtp96();
                    myWorkSheet.Cells[i + 5, 760] = myCam.GetSPtp97();
                    myWorkSheet.Cells[i + 5, 759] = myCam.GetSPtp98();
                    myWorkSheet.Cells[i + 5, 758] = myCam.GetSPtp99();
                    myWorkSheet.Cells[i + 5, 757] = myCam.GetSPtp100();
                    myWorkSheet.Cells[i + 5, 756] = myCam.GetSPtp101();
                    myWorkSheet.Cells[i + 5, 755] = myCam.GetSPtp102();
                    myWorkSheet.Cells[i + 5, 754] = myCam.GetSPtp103();
                    myWorkSheet.Cells[i + 5, 753] = myCam.GetSPtp104();
                    myWorkSheet.Cells[i + 5, 751] = myCam.GetSPtp105();
                    myWorkSheet.Cells[i + 5, 750] = myCam.GetSPtp106();
                    myWorkSheet.Cells[i + 5, 749] = myCam.GetSPtp107();
                    myWorkSheet.Cells[i + 5, 748] = myCam.GetSPtp108();
                    myWorkSheet.Cells[i + 5, 747] = myCam.GetSPtp109();
                    myWorkSheet.Cells[i + 5, 746] = myCam.GetSPtp110();
                    myWorkSheet.Cells[i + 5, 745] = myCam.GetSPtp111();
                    myWorkSheet.Cells[i + 5, 718] = myCam.GetSPtp136();
                    myWorkSheet.Cells[i + 5, 717] = myCam.GetSPtp137();
                    myWorkSheet.Cells[i + 5, 713] = myCam.GetSPtp141();
                    myWorkSheet.Cells[i + 5, 714] = myCam.GetSPtp142();
                    myWorkSheet.Cells[i + 5, 705] = myCam.GetSPtp143();
                    myWorkSheet.Cells[i + 5, 706] = myCam.GetSPtp144();
                    myWorkSheet.Cells[i + 5, 707] = myCam.GetSPtp145();
                    myWorkSheet.Cells[i + 5, 708] = myCam.GetSPtp146();
                    myWorkSheet.Cells[i + 5, 889] = myCam.GetSPtp147();
                    myWorkSheet.Cells[i + 5, 891] = myCam.GetSPtp148();
                    myWorkSheet.Cells[i + 5, 867] = myCam.GetTPcq();
                    myWorkSheet.Cells[i + 5, 808] = myCam.GetTPcv();
                    myWorkSheet.Cells[i + 5, 861] = myCam.GetTPcx();
                    myWorkSheet.Cells[i + 5, 802] = myCam.GetTPcy();
                    myWorkSheet.Cells[i + 5, 803] = myCam.GetTPex();
                    myWorkSheet.Cells[i + 5, 804] = myCam.GetTPrs();
                    myWorkSheet.Cells[i + 5, 804] = myCam.GetTPrs();
                    myWorkSheet.Cells[i + 5, 804] = myCam.GetTPrs();
                    myWorkSheet.Cells[i + 5, 805] = myCam.GetTPem();
                    myWorkSheet.Cells[i + 5, 806] = myCam.GetTPnt();
                    myWorkSheet.Cells[i + 5, 807] = myCam.GetTPml();
                    myWorkSheet.Cells[i + 5, 798] = myCam.GetTPdtr();
                    myWorkSheet.Cells[i + 5, 799] = myCam.GetTPdtc();
                    myWorkSheet.Cells[i + 5, 800] = myCam.GetTPdt2();
                    myWorkSheet.Cells[i + 5, 801] = myCam.GetTPdt1();
                    myWorkSheet.Cells[i + 5, 849] = myCam.GetTPlq();
                    myWorkSheet.Cells[i + 5, 851] = myCam.GetTPvq();
                    myWorkSheet.Cells[i + 5, 853] = myCam.GetTPls();
                    myWorkSheet.Cells[i + 5, 855] = myCam.GetTPosc();
                    myWorkSheet.Cells[i + 5, 865] = myCam.GetTPsq();
                    myWorkSheet.Cells[i + 5, 845] = myCam.GetTQevp();
                    myWorkSheet.Cells[i + 5, 843] = myCam.GetTTint();
                    myWorkSheet.Cells[i + 5, 847] = myCam.GetTEquip();
                    myWorkSheet.Cells[i + 5, 837] = myCam.GetTCmce();
                    myWorkSheet.Cells[i + 5, 838] = myCam.GetTPmce();
                    myWorkSheet.Cells[i + 5, 836] = myCam.GetTDmce();
                    myWorkSheet.Cells[i + 5, 823] = myCam.GetCantEv();
                    myWorkSheet.Cells[i + 5, 824] = myCam.GetDECF();// LN LINEA DE LIQUIDO -30°C
                    myWorkSheet.Cells[i + 5, 825] = myCam.GetTDmc11();// LN LINE DE SUCCION -10°C
                    myWorkSheet.Cells[i + 5, 834] = myCam.GetTDmc12();// LN LINE DE SUCCION +5°C
                    myWorkSheet.Cells[i + 5, 826] = myCam.GetTDmc8();// DN LINEA DE SUCCION -10°C
                    myWorkSheet.Cells[i + 5, 827] = myCam.GetTDmc1();// LN LINEA DE SUCCION -30°C
                    myWorkSheet.Cells[i + 5, 828] = myCam.GetTDmc6();// DN LINEA DE SUCCION -30°C
                    myWorkSheet.Cells[i + 5, 829] = myCam.GetTDmc3();// DN LINEA DE LIQUIDO -30°C
                    myWorkSheet.Cells[i + 5, 830] = myCam.GetTDmc2();// LN LIQUIDO + SUCCION EVAPORADORES
                    myWorkSheet.Cells[i + 5, 831] = myCam.GetTDmc5();// DN EVAPORADOR LIQUIDO
                    myWorkSheet.Cells[i + 5, 832] = myCam.GetTDmc4(); // DN EVAPORADOR SUCCION
                    myWorkSheet.Cells[i + 5, 833] = myCam.GetTDmc7();// DN LINEA DE SUCCION +5°C
                    myWorkSheet.Cells[i + 5, 821] = myCam.GetTLcc();
                    myWorkSheet.Cells[i + 5, 822] = myCam.GetTLss();
                    myWorkSheet.Cells[i + 5, 818] = myCam.GetTIned();
                    myWorkSheet.Cells[i + 5, 752] = myCam.GetTIncd();
                    myWorkSheet.Cells[i + 5, 819] = myCam.GetTInev();
                    myWorkSheet.Cells[i + 5, 820] = myCam.GetTInc();
                    myWorkSheet.Cells[i + 5, 817] = myCam.GetCSumi();
                    myWorkSheet.Cells[i + 5, 816] = myCam.GetDECF();
                    myWorkSheet.Cells[i + 5, 815] = myCam.GetDECH();
                    myWorkSheet.Cells[i + 5, 814] = myCam.GetDECE();
                    myWorkSheet.Cells[i + 5, 813] = myCam.GetDEC();
                    myWorkSheet.Cells[i + 5, 812] = myCam.GetVol();
                    myWorkSheet.Cells[i + 5, 811] = myCam.GetTemp();
                    myWorkSheet.Cells[i + 5, 810] = myCam.GetSUP();
                    myWorkSheet.Cells[i + 5, 716] = myCam.GetQfw();
                    myWorkSheet.Cells[i + 5, 877] = myCam.GetTCsist();
                    myWorkSheet.Cells[i + 465, 715] = myCam.GetCBcm();
                    myWorkSheet.Cells[i + 465, 716] = myCam.GetCCmci();
                    myWorkSheet.Cells[i + 5, 894] = myCam.GetTQevpd();
                    myWorkSheet.Cells[i + 5, 895] = myCam.GetTQevpc();
                    myWorkSheet.Cells[i + 5, 908] = myCam.GetTDlq1();
                    myWorkSheet.Cells[i + 5, 909] = myCam.GetTDlq2();
                    myWorkSheet.Cells[i + 5, 910] = myCam.GetTDlq3();
                    myWorkSheet.Cells[i + 5, 896] = myCam.GetTDlq11();//Liquido -30°C T1:
                    myWorkSheet.Cells[i + 5, 897] = myCam.GetTDsu130();//Succion -30°C T1:
                    myWorkSheet.Cells[i + 5, 898] = myCam.GetTDsu105();//Succiön +5°C T1:
                    myWorkSheet.Cells[i + 5, 899] = myCam.GetTDsu110();//Succión -10°C T1:
                    myWorkSheet.Cells[i + 5, 900] = myCam.GetTDlq21();//Liquido -30°C T2:
                    myWorkSheet.Cells[i + 5, 901] = myCam.GetTDsu230();//Succion -30°C T2:
                    myWorkSheet.Cells[i + 5, 902] = myCam.GetTDsu205();//Succiön +5°C T12
                    myWorkSheet.Cells[i + 5, 903] = myCam.GetTDsu210();//Succión -10°C T2:
                    myWorkSheet.Cells[i + 5, 904] = myCam.GetTDlq31();//Liquido -30°C T3:
                    myWorkSheet.Cells[i + 5, 905] = myCam.GetTDsu330();//Succion -30°C T3:
                    myWorkSheet.Cells[i + 5, 906] = myCam.GetTDsu305();//Succiön +5°C T3:
                    myWorkSheet.Cells[i + 5, 907] = myCam.GetTDsu310();//Succión -10°C T3:

                    //*************************************************************************************************************
                    myWorkSheet.Cells[i + 5, 911] = myCam.GetTIn1();//TIn1 (165, 1 ) // Abrasaderas sifonicas LIQ LIN- CTRAL.
                    myWorkSheet.Cells[i + 5, 912] = myCam.GetTIn2();//TIn2 (167, 1 ) // Abrasaderas sifonicas SUCC LIN- CTRAL
                    myWorkSheet.Cells[i + 5, 913] = myCam.GetTIn3();//TIn3 (168, 1 ) // Abrasaderas sifonicas SUCC 1er EVAP.
                    myWorkSheet.Cells[i + 5, 914] = myCam.GetTIn4();//TIn4 (169, 1 ) //Abrasaderas sifonicas SUCC 2do EVAP
                    myWorkSheet.Cells[i + 5, 915] = myCam.GetTIn5();//TIn5 (170, 1 ) //Abrasaderas sifonicas SUCC 3er EVAP.
                    myWorkSheet.Cells[i + 5, 916] = myCam.GetTIn6();//TIn6 (171, 1 ) //Abrasaderas sifonicas SUCC LIN T1 (-30°C ) .
                    myWorkSheet.Cells[i + 5, 917] = myCam.GetTIn7();//TIn7 (172, 1 ) //Abrasaderas sifonicas SUCC LIN T2  (-30°C ) .
                    myWorkSheet.Cells[i + 5, 918] = myCam.GetTIn8();//TIn8 (173, 1 ) // Abrasaderas sifonicas SUCC LIN T3 (-30°C )
                    myWorkSheet.Cells[i + 5, 919] = myCam.GetTIn9();//TIn9 (174, 1 ) // Abrasaderas sifonicas SUCC LIN T1 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 920] = myCam.GetTIn10();//TIn10 (175, 1 ) // Abrasaderas sifonicas SUCC LIN T2 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 921] = myCam.GetTIn11();//TIn11 (176, 1 ) // Abrasaderas sifonicas SUCC LIN T3 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 922] = myCam.GetTIn12();//TIn12 (177, 1 ) // Abrasaderas sifonicas SUCC LIN T1 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 923] = myCam.GetTIn13();//TIn13 (178, 1 ) // Abrasaderas sifonicas SUCC LIN T2 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 924] = myCam.GetTIn14();//TIn14 (179, 1 ) // Abrasaderas sifonicas SUCC LIN T3 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 925] = myCam.GetTIn15();//TIn15 (193, 1 ) // BARILLA ROSCADA M10-M8
                    myWorkSheet.Cells[i + 5, 926] = myCam.GetTIn16();//TIn16 (194, 1 ) // TUERCAS M10-M8
                    myWorkSheet.Cells[i + 5, 927] = myCam.GetTIn17();//TIn17 (195, 1 ) // EXPANCIONES M10-M8
                    myWorkSheet.Cells[i + 5, 928] = myCam.GetTIn18();//TIn18 (196, 1 ) // ARANDELAS CUAD M10-M8
                    myWorkSheet.Cells[i + 5, 929] = myCam.GetTIn19();//TIn19 (197, 1 ) // ARANDELAS M10-M8
                    myWorkSheet.Cells[i + 5, 930] = myCam.GetTIn20();//TIn20 (198, 1 ) // PERFIL DE CARGA
                    myWorkSheet.Cells[i + 5, 931] = myCam.GetTIn21();//TIn21 (205, 1 ) // BANDEJA DE 300MM
                    myWorkSheet.Cells[i + 5, 932] = myCam.GetTIn22();//TIn22 (206, 1 ) // TAPA BANDEJA 300MM
                    myWorkSheet.Cells[i + 5, 954] = myCam.GetTIn23();//TIn23 (208, 1 ) // Abrasadera LIQ LIN EVP

                    myWorkSheet.Cells[i + 5, 933] = myCam.GetTIp1();//TIp1 (166, 1 ) // CANTIDAD Abrasaderas sifonicas LIQ LIN- CTRAL.
                    myWorkSheet.Cells[i + 5, 934] = myCam.GetTIp2();//TIp2 (180, 1 ) //Cantidad de abrasaderas sifonicas SUCC LIN- CTRAL.
                    myWorkSheet.Cells[i + 5, 935] = myCam.GetTIp3();//TIp3 (181, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC 1er EVAP.
                    myWorkSheet.Cells[i + 5, 936] = myCam.GetTIp4();//TIp4 (182, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC 2do EVAP
                    myWorkSheet.Cells[i + 5, 937] = myCam.GetTIp5();//TIp5 (183, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC 3er EVAP.
                    myWorkSheet.Cells[i + 5, 938] = myCam.GetTIp6();//TIp6 (184, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T1 (-30°C ) .
                    myWorkSheet.Cells[i + 5, 939] = myCam.GetTIp7();//TIp7 (185, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T2  (-30°C ) .
                    myWorkSheet.Cells[i + 5, 940] = myCam.GetTIp8();//TIp8 (186, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T3 (-30°C ) .
                    myWorkSheet.Cells[i + 5, 941] = myCam.GetTIp9();// TIp9 (187, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T1 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 942] = myCam.GetTIp10();//TIp10 (188, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T3 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 943] = myCam.GetTIp11();//TIp11 (189, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T3 (-10°C ) .
                    myWorkSheet.Cells[i + 5, 944] = myCam.GetTIp12();//TIp12 (190, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T1 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 945] = myCam.GetTIp13();//TIp13 (191, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T2 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 946] = myCam.GetTIp14();//TIp14 (192, 1 ) // CANTIDAD Abrasaderas sifonicas SUCC LIN T3 ( 5°C ) .
                    myWorkSheet.Cells[i + 5, 947] = myCam.GetTIp15();//TIp15 (199, 1 ) // CANTIDAD BARILLA ROSCADA M10-M8
                    myWorkSheet.Cells[i + 5, 948] = myCam.GetTIp16();//TIp16 (200, 1 ) // CANTIDAD TUERCAS M10-M8
                    myWorkSheet.Cells[i + 5, 949] = myCam.GetTIp17();//TIp17 (201, 1 ) // CANTIDAD EXPANCIONES M10-M8
                    myWorkSheet.Cells[i + 5, 950] = myCam.GetTIp18();//TIp18 (202, 1 ) // CANTIDAD ARANDELAS CUAD M10-M8
                    myWorkSheet.Cells[i + 5, 951] = myCam.GetTIp19();//TIp19 (203, 1 ) // CANTIDAD ARANDELAS M10-M8
                    myWorkSheet.Cells[i + 5, 952] = myCam.GetTIp20();//TIp20 (204, 1 ) // CANTIDAD PERFIL DE CARGA
                    myWorkSheet.Cells[i + 5, 953] = myCam.GetTIp21();//TIp21 (207, 1 ) // CANTIDAD BANDEJA DE 300MM + TAPA BANDEJA 300MM
                    myWorkSheet.Cells[i + 5, 955] = myCam.GetTIp21();//TIp23 (209, 1 ) // CANTIDAD Abrasadera LIQ LIN EVP

                    if (myCam.GetKmod() == true)
                    {
                        myWorkSheet.Cells[i + 129, 728] = myCam.GetCmod();
                        myWorkSheet.Cells[i + 129, 729] = myCam.GetCmodd();
                        myWorkSheet.Cells[i + 129, 730] = myCam.GetCmodp();
                        myWorkSheet.Cells[i + 190, 740] = myCam.GetCmod();
                        myWorkSheet.Cells[i + 190, 741] = myCam.GetDesc();
                        myWorkSheet.Cells[i + 190, 742] = myCam.GetPrec();
                    }
                    myWorkSheet.Cells[i + 129, 723] = myCam.GetCnoff8();
                    myWorkSheet.Cells[i + 129, 724] = myCam.GetCoff8();
                    myWorkSheet.Cells[i + 5, 79] = myCam.GetCSumi();
                    myWorkSheet.Cells[i + 5, 40] = myCam.GetTValv();
                    myWorkSheet.Cells[i + 5, 42] = myCam.GetCodValv();
                    myWorkSheet.Cells[i + 5, 857] = myCam.GetCodValv();
                    myWorkSheet.Cells[i + 5, 859] = myCam.GetTValv();
                    myWorkSheet.Cells[i + 5, 29] = myCam.GetTCvta();
                    myWorkSheet.Cells[i + 5, 834] = myCam.GetTmos();

                    if (myCam.GetCSumi()=="DORIN")
                        myWorkSheet.Cells[i + 5, 62] = myCam.GetTInc();
                        myWorkSheet.Cells[i + 5, 70] = myCam.GetTInev();
                        myWorkSheet.Cells[i + 5, 72] = myCam.GetTIned();
                        myWorkSheet.Cells[i + 5, 63] = myCam.GetTIncd();

                    if (myCam.GetIE())
                        myWorkSheet.Cells[i + 5, 28] = "I";
                    else
                        myWorkSheet.Cells[i + 5, 28] = "E";
                    if (myCam.GetKP())
                        myWorkSheet.Cells[i + 129, 515] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 515] = "0";

                    if (myCam.GetKN())
                        myWorkSheet.Cells[i + 129, 516] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 516] = "0";

                    if (myCam.GetKM())
                        myWorkSheet.Cells[i + 129, 517] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 517] = "0";

                    if (myCam.GetKC())
                        myWorkSheet.Cells[i + 129, 518] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 518] = "0";

                    if (myCam.GetKU())
                        myWorkSheet.Cells[i + 129, 519] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 519] = "0";

                    if (myCam.GetKPR())
                        myWorkSheet.Cells[i + 129, 520] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 520] = "0";

                    if(myCam.GetKexpo())
                        myWorkSheet.Cells[i + 129, 558] = "0";
                    else
                        myWorkSheet.Cells[i + 129, 558] = "1";

                    if (myCam.GetKepiso())
                        myWorkSheet.Cells[i + 129, 561] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 561] = "0";

                    if (myCam.GetKD())
                        myWorkSheet.Cells[i + 129, 521] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 521] = "0";

                    if (myCam.GetKS())
                        myWorkSheet.Cells[i + 129, 522] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 522] = "0";

                    if (myCam.GetKCE())
                        myWorkSheet.Cells[i + 129, 523] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 523] = "0";

                    if (myCam.GetKB())
                        myWorkSheet.Cells[i + 129, 524] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 524] = "0";

                    if (myCam.GetKCO())
                        myWorkSheet.Cells[i + 129, 525] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 525] = "0";

                    if (myCam.GetKTO())
                        myWorkSheet.Cells[i + 129, 526] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 526] = "0";

                    if (myCam.GetKTC())
                        myWorkSheet.Cells[i + 129, 527] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 527] = "0";

                    if (myCam.GetKRE())
                        myWorkSheet.Cells[i + 129, 528] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 528] = "0";

                    if (myCam.GetKSO())
                        myWorkSheet.Cells[i + 129, 529] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 529] = "0";

                    if (myCam.GetKVA())
                        myWorkSheet.Cells[i + 129, 530] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 530] = "0";

                    if (myCam.GetKCL())
                        myWorkSheet.Cells[i + 129, 531] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 531] = "0";

                    if (myCam.GetKPE())
                        myWorkSheet.Cells[i + 129, 532] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 532] = "0";

                    if (myCam.GetKUA())
                        myWorkSheet.Cells[i + 129, 533] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 533] = "0";

                    if (myCam.GetKAL())
                        myWorkSheet.Cells[i + 129, 534] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 534] = "0";

                    if (myCam.GetKMO())
                        myWorkSheet.Cells[i + 129, 535] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 535] = "0";

                    if (myCam.GetKSD())
                        myWorkSheet.Cells[i + 129, 536] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 536] = "0";

                    if (myCam.GetKSMin())
                        myWorkSheet.Cells[i + 129, 537] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 537] = "0";

                    if (myCam.GetKPAI())
                        myWorkSheet.Cells[i + 129, 538] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 538] = "0";

                    if (myCam.GetKpmtal())
                        myWorkSheet.Cells[i + 129, 514] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 514] = "0";

                    if (myCam.GetKlux())
                        myWorkSheet.Cells[i + 129, 787] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 787] = "0";

                    if (myCam.GetKvsol())
                        myWorkSheet.Cells[i + 129, 788] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 788] = "0";

                    if (myCam.GetKp10())
                        myWorkSheet.Cells[i + 129, 789] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 789] = "0";

                    if (myCam.GetKp12())
                        myWorkSheet.Cells[i + 129, 790] = "2";
                    else
                        myWorkSheet.Cells[i + 129, 790] = "0";

                    if (myCam.GetKp15())
                        myWorkSheet.Cells[i + 129, 791] = "3";
                    else
                        myWorkSheet.Cells[i + 129, 791] = "0";

                    if (myCam.GetKp15t())
                        myWorkSheet.Cells[i + 129, 792] = "4";
                    else
                        myWorkSheet.Cells[i + 129, 792] = "0";

                    if (myCam.GetKdt())
                        myWorkSheet.Cells[i + 129, 727] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 727] = "0";

                    if (myCam.GetKat())
                        myWorkSheet.Cells[i + 129, 792] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 792] = "0";

                    if (myCam.GetKbt())
                        myWorkSheet.Cells[i + 129, 793] = "2";
                    else
                        myWorkSheet.Cells[i + 129, 793] = "0";

                    if (myCam.GetKmt())
                        myWorkSheet.Cells[i + 129, 794] = "3";
                    else
                        myWorkSheet.Cells[i + 129, 794] = "0";

                    if (myCam.GetKmod())
                        myWorkSheet.Cells[i + 129, 795] = "1";
                    else
                        myWorkSheet.Cells[i + 129, 795] = "0";

                    if (myCam.GetRvent())
                        myWorkSheet.Cells[i + 5, 209] = "1";
                    else
                        myWorkSheet.Cells[i + 5, 209] = "0";

                    if(myCam.GetKantc())
                        myWorkSheet.Cells[i + 5, 715] = "1";
                    else
                        myWorkSheet.Cells[i + 5, 715] = "0";


                    myWorkSheet.Cells[i + 5, 65] = myCam.GetVol();
                    myWorkSheet.Cells[i + 5, 66] = myCam.GetFase();
                    myWorkSheet.Cells[i + 5, 2] = (i).ToString();
                    myWorkSheet.Cells[i + 5, 705] = myCam.GetSPtp143();
                    myWorkSheet.Cells[i + 5, 706] = myCam.GetSPtp144();
                    myWorkSheet.Cells[i + 5, 707] = myCam.GetSPtp145();
                    myWorkSheet.Cells[i + 5, 708] = myCam.GetSPtp146();
                    myWorkSheet.Cells[i + 5, 889] = myCam.GetSPtp147();
                    myWorkSheet.Cells[i + 5, 891] = myCam.GetSPtp148();
            
                    percent += perby;
                    thisForm.LEstado.Text = "Actualizando base de datos... " + percent.ToString() + "%";
                }
                //=======================================================================================
                this.ChangeSheet(5);
                if(Oferta.GetBun9() == true)
                {
                    myWorkSheet.Cells[9, 10] = "1";
                }
                else
                {
                    myWorkSheet.Cells[9, 10] = "0";
                }

                this.ChangeSheet(5);
                myRange = myWorkSheet.get_Range("A1", "K870");
                myValues = (Array)myRange.Value2;
               
                for (int i = 1; i <= cc; i++)
                {
                    CCam myCam = Oferta.GetCam(i - 1);
                    if (myCam.GetRvent())
                        myWorkSheet.Cells[i + 832, 5] = "X";
                    else
                        myWorkSheet.Cells[i + 832, 5] = "";
                    if (myCam.GetKepc())
                        myWorkSheet.Cells[i + 832, 6] = "X";
                    else
                        myWorkSheet.Cells[i + 832, 6] = "";
                }

                 
                //========================================================================================
                this.ChangeSheet(2);
                myWorkSheet.Cells[1, 10] = Oferta.GetDesE();
                try
                {
                    for (int i = 1; i <= 6; i++)
                    {
                        CCam myCam = Oferta.GetCam(i - 1);
                        if (myCam.GetCSumi() == "MULT-COMP-INT")
                        {
                            myWorkSheet.Cells[i + 88, 8] = "1";
                        }
                        else
                        {
                            myWorkSheet.Cells[i + 88, 8] = "0";
                        }

                    }
                }
                catch { }
                
                //=======================================================================================
                //pinchando con 5
                CCam Cam = Oferta.GetCam(0);
                this.ChangeSheet(5);
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetIE() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKP() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKN() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKM() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKC() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKU() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKPR() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKexpo()== true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKepiso()== true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKD() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKS() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKCE() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKB() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKCO() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKTO() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKTC() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKRE() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKSO() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKVA() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKCL() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKPE() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKUA() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKAL() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKMO() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKSD() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKSMin() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKPAI() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKpmtal() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKlux() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKvsol() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKp10() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKp12() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKp15() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKp15t() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKdt() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKat() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKbt() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKmt() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKmod()== true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetRvent() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKantc() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKppc() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKepc() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKcion() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKpcion() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKastre() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKfrio() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKeq1() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKeq2() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKeq3() == true))
                if((Oferta.GetCantCam()) <= 1 && (Cam.GetKsu1() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKsu2() == true))
                if ((Oferta.GetCantCam()) <= 1 && (Cam.GetKsu3() == true))
                    
                {
                    myWorkSheet.Cells[18, 2] = "";
                    myWorkSheet.Cells[18, 3] = "";
                    myWorkSheet.Cells[18, 4] = "";
                    myWorkSheet.Cells[18, 5] = "";
                    myWorkSheet.Cells[18, 6] = "";
                    myWorkSheet.Cells[18, 7] = "";
                    myWorkSheet.Cells[18, 8] = "";
                    myWorkSheet.Cells[18, 9] = "";
                    myWorkSheet.Cells[18, 10] = "";
                    myWorkSheet.Cells[18, 11] = "";
                    myWorkSheet.Cells[18, 12] = "";
                    myWorkSheet.Cells[19, 2] = "2";

                    myWorkSheet.Cells[20, 2] = "2,01";
                    myWorkSheet.Cells[21, 2] = "2,02";
                    myWorkSheet.Cells[22, 2] = "2,03";
                    myWorkSheet.Cells[23, 2] = "2,04";
                    myWorkSheet.Cells[24, 2] = "2,05";
                    myWorkSheet.Cells[25, 2] = "2,06";
                    myWorkSheet.Cells[26, 2] = "2,07";
                    myWorkSheet.Cells[27, 2] = "2,08";
                    myWorkSheet.Cells[28, 2] = "2,09";
                    myWorkSheet.Cells[29, 2] = "2,10";
                    myWorkSheet.Cells[30, 2] = "2,11"; 
                }
                
                //******* INFORM GNRAL  OFRTA***************************************
                this.ChangeSheet(2);
                myWorkSheet.Cells[850, 4] = "IMPORTE OFERTA " + Oferta.GetCdc() + " (" + Oferta.GetMon() + ")";
                myWorkSheet.Cells[850, 10] = "=SUMA(J9:J848)";

                if(Oferta.GetDesct().ToString() != "0")
                {
                    myWorkSheet.Cells[851, 4] = "DESCUENTO COMERCIAL " + Oferta.GetCdc() + " (" + Oferta.GetMon() + ")";
                    myWorkSheet.Cells[851, 10] = "=J850*" + Oferta.GetDesct() + "%";
                    if(Oferta.GetFlet().ToString() != "0")
                    {
                        myWorkSheet.Cells[852, 4] = "VALOR " + Oferta.GetCdc() + " CON DESCUENTO (" + Oferta.GetMon() + ")";
                        myWorkSheet.Cells[852, 10] = "=REDONDEAR(J850-J851;2)";
                    }
                   
                }
                else
                {
                    myWorkSheet.Cells[853, 4] = "0";
                    myWorkSheet.Cells[853, 10] = "0";
                    myWorkSheet.Cells[854, 4] = "0";
                    myWorkSheet.Cells[854, 12] = "=REDONDEAR(J786-J787;2)";
                }
                if (Oferta.GetFlet().ToString() != "0")
                {
                    myWorkSheet.Cells[855, 4] = "FLETE ESTIMADO (" + Oferta.GetMon() + ")";
                        if (Oferta.GetFlet().ToString() != "0")
                        {
                            if (Oferta.GetFlet().ToString() == "AGRUP-ITALIA")
                                myWorkSheet.Cells[855, 10] = "70";

                            if (Oferta.GetFlet().ToString() == "AGRUP-GENOVA")
                                myWorkSheet.Cells[855, 10] = "95";

                            if (Oferta.GetFlet().ToString() == "AGRUP-PANAMA")
                                myWorkSheet.Cells[855, 10] = "75";

                            if (Oferta.GetFlet().ToString() == "AGRUP-MEXICO")
                                myWorkSheet.Cells[855, 10] = "60";

                            if (Oferta.GetFlet().ToString() == "BOX 20´ ESPAÑA")
                                myWorkSheet.Cells[855, 10] = "1776*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 20´ ITALIA")
                                myWorkSheet.Cells[855, 10] = "1870*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 20´ PANAMA")
                                myWorkSheet.Cells[855, 10] = "1776*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 20´ MEX ALTMIRA")
                                myWorkSheet.Cells[855, 10] = "1210*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 20´ MEX PROGRESO")
                                myWorkSheet.Cells[855, 10] = "1335*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 40´ ESPAÑA")
                                myWorkSheet.Cells[855, 10] = "3036*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 40´ ITALIA")
                                myWorkSheet.Cells[855, 10] = "3180*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 40´ PANAMA")
                                myWorkSheet.Cells[855, 10] = "1670*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 40´ MEX ALTMIRA")
                                myWorkSheet.Cells[855, 10] = "1840*" + Oferta.GetNcont().ToString();

                            if (Oferta.GetFlet().ToString() == "BOX 40´ MEX PROGRESO")
                                myWorkSheet.Cells[855, 10] = "1810*" + Oferta.GetNcont().ToString();
                        }
                        myWorkSheet.Cells[855, 10] = Oferta.GetFlet().ToString(); 
                }
                else
                {
                    myWorkSheet.Cells[855, 4] = "0"; 
                }

                if (Oferta.GetCgr().ToString() !="0")
                {
                    myWorkSheet.Cells[856, 4] = "SEGURO TRANSPORTACION (" + Oferta.GetMon() + ")";
                    myWorkSheet.Cells[856, 10] = "=REDONDEAR(L788 *" + Oferta.GetCgr().ToString() + "%;2)";
                }
                else
                {
                    myWorkSheet.Cells[856, 4] = "0";
                    myWorkSheet.Cells[856, 10] = "0";
                }
                
                if (Oferta.GetIntr().ToString() != "0")
                {
                    myWorkSheet.Cells[857, 4] = "INTERES ANUAL " + Oferta.GetIntr() + "%  (" + Oferta.GetMon() + ")";
                    myWorkSheet.Cells[857, 10] = "=REDONDEAR(" + Oferta.GetIntr() + "%*(L788+H169+H170);2)";
                    //myWorkSheet.Cells[172, 4] = "IMPORTE TOTAL " + Oferta.GetCdc() + "(" + Oferta.GetMond() + ")";
                    //myWorkSheet.Cells[172, 8] = "=H171+H168+H169+H170";
                }
                else
                {
                    myWorkSheet.Cells[857, 4] = "0";
                    myWorkSheet.Cells[857, 10] = "0";
                }
                //********************** VALOR TOTAL ***********************
                
                if(Oferta.GetFlet().ToString() != "0")
                    if(Oferta.GetCgr().ToString() != "0")
                        if(Oferta.GetIntr().ToString() != "0")
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "IMPORTE TOTAL CIF (" + Oferta.GetMon() + ")";
                            myWorkSheet.Cells[857, 10] = "=REDONDEAR(L788+H169+H170+H171;2)";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[822, 4] = "='Oferta F'!D850";
                            myWorkSheet.Cells[823, 4] = "='Oferta F'!D851";
                            myWorkSheet.Cells[824, 4] = "='Oferta F'!D852";
                            myWorkSheet.Cells[825, 4] = "='Oferta F'!D853";
                            myWorkSheet.Cells[826, 4] = "='Oferta F'!D854";
                            myWorkSheet.Cells[827, 4] = "='Oferta F'!D855";
                            myWorkSheet.Cells[828, 4] = "='Oferta F'!D856";
                            myWorkSheet.Cells[829, 4] = "='Oferta F'!D857";
                            myWorkSheet.Cells[822, 8] = "='Oferta F'!J850";
                            myWorkSheet.Cells[823, 8] = "='Oferta F'!J851";
                            myWorkSheet.Cells[824, 8] = "='Oferta F'!J852";
                            myWorkSheet.Cells[825, 8] = "='Oferta F'!J853";
                            myWorkSheet.Cells[826, 8] = "='Oferta F'!J854";
                            myWorkSheet.Cells[827, 8] = "='Oferta F'!J855";
                            myWorkSheet.Cells[828, 8] = "='Oferta F'!J856";
                            myWorkSheet.Cells[829, 8] = "='Oferta F'!J857";
                        }
                else
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "0";
                            myWorkSheet.Cells[857, 10] = "0";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[759, 4] = "0";
                            myWorkSheet.Cells[759, 8] = "0";
                        }
                if(Oferta.GetCgr().ToString() == "0")
                {
                    if (Oferta.GetFlet().ToString() != "0")
                        if (Oferta.GetIntr().ToString() != "0")
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "IMPORTE TOTAL CFR (" + Oferta.GetMon() + ")";
                            myWorkSheet.Cells[857, 10] = "=REDONDEAR(L788+H169+H171;2)";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[822, 4] = "='Oferta F'!D850";
                            myWorkSheet.Cells[823, 4] = "='Oferta F'!D851";
                            myWorkSheet.Cells[824, 4] = "='Oferta F'!D852";
                            myWorkSheet.Cells[825, 4] = "='Oferta F'!D853";
                            myWorkSheet.Cells[826, 4] = "='Oferta F'!D854";
                            myWorkSheet.Cells[827, 4] = "='Oferta F'!D855";
                            myWorkSheet.Cells[828, 4] = "='Oferta F'!D856";
                            myWorkSheet.Cells[829, 4] = "='Oferta F'!D857";
                            myWorkSheet.Cells[822, 8] = "='Oferta F'!J850";
                            myWorkSheet.Cells[823, 8] = "='Oferta F'!J851";
                            myWorkSheet.Cells[824, 8] = "='Oferta F'!J852";
                            myWorkSheet.Cells[825, 8] = "='Oferta F'!J853";
                            myWorkSheet.Cells[826, 8] = "='Oferta F'!J854";
                            myWorkSheet.Cells[827, 8] = "='Oferta F'!J855";
                            myWorkSheet.Cells[828, 8] = "='Oferta F'!J856";
                            myWorkSheet.Cells[829, 8] = "='Oferta F'!J857";
                        }
                        else
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "0";
                            myWorkSheet.Cells[857, 10] = "0";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[759, 4] = "0";
                            myWorkSheet.Cells[759, 8] = "0";
                        }
                }
                if (Oferta.GetDesct().ToString() != "0")
                   if (Oferta.GetFlet().ToString() == "0")
                      if (Oferta.GetCgr().ToString() == "0")
                        if (Oferta.GetIntr().ToString() == "0")
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "IMPORTE TOTAL FCA (" + Oferta.GetMon() + ")";
                            myWorkSheet.Cells[857, 10] = "=REDONDEAR(J786-J787;2)";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[822, 4] = "='Oferta F'!D850";
                            myWorkSheet.Cells[823, 4] = "='Oferta F'!D851";
                            myWorkSheet.Cells[824, 4] = "='Oferta F'!D852";
                            myWorkSheet.Cells[825, 4] = "='Oferta F'!D853";
                            myWorkSheet.Cells[826, 4] = "='Oferta F'!D854";
                            myWorkSheet.Cells[827, 4] = "='Oferta F'!D855";
                            myWorkSheet.Cells[828, 4] = "='Oferta F'!D856";
                            myWorkSheet.Cells[829, 4] = "='Oferta F'!D857";
                            myWorkSheet.Cells[822, 8] = "='Oferta F'!J850";
                            myWorkSheet.Cells[823, 8] = "='Oferta F'!J851";
                            myWorkSheet.Cells[824, 8] = "='Oferta F'!J852";
                            myWorkSheet.Cells[825, 8] = "='Oferta F'!J853";
                            myWorkSheet.Cells[826, 8] = "='Oferta F'!J854";
                            myWorkSheet.Cells[827, 8] = "='Oferta F'!J855";
                            myWorkSheet.Cells[828, 8] = "='Oferta F'!J856";
                            myWorkSheet.Cells[829, 8] = "='Oferta F'!J857";
                        }
                        else
                        {
                            this.ChangeSheet(2);
                            myWorkSheet.Cells[857, 4] = "0";
                            myWorkSheet.Cells[857, 10] = "0";
                            this.ChangeSheet(5);
                            myWorkSheet.Cells[759, 4] = "0";
                            myWorkSheet.Cells[759, 8] = "0";
                        }

                if (Oferta.GetDesct().ToString() == "0")
                    if (Oferta.GetFlet().ToString() == "0")
                        if (Oferta.GetCgr().ToString() == "0")
                            if (Oferta.GetIntr().ToString() == "0")
                            {
                                this.ChangeSheet(2);
                                myWorkSheet.Cells[857, 4] = "IMPORTE TOTAL FCA (" + Oferta.GetMon() + ")";
                                //myWorkSheet.Cells[857, 10] = "=REDONDEAR(J786-J787;2)";
                                this.ChangeSheet(5);
                                myWorkSheet.Cells[822, 4] = "='Oferta F'!D850";
                                myWorkSheet.Cells[823, 4] = "='Oferta F'!D851";
                                myWorkSheet.Cells[824, 4] = "='Oferta F'!D852";
                                myWorkSheet.Cells[825, 4] = "='Oferta F'!D853";
                                myWorkSheet.Cells[826, 4] = "='Oferta F'!D854";
                                myWorkSheet.Cells[827, 4] = "='Oferta F'!D855";
                                myWorkSheet.Cells[828, 4] = "='Oferta F'!D856";
                                myWorkSheet.Cells[829, 4] = "='Oferta F'!D857";
                                myWorkSheet.Cells[822, 8] = "='Oferta F'!J850";
                                myWorkSheet.Cells[823, 8] = "='Oferta F'!J851";
                                myWorkSheet.Cells[824, 8] = "='Oferta F'!J852";
                                myWorkSheet.Cells[825, 8] = "='Oferta F'!J853";
                                myWorkSheet.Cells[826, 8] = "='Oferta F'!J854";
                                myWorkSheet.Cells[827, 8] = "='Oferta F'!J855";
                                myWorkSheet.Cells[828, 8] = "='Oferta F'!J856";
                                myWorkSheet.Cells[829, 8] = "='Oferta F'!J857";
                            }
                           /* else
                            {
                                this.ChangeSheet(2);
                                myWorkSheet.Cells[857, 4] = "0";
                                myWorkSheet.Cells[857, 10] = "0";
                                this.ChangeSheet(5);
                                myWorkSheet.Cells[759, 4] = "0";
                                myWorkSheet.Cells[759, 8] = "0";
                            }*/

                //******* 06/08/2017 **********************************************         
                
                //PROGRAMA GENERADOR DE LA OFERTA
                this.ChangeSheet(7);
                thisForm.LEstado.Text = "Generando Fichas Técnicas..." + percent.ToString() + "%";
                myWorkSheet.Cells[2, 5] = Oferta.GetLugar();
                myWorkSheet.Cells[49, 5] = Oferta.GetBmoni();
                myWorkSheet.Cells[49, 6] = Oferta.GetB60H();
                myWorkSheet.Cells[49, 7] = Oferta.GetBinvert();
                myWorkSheet.Cells[56, 5] = Oferta.GetNO();
               
                myRange = myWorkSheet.get_Range("A1", "GX" + 190.ToString());
                myValues = (Array)myRange.Value2;
                myWorkSheet.Cells[2, 3] = myValues.GetValue(2, 3).ToString();
                
                /*
                for (int l = cc + 6; l <= 42; l++)
                    for (int m = 2; m <= 206; m++)
                        myWorkSheet.Cells[l, m] = "";
                myWorkSheet.Calculate();
                
                myRange = myWorkSheet.get_Range("A1", "GX" + 190.ToString());
                myValues = (Array)myRange.Value2;
               
                myWorkSheet.Cells[112, 10] = "=SI.ERROR(SUMA('Oferta F'!J7:J599);0)";
                myWorkSheet.Cells[96, 26] = myValues.GetValue(96, 26).ToString();
                myWorkSheet.Cells[97, 26] = myValues.GetValue(97, 26).ToString();
                myWorkSheet.Cells[2, 3] = myValues.GetValue(2, 3).ToString();
                myWorkSheet.Cells[78, 13] = myValues.GetValue(78, 13).ToString();
               

                for (int e = 1; e <= cc; e++)
                    for (int p = 201; p <= 204; p++)
                        myWorkSheet.Cells[e + 5, p] = myValues.GetValue(e + 5, p).ToString();
                        
                */
                

                //==========================================================================               
                
                int counting = 0;
                //FIjando el valor de condiciones
                /* 01/04/2017 ***************************************
                this.ChangeSheet(6);
                myRange = myWorkSheet.get_Range("A1", "K" + 30.ToString());
                myValues = (Array)myRange.Value2;
                myWorkSheet.Cells[24, 6] = myValues.GetValue(24, 6);
                myWorkSheet.Cells[24, 7] = myValues.GetValue(24, 7);
                myWorkSheet.Cells[26, 6] = myValues.GetValue(26, 6);
                myWorkSheet.Cells[26, 7] = myValues.GetValue(26, 7);
                myWorkSheet.Cells[22, 6] = myValues.GetValue(22, 6);
                myWorkSheet.Cells[22, 7] = myValues.GetValue(22, 7);
                */ //**************************************************
                this.ChangeSheet(5);
                //myWorkSheet.Cells[1, 10] = Oferta.GetLugar();
                myWorkSheet.Cells[2, 4] = Oferta.GetNO() + "-" + String.Format("{0:yyyy}", DateTime.Now);
                myWorkSheet.Cells[3, 4] = Oferta.GetNP();
                //myWorkSheet.Cells[2, 4] = "GEM Delegación Cuba";
                //myWorkSheet.Cells[3, 4] = "Jesus Alvarez Hernández";
                //myWorkSheet.Cells[4, 4] = "Vendedor - Comprador";
                //***************************************************************************
                
                this.ChangeSheet(2);
                //myWorkSheet.Cells[785, 16] = Oferta.GetCdc();
                //myWorkSheet.Cells[786, 16] = Oferta.GetDesct();
                //myWorkSheet.Cells[787, 16] = Oferta.GetFlet();
               //myWorkSheet.Cells[787, 15] = Oferta.GetNcont();
                //myWorkSheet.Cells[788, 16] = Oferta.GetCgr();
                //myWorkSheet.Cells[789, 16] = Oferta.GetIntr();
                 
                myWorkSheet.Cells[2, 14] = Oferta.GetMon();
                
                //***************************************************************************
                //myWorkSheet.Cells[964, 2] = "Calle 30 e/ 1ra y 3ra,  #108, Miramar, Ciudad de la Habana";
                //myWorkSheet.Cells[965, 2] = "Tel:204-9358    Fax:204-9359";
                //myWorkSheet.Cells[966, 2] = "Email:Gem-cuba@gem.co.cu";
                //myWorkSheet.Cells[6, 6] = "Software ProJDC v.5.2";
                //===============================-=========================================
                myWorkBook.SaveCopyAs(@"C:\JOSE\cero.xlsx");
                myWorkBook.SaveCopyAs(@"C:\JOSE\" + Oferta.GetNO() +"-"+ Oferta.GetNP() +".xlsx");
                
                //FIjando el valor de condiciones
                this.ChangeSheet(6);
                myRange = myWorkSheet.get_Range("A1", "K" + 30.ToString());
                myValues = (Array)myRange.Value2;
                myWorkSheet.Cells[24, 6] = myValues.GetValue(24, 6);
                myWorkSheet.Cells[24, 7] = myValues.GetValue(24, 7);
                myWorkSheet.Cells[26, 6] = myValues.GetValue(26, 6);
                myWorkSheet.Cells[26, 7] = myValues.GetValue(26, 7);
                myWorkSheet.Cells[27, 6] = myValues.GetValue(27, 6);
                myWorkSheet.Cells[27, 7] = myValues.GetValue(27, 7);

                // ********************* CONTROL DE CARTA ****************
                this.ChangeSheet(3);
                myRange = myWorkSheet.get_Range("A1", "K45");
                myValues = (Array)myRange.Value2;

                for (int i = 1; i <= 40; i++)
                    for (int j = 2; j <= 11; j++)
                    {
                        try
                        {

                            {

                                {
                                    myWorkSheet.Cells[i, j] = Convert.ToString(myValues.GetValue(i, j), new CultureInfo("es-Es"));
                                    percent = (counting * 100) / 37318;
                                    thisForm.LEstado.Text = "Procesando Datos... " + percent.ToString() + "%";
                                    counting++;

                                }
                            }
                        }
                        catch { }
                    }
                this.ChangeSheet(8);
                myRange = myWorkSheet.get_Range("A1", "K118");
                myValues = (Array)myRange.Value2;

                for (int i = 1; i <= 118; i++)
                    for (int j = 2; j <= 11; j++)
                    {
                        try
                        {

                            {

                                {
                                    myWorkSheet.Cells[i, j] = Convert.ToString(myValues.GetValue(i, j), new CultureInfo("es-Es"));
                                    percent = (counting * 100) / 37318;
                                    thisForm.LEstado.Text = "Procesando Datos... " + percent.ToString() + "%";
                                    counting++;

                                }
                            }
                        }
                        catch { }
                    }
                // *********************************************************
                if (thisForm.RODC.Checked)
                {
                    this.ChangeSheet(9);
                    myRange = myWorkSheet.get_Range("A1", "AC157");
                    myValues = (Array)myRange.Value2;

                    for (int i = 2; i <= 157; i++)
                        for (int j = 1; j <= 27; j++)
                        {
                            try
                            {

                                {

                                    {
                                        myWorkSheet.Cells[i, j] = Convert.ToString(myValues.GetValue(i, j), new CultureInfo("es-Es"));
                                        percent = (counting * 100) / 37318;
                                        thisForm.LEstado.Text = "Procesando Detalles Electricos... " + percent.ToString() + "%";
                                        counting++;

                                    }


                                }
                            }
                            catch { }
                        }

                    bool vcaninc = true;
                    int rv = 6;
                    int bv = 157;
                    int yv = 0;


                    while (rv <= bv)
                    {
                        try
                        {
                            if (myValues.GetValue(rv, 2).ToString() == "0")
                            {
                                Range newRange = myWorkSheet.get_Range("A" + rv.ToString(), "AA" + rv.ToString());
                                newRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                                bv--;
                                if (rv < 157)

                                    yv++;
                                vcaninc = false;
                                myValues = (Array)myRange.Value2;
                            }
                            else
                                vcaninc = true;

                            if (vcaninc)
                                rv++;
                        }
                        catch { rv++; }
                    }

                    

                    this.ChangeSheet(8);
                    myWorkSheet.Cells[24, 2] = Oferta.GetLugar();
                    
                    
                    this.ChangeSheet(2);
                    thisForm.LEstado.Text = "Convirtiendo... ";
                    myRange = myWorkSheet.get_Range("B1", "M998");
                    thisForm.LEstado.Text = "Convitiendo... ";
                    myValues = (Array)myRange.Value2;
                    thisForm.LEstado.Text = "Convirtiendo... " + percent.ToString() + "%";
                    this.ChangeSheet(7);
                    myRange = myWorkSheet.get_Range("A1", "GX" + 190.ToString());
                    myValues = (Array)myRange.Value2;
                    //myWorkSheet.Cells[112, 10] = myValues.GetValue(112, 10).ToString();
                    myval = (20 / (403 - 8) - (cc * 8)).ToString();
                    perby = int.Parse(myval);
                    percent = (counting * 100) / 37318;
                    thisForm.LEstado.Text = "Actualizando listado... " + percent.ToString() + "%";

                    //=================================================================================

                    this.ChangeSheet(8);
                    myRange = myWorkSheet.get_Range("B1", "K118");
                    myValues = (Array)myRange.Value2;
                    bool eliminador5 = true;
                    int r = 3;
                    int b = 118;
                    int y = 0;


                    while (r <= b)
                    {
                        try
                        {
                            if (myValues.GetValue(r, 2).ToString() == "0")
                            {
                                Range newRange = myWorkSheet.get_Range("B" + r.ToString(), "P" + r.ToString());
                                newRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                                b--;
                                if (r < 118)

                                    y++;
                                eliminador5 = false;
                                myValues = (Array)myRange.Value2;
                            }
                            else
                                eliminador5 = true;

                            if (eliminador5)
                                r++;
                        }
                        catch { r++; }
                    }
                    //myWorkBook.SaveCopyAs(@"D:\test.xlsx");
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    thisForm.LEstado.Text = "Finalizando... ";
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    thisForm.LEstado.Text = "Finalizando... ";
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[3];
                    myWorkSheet.Delete();

                    thisForm.LEstado.Text = "Finalizando... ";


                    int q = 1;

                    while (q <= 20)
                    {
                        try
                        {
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[3];
                            myWorkSheet.Delete();
                        }
                        catch { }
                        q++;

                    }
                }
                else
                {
                    //IMPRIMENDO//
                    this.ChangeSheet(7);
                    if (thisForm.RCX.Checked)
                    {

                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[10];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[11];
                        myWorkSheet.PrintOut(1, 3, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[12];
                        myWorkSheet.PrintOut(1, 7, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[13];
                        myWorkSheet.PrintOut(1, 6, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[14];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[15];
                        myWorkSheet.PrintOut(1, 3, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[16];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[17];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[18];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[19];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[20];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[21];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[22];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[23];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[24];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[25];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[26];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[27];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                        myWorkSheet = (_Worksheet)myWorkBook.Worksheets[28];
                        myWorkSheet.PrintOut(1, 1, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value, Missing.Value);
                    }
                   

                    //===================21/12/2013==============================================
                    /*
                    this.ChangeSheet(7);
                    myRange = myWorkSheet.get_Range("A1", "GX" + 206.ToString());
                    myValues = (Array)myRange.Value2;
                
                    for (int l = 5; l <= cc + 5; l++)
                        for (int m = 13; m <= 206; m++)
                        {
                            try
                            {

                                {
                                    myWorkSheet.Cells[l, m] = myValues.GetValue(l, m).ToString();
                                    percent = (counting * 100) / 37318;
                                    thisForm.LEstado.Text = "Procesando Sistema... " + percent.ToString() + "%";
                                    counting++;
                                }
                            }
                            catch { }
                        }
                    
                    //22.01.2015
                    if (thisForm.RGX.Checked)
                    {
                        this.ChangeSheet(9);
                        myRange = myWorkSheet.get_Range("A1", "AC157");
                        myValues = (Array)myRange.Value2;

                        for (int i = 4; i <= 157; i++)
                            for (int j = 1; j <= 26; j++)
                            {
                                try
                                {

                                    {

                                        {
                                            myWorkSheet.Cells[i, j] = myValues.GetValue(i, j).ToString();
                                            percent = (counting * 100) / 37318;
                                            thisForm.LEstado.Text = "Procesando Evaporadores... " + percent.ToString() + "%";
                                            counting++;

                                        }


                                    }
                                }
                                catch { }
                            }

                        bool vcaninc = true;
                        int rv = 6;
                        int bv = 157;
                        int yv = 0;


                        while (rv <= bv)
                        {
                            try
                            {
                                if (myValues.GetValue(rv, 2).ToString() == "0")
                                {
                                    Range newRange = myWorkSheet.get_Range("A" + rv.ToString(), "Z" + rv.ToString());
                                    newRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                                    bv--;
                                    if (rv < 157)

                                        yv++;
                                    vcaninc = false;
                                    myValues = (Array)myRange.Value2;
                                }
                                else
                                    vcaninc = true;

                                if (vcaninc)
                                    rv++;
                            }
                            catch { rv++; }
                        }
                    }


                    this.ChangeSheet(5);
                    myWorkSheet.Cells[1, 10] = Oferta.GetLugar();
                    myWorkSheet.Cells[2, 5] = Oferta.GetNO();
                    myWorkSheet.Cells[3, 5] = Oferta.GetNP();
                    myWorkSheet.Cells[2, 6] = "GEM Delegación Cuba";
                    myWorkSheet.Cells[3, 6] = "Vivian Castro Sardiñas";
                    myWorkSheet.Cells[4, 6] = "Vendedor - Comprador";
                    //myWorkSheet.Cells[964, 2] = "Calle 30 e/ 1ra y 3ra,  #108, Miramar, Ciudad de la Habana";
                    //myWorkSheet.Cells[965, 2] = "Tel:204-9358    Fax:204-9359";
                    //myWorkSheet.Cells[966, 2] = "Email:Gem-cuba@gem.co.cu";
                    myWorkSheet.Cells[6, 6] = "Software ProJDC v.5.2";
                    */
                    /*
                    this.ChangeSheet(2);

                    if(Oferta.GetCdc() == "CIF")
                    {
                        myWorkSheet.Cells[785, 8] = "1";
                        myWorkSheet.Cells[786, 8] = "1";
                        myWorkSheet.Cells[787, 8] = "1";
                        myWorkSheet.Cells[788, 8] = "1";
                        myWorkSheet.Cells[789, 8] = "1";
                    }

                    if (Oferta.GetCdc() == "DAP")
                    {

                        myWorkSheet.Cells[785, 8] = "0";
                        myWorkSheet.Cells[786, 8] = "0";
                        myWorkSheet.Cells[787, 8] = "0";
                        myWorkSheet.Cells[788, 8] = "0";

                    }
                    if (Oferta.GetDesct() == "0")
                    {
                        if (Oferta.GetCdc() == "FCA")
                        {
                            myWorkSheet.Cells[786, 8] = "0";
                            myWorkSheet.Cells[787, 8] = "0";
                            myWorkSheet.Cells[788, 8] = "0";
                            myWorkSheet.Cells[789, 8] = "0";

                        }
                    }
                    else
                    {
                        if (Oferta.GetCdc() == "FCA")
                        {
                            myWorkSheet.Cells[786, 8] = "0";
                            myWorkSheet.Cells[787, 8] = "0";
                            myWorkSheet.Cells[788, 8] = "0";
                            myWorkSheet.Cells[789, 8] = "1";

                        }
                    }
                    if (Oferta.GetDesct() == "0")
                    {
                        if (Oferta.GetCdc() == "FOB")
                        {
                            myWorkSheet.Cells[786, 8] = "0";
                            myWorkSheet.Cells[787, 8] = "0";
                            myWorkSheet.Cells[788, 8] = "0";
                            myWorkSheet.Cells[789, 8] = "0";

                        }
                    }
                    else
                    {
                        if (Oferta.GetCdc() == "FOB")
                        {
                            myWorkSheet.Cells[786, 8] = "0";
                            myWorkSheet.Cells[787, 8] = "0";
                            myWorkSheet.Cells[788, 8] = "0";
                            myWorkSheet.Cells[789, 8] = "1";

                        }
                    }
                    */
                    this.ChangeSheet(5);
                    myWorkSheet.Cells[821, 7] = "=REDONDEAR(SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-25\")+SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-46\")+SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-57\")+SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-67\")+SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-80\")+SUMAR.SI.CONJUNTO(J10:J811;F10:F811;\"2w-120\")+(CONTAR.SI.CONJUNTO(F10:F811;\"2w-25\")+CONTAR.SI.CONJUNTO(F10:F811;\"2w-46\")+CONTAR.SI.CONJUNTO(F10:F811;\"2w-57\")+CONTAR.SI.CONJUNTO(F10:F811;\"2w-67\")+CONTAR.SI.CONJUNTO(F10:F811;\"2w-80\")+CONTAR.SI.CONJUNTO(F10:F811;\"2w-120\"))*102,91;2)";
                    myRange = myWorkSheet.get_Range("A1", "N890");
                    myValues = (Array)myRange.Value2;

                    for (int i = 1; i <= 869; i++)
                        for (int j = 1; j <= 11; j++)
                        {
                            try
                            {

                                {

                                    {
                                        myWorkSheet.Cells[i, j] = Convert.ToString(myValues.GetValue(i, j), new CultureInfo("es-Es"));
                                        percent = (counting * 100) / 37318;
                                        thisForm.LEstado.Text = "Procesando Und.Compresoras... " + percent.ToString() + "%";
                                        counting++;

                                    }
                                }
                            }
                            catch { }
                        }
                    //********************************************************************************************** 20/03/2016
                    

                    //********************************************************************************************** 20/03/2016
                    this.ChangeSheet(2);
                    thisForm.LEstado.Text = "Descargando Evaporadores... ";
                    myRange = myWorkSheet.get_Range("B1", "M998");
                    thisForm.LEstado.Text = "Descargando... ";
                    myValues = (Array)myRange.Value2;
                    thisForm.LEstado.Text = "Descargando Compresores... " + percent.ToString() + "%";
                    this.ChangeSheet(7);
                    myRange = myWorkSheet.get_Range("A1", "GX" + 190.ToString());
                    myValues = (Array)myRange.Value2;
                    //myWorkSheet.Cells[112, 10] = myValues.GetValue(112, 10).ToString();
                    myval = (20 / (403 - 8) - (cc * 8)).ToString();
                    perby = int.Parse(myval);
                    percent = (counting * 100) / 37318;
                    thisForm.LEstado.Text = "Descargando Monile... " + percent.ToString() + "%";

                    //=================================================================================
                    /*
                    this.ChangeSheet(5);
                    myRange = myWorkSheet.get_Range("A1", "P788");
                    myValues = (Array)myRange.Value2;
                    bool eliminador6 = true;
                    int r = 784;
                    int b = 788;
                    int y = 0;


                    while (r <= b)
                    {
                        try
                        {
                            if (myValues.GetValue(r, 10).ToString() == "0")
                            {
                                Range newRange = myWorkSheet.get_Range("A" + r.ToString(), "P" + r.ToString());
                                newRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                                b--;
                                if (r < 788)

                                    y++;
                                eliminador6 = false;
                                myValues = (Array)myRange.Value2;
                            }
                            else
                                eliminador6 = true;

                            if (eliminador6)
                                r++;
                        }
                        catch { r++; }
                    }
                    */
                    //==========================================================================================
                    this.ChangeSheet(5);
                    myRange = myWorkSheet.get_Range("A1", "P888");
                    myValues = (Array)myRange.Value2;
                    bool eliminador5 = true;
                    int u = 6;
                    int v = 869;
                    int w = 0;


                    while (u <= v)
                    {
                        try
                        {
                            if (myValues.GetValue(u, 7).ToString() == "0")
                            {
                                Range newRange = myWorkSheet.get_Range("A" + u.ToString(), "P" + u.ToString());
                                newRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftUp);

                                v--;
                                if (u < 869)

                                    w++;
                                eliminador5 = false;
                                myValues = (Array)myRange.Value2;
                            }
                            else
                                eliminador5 = true;

                            if (eliminador5)
                                u++;
                        }
                        catch { u++; }
                    }
                    /*
                    this.ChangeSheet(2);
                    if (thisForm.RGX.Checked)
                    {
                        thisForm.LEstado.Text = "Descargando RGX... " + percent.ToString() + "%";
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    }
                    this.ChangeSheet(2);
                    if (thisForm.RCX.Checked)
                    {
                        thisForm.LEstado.Text = "Descargando RCX... " + percent.ToString() + "%";
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                        myRange = myWorkSheet.get_Range("L1", "L1010");
                        myRange.Delete(Microsoft.Office.Interop.Excel.XlDeleteShiftDirection.xlShiftToLeft);
                    }
                    */
                    this.ChangeSheet(3);
                    myRange = myWorkSheet.get_Range("A1", "C26");
                    myValues = (Array)myRange.Value2;
                    myWorkSheet.Cells[12, 3] = DateTime.Now.ToString();
                    myWorkSheet.Cells[21, 2] = "Att:";
                    //myWorkSheet.Cells[21, 2] = "Comercial:";
                    myWorkSheet.Cells[15, 3] = Oferta.GetNO() + "-" + String.Format("{0:yyyy}", DateTime.Now);
                    myWorkSheet.Cells[18, 3] = Oferta.GetNP();
                    /*
                    if (Oferta.GetClit() != "")
                    {
                        myWorkSheet.Cells[21, 3] = "Sr(a). " + Oferta.GetClit() + "";
                        myWorkSheet.Cells[18, 3] = "TECNOTEX - " + Oferta.GetREF();

                    }
                    else
                    {
                        myWorkSheet.Cells[21, 3] = "Estimado Comercial";
                    }
                    */
                   
                    //=================================================
                    /*
                    this.ChangeSheet(11);
                    myWorkSheet.Cells[30, 6] = "=SI.ERROR(PRESUPUESTO!F120+PRESUPUESTO!J112+PRESUPUESTO!P65;0)";
                    myWorkSheet.Cells[48, 6] = "=SI.ERROR(F46+F47-PRESUPUESTO!J112;0)";
                    myWorkSheet.Cells[48, 7] = "=SI.ERROR(G46+G47-PRESUPUESTO!J112;0)";
                    */
                    // Fijar valore y fijar Formulas con RCUC.......

                    this.ChangeSheet(2);
                    myWorkSheet.Cells[2, 8] = "=suma(j7:j998)";
                    //myWorkSheet.Cells[2, 6] = "IMPORTE ESTIMADO TOTAL DE LA OFERTA FCA";
                    this.ChangeSheet(7);
                    //myWorkSheet.Cells[112, 10] = "=SI.ERROR(SUMA('Oferta F'!H2);0)";

                    myRange = myWorkSheet.get_Range("AN6", "AP42");
                    myValues = (Array)myRange.Value2;
                    for (int i = 1; i <= 37; i++)
                        try
                        {

                            {
                                myWorkSheet.Cells[i + 5, 40] = myValues.GetValue(i, 1).ToString();
                                myWorkSheet.Cells[i + 5, 42] = myValues.GetValue(i, 3).ToString();
                                percent = (counting * 100) / 37318;
                                thisForm.LEstado.Text = "Procesando Accesorios... " + percent.ToString() + "%";
                                counting++;
                            }
                        }
                        catch { }
                    
                }
                

                //=================================================
                /*
                this.ChangeSheet(11);
                myWorkSheet.Cells[30, 6] = "=SI.ERROR(PRESUPUESTO!J114+PRESUPUESTO!J112+PRESUPUESTO!P65;0)";
                myWorkSheet.Cells[30, 7] = "=F30-(PRESUPUESTO!P65)";
                myWorkSheet.Cells[48, 6] = "=SI.ERROR(F46+F47-PRESUPUESTO!J112;0)";
                myWorkSheet.Cells[48, 7] = "=SI.ERROR(G46+G47-PRESUPUESTO!J112;0)";

                this.ChangeSheet(12);
                myWorkSheet.Cells[7, 7] = "='FORMULA TIPICA DE CALCULO'!G30/'FORMULA TIPICA DE CALCULO'!F30";
                
                // Fijar valore y fijar Formulas con RCUC.......
                
                this.ChangeSheet(11);
                myWorkSheet.Cells[30, 6] = "=SI.ERROR(PRESUPUESTO!J114+PRESUPUESTO!J112+PRESUPUESTO!P65;0)";
                myWorkSheet.Cells[30, 7] = "=F30-(PRESUPUESTO!P65)";
                myWorkSheet.Cells[48, 6] = "=SI.ERROR(F46+F47-PRESUPUESTO!J112;0)";
                myWorkSheet.Cells[48, 7] = "=SI.ERROR(G46+G47-PRESUPUESTO!J112;0)";
                myRange = myWorkSheet.get_Range("A1", "K57");
                myValues = (Array)myRange.Value2;
                for (int i = 1; i <= 27; i++)
                    for (int j = 1; j <= 11; j++)
                    {
                        try
                        {

                            {
                                myWorkSheet.Cells[i, j] = myValues.GetValue(i, j).ToString();
                                percent = (counting * 100) / 37318;
                                thisForm.LEstado.Text = "Procesando Paneles... " + percent.ToString() + "%";
                                counting++;
                            }
                        }
                        catch { }
                    }

                this.ChangeSheet(12);
                myWorkSheet.Cells[7, 7] = "='FORMULA TIPICA DE CALCULO'!G30/'FORMULA TIPICA DE CALCULO'!F30";
                  
                */
                
                //=====================================================================
                
                if (thisForm.RFTX.Checked)
                {
                    
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    thisForm.LEstado.Text = "Finalizando... ";
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    thisForm.LEstado.Text = "Finalizando... ";
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[2];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[3];
                    myWorkSheet.Delete();
                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[3];
                    myWorkSheet.Delete();
                    thisForm.LEstado.Text = "Finalizando... ";
                    /*
                    int q = 1;

                    while (q <= 61)
                    {
                        try
                        {
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[6];
                            myWorkSheet.Delete();
                        }
                        catch { }
                        q++;

                    }
                    */
                }
                else                    
                    {
                        if (thisForm.RCX.Checked)
                        {
                            //myWorkBook.SaveCopyAs(@"D:\test.xlsx");
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                            myWorkSheet.Delete();
                            thisForm.LEstado.Text = "Finalizando... ";
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                            myWorkSheet.Delete();
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[2];
                            myWorkSheet.Delete();
                            thisForm.LEstado.Text = "Finalizando... ";
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                            myWorkSheet.Delete();
                            
                            thisForm.LEstado.Text = "Finalizando... ";
                            

                            int q = 1;

                            while (q <= 20)
                            {
                                try
                                {
                                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                                    myWorkSheet.Delete();
                                }
                                catch { }
                                q++;

                            }
                            
                        }
                        if (thisForm.RGX.Checked)
                        {
                            //myWorkBook.SaveCopyAs(@"D:\test.xlsx");
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                            myWorkSheet.Delete();
                            thisForm.LEstado.Text = "Finalizando... ";
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[1];
                            myWorkSheet.Delete();
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[2];
                            myWorkSheet.Delete();
                            thisForm.LEstado.Text = "Finalizando... ";
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                            myWorkSheet.Delete();
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                            myWorkSheet.Delete();
                            myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                            myWorkSheet.Delete();


                            thisForm.LEstado.Text = "Finalizando... ";


                            int q = 1;

                            while (q <= 20)
                            {
                                try
                                {
                                    myWorkSheet = (_Worksheet)myWorkBook.Worksheets[4];
                                    myWorkSheet.Delete();
                                }
                                catch { }
                                q++;

                            }

                        }

                        
                    }
                
                thisForm.LEstado.Text = "Finalizando... ";
                thisForm.generado = true;                
                thisForm.LEstado.Text = "Exportando a *.xlsx...";
                myWorkBook.SaveCopyAs(thisForm.myXlsxSaveDialog.FileName);
                ExcelApp.ActiveWorkbook.Close(false, @"fnstpte.tpt", Type.Missing);
                ExcelApp.Quit();                
                thisForm.LEstado.Text = "Exportar terminado.";
                try
                {
                    thisForm.BExportar.Enabled = true;

                    thisForm.BAbrir.Enabled = true;
                    thisForm.MenuAbrir.Enabled = true;
                    thisForm.actualizarCámaraActualToolStripMenuItem.Enabled = true;
                    thisForm.actualizarOfertaActualToolStripMenuItem.Enabled = true;
                    thisForm.BAdd.Enabled = true;
                    thisForm.borrarCámaraActualToolStripMenuItem.Enabled = false;
                }
                catch { }
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message,
                    "Error al almacenar datos:", MessageBoxButtons.OK, MessageBoxIcon.Error);
                thisForm.LEstado.Text = "Error.";
            }
            
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
        string CantEv;            
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
        string Txm2;
        string TCuadro;
        string CBint;
        string CBcm;
        string CCmci;
        string CBexpo;
        string Cevap;
        string Cmodex;
        string Ctxv;
        string CSumi;
        string CTCond;
        string CTamb;
        string Ctpd;

        string Cnoff1;
        string Coff1;
        string Cnoff2;
        string Coff2;
        string Cnoff3;
        string Coff3;
        string Cnoff4;
        string Coff4;
        string Cnoff5;
        string Coff5;
        string Cnoff6;
        string Coff6;
        string Cnoff7;
        string Coff7;
        string Cnoff8;
        string Coff8;

        string CTEvap;
        
        bool IE;
        bool KP;
        bool KN;
        bool KM;
        bool KC;
        bool KU;
        bool KPR;
        bool Kexpo;
        bool Kepiso;
        bool KD;
        bool KS;
        bool KCE;
        bool KB;
        bool KCO;
        bool KTO;
        bool KTC;
        bool KRE;
        bool KSO;
        bool KVA;
        bool KCL;
        bool KPE;
        bool KUA;
        bool KAL;
        bool KMO;
        bool KSD;
        bool KSMin;
        bool KPAI;
        bool Kpmtal;
        bool Klux;
        bool Kvsol;
        bool Kp10;
        bool Kp12;
        bool Kp15;
        bool Kp15t;
        bool Kdt;
        bool Kat;
        bool Kbt;
        bool Kmt;
        bool Kmod;
        bool Rvent;
        bool Kantc;
        bool Kppc;
        bool Kepc;
        bool Kcion;
        bool Kastre;
        bool Kpcion;
        bool Kfrio;
        bool Keq1;
        bool Keq2;
        bool Keq3;
        bool Ksu1;
        bool Ksu2;
        bool Ksu3;
        bool Kpps;
        string Vol;
        string Fase;
        string ConD;
        string CodValv;
        string Tmos;
        string TInc;
        string TInev;
        string TIned;
        string TIncd;
        string TIpv;
        string TIcc;
        string TQevp;
        string TQevpd;
        string TQevpc;
        string TTint;
        string TEquip;
        string TCint1;
        string TCint2;
        string TCint3;
        string TMcc;
        string TCmce;
        string TPmce;
        string TDmce;
        string TDmc1;
        string TDmc11;
        string TDmc12;
        string TDlq1;
        string TDlq2;
        string TDlq3;
        string TDlq11;
        string TDlq21;
        string TDlq31;
        string TDsu130;
        string TDsu230;
        string TDsu330;
        string TDsu110;
        string TDsu210;
        string TDsu310;
        string TDsu105;
        string TDsu205;
        string TDsu305;
        string TDmc2;
        string TDmc3;
        string TDmc4;
        string TDmc5;
        string TDmc6;
        string TDmc7;
        string TDmc8;
        string TLcc;
        string TLss;
        string TPcmc;
        string TPcmc1;
        string TPcmc2;
        string TPcmc3;
        string TPcond;
        string TPlq;
        string TPvq;
        string TPls;
        string TPosc;
        string TPsq;
        string TPcq;
        string TPcv;
        string TPcx;
        string TPcy;
        string TPcp;
        string TPex;
        string TPrs;
        string TCsist;
        string TPem;
        string TPnt;
        string TPml;
        string TPdtr;
        string TPdtc;
        string TPdt2;
        string TPdt1;
        string TCp80;
        string TCp100;
        string TCp120;
        string TCp150;
        string TCt80;
        string TCt100;
        string TCt120;
        string TCt150;
        string TCp80m;
        string TCp100m;
        string TCp120m;
        string TCp150m;
        string TCt80m;
        string TCt100m;
        string TCt120m;
        string TCt150m;
        string SPtp84;
        string SPtp83;
        string SPtp78;
        string SPtp76;
        string SPtp75;
        string SPtp74;
        string SPtp73;
        string SPtp72;
        string SPtp85;
        string SPtp86;
        string SPtp87;
        string SPtp88;
        string SPtp89;
        string SPtp90;
        string SPtp91;
        string SPtp92;
        string SPtp93;
        string SPtp94;
        string SPtp95;
        string SPtp96;
        string SPtp97;
        string SPtp98;
        string SPtp99;
        string SPtp100;
        string SPtp101;
        string SPtp102;
        string SPtp103;
        string SPtp104;
        string SPtp105;
        string SPtp106;
        string SPtp107;
        string SPtp108;
        string SPtp109;
        string SPtp110;
        string SPtp111;
        string SPtp136;
        string SPtp137;
        string SPtp141;
        string SPtp142;
        string SPtp143;
        string SPtp144;
        string SPtp145;
        string SPtp146;
        string SPtp147;
        string SPtp148;
        string SPtp149;
        string SPtp150;
        string SPtp151;
        string TIn1;
        string TIn2;
        string TIn3;
        string TIn4;
        string TIn5;
        string TIn6;
        string TIn7;
        string TIn8;
        string TIn9;
        string TIn10;
        string TIn11;
        string TIn12;
        string TIn13;
        string TIn14;
        string TIn15;
        string TIn16;
        string TIn17;
        string TIn18;
        string TIn19;
        string TIn20;
        string TIn21;
        string TIn22;
        string TIn23;
        string TIp1;
        string TIp2;
        string TIp3;
        string TIp4;
        string TIp5;
        string TIp6;
        string TIp7;
        string TIp8;
        string TIp9;
        string TIp10;
        string TIp11;
        string TIp12;
        string TIp13;
        string TIp14;
        string TIp15;
        string TIp16;
        string TIp17;
        string TIp18;
        string TIp19;
        string TIp20;
        string TIp21;
        string TIp23;
        /// <summary>
        /// Contador que me lleva la constancia de las imagenes
        /// </summary>
        int pk;


        public CCam(string nc,
            string temp, string largo, string ancho, string alto,
            string tp, string dt, string ce, string cf, string fw, string qfw, string cmod, string cmodd, string cmodp, string desc, string prec, string qfep, string scdro, string spsi, string stemp, string apsi, string emevp, string sup,
            string it, string dec, string dece, string dech, string decf, string cantev, string centx, string cdin, string cxp, string refrig, string digt, string ccion, string caster, string cpcion, string cfrio, string ceq1, string ceq2, string ceq3, bool ie, bool kp, bool kn, bool km, bool kc, bool ku, 
            bool kpr, bool kexpo, bool kepiso, bool kd, bool ks, bool kce, bool kb,bool kco, bool kto, bool ktc, bool kre, bool kso, 
            bool kva, bool kcl, bool kpe, bool kua, bool kal, bool kmo, bool ksd, bool ksmin, bool kpai, bool kpmtal, bool klux, bool kvsol, bool kp10, bool kp12, bool kp15, bool kp15t,
            bool kdt, bool kat, bool kbt, bool kmt, bool kmod, bool rvent, bool kantc, bool kppc, bool kepc, bool kcion, bool kastre, bool kpcion, bool kfrio, bool keq1, bool keq2, bool keq3, bool ksu1, bool ksu2, bool ksu3, bool kpps, string vol, string fase, string cd,
            string tmuc, string tmevp, string tsol, string tvalv, string tcvta, string txm2, string tcuadro, string cbint, string cbcm, string ccmci, string cbexpo,
            string cevap, string cmodex, string ctxv, string csumi, string ctcond, string ctamb, string ctpd,
            string cnoff1, string coff1, string cnoff2, string coff2, string cnoff3, string coff3, string cnoff4, string coff4, string cnoff5,
            string coff5, string cnoff6, string coff6, string cnoff7, string coff7, string cnoff8, string coff8, string ctevap, string codvalv, string tmos, string tinc, string tinev, string tined, string tincd, string tipv, string ticc,
            string tqevp, string tqevpd, string tqevpc, string ttint, string tequip,
            string tcint1, string tcint2, string tcint3, string tmcc, string tcmce, string tpmce, string tdmce, string tdmc1, string tdmc11, string tdmc12, string tdlq1, string tdlq2, string tdlq3, string tdlq11, string tdlq21, string tdlq31, string tdsu130, string tdsu230, string tdsu330, string tdsu110, string tdsu210, string tdsu310,
            string tdsu105, string tdsu205, string tdsu305, string tdmc2, string tdmc3, string tdmc4, string tdmc5, string tdmc6, string tdmc7,
            string tdmc8, string tlcc, string tlss, string tpcmc, string tpcmc1, string tpcmc2, string tpcmc3, string tpcond, string tplq, string tpvq, string tpls, string tposc, string tpsq, string tpcq, string tpcv, string tpcx, string tpcy, string tpcp,
            string tpex, string tprs, string tcsist, string tpem, string tpnt, string tpml, string tpdtr, string tpdtc, string tpdt2, string tpdt1,
            string tcp80, string tcp100, string tcp120, string tcp150, string tct80, string tct100, string tct120, string tct150, string tcp80m, string tcp100m, string tcp120m, string tcp150m, string tct80m, string tct100m, string tct120m, string tct150m,
            string sptp84, string sptp83, string sptp78, string sptp76, string sptp75, string sptp74, string sptp73, string sptp72, string sptp85, string sptp86, string sptp87, string sptp88, string sptp89, string sptp90,
            string sptp91, string sptp92, string sptp93, string sptp94, string sptp95, string sptp96, string sptp97, string sptp98, string sptp99, string sptp100, string sptp101, string sptp102, string sptp103, string sptp104,
            string sptp105, string sptp106, string sptp107, string sptp108, string sptp109, string sptp110, string sptp111, string sptp136, string sptp137, string sptp141, string sptp142, string sptp143, string sptp144, string sptp145, string sptp146, string sptp147, string sptp148, string sptp149, string sptp150, string sptp151,
            string tin1, string tin2, string tin3, string tin4, string tin5, string tin6, string tin7, string tin8, string tin9, string tin10, string tin11, string tin12, string tin13, string tin14, string tin15, string tin16, string tin17, 
            string tin18, string tin19, string tin20, string tin21, string tin22, string tin23,
            string tip1, string tip2, string tip3, string tip4, string tip5, string tip6, string tip7, string tip8, string tip9, string tip10, string tip11, string tip12, string tip13, string tip14, string tip15, string tip16, string tip17,
            string tip18, string tip19, string tip20, string tip21, string tip23, int PK)
        {
            NC = nc;
            Temp = temp;
            Largo = largo;
            Ancho = ancho;
            Alto = alto;
            TP = tp;
            DT = dt;
            CE = ce;
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
            IT = it;
            DEC = dec;
            DECE = dece;
            DECH = dech;
            DECF = decf;
            CantEv = cantev;
            Centx = centx;
            Cdin = cdin;
            Cxp = cxp;
            Centx = centx;
            Refrig = refrig;
            Digt = digt;
            Ccion = ccion;
            Castre = caster;
            Cpcion = cpcion;
            Cfrio = cfrio;
            Ceq1 = ceq1;
            Ceq2 = ceq2;
            Ceq3 = ceq3;
            IE = ie;
            KP = kp;
            KN = kn;
            KM = km;
            KC = kc;
            KU = ku;
            KPR = kpr;
            Kexpo = kexpo;
            Kepiso = kepiso;
            KD = kd;
            KS = ks;
            KCE = kce;
            KB = kb;
            KCO = kco;
            KTO = kto;
            KTC = ktc;
            KRE = kre;
            KSO = kso;
            KVA = kva;
            KCL = kcl;
            KPE = kpe;
            KUA= kua;
            KAL = kal;
            KMO = kmo;
            KSD = ksd;
            KSMin = ksmin;
            KPAI = kpai;
            Kpmtal = kpmtal;
            Klux = klux;
            Kvsol = kvsol;
            Kp10 = kp10;
            Kp12 = kp12;
            Kp15 = kp15;
            Kp15t = kp15t;
            Kdt = kdt;
            Kat = kat;
            Kbt = kbt;
            Kmt = kmt;
            Kmod = kmod;
            Rvent = rvent;
            Kantc = kantc;
            Kppc = kppc;
            Kepc = kepc;
            Kcion = kcion;
            Kastre = kastre;
            Kpcion = kpcion;
            Kfrio = kfrio;
            Keq1 = keq1;
            Keq2 = keq2;
            Keq3 = keq3;
            Ksu1 = ksu1;
            Ksu2 = ksu2;
            Ksu3 = ksu3;
            Kpps = kpps;
            Vol = vol;
            Fase = fase;
            ConD = cd;
            TMuc = tmuc;
            TMevp = tmevp;
            
            TSol = tsol;
            TValv = tvalv;
            TCvta = tcvta;
            Txm2 = txm2;
            TCuadro = tcuadro;
            CBint = cbint;
            CBcm = cbcm;
            CCmci = ccmci;
            CBexpo = cbexpo;
            Cevap = cevap;
            Cmodex = cmodex;
            Ctxv = ctxv;
            CSumi = csumi;
            CTCond = ctcond;
            CTamb = ctamb;
            Ctpd = ctpd;
            Cnoff1 = cnoff1;
            Coff1 = coff1;
            Cnoff2 = cnoff2;
            Coff2 = coff2;
            Cnoff3 = cnoff3;
            Coff3 = coff3;
            Cnoff4 = cnoff4;
            Coff4 = coff4;
            Cnoff5 = cnoff5;
            Coff5 = coff5;
            Cnoff6 = cnoff6;
            Coff6 = coff6;
            Cnoff7 = cnoff7;
            Coff7 = coff7;
            Cnoff8 = cnoff8;
            Coff8 = coff8;
            CTEvap = ctevap;
            CodValv = codvalv;
            Tmos = tmos;
            TInc = tinc;
            TInev = tinev;
            TIned = tined;
            TIncd = tincd;
            TIpv = tipv;
            TIcc = ticc;
            TQevp = tqevp;
            TQevpd = tqevpd;
            TQevpc = tqevpc;
            TTint = ttint;
            TEquip = tequip;
            TCint1 = tcint1;
            TCint2 = tcint2;
            TCint3 = tcint3;
            TMcc = tmcc;
            TCmce = tcmce;
            TPmce = tpmce;
            TDmce = tdmce;
            TDmc1 = tdmc1;
            TDmc11 = tdmc11;
            TDmc12 = tdmc12;
            TDlq1 = tdlq1;
            TDlq2 = tdlq2;
            TDlq3 = tdlq3;
            TDlq11 = tdlq11;
            TDlq21 = tdlq21;
            TDlq31 = tdlq31;
            TDsu130 = tdsu130;
            TDsu230 = tdsu230;
            TDsu330 = tdsu330;
            TDsu110 = tdsu110;
            TDsu210 = tdsu210;
            TDsu310 = tdsu310;
            TDsu105 = tdsu105;
            TDsu205 = tdsu205;
            TDsu305 = tdsu305;
            TDmc2 = tdmc2;
            TDmc3 = tdmc3;
            TDmc4 = tdmc4;
            TDmc5 = tdmc5;
            TDmc6 = tdmc6;
            TDmc7 = tdmc7;
            TDmc8 = tdmc8;
            TLcc = tlcc;
            TLss = tlss;
            TPcmc = tpcmc;
            TPcmc1 = tpcmc1;
            TPcmc2 = tpcmc2;
            TPcmc3 = tpcmc3;
            TPcond = tpcond;
            TPlq = tplq;
            TPvq = tpvq;
            TPls = tpls;
            TPosc = tposc;
            TPsq = tpsq;
            TPcq = tpcq;
            TPcv = tpcv;
            TPcx = tpcx;
            TPcy = tpcy;
            TPcp = tpcp;
            TPex = tpex;
            TPrs = tprs;
            TCsist = tcsist;
            TPem = tpem;
            TPnt = tpnt;
            TPml = tpml;
            TPdtr = tpdtr;
            TPdtc = tpdtc;
            TPdt2 = tpdt2;
            TPdt1 = tpdt1;
            TCp80 = tcp80;
            TCp100 = tcp100;
            TCp120 = tcp120;
            TCp150 = tcp150;
            TCt80 = tct80;
            TCt100 = tct100;
            TCt120 = tct120;
            TCt150 = tct150;
            TCp80m = tcp80m;
            TCp100m = tcp100m;
            TCp120m = tcp120m;
            TCp150m = tcp150m;
            TCt80m = tct80m;
            TCt100m = tct100m;
            TCt120m = tct120m;
            TCt150m = tct150m;
            SPtp84 = sptp84;
            SPtp83 = sptp83;
            SPtp78 = sptp78;
            SPtp76 = sptp76;
            SPtp75 = sptp75;
            SPtp74 = sptp74;
            SPtp73 = sptp73;
            SPtp72 = sptp72;
            SPtp85 = sptp85;
            SPtp86 = sptp86;
            SPtp87 = sptp87;
            SPtp88 = sptp88;
            SPtp89 = sptp89;
            SPtp90 = sptp90;
            SPtp91 = sptp91;
            SPtp92 = sptp92;
            SPtp93 = sptp93;
            SPtp94 = sptp94;
            SPtp95 = sptp95;
            SPtp96 = sptp96;
            SPtp97 = sptp97;
            SPtp98 = sptp98;
            SPtp99 = sptp99;
            SPtp100 = sptp100;
            SPtp101 = sptp101;
            SPtp102 = sptp102;
            SPtp103 = sptp103;
            SPtp104 = sptp104;
            SPtp105 = sptp105;
            SPtp106 = sptp106;
            SPtp107 = sptp107;
            SPtp108 = sptp108;
            SPtp109 = sptp109;
            SPtp110 = sptp110;
            SPtp111 = sptp111;
            SPtp136 = sptp136;
            SPtp137 = sptp137;
            SPtp141 = sptp141;
            SPtp142 = sptp142;
            SPtp143 = sptp143;
            SPtp144 = sptp144;
            SPtp145 = sptp145;
            SPtp146 = sptp146;
            SPtp147 = sptp147;
            SPtp148 = sptp148;
            SPtp149 = sptp149;
            SPtp150 = sptp150;
            SPtp151 = sptp151;
            TIn1 = tin1;
            TIn2 = tin2;
            TIn3 = tin3;
            TIn4 = tin4;
            TIn5 = tin5;
            TIn6 = tin6;
            TIn7 = tin7;
            TIn8 = tin8;
            TIn9 = tin9;
            TIn10 = tin10;
            TIn11 = tin11;
            TIn12 = tin12;
            TIn13 = tin13;
            TIn14 = tin14;
            TIn15 = tin15;
            TIn16 = tin16;
            TIn17 = tin17;
            TIn18 = tin18;
            TIn19 = tin19;
            TIn20 = tin20;
            TIn21 = tin21;
            TIn22 = tin22;
            TIn23 = tin23;
            TIp1 = tip1;
            TIp2 = tip2;
            TIp3 = tip3;
            TIp4 = tip4;
            TIp5 = tip5;
            TIp6 = tip6;
            TIp7 = tip7;
            TIp8 = tip8;
            TIp9 = tip9;
            TIp10 = tip10;
            TIp11 = tip11;
            TIp12 = tip12;
            TIp13 = tip13;
            TIp14 = tip14;
            TIp15 = tip15;
            TIp16 = tip16;
            TIp17 = tip17;
            TIp18 = tip18;
            TIp19 = tip19;
            TIp20 = tip20;
            TIp21 = tip21;
            TIp23 = tip23;
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

        public string GetCantEv()
        {
            return CantEv;
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
        public string GetConD()
        {
            return ConD;
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
        public bool GetIE()
        {
            return IE;
        }
        public bool GetKP()
        {
            return KP;
        }
        public bool GetKN()
        {
            return KN;
        }
        public bool GetKM()
        {
            return KM;
        }
        public bool GetKC()
        {
            return KC;
        }
        public bool GetKU()
        {
            return KU;
        }
        public bool GetKPR()
        {
            return KPR;
        }
        public bool GetKexpo()
        {
            return Kexpo;
        }
        public bool GetKepiso()
        {
            return Kepiso;
        }
        public bool GetKD()
        {
            return KD;
        }
        public bool GetKS()
        {
            return KS;
        }
        public bool GetKCE()
        {
            return KCE;
        }
        public bool GetKB()
        {
            return KB;
        }
        public bool GetKCO()
        {
            return KCO;
        }
        public bool GetKTO()
        {
            return KTO;
        }
        public bool GetKTC()
        {
            return KTC;
        }
        public bool GetKRE()
        {
            return KRE;
        }
        public bool GetKSO()
        {
            return KSO;
        }
        public bool GetKVA()
        {
            return KVA;
        }
        public bool GetKCL()
        {
            return KCL;
        }
        public bool GetKPE()
        {
            return KPE;
        }
        public bool GetKUA()
        {
            return KUA;
        }
        public bool GetKAL()
        {
            return KAL;
        }
        public bool GetKMO()
        {
            return KMO;
        }
        public bool GetKSD()
        {
            return KSD;
        }
        public bool GetKSMin()
        {
            return KSMin;
        }
        public bool GetKpmtal()
        {
            return Kpmtal;
        }
        
        public bool GetKPAI()
        {
            return KPAI;
        }

        public bool GetKlux()
        {
            return Klux;
        }

        public bool GetKvsol()
        {
            return Kvsol;
        }

        public bool GetKp10()
        {
            return Kp10;
        }

        public bool GetKp12()
        {
            return Kp12;
        }

        public bool GetKp15()
        {
            return Kp15;
        }
        public bool GetKp15t()
        {
            return Kp15t;
        }
        public bool GetKdt()
        {
            return Kdt;
        }
        public bool GetKantc()
        {
            return Kantc;
        }
        public bool GetKppc()
        {
            return Kppc;
        }
        
        public bool GetKcion()
        {
            return Kcion;
        }
        public bool GetKastre()
        {
            return Kastre;
        }
        public bool GetKpcion()
        {
            return Kpcion;
        }
        public bool GetKfrio()
        {
            return Kfrio;
        }
        public bool GetKeq1()
        {
            return Keq1;
        }

        public bool GetKeq2()
        {
            return Keq2;
        }
        public bool GetKeq3()
        {
            return Keq3;
        }
        public bool GetKat()
        {
            return Kat;
        }

        public bool GetKbt()
        {
            return Kbt;
        }

        public bool GetKmt()
        {
            return Kmt;
        }

        public bool GetKmod()
        {
            return Kmod;
        }

        public bool GetRvent()
        {
            return Rvent;
        }
        public bool GetKepc()
        {
            return Kepc;
        }
        public bool GetKsu1()
        {
            return Ksu1;
        }
        public bool GetKsu2()
        {
            return Ksu2;
        }
        public bool GetKsu3()
        {
            return Ksu3;
        }
        public bool GetKpps()
        {
            return Kpps;
        }
        public string GetVol()
        {
            return Vol;
        }

        public string GetFase()
        {
            return Fase;
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

        public string GetTxm2()
        {
            return Txm2;
        }
        public string GetTCuadro()
        {
            return TCuadro;
        }
        public string GetCBint()
        {
            return CBint;
        }
        public string GetCBcm()
        {
            return CBcm;
        }
        public string GetCCmci()
        {
            return CCmci;
        }

        public string GetCBexpo()
        {
            return CBexpo;
        }

        public string GetCevap()
        {
            return Cevap;
        }

        
        public string GetCmodex()
        {
            return Cmodex;
        }
        public string GetCtxv()
        {
            return Ctxv;
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
        public string GetCoff2()
        {
            return Coff2;
        }
        public string GetCnoff3()
        {
            return Cnoff3;
        }
        public string GetCoff3()
        {
            return Coff3;
        }
        public string GetCnoff4()
        {
            return Cnoff4;
        }
        public string GetCoff4()
        {
            return Coff4;
        }
        public string GetCnoff5()
        {
            return Cnoff5;
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

        public string GetCnoff7()
        {
            return Cnoff7;
        }
        public string GetCoff7()
        {
            return Coff7;
        }


        public string GetCnoff8()
        {
            return Cnoff8;
        }
        public string GetCoff8()
        {
            return Coff8;
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
        public string GetTQevpc()
        {
            return TQevpc;
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
        public string GetTDmc1()
        {
            return TDmc1;
        }
        public string GetTDmc11()
        {
            return TDmc11;
        }
        public string GetTDmc12()
        {
            return TDmc12;
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
        public string GetTDsu130()
        {
            return TDsu130;
        }
        public string GetTDsu230()
        {
            return TDsu230;
        }
        public string GetTDsu330()
        {
            return TDsu330;
        }
        public string GetTDsu110()
        {
            return TDsu110;
        }
        public string GetTDsu210()
        {
            return TDsu210;
        }
        public string GetTDsu310()
        {
            return TDsu310;
        }
        public string GetTDsu105()
        {
            return TDsu105;
        }
        public string GetTDsu205()
        {
            return TDsu205;
        }
        public string GetTDsu305()
        {
            return TDsu305;
        }
        public string GetTDmc2()
        {
            return TDmc2;
        }
        public string GetTDmc3()
        {
            return TDmc3;
        }
        public string GetTDmc4()
        {
            return TDmc4;
        }
        public string GetTDmc5()
        {
            return TDmc5;
        }
        public string GetTDmc6()
        {
            return TDmc6;
        }
        public string GetTDmc7()
        {
            return TDmc7;
        }
        public string GetTDmc8()
        {
            return TDmc8;
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
        public string GetTPcmc1()
        {
            return TPcmc1;
        }
        public string GetTPcmc2()
        {
            return TPcmc2;
        }
        public string GetTPcmc3()
        {
            return TPcmc3;
        }
        public string GetTPcond()
        {
            return TPcond;
        }
        public string GetTPlq()
        {
            return TPlq;
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
        public string GetTPcv()
        {
            return TPcv;
        }
        public string GetTPcx()
        {
            return TPcx;
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
        public string GetTPdtr()
        {
            return TPdtr;
        }
        public string GetTPdtc()
        {
            return TPdtc;
        }
        public string GetTPdt2()
        {
            return TPdt2;
        }
        public string GetTPdt1()
        {
            return TPdt1;
        }
        public string GetTCp80()
        {
            return TCp80;
        }
        public string GetTCp100()
        {
            return TCp100;
        }
        public string GetTCp120()
        {
            return TCp120;
        }
        public string GetTCp150()
        {
            return TCp150;
        }
        public string GetTCt80()
        {
            return TCt80;
        }
        public string GetTCt100()
        {
            return TCt100;
        }
        public string GetTCt120()
        {
            return TCt120;
        }
        public string GetTCt150()
        {
            return TCt150;
        }
        public string GetTCp80m()
        {
            return TCp80m;
        }
        public string GetTCp100m()
        {
            return TCp100m;
        }
        public string GetTCp120m()
        {
            return TCp120m;
        }
        public string GetTCp150m()
        {
            return TCp150m;
        }
        public string GetTCt80m()
        {
            return TCt80m;
        }
        public string GetTCt100m()
        {
            return TCt100m;
        }
        public string GetTCt120m()
        {
            return TCt120m;
        }
        public string GetTCt150m()
        {
            return TCt150m;
        }

        public string GetSPtp84()
        {
            return SPtp84;
        }
        public string GetSPtp83()
        {
            return SPtp83;
        }
        public string GetSPtp78()
        {
            return SPtp78;
        }
        public string GetSPtp76()
        {
            return SPtp76;
        }
        public string GetSPtp75()
        {
            return SPtp75;
        }
        public string GetSPtp74()
        {
            return SPtp74;
        }
        public string GetSPtp73()
        {
            return SPtp73;
        }
        public string GetSPtp72()
        {
            return SPtp72;
        }
        public string GetSPtp85()
        {
            return SPtp85;
        }
        public string GetSPtp86()
        {
            return SPtp86;
        }
        public string GetSPtp87()
        {
            return SPtp87;
        }
        public string GetSPtp88()
        {
            return SPtp88;
        }
        public string GetSPtp89()
        {
            return SPtp89;
        }
        public string GetSPtp90()
        {
            return SPtp90;
        }
        public string GetSPtp91()
        {
            return SPtp91;
        }
        public string GetSPtp92()
        {
            return SPtp92;
        }
        public string GetSPtp93()
        {
            return SPtp93;
        }
        public string GetSPtp94()
        {
            return SPtp94;
        }
        public string GetSPtp95()
        {
            return SPtp95;
        }
        public string GetSPtp96()
        {
            return SPtp96;
        }
        public string GetSPtp97()
        {
            return SPtp97;
        }
        public string GetSPtp98()
        {
            return SPtp98;
        }
        public string GetSPtp99()
        {
            return SPtp99;
        }
        public string GetSPtp100()
        {
            return SPtp100;
        }
        public string GetSPtp101()
        {
            return SPtp101;
        }
        public string GetSPtp102()
        {
            return SPtp102;
        }
       
        public string GetSPtp103()
        {
            return SPtp103;
        }
        public string GetSPtp104()
        {
            return SPtp104;
        }
        public string GetSPtp105()
        {
            return SPtp105;
        }
        public string GetSPtp106()
        {
            return SPtp106;
        }
        public string GetSPtp107()
        {
            return SPtp107;
        }
        public string GetSPtp108()
        {
            return SPtp108;
        }
        public string GetSPtp109()
        {
            return SPtp109;
        }
        public string GetSPtp110()
        {
            return SPtp110;
        }
        public string GetSPtp111()
        {
            return SPtp111;
        }
        public string GetSPtp136()
        {
            return SPtp136;
        }
        public string GetSPtp137()
        {
            return SPtp137;
        }
        public string GetSPtp141()
        {
            return SPtp141;
        }
        public string GetSPtp142()
        {
            return SPtp142;
        }
        public string GetSPtp143()
        {
            return SPtp143;
        }
        public string GetSPtp144()
        {
            return SPtp144;
        }
        public string GetSPtp145()
        {
            return SPtp145;
        }
        public string GetSPtp146()
        {
            return SPtp146;
        }
        public string GetSPtp147()
        {
            return SPtp147;
        }
        public string GetSPtp148()
        {
            return SPtp148;
        }
        public string GetSPtp149()
        {
            return SPtp149;
        }
        public string GetSPtp150()
        {
            return SPtp150;
        }
        public string GetSPtp151()
        {
            return SPtp151;
        }

        public string GetTIn1()
        {
            return TIn1;
        }
        public string GetTIn2()
        {
            return TIn2;
        }
        public string GetTIn3()
        {
            return TIn3;
        }
        public string GetTIn4()
        {
            return TIn4;
        }
        public string GetTIn5()
        {
            return TIn5;
        }
        public string GetTIn6()
        {
            return TIn6;
        }
        public string GetTIn7()
        {
            return TIn7;
        }
        public string GetTIn8()
        {
            return TIn8;
        }
        public string GetTIn9()
        {
            return TIn9;
        }
        public string GetTIn10()
        {
            return TIn10;
        }
        public string GetTIn11()
        {
            return TIn11;
        }
        public string GetTIn12()
        {
            return TIn12;
        }
        public string GetTIn13()
        {
            return TIn13;
        }
        public string GetTIn14()
        {
            return TIn14;
        }
        public string GetTIn15()
        {
            return TIn15;
        }
        public string GetTIn16()
        {
            return TIn16;
        }
        public string GetTIn17()
        {
            return TIn17;
        }
        public string GetTIn18()
        {
            return TIn18;
        }
        public string GetTIn19()
        {
            return TIn19;
        }
        public string GetTIn20()
        {
            return TIn20;
        }
        public string GetTIn21()
        {
            return TIn21;
        }
        public string GetTIn22()
        {
            return TIn22;
        }
        public string GetTIn23()
        {
            return TIn23;
        }
        public string GetTIp1()
        {
            return TIp1;
        }
        public string GetTIp2()
        {
            return TIp2;
        }
        public string GetTIp3()
        {
            return TIp3;
        }
        public string GetTIp4()
        {
            return TIp4;
        }
        public string GetTIp5()
        {
            return TIp5;
        }
        public string GetTIp6()
        {
            return TIp6;
        }
        public string GetTIp7()
        {
            return TIp7;
        }
        public string GetTIp8()
        {
            return TIp8;
        }
        public string GetTIp9()
        {
            return TIp9;
        }
        public string GetTIp10()
        {
            return TIp10;
        }
        public string GetTIp11()
        {
            return TIp11;
        }
        public string GetTIp12()
        {
            return TIp12;
        }
        public string GetTIp13()
        {
            return TIp13;
        }
        public string GetTIp14()
        {
            return TIp14;
        }
        public string GetTIp15()
        {
            return TIp15;
        }
        public string GetTIp16()
        {
            return TIp16;
        }
        public string GetTIp17()
        {
            return TIp17;
        }
        public string GetTIp18()
        {
            return TIp18;
        }
        public string GetTIp19()
        {
            return TIp19;
        }
        public string GetTIp20()
        {
            return TIp20;
        }
        public string GetTIp21()
        {
            return TIp21;
        }
        public string GetTIp23()
        {
            return TIp23;
        }
        public int GetPK()
        {
            return pk;
        }

        //*********************************
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
        public void setCantEv(string cantev)
        {
            CantEv = cantev;
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

        public void SetConD(string cd)
        {
            ConD = cd;
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
        public void SetTxm2(string txm2)
        {
            Txm2 = txm2;
        }
        public void SetCuadro(string cuadro)
        {
            TCuadro = cuadro;
        }
        public void Setint(string cbint)
        {
            CBint = cbint;
        }
        public void Setcbmc(string cbmc)
        {
            CBcm = cbmc;
        }
        public void Setccmci(string ccmci)
        {
            CCmci = ccmci;
        }

        public void Setcbexpo(string cbexpo)
        {
            CBexpo = cbexpo;
        }

        public void Setcevap(string cevap)
        {
            Cevap = cevap;
        }

        
        public void Setcmodex(string cmodex)
        {
            Cmodex = cmodex;
        }
        public void Setctxv(string ctxv)
        {
            Ctxv = ctxv;
        }
        
        
        public void SetCSumi(string csumi)
        {
            CSumi = csumi;
        }
        
        public void SetCTCond(string ctcond)
        {
            CTCond = ctcond;
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

        public void SetCoff2(string coff2)
        {
            Coff2 = coff2;
        }
        public void SetCnoff3(string cnoff3)
        {
            Cnoff3 = cnoff3;
        }

        public void SetCoff3(string coff3)
        {
            Coff3 = coff3;
        }
        public void SetCnoff4(string cnoff4)
        {
            Cnoff4 = cnoff4;
        }
        public void SetCoff4(string coff4)
        {
            Coff4 = coff4;
        }
        public void SetCnoff5(string cnoff5)
        {
            Cnoff5 = cnoff5;
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

        public void SetCnoff7(string cnoff7)
        {
            Cnoff7 = cnoff7;
        }
        public void SetCoff7(string coff7)
        {
            Coff7 = coff7;
        }

        public void SetCnoff8(string cnoff8)
        {
            Cnoff8 = cnoff8;
        }
        public void SetCoff8(string coff8)
        {
            Coff8 = coff8;
        }

        public void SetCTEvap(string ctevap)
        {
            CTEvap = ctevap;
        }
        
        public void SetCodValv(string codvalv)
        {
            CodValv = codvalv;
        }
        public void SetTmos(string tmos)
        {
            Tmos = tmos;
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
        public void SetTQevpc(string tqevpc)
        {
            TQevpc = tqevpc;
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
        public void SetTDmc1(string tdmc1)
        {
            TDmc1 = tdmc1;
        }
        public void SetTDmc11(string tdmc11)
        {
            TDmc11 = tdmc11;
        }
        public void SetTDmc12(string tdmc12)
        {
            TDmc12 = tdmc12;
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
        public void SetTDmc2(string tdmc2)
        {
            TDmc2 = tdmc2;
        }
        public void SetTDmc3(string tdmc3)
        {
            TDmc3 = tdmc3;
        }
        public void SetTDmc4(string tdmc4)
        {
            TDmc4 = tdmc4;
        }
        public void SetTDmc5(string tdmc5)
        {
            TDmc5 = tdmc5;
        }
        public void SetTDmc6(string tdmc6)
        {
            TDmc6 = tdmc6;
        }
        public void SetTDmc7(string tdmc7)
        {
            TDmc7 = tdmc7;
        }
        public void SetTDmc8(string tdmc8)
        {
            TDmc8 = tdmc8;
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
        public void SetTPcmc1(string tpcmc1)
        {
            TPcmc1 = tpcmc1;
        }
        public void SetTPcmc2(string tpcmc2)
        {
            TPcmc2 = tpcmc2;
        }
        public void SetTPcmc3(string tpcmc3)
        {
            TPcmc3 = tpcmc3;
        }
        public void SetTPcond(string tpcond)
        {
            TPcond = tpcond;
        }
        public void SetTPlq(string tplq)
        {
            TPlq = tplq;
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
        public void SetTPcv(string tpcv)
        {
            TPcv = tpcv;
        }
        public void SetTPcx(string tpcx)
        {
            TPcx = tpcx;
        }
        public void SetTPcy(string tpcy)
        {
            TPcy = tpcy;
        }
        public void SetTPcp(string tpcp)
        {
            TPcp = tpcp;
        }
        public void SetTPcex(string tpex)
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
        public void SetTPdtr(string tpdtr)
        {
            TPdtr = tpdtr;
        }
        public void SetTPdtc(string tpdtc)
        {
            TPdtc = tpdtc;
        }
        public void SetTPdt2(string tpdt2)
        {
            TPdt2 = tpdt2;
        }
        public void SetTPdt1(string tpdt1)
        {
            TPdt1 = tpdt1;
        }
        public void SetTCp80(string tcp80)
        {
            TCp80 = tcp80;
        }
        public void SetTCp100(string tcp100)
        {
            TCp100 = tcp100;
        }
        public void SetTCp120(string tcp120)
        {
            TCp120 = tcp120;
        }
        public void SetTC150(string tcp150)
        {
            TCp150m = tcp150;
        }
        public void SetTCt80(string tct80)
        {
            TCt80 = tct80;
        }
        public void SetTCt100(string tct100)
        {
            TCt100 = tct100;
        }
        public void SetTCt120(string tct120)
        {
            TCt120 = tct120;
        }
        public void SetTCt150(string tct150)
        {
            TCt150 = tct150;
        }
        public void SetTCp80m(string tcp80m)
        {
            TCp80m = tcp80m;
        }
        public void SetTCp100m(string tcp100m)
        {
            TCp100m = tcp100m;
        }
        public void SetTCp120m(string tcp120m)
        {
            TCp120m = tcp120m;
        }
        public void SetTCp150m(string tcp150m)
        {
            TCp150m = tcp150m;
        }
        public void SetTCt80m(string tct80m)
        {
            TCt80m = tct80m;
        }
        public void SetTCt100m(string tct100m)
        {
            TCt100m = tct100m;
        }
        public void SetTCt120m(string tct120m)
        {
            TCt120m = tct120m;
        }
        public void SetTCt150m(string tct150m)
        {
            TCt150m = tct150m;
        }

        public void SetSPtp84(string sptp84)
        {
            SPtp84 = sptp84;
        }
        public void SetSPtp83(string sptp83)
        {
            SPtp83 = sptp83;
        }
        public void SetSPtp78(string sptp78)
        {
            SPtp78 = sptp78;
        }
        public void SetSPtp76(string sptp76)
        {
            SPtp76 = sptp76;
        }
        public void SetSPtp75(string sptp75)
        {
            SPtp75 = sptp75;
        }
        public void SetSPtp74(string sptp74)
        {
            SPtp74 = sptp74;
        }
        public void SetSPtp73(string sptp73)
        {
            SPtp73 = sptp73;
        }
        public void SetSPtp72(string sptp72)
        {
            SPtp72 = sptp72;
        }
        public void SetSPtp85(string sptp85)
        {
            SPtp85 = sptp85;
        }
        public void SetSPtp86(string sptp86)
        {
            SPtp86 = sptp86;
        }
        public void SetSPtp87(string sptp87)
        {
            SPtp87 = sptp87;
        }
        public void SetSPtp88(string sptp88)
        {
            SPtp87 = sptp88;
        }
        public void SetSPtp89(string sptp89)
        {
            SPtp89 = sptp89;
        }
        public void SetSPtp90(string sptp90)
        {
            SPtp90 = sptp90;
        }
        public void SetSPtp91(string sptp91)
        {
            SPtp91 = sptp91;
        }
        public void SetSPtp92(string sptp92)
        {
            SPtp92 = sptp92;
        }
        public void SetSPtp93(string sptp93)
        {
            SPtp93 = sptp93;
        }
        public void SetSPtp94(string sptp94)
        {
            SPtp94 = sptp94;
        }
        public void SetSPtp95(string sptp95)
        {
            SPtp95 = sptp95;
        }
        public void SetSPtp96(string sptp96)
        {
            SPtp96 = sptp96;
        }
        public void SetSPtp97(string sptp97)
        {
            SPtp97 = sptp97;
        }
        public void SetSPtp98(string sptp98)
        {
            SPtp98 = sptp98;
        }
        public void SetSPtp99(string sptp99)
        {
            SPtp99 = sptp99;
        }
        public void SetSPtp100(string sptp100)
        {
            SPtp100 = sptp100;
        }
        public void SetSPtp101(string sptp101)
        {
            SPtp101 = sptp101;
        }
        public void SetSPtp102(string sptp102)
        {
            SPtp102 = sptp102;
        }
        public void SetSPtp103(string sptp103)
        {
            SPtp103 = sptp103;
        }
        public void SetSPtp104(string sptp104)
        {
            SPtp104 = sptp104;
        }
        public void SetSPtp105(string sptp105)
        {
            SPtp105 = sptp105;
        }
        public void SetSPtp106(string sptp106)
        {
            SPtp106 = sptp106;
        }
        public void SetSPtp107(string sptp107)
        {
            SPtp107 = sptp107;
        }
        public void SetSPtp108(string sptp108)
        {
            SPtp108 = sptp108;
        }
        public void SetSPtp109(string sptp109)
        {
            SPtp109 = sptp109;
        }
        public void SetSPtp110(string sptp110)
        {
            SPtp110 = sptp110;
        }
        public void SetSPtp111(string sptp111)
        {
            SPtp111 = sptp111;
        }
        public void SetSPtp136(string sptp136)
        {
            SPtp136 = sptp136;
        }
        public void SetSPtp137(string sptp137)
        {
            SPtp137 = sptp137;
        }
        public void SetSPtp141(string sptp141)
        {
            SPtp141 = sptp141;
        }
        public void SetSPtp142(string sptp142)
        {
            SPtp142 = sptp142;
        }
        public void SetSPtp143(string sptp143)
        {
            SPtp143 = sptp143;
        }
        public void SetSPtp144(string sptp144)
        {
            SPtp144 = sptp144;
        }
        public void SetSPtp145(string sptp145)
        {
            SPtp145 = sptp145;
        }
        public void SetSPtp146(string sptp146)
        {
            SPtp146 = sptp146;
        }
        public void SetSPtp147(string sptp147)
        {
            SPtp147 = sptp147;
        }
        public void SetSPtp148(string sptp148)
        {
            SPtp148 = sptp148;
        }
        public void SetSPtp149(string sptp149)
        {
            SPtp149 = sptp149;
        }
        public void SetSPtp150(string sptp150)
        {
            SPtp150 = sptp150;
        }
        public void SetTIn1(string tin1)
        {
            TIn1 = tin1;
        }
        public void SetTIn2(string tin2)
        {
            TIn2 = tin2;
        }
        public void SetTIn3(string tin3)
        {
            TIn3 = tin3;
        }
        public void SetTIn4(string tin4)
        {
            TIn4 = tin4;
        }
        public void SetTIn5(string tin5)
        {
            TIn5 = tin5;
        }
        public void SetTIn6(string tin6)
        {
            TIn6 = tin6;
        }
        public void SetTIn7(string tin7)
        {
            TIn7 = tin7;
        }
        public void SetTIn8(string tin8)
        {
            TIn8 = tin8;
        }
        public void SetTIn9(string tin9)
        {
            TIn9 = tin9;
        }
        public void SetTIn10(string tin10)
        {
            TIn10 = tin10;
        }
        public void SetTIn11(string tin11)
        {
            TIn11 = tin11;
        }
        public void SetTIn12(string tin12)
        {
            TIn12 = tin12;
        }
        public void SetTIn13(string tin13)
        {
            TIn13 = tin13;
        }
        public void SetTIn14(string tin14)
        {
            TIn14 = tin14;
        }
        public void SetTIn15(string tin15)
        {
            TIn15 = tin15;
        }
        public void SetTIn16(string tin16)
        {
            TIn16 = tin16;
        }
        public void SetTIn17(string tin17)
        {
            TIn17 = tin17;
        }
        public void SetTIn18(string tin18)
        {
            TIn18 = tin18;
        }
        public void SetTIn19(string tin19)
        {
            TIn19 = tin19;
        }
        public void SetTIn20(string tin20)
        {
            TIn20 = tin20;
        }
        public void SetTIn21(string tin21)
        {
            TIn21 = tin21;
        }
        public void SetTIn22(string tin22)
        {
            TIn22 = tin22;
        }
        public void SetTIn23(string tin23)
        {
            TIn23 = tin23;
        }
        public void SetTIp1(string tip1)
        {
            TIp1 = tip1;
        }
        public void SetTIp2(string tip2)
        {
            TIp2 = tip2;
        }
        public void SetTIp3(string tip3)
        {
            TIp3 = tip3;
        }
        public void SetTIp4(string tip4)
        {
            TIp4 = tip4;
        }
        public void SetTIp5(string tip5)
        {
            TIp5 = tip5;
        }
        public void SetTIp6(string tip6)
        {
            TIp6 = tip6;
        }
        public void SetTIp7(string tip7)
        {
            TIp7 = tip7;
        }
        public void SetTIp8(string tip8)
        {
            TIp8 = tip8;
        }
        public void SetTIp9(string tip9)
        {
            TIp9 = tip9;
        }
        public void SetTIp10(string tip10)
        {
            TIp10 = tip10;
        }
        public void SetTIp11(string tip11)
        {
            TIp11 = tip11;
        }
        public void SetTIp12(string tip12)
        {
            TIp12 = tip12;
        }
        public void SetTIp13(string tip13)
        {
            TIp13 = tip13;
        }
        public void SetTIp14(string tip14)
        {
            TIp14 = tip14;
        }
        public void SetTIp15(string tip15)
        {
            TIp15 = tip15;
        }
        public void SetTIp16(string tip16)
        {
            TIp16 = tip16;
        }
        public void SetTIp17(string tip17)
        {
            TIp17 = tip17;
        }
        public void SetTIp18(string tip18)
        {
            TIp18 = tip18;
        }
        public void SetTIp19(string tip19)
        {
            TIp19 = tip19;
        }
        public void SetTIp20(string tip20)
        {
            TIp20 = tip20;
        }
        public void SetTIp21(string tip21)
        {
            TIp21 = tip21;
        }
        public void SetTIp23(string tip23)
        {
            TIp23 = tip23;
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
        string Dmc1;
        string Dmc11;
        string Dmc12;
        string Dlq1;
        string Dlq2;
        string Dlq3;
        string Dlq11;
        string Dlq21;
        string Dlq31;
        string Dsu130;
        string Dsu230;
        string Dsu330;
        string Dsu110;
        string Dsu210;
        string Dsu310;
        string Dsu105;
        string Dsu205;
        string Dsu305;
        string Dmc2;
        string Dmc3;
        string Dmc4;
        string Dmc5;
        string Dmc6;
        string Dmc7;
        string Dmc8;
        string Lcc;
        string Lss;
        string Pcmc;
        string Pcmc1;
        string Pcmc2;
        string Pcmc3;
        string Pcond;
        string Plq;
        string Pvq;
        string Pls;
        string Posc;
        string Psq;
        string Pcq;
        string Pcv;
        string Pcx;
        string Pcy;
        string Pcp;
        string Pex;
        string Prs;
        string Csist;
        string Pem;
        string Pnt;
        string Pml;
        string Pdtr;
        string Pdtc;
        string Pdt2;
        string Pdt1;
        string Cp80;
        string Cp100;
        string Cp120;
        string Cp150;
        string Ct80;
        string Ct100;
        string Ct120;
        string Ct150;
        string Cp80m;
        string Cp100m;
        string Cp120m;
        string Cp150m;
        string Ct80m;
        string Ct100m;
        string Ct120m;
        string Ct150m;
        string Ptp84;
        string Ptp83;
        string Ptp78;
        string Ptp76;
        string Ptp75;
        string Ptp74;
        string Ptp73;
        string Ptp72;
        string Ptp85;
        string Ptp86;
        string Ptp87;
        string Ptp88;
        string Ptp89;
        string Ptp90;
        string Ptp91;
        string Ptp92;
        string Ptp93;
        string Ptp94;
        string Ptp95;
        string Ptp96;
        string Ptp97;
        string Ptp98;
        string Ptp99;
        string Ptp100;
        string Ptp101;
        string Ptp102;
        string Ptp103;
        string Ptp104;
        string Ptp105;
        string Ptp106;
        string Ptp107;
        string Ptp108;
        string Ptp109;
        string Ptp110;
        string Ptp111;
        string Ptp136;
        string Ptp137;
        string Ptp141;
        string Ptp142;
        string Ptp143;
        string Ptp144;
        string Ptp145;
        string Ptp146;
        string Ptp147;
        string Ptp148;
        string Ptp149;
        string Ptp150;
        string Ptp151;
        string Inc;
        string Inev;
        string Ined;
        string Incd;
        string Ipv;
        string Icc;
        string Qevp;
        string Qevpd;
        string Qevpc;
        string Tint;
        string Equip;
        string Sumi;
        string DT;
        CCam[] Camaras;
        int CantCam;
        int cont;
        string Fecha;
        string Digt;
        string Ccion;
        string Castre;
        string Cpcion;
        string Cfrio;
        string Ceq1;
        string Ceq2;
        string Ceq3;
        string ResinaM;
        string inc;
        string ConstCivPan;
        string EquipFrig;
        string PuertasFrig;
        string DesE;
        string Tasa;
        string Dsc;
        string GastosAdmObra;
        string GastosIndObra;
        string GastosIndObracuc;
        string Credito;
        string Creditocup;
        string CRcivil;
        string CRpiso;
        string Lugar;
        string Clit;
        string Clit1;
        string Clitm;
        bool Bmoni;
        bool B60H;
        bool Bun;
        bool Bun2;
        bool Bun3;
        bool Bun4;
        bool Bun5;
        bool Bun6;
        bool Bun7;
        bool Bun8;
        bool Bun9;
        bool Binvert;
        bool B360;
        bool Keur;
        bool Kmod;
        bool Bsup;
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
        string Mon;
        string In1;
        string In2;
        string In3;
        string In4;
        string In5;
        string In6;
        string In7;
        string In8;
        string In9;
        string In10;
        string In11;
        string In12;
        string In13;
        string In14;
        string In15;
        string In16;
        string In17;
        string In18;
        string In19;
        string In20;
        string In21;
        string In22;
        string In23;
        string Ip1;
        string Ip2;
        string Ip3;
        string Ip4;
        string Ip5;
        string Ip6;
        string Ip7;
        string Ip8;
        string Ip9;
        string Ip10;
        string Ip11;
        string Ip12;
        string Ip13;
        string Ip14;
        string Ip15;
        string Ip16;
        string Ip17;
        string Ip18;
        string Ip19;
        string Ip20;
        string Ip21;
        string Ip23;
       
        public COferta(string np, string REF, string no, string cmat, string dmc1, string dmc11, string dmc12, string dlq1, string dlq2, string dlq3, string dlq11, string dlq21, string dlq31, string dsu130, string dsu230, string dsu330, string dsu110, string dsu210, string dsu310, string dsu105, string dsu205, string dsu305, string dmc2, string dmc3, string dmc4, string dmc5, string dmc6, string dmc7, string dmc8, string lcc, string lss, string pcmc, string pcmc1, string pcmc2, string pcmc3, string pcond,
            string plq, string pvq, string pls, string posc, string psq, string pcq, string pcv, string pcx, string pcy, string pcp, string pex, string prs, string csist, string pem, string pnt, string pml, string pdtr, string pdtc, string pdt2, string pdt1,
            string cp80, string cp100, string cp120, string cp150, string ct80, string ct100, string ct120, string ct150, string cp80m, string cp100m, string cp120m, string cp150m, string ct80m, string ct100m, string ct120m, string ct150m, string ptp84, string ptp83, string ptp78,
            string ptp76, string ptp75, string ptp74, string ptp73, string ptp72, string ptp85, string ptp86, string ptp87, string ptp88, string ptp89, string ptp90, string ptp91, string ptp92, string ptp93, string ptp94, string ptp95,
            string ptp96, string ptp97, string ptp98, string ptp99, string ptp100, string ptp101, string ptp102, string ptp103, string ptp104, string ptp105, string ptp106, string ptp107, string ptp108, string ptp109, string ptp110, string ptp111, string ptp136, string ptp137,
            string ptp141, string ptp142, string ptp143, string ptp144, string ptp145, string ptp146, string ptp147, string ptp148, string ptp149, string ptp150, string ptp151,
            string inc, string inev, string ined, string incd, string ipv, string icc, string qevp, string qevpd, string qevpc, string tint, string equip, string sumi, string dt, int cantcam, string fecha, string resinam, string digt, string ccion, string castre, string cpcion, string cfrio, string ceq1, string ceq2, string ceq3, string Inc,
            string constcivpan, string equipfrig, string puertasfrig, string desE, string tasa, string dsc, string gastosadmobra,
            string gastosindobra, string gastosindobracuc, string credito, string creditocup, string crcivil, string crpiso, string lugar, string clit, string clit1, string clitm, bool bmoni, bool b60h, bool bun, bool bun2, bool bun3, bool bun4, bool bun5, bool bun6, bool bun7, bool rbun8, bool bun9, bool binvert, bool b360, bool keur, bool kmod, bool bsup, string bscu,
            string bcont, string bcos, string bdir, string benv, string bpo, string bfec, string bdes, string cdc, string flet, string cgr, string intr, string desct, string ncont, string mon, string in1, string in2, string in3, string in4, string in5, string in6, string in7, string in8, string in9, string in10, string in11, string in12, string in13, string in14, string in15, string in16, string in17, string in18, string in19, string in20, string in21,
            string in22, string in23, string ip1, string ip2, string ip3, string ip4, string ip5, string ip6, string ip7, string ip8, string ip9, string ip10, string ip11, string ip12, string ip13, string ip14, string ip15, string ip16, string ip17, string ip18, string ip19, string ip20, string ip21, string ip23)
        {
            NP = np;
            Ref = REF;
            NO = no;
            Cmat = cmat;
            Dmc1 = dmc1;
            Dmc11 = dmc11;
            Dmc12 = dmc12;
            Dlq1 = dlq1;
            Dlq2 = dlq2;
            Dlq3 = dlq3;
            Dlq11 = dlq11;
            Dlq21 = dlq21;
            Dlq31 = dlq31;
            Dsu130 = dsu130;
            Dsu230 = dsu230;
            Dsu330 = dsu330;
            Dsu110 = dsu110;
            Dsu210 = dsu210;
            Dsu310 = dsu310;
            Dsu105 = dsu105;
            Dsu205 = dsu205;
            Dsu305 = dsu305;
            Dmc2 = dmc2;
            Dmc3 = dmc3;
            Dmc4 = dmc4;
            Dmc5 = dmc5;
            Dmc6 = dmc6;
            Dmc7 = dmc7;
            Dmc8 = dmc8;
            Lcc = lcc;
            Lss = lss;
            Pcmc = pcmc;
            Pcmc1 = pcmc1;
            Pcmc2 = pcmc2;
            Pcmc3 = pcmc3;
            Pcond = pcond;
            Plq = plq;
            Pvq = pvq;
            Pls = pls;
            Posc = posc;
            Psq = psq;
            Pcq = pcq;
            Pcv = pcv;
            Pcx = pcx;
            Pcy = pcy;
            Pcp = pcp;
            Pex = pex;
            Prs = prs;
            Csist = csist;
            Pem = pem;
            Pnt = pnt;
            Pdtr = pdtr;
            Pdtc = pdtc;
            Pdt2 = pdt2;
            Pdt1 = pdt1;
            Cp80 = cp80;
            Cp100 = cp100;
            Cp120 = cp120;
            Cp150 = cp150;
            Ct80 = ct80;
            Ct100 = ct100;
            Ct120 = ct120;
            Ct150 = ct150;
            Cp80m = cp80m;
            Cp100m = cp100m;
            Cp120m = cp120m;
            Cp150m = cp150m;
            Ct80m = ct80m;
            Ct100m = ct100m;
            Ct120m = ct120m;
            Ct150m = ct150m;
            Ptp84 = ptp84;
            Ptp83 = ptp83;
            Ptp78 = ptp78;
            Ptp76 = ptp76;
            Ptp75 = ptp75;
            Ptp74 = ptp74;
            Ptp73 = ptp73;
            Ptp72 = ptp72;
            Ptp85 = ptp85;
            Ptp86 = ptp86;
            Ptp87 = ptp87;
            Ptp88 = ptp88;
            Ptp89 = ptp89;
            Ptp90 = ptp90;
            Ptp91 = ptp91;
            Ptp92 = ptp92;
            Ptp93 = ptp93;
            Ptp94 = ptp94;
            Ptp95 = ptp95;
            Ptp96 = ptp96;
            Ptp97 = ptp97;
            Ptp98 = ptp98;
            Ptp99 = ptp99;
            Ptp100 = ptp100;
            Ptp101 = ptp101;
            Ptp102 = ptp102;
            Ptp103 = ptp103;
            Ptp104 = ptp104;
            Ptp105 = ptp105;
            Ptp106 = ptp106;
            Ptp107 = ptp107;
            Ptp108 = ptp108;
            Ptp109 = ptp109;
            Ptp110 = ptp110;
            Ptp111 = ptp111;
            Ptp136 = ptp136;
            Ptp137 = ptp137;
            Ptp141 = ptp141;
            Ptp142 = ptp142;
            Ptp143 = ptp143;
            Ptp144 = ptp144;
            Ptp145 = ptp145;
            Ptp146 = ptp146;
            Ptp147 = ptp147;
            Ptp148 = ptp148;
            Ptp149 = ptp149;
            Ptp150 = ptp150;
            Ptp151 = ptp151;
            Inc = inc;
            Inev = inev;
            Ined = ined;
            Incd = incd;
            Ipv = ipv;
            Icc = icc;
            Qevp = qevp;
            Qevpd = qevpd;
            Qevpc = qevpc;
            Tint = tint;
            Equip = equip;
            Sumi = sumi;
            DT = dt;
            Camaras = new CCam[100];
            CantCam = cantcam;
            cont = 0;
            Fecha = fecha;
            Digt = digt;
            Ccion = ccion;
            Castre = castre;
            Cpcion = cpcion;
            Cfrio = cfrio;
            Ceq1 = ceq1;
            Ceq2 = ceq2;
            Ceq3 = ceq3;
            ResinaM = resinam;
            inc = Inc;
            ConstCivPan = constcivpan;
            EquipFrig  = equipfrig;
            PuertasFrig = puertasfrig;
            DesE = desE;
            Tasa  = tasa;
            Dsc = dsc;
            GastosAdmObra = gastosadmobra;
            GastosIndObra = gastosindobra;
            GastosIndObracuc = gastosindobracuc;
            Credito = credito;
            Creditocup = creditocup;
            CRcivil = crcivil;
            CRpiso = crpiso;
            Lugar = lugar;
            Clit = clit;
            Clit1 = clit1;
            Clitm = clitm;
            Bmoni  = bmoni;
            B60H = b60h;
            Bun = bun;
            Bun2 = bun2;
            Bun3 = bun3;
            Bun4 = bun4;
            Bun4 = bun5;
            Bun6 = bun6;
            Bun7 = bun7;
            //Bun8 = bun8;
            Bun9 = bun9;
            Binvert = binvert;
            B360 = b360;
            Keur = keur;
            Kmod = kmod;
            Bsup = bsup;
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
            Mon = mon;
            In1 = in1;
            In2 = in2;
            In3 = in3;
            In4 = in14;
            In5 = in5;
            In6 = in6;
            In7 = in7;
            In8 = in8;
            In9 = in9;
            In10 = in10;
            In11 = in11;
            In12 = in12;
            In13 = in13;
            In14 = in14;
            In15 = in15;
            In16 = in16;
            In17 = in17;
            In18 = in18;
            In19 = in19;
            In20 = in20;
            In21 = in21;
            In22 = in22;
            In23 = in23;
            Ip1 = ip1;
            Ip2 = ip2;
            Ip3 = ip3;
            Ip4 = ip4;
            Ip5 = ip5;
            Ip6 = ip6;
            Ip7 = ip7;
            Ip8 = ip8;
            Ip9 = ip9;
            Ip10 = ip10;
            Ip11 = ip11;
            Ip12 = ip12;
            Ip13 = ip13;
            Ip14 = ip14;
            Ip15 = ip15;
            Ip16 = ip16;
            Ip17 = ip17;
            Ip18 = ip18;
            Ip19 = ip19;
            Ip20 = ip20;
            Ip21 = ip21;
            Ip23 = ip23;

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
        public string GetDmc1()
        {
            return Dmc1;
        }
        public string GetDmc11()
        {
            return Dmc11;
        }
        public string GetDmc12()
        {
            return Dmc12;
        }
        public string GetDlq1()
        {
            return Dlq1;
        }
        public string GetDlq2()
        {
            return Dlq2;
        }
        public string GetDlq3()
        {
            return Dlq3;
        }
        public string GetDlq11()
        {
            return Dlq11;
        }
        public string GetDlq21()
        {
            return Dlq21;
        }
        public string GetDlq31()
        {
            return Dlq31;
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
        public string GetDsu105()
        {
            return Dsu105;
        }
        public string GetDsu205()
        {
            return Dsu205;
        }
        public string GetDsu305()
        {
            return Dsu305;
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
        public string GetDmc7()
        {
            return Dmc7;
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
        public string GetPcmc1()
        {
            return Pcmc1;
        }
        public string GetPcmc2()
        {
            return Pcmc2;
        }
        public string GetPcmc3()
        {
            return Pcmc3;
        }
        public string GetPcond()
        {
            return Pcond;
        }
        public string GetPlq()
        {
            return Plq;
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
        public string GetPcv()
        {
            return Pcv;
        }
        public string GetPcx()
        {
            return Pcx;
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
        public string GetPdtr()
        {
            return Pdtr;
        }
        public string GetPdtc()
        {
            return Pdtc;
        }
        public string GetPdt2()
        {
            return Pdt2;
        }
        public string GetPdt1()
        {
            return Pdt1;
        }
        public string GetCp80()
        {
            return Cp80;
        }
        public string GetCp100()
        {
            return Cp100;
        }
        public string GetCp120()
        {
            return Cp120;
        }
        public string GetCp150()
        {
            return Cp150;
        }
        public string GetCt80()
        {
            return Ct80;
        }
        public string GetCt100()
        {
            return Ct100;
        }
        public string GetCt120()
        {
            return Ct120;
        }
        public string GetCt150()
        {
            return Ct150;
        }
        public string GetCp80m()
        {
            return Cp80m;
        }
        public string GetCp100m()
        {
            return Cp100m;
        }
        public string GetCp120m()
        {
            return Cp120m;
        }
        public string GetCp150m()
        {
            return Cp150m;
        }
        public string GetCt80m()
        {
            return Ct80m;
        }
        public string GetCt100m()
        {
            return Ct100m;
        }
        public string GetCt120m()
        {
            return Ct120m;
        }
        public string GetCt150m()
        {
            return Ct150m;
        }

        public string GetPtp84()
        {
            return Ptp84;
        }
        public string GetPtp83()
        {
            return Ptp83;
        }
        public string GetPtp78()
        {
            return Ptp78;
        }
        public string GetPtp76()
        {
            return Ptp76;
        }
        public string GetPtp75()
        {
            return Ptp75;
        }
        public string GetPtp74()
        {
            return Ptp74;
        }
        public string GetPtp73()
        {
            return Ptp73;
        }
        public string GetPtp72()
        {
            return Ptp72;
        }
        public string GetPtp85()
        {
            return Ptp85;
        }
        public string GetPtp86()
        {
            return Ptp86;
        }
        public string GetPtp87()
        {
            return Ptp87;
        }
        public string GetPtp88()
        {
            return Ptp88;
        }
        public string GetPtp89()
        {
            return Ptp89;
        }
        public string GetPtp90()
        {
            return Ptp90;
        }
        public string GetPtp91()
        {
            return Ptp91;
        }
        public string GetPtp92()
        {
            return Ptp92;
        }
        public string GetPtp93()
        {
            return Ptp93;
        }
        public string GetPtp94()
        {
            return Ptp94;
        }
        public string GetPtp95()
        {
            return Ptp95;
        }
        public string GetPtp96()
        {
            return Ptp96;
        }
        public string GetPtp97()
        {
            return Ptp97;
        }
        public string GetPtp98()
        {
            return Ptp98;
        }
        public string GetPtp99()
        {
            return Ptp99;
        }
        public string GetPtp100()
        {
            return Ptp100;
        }
        public string GetPtp101()
        {
            return Ptp101;
        }
        public string GetPtp102()
        {
            return Ptp102;
        }
        public string GetPtp103()
        {
            return Ptp103;
        }
        public string GetPtp104()
        {
            return Ptp104;
        }
        public string GetPtp105()
        {
            return Ptp105;
        }
        public string GetPtp106()
        {
            return Ptp106;
        }
        public string GetPtp107()
        {
            return Ptp107;
        }
        public string GetPtp108()
        {
            return Ptp108;
        }
        public string GetPtp109()
        {
            return Ptp109;
        }
        public string GetPtp110()
        {
            return Ptp110;
        }
        public string GetPtp111()
        {
            return Ptp111;
        }
        public string GetPtp136()
        {
            return Ptp136;
        }
        public string GetPtp137()
        {
            return Ptp137;
        }
        public string GetPtp141()
        {
            return Ptp141;
        }
        public string GetPtp142()
        {
            return Ptp142;
        }
        public string GetPtp143()
        {
            return Ptp143;
        }
        public string GetPtp144()
        {
            return Ptp144;
        }
        public string GetPtp145()
        {
            return Ptp145;
        }
        public string GetPtp146()
        {
            return Ptp146;
        }
        public string GetPtp147()
        {
            return Ptp147;
        }
        public string GetPtp148()
        {
            return Ptp148;
        }
        public string GetPtp149()
        {
            return Ptp149;
        }
        public string GetPtp150()
        {
            return Ptp150;
        }
        public string GetPtp151()
        {
            return Ptp151;
        }
        public string GetInc()
        {
            return Inc;
        }

        public string GetInev()
        {
            return Inev;
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
        public string GetQevpc()
        {
            return Qevpc;
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
        public string GetMon()
        {
            return Mon;
        }

        public string GetIn1()
        {
            return In1;
        }
        public string GetIn2()
        {
            return In2;
        }
        public string GetIn3()
        {
            return In3;
        }
        public string GetIn4()
        {
            return In4;
        }
        public string GetIn5()
        {
            return In5;
        }
        public string GetIn6()
        {
            return In6;
        }
        public string GetIn7()
        {
            return In7;
        }
        public string GetIn8()
        {
            return In8;
        }
        public string GetIn9()
        {
            return In9;
        }
        public string GetIn10()
        {
            return In10;
        }
        public string GetIn11()
        {
            return In11;
        }
        public string GetIn12()
        {
            return In12;
        }
        public string GetIn13()
        {
            return In13;
        }
        public string GetIn14()
        {
            return In14;
        }
        public string GetIn15()
        {
            return In15;
        }
        public string GetIn16()
        {
            return In16;
        }
        public string GetIn17()
        {
            return In17;
        }
        public string GetIn18()
        {
            return In18;
        }
        public string GetIn19()
        {
            return In19;
        }
        public string GetIn20()
        {
            return In20;
        }
        public string GetIn21()
        {
            return In21;
        }
        public string GetIn22()
        {
            return In22;
        }
        public string GetIn23()
        {
            return In23;
        }

        public string GetIp1()
        {
            return Ip1;
        }
        public string GetIp2()
        {
            return Ip2;
        }
        public string GetIp3()
        {
            return Ip3;
        }
        public string GetIp4()
        {
            return Ip4;
        }
        public string GetIp5()
        {
            return Ip5;
        }
        public string GetIp6()
        {
            return Ip6;
        }
        public string GetIp7()
        {
            return Ip7;
        }
        public string GetIp8()
        {
            return Ip8;
        }
        public string GetIp9()
        {
            return Ip9;
        }
        public string GetIp10()
        {
            return Ip10;
        }
        public string GetIp11()
        {
            return Ip11;
        }
        public string GetIp12()
        {
            return Ip12;
        }
        public string GetIp13()
        {
            return Ip13;
        }
        public string GetIp14()
        {
            return Ip14;
        }
        public string GetIp15()
        {
            return Ip15;
        }
        public string GetIp16()
        {
            return Ip16;
        }
        public string GetIp17()
        {
            return Ip17;
        }
        public string GetIp18()
        {
            return Ip18;
        }
        public string GetIp19()
        {
            return Ip19;
        }
        public string GetIp20()
        {
            return Ip20;
        }
        public string GetIp21()
        {
            return Ip21;
        }
        public string GetIp23()
        {
            return Ip23;
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
        public void SetDmc1(string dmc1)
        {
            Dmc1 = dmc1;
        }
        public void SetDmc11(string dmc11)
        {
            Dmc11 = dmc11;
        }
        public void SetDmc12(string dmc12)
        {
            Dmc12 = dmc12;
        }
        public void SetDlq1(string dlq1)
        {
            Dlq1 = dlq1;
        }
        public void SetDlq2(string dlq2)
        {
            Dlq2 = dlq2;
        }
        public void SetDlq3(string dlq3)
        {
            Dlq3 = dlq3;
        }
        public void SetDlq11(string dlq11)
        {
            Dlq11 = dlq11;
        }
        public void SetDlq21(string dlq21)
        {
            Dlq21 = dlq21;
        }
        public void SetDlq31(string dlq31)
        {
            Dlq31 = dlq31;
        }
        public void SetDsu130(string dsu130)
        {
            Dsu130 = dsu130;
        }
        public void SetDsu230(string dsu230)
        {
            Dsu230 = dsu230;
        }
        public void SetDsu330(string dsu330)
        {
            Dsu330 = dsu330;
        }
        public void SetDsu110(string dsu110)
        {
            Dsu110 = dsu110;
        }
        public void SetDsu210(string dsu210)
        {
            Dsu210 = dsu210;
        }
        public void SetDsu310(string dsu310)
        {
            Dsu310 = dsu310;
        }
        public void SetDsu105(string dsu105)
        {
            Dsu105 = dsu105;
        }
        public void SetDsu205(string dsu205)
        {
            Dsu205 = dsu205;
        }
        public void SetDsu305(string dsu305)
        {
            Dsu305 = dsu305;
        }
        
        public void SetDmc2(string dmc2)
        {
            Dmc2 = dmc2;
        }
        public void SetDmc3(string dmc3)
        {
            Dmc3 = dmc3;
        }

        public void SetDmc4(string dmc4)
        {
            Dmc4 = dmc4;
        }
        public void SetDmc5(string dmc5)
        {
            Dmc5 = dmc5;
        }
        public void SetDmc6(string dmc6)
        {
            Dmc6 = dmc6;
        }
        public void SetDmc7(string dmc7)
        {
            Dmc7 = dmc7;
        }
        public void SetDmc8(string dmc8)
        {
            Dmc8 = dmc8;
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
        public void SetPcmc1(string pcmc1)
        {
            Pcmc1 = pcmc1;
        }
        public void SetPcmc2(string pcmc2)
        {
            Pcmc2 = pcmc2;
        }
        public void SetPcmc3(string pcmc3)
        {
            Pcmc3 = pcmc3;
        }
        public void SetPcond(string pcond)
        {
            Pcond = pcond;
        }
        public void SetPlq(string plq)
        {
            Plq = plq;
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
        public void SetPcv(string pcv)
        {
            Pcv = pcv;
        }
        public void SetPcx(string pcx)
        {
            Pcx = pcx;
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
        public void SetPdtr(string pdtr)
        {
            Pdtr = pdtr;
        }
        public void SetPdtc(string pdtc)
        {
            Pdtc = pdtc;
        }
        public void SetPdt2(string pdt2)
        {
            Pdt2 = pdt2;
        }
        public void SetPdt1(string pdt1)
        {
            Pdt1 = pdt1;
        }
        public void SetCp80(string cp80)
        {
            Cp80 = cp80;
        }
        public void SetCp100(string cp100)
        {
            Cp100 = cp100;
        }
        public void SetCp120(string cp120)
        {
            Cp120 = cp120;
        }
        public void SetCp150(string cp150)
        {
            Cp150 = cp150;
        }
        public void SetCt80(string ct80)
        {
            Ct80 = ct80;
        }
        public void SetCt100(string ct100)
        {
            Ct100 = ct100;
        }
        public void SetCt120(string ct120)
        {
            Ct120 = ct120;
        }
        public void SetCt150(string ct150)
        {
            Ct150 = ct150;
        }
        public void SetCp80m(string cp80m)
        {
            Cp80m = cp80m;
        }
        public void SetCp100m(string cp100m)
        {
            Cp100m = cp100m;
        }
        public void SetCp120m(string cp120m)
        {
            Cp120m = cp120m;
        }
        public void SetCp150m(string cp150m)
        {
            Cp150m = cp150m;
        }
        public void SetCt80m(string ct80m)
        {
            Ct80m = ct80m;
        }
        public void SetCt100m(string ct100m)
        {
            Ct100m = ct100m;
        }
        public void SetCt120m(string ct120m)
        {
            Ct120m = ct120m;
        }
        public void SetCt150m(string ct150m)
        {
            Ct150m = ct150m;
        }

        public void SetPtp84(string ptp84)
        {
            Ptp84 = ptp84;
        }
        public void SetPtp83(string ptp83)
        {
            Ptp83 = ptp83;
        }
        public void SetPtp78(string ptp78)
        {
            Ptp78 = ptp78;
        }
        public void SetPtp76(string ptp76)
        {
            Ptp76 = ptp76;
        }
        public void SetPtp75(string ptp75)
        {
            Ptp75 = ptp75;
        }
        public void SetPtp74(string ptp74)
        {
            Ptp74 = ptp74;
        }
        public void SetPtp73(string ptp73)
        {
            Ptp73 = ptp73;
        }
        public void SetPtp72(string ptp72)
        {
            Ptp72 = ptp72;
        }
        public void SetPtp85(string ptp85)
        {
            Ptp85 = ptp85;
        }
        public void SetPtp86(string ptp86)
        {
            Ptp86 = ptp86;
        }
        public void SetPtp87(string ptp87)
        {
            Ptp87 = ptp87;
        }
        public void SetPtp88(string ptp88)
        {
            Ptp88 = ptp88;
        }
        public void SetPtp89(string ptp89)
        {
            Ptp89 = ptp89;
        }
        public void SetPtp90(string ptp90)
        {
            Ptp90 = ptp90;
        }
        public void SetPtp91(string ptp91)
        {
            Ptp91 = ptp91;
        }
        public void SetPtp92(string ptp92)
        {
            Ptp92 = ptp92;
        }
        public void SetPtp93(string ptp93)
        {
            Ptp93 = ptp93;
        }
        public void SetPtp94(string ptp94)
        {
            Ptp94 = ptp94;
        }
        public void SetPtp95(string ptp95)
        {
            Ptp95 = ptp95;
        }
        public void SetPtp96(string ptp96)
        {
            Ptp96 = ptp96;
        }
        public void SetPtp97(string ptp97)
        {
            Ptp97 = ptp97;
        }
        public void SetPtp98(string ptp98)
        {
            Ptp98 = ptp98;
        }
        public void SetPtp99(string ptp99)
        {
            Ptp99 = ptp99;
        }
        public void SetPtp100(string ptp100)
        {
            Ptp100 = ptp100;
        }
        public void SetPtp101(string ptp101)
        {
            Ptp101 = ptp101;
        }
        public void SetPtp102(string ptp102)
        {
            Ptp102 = ptp102;
        }
        public void SetPtp103(string ptp103)
        {
            Ptp103 = ptp103;
        }
        public void SetPtp104(string ptp104)
        {
            Ptp104 = ptp104;
        }
        public void SetPtp105(string ptp105)
        {
            Ptp105 = ptp105;
        }
        public void SetPtp106(string ptp106)
        {
            Ptp106 = ptp106;
        }
        public void SetPtp107(string ptp107)
        {
            Ptp107 = ptp107;
        }
        public void SetPtp108(string ptp108)
        {
            Ptp108 = ptp108;
        }
        public void SetPtp109(string ptp109)
        {
            Ptp109 = ptp109;
        }
        public void SetPtp110(string ptp110)
        {
            Ptp110 = ptp110;
        }
        public void SetPtp111(string ptp111)
        {
            Ptp111 = ptp111;
        }
        public void SetPtp136(string ptp136)
        {
            Ptp136 = ptp136;
        }
        public void SetPtp137(string ptp137)
        {
            Ptp137 = ptp137;
        }
        public void SetPtp141(string ptp141)
        {
            Ptp141 = ptp141;
        }
        public void SetPtp142(string ptp142)
        {
            Ptp142 = ptp142;
        }
        public void SetPtp143(string ptp143)
        {
            Ptp143 = ptp143;
        }
        public void SetPtp144(string ptp144)
        {
            Ptp144 = ptp144;
        }
        public void SetPtp145(string ptp145)
        {
            Ptp145 = ptp145;
        }
        public void SetPtp146(string ptp146)
        {
            Ptp146 = ptp146;
        }
        public void SetPtp147(string ptp147)
        {
            Ptp147 = ptp147;
        }
        public void SetPtp148(string ptp148)
        {
            Ptp148 = ptp148;
        }
        public void SetPtp149(string ptp149)
        {
            Ptp149 = ptp149;
        }
        public void SetPtp150(string ptp150)
        {
            Ptp150 = ptp150;
        }
        public void SetPtp151(string ptp151)
        {
            Ptp151 = ptp151;
        }
        public void SetInc(string inc)
        {
            Inc = inc;
        }
        public void SetInev(string inev)
        {
            Inev = inev;
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
        public void SetQevpc(string qevpc)
        {
            Qevpc = qevpc;
        }

        public void SetTint(string tint)
        {
            Tint = tint;
        }
        public void SetEquip(string equip)
        {
            Equip = equip;
        }

        public void SetDT(string dt)
        {
            DT = dt;
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
        public void SetMon(string mon)
        {
            Mon = mon;
        }

        public void SetIn1(string in1)
        {
            In1 = in1;
        }
        public void SetIn2(string in2)
        {
            In2 = in2;
        }
        public void SetIn3(string in3)
        {
            In3 = in3;
        }
        public void SetIn4(string in4)
        {
            In4 = in4;
        }
        public void SetIn5(string in5)
        {
            In5 = in5;
        }
        public void SetIn6(string in6)
        {
            In6 = in6;
        }
        public void SetIn7(string in7)
        {
            In7 = in7;
        }
        public void SetIn8(string in8)
        {
            In8 = in8;
        }
        public void SetIn9(string in9)
        {
            In9 = in9;
        }
        public void SetIn10(string in10)
        {
            In10 = in10;
        }
        public void SetIn11(string in11)
        {
            In11 = in11;
        }
        public void SetIn12(string in12)
        {
            In12 = in12;
        }
        public void SetIn13(string in13)
        {
            In13 = in13;
        }
        public void SetIn14(string in14)
        {
            In14 = in14;
        }
        public void SetIn15(string in15)
        {
            In15 = in15;
        }
        public void SetIn16(string in16)
        {
            In16 = in16;
        }
        public void SetIn17(string in17)
        {
            In17 = in17;
        }
        public void SetIn18(string in18)
        {
            In18 = in18;
        }
        public void SetIn19(string in19)
        {
            In19 = in19;
        }
        public void SetIn20(string in20)
        {
            In20 = in20;
        }
        public void SetIn21(string in21)
        {
            In21 = in21;
        }
        public void SetIn22(string in22)
        {
            In22 = in22;
        }
        public void SetIn23(string in23)
        {
            In23 = in23;
        }
        public void SetIp1(string ip1)
        {
            Ip1 = ip1;
        }
        public void SetIp2(string ip2)
        {
            Ip2 = ip2;
        }
        public void SetIp3(string ip3)
        {
            Ip3 = ip3;
        }
        public void SetIp4(string ip4)
        {
            Ip4 = ip4;
        }
        public void SetIp5(string ip5)
        {
            Ip5 = ip5;
        }
        public void SetIp6(string ip6)
        {
            Ip6 = ip6;
        }
        public void SetIp7(string ip7)
        {
            Ip7 = ip7;
        }
        public void SetIp8(string ip8)
        {
            Ip8 = ip8;
        }
        public void SetIp9(string ip9)
        {
            Ip9 = ip9;
        }
        public void SetIp10(string ip10)
        {
            Ip10 = ip10;
        }
        public void SetIp11(string ip11)
        {
            Ip11 = ip11;
        }
        public void SetIp12(string ip12)
        {
            Ip12 = ip12;
        }
        public void SetIp13(string ip13)
        {
            Ip13 = ip13;
        }
        public void SetIp14(string ip14)
        {
            Ip14 = ip14;
        }
        public void SetIp15(string ip15)
        {
            Ip15 = ip15;
        }
        public void SetIp16(string ip16)
        {
            Ip16 = ip16;
        }
        public void SetIp17(string ip17)
        {
            Ip17 = ip17;
        }
        public void SetIp18(string ip18)
        {
            Ip18 = ip18;
        }
        public void SetIp19(string ip19)
        {
            Ip19 = ip19;
        }
        public void SetIp20(string ip20)
        {
            Ip20 = ip20;
        }
        public void SetIp21(string ip21)
        {
            Ip21 = ip21;
        }
        public void SetIp23(string ip23)
        {
            Ip23 = ip23;
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
        public string GetResinaM()
        { 
            return ResinaM;
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

        public string GetConstCivPan()
        {
            return ConstCivPan;
        }

        public string GetEquipFrig()
        {
            return EquipFrig;
        }
        public string GetPuertasFrig()
        {
            return PuertasFrig;
        }
        public string GetDesE()
        {
            return DesE;
        }

        public string GetTasa()
        {
            return Tasa;
        }

        public string GetDsc()
        {
            return Dsc;
        }

        public string GetGastosAdmObra()
        {
            return GastosAdmObra;
        }

        public string GetGastosIndObra()
        {
            return GastosIndObra;
        }
        public string GetGastosIndObracuc()
        {
            return GastosIndObracuc;
        }
        public string GetCredito()
        {
            return Credito;
        }
        public string GetCreditocup()
        {
            return Creditocup;
        }
        public string GetCRcivil()
        {
            return CRcivil;
        }
        public string GetCRpiso()
        {
            return CRpiso;
        }



        //------------------------------------------------------------------
        public void SetDigt(string digt)
        {
            Digt = digt;
        }
        
        public void SetCcion(string ccion)
        {
            Ccion = ccion;
        }
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
        public void SetResinaM(string resina)
        {
            ResinaM = resina;
        }

        public void Setinc(string Inc)
        {
            inc = Inc;
        }
        
        public void SetConstCivPan(string constCiv)
        {
            ConstCivPan = constCiv;
        }

        public void SetEquipFrig(string equip)
        {
            EquipFrig = equip;
        }
        public void SetPuertasFrig(string puertas)
        {
            PuertasFrig = puertas;
        }
        public void SetDesE(string desE)
        {
            DesE = desE;
        }

        public void SetTasa(string tasa)
        {
            Tasa = tasa;
        }
        public void SetDsc(string dsc)
        {
            Dsc = dsc;
        }

        public void SetGastosAdmObra(string gastosa)
        {
            GastosAdmObra = gastosa;
        }

        public void SetGastosIndObra(string gastos)
        {
            GastosIndObra = gastos;
        }

        public void SetGastosIndObracuc(string gastoscuc)
        {
            GastosIndObracuc = gastoscuc;
        }
        public void SetCredito(string credito)
        {
            Credito = credito;
        }
        public void SetCreditocup(string creditocup)
        {
            Creditocup = creditocup;
        }
        public void SetCRcivil(string crcivil)
        {
            CRcivil = crcivil;
        }
        public void SetCRpiso(string crpiso)
        {
            CRpiso = crpiso;
        }

        public void SetLugar(string lugar)
        {
            Lugar = lugar;
        }
       
        public void SetBmoni(bool bmoni)
        {
            Bmoni = bmoni;
        }
        public void SetB60H(bool b60h)
        {
            B60H = b60h;
        }
        public void SetBun(bool bun)
        {
            Bun = bun;
        }
        public void SetBun2(bool bun2)
        {
            Bun2 = bun2;
        }
        public void SetBun3(bool bun3)
        {
            Bun3 = bun3;
        }
        public void SetBun4(bool bun4)
        {
            Bun4 = bun4;
        }
        public void SetBun5(bool bun5)
        {
            Bun5 = bun5;
        }
        public void SetBun6(bool bun6)
        {
            Bun6 = bun6;
        }
        public void SetBun7(bool bun7)
        {
            Bun7 = bun7;
        }
        public void SetBun8(bool bun8)
        {
            Bun8 = bun8;
        }
        public void SetBun9(bool bun9)
        {
            Bun9 = bun9;
        }
        public void SetBinvert(bool binvert)
        {
            Binvert = binvert;
        }
        public void SetB360(bool b360)
        {
            B360 = b360;
        }
            
        public void SetKeur(bool keur)
        {
            Keur = keur;
        }
        public void SetBsup(bool bsup)
        {
            Bsup = bsup;
        }
        public string GetLugar()
        {
           return Lugar;
        }
        
        public bool GetBmoni()
        {
            return Bmoni;
        }
        public bool GetB60H()
        {
            return B60H;
        }
        public bool GetBun()
        {
            return Bun;
        }
        public bool GetBun2()
        {
            return Bun2;
        }
        public bool GetBun3()
        {
            return Bun3;
        }
        public bool GetBun4()
        {
            return Bun4;
        }
        public bool GetBun5()
        {
            return Bun5;
        }
        public bool GetBun6()
        {
            return Bun6;
        }
        public bool GetBun7()
        {
            return Bun7;
        }
        public bool GetBun8()
        {
            return Bun8;
        }
        public bool GetBun9()
        {
            return Bun9;
        }
        public bool GetBinvert()
        {
            return Binvert;
        }
        public bool GetB360()
        {
            return B360;
        }
        public bool GetKeur()
        {
            return Keur;
        }
        public bool GetBsup()
        {
            return Bsup;
        }
        
                
      } 
}
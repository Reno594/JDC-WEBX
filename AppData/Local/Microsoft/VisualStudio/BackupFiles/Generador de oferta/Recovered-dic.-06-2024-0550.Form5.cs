using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Security.Cryptography;
using System.Text;
using System.Windows.Forms;

using static System.Windows.Forms.VisualStyles.VisualStyleElement;

using System.Data.SqlClient;
using static System.ComponentModel.Design.ObjectSelectorEditor;
using System.IO;
using System.Data.SqlTypes;
using Microsoft.Azure.Amqp.Transaction;
using System.Globalization;
using Microsoft.Data.SqlClient.Server;
using iTextSharp.text;
using Microsoft.Azure.Amqp.Framing;
using Microsoft.Office.Interop.Excel;
using System.Net.NetworkInformation;
using System.Text.RegularExpressions;

namespace Generador_de_oferta
{

    public partial class Form5 : Form
    {

        // Declaración de las variables Hur1 y Hur2
        private string Hur1;
        private string Hur2;

        private Form1 _localForm1;

        public object PTevapS { get; private set; }

        // Método para manejar el evento TextChanged


        public Form5(Form1 local)
        {
            InitializeComponent();
            _localForm1 = local;
            Datos.ActualcamChanged += OnActualcamChanged;
            CargarValores();
        }

        public Form5()
        {
        }
        private void OnActualcamChanged()
        {
            Console.WriteLine("OnActualcamChanged called in Form5"); // Mensaje de depuración
            CargarValores();
        }
        private void CargarValores()
        {

            Console.WriteLine($"Datos.actualcam: {Datos.actualcam}");
            TNcam5.Text = Datos.actualcam;
            CSerie.Text = Datos.CSerie;
            TPcp.Text = Datos.TFWe.ToString();
            PTevap.Text = Datos.CTEvap;
            PEcalc.Text = Datos.PEcalc;
            TRdin.Text = Datos.TRdin;
            TCentx.Text = Datos.TCentd.ToString();
            TTcmc8.Text = Datos.Tcmc8;
            TeVP.Text = Datos.TeVP;
            NEvap.Text = Datos.NEvap;
            DTpd.Text = Datos.DTpd;
            DTsl.Text = Datos.DTsl;
            Rfcalx.Text = Datos.Rfcalx;
            SCdt.Text = Datos.SCdt;
            TTcmc6.Text = Datos.Tcmc6;
            TValv.Text = Datos.TValv;
            TQevp.Text = Datos.TQevp;
            Hur1 = Datos.Hur11.ToString();
            Hur2 = Datos.Hur21.ToString();
            TNcam5.Text = Datos.TPem;
            FbcT.Text = Datos.TCsist1;
            AutR.Text = Datos.AutR;
            SCdt.Text = Datos.SCdt;
        }

        private void Guardar_Click(object sender, EventArgs e)
        {
            Datos.TQevp = TQevp.Text;
            Datos.CSerie = CSerie.Text;
            Datos.CTEvap = PTevap.Text;
            Datos.PEcalc = PEcalc.Text;
            Datos.TRdin = TRdin.Text;
            Datos.TCentd = Convert.ToInt32(TCentx.Text);
            Datos.TeVP = TeVP.Text;
            Datos.NEvap = NEvap.Text;
            Datos.DTpd = DTpd.Text;
            Datos.DTsl = DTsl.Text;
            Datos.Rfcalx = Rfcalx.Text;
            Datos.SCdt = SCdt.Text;
            Datos.Tcmc6 = TTcmc6.Text;
            Datos.TValv = TValv.Text;
            Datos.Tcmc8 = TTcmc8.Text;
            Datos.Hur11 = Convert.ToInt16(Hur1);
            Datos.Hur21 = Convert.ToInt16(Hur2);
            Datos.TPcp = Convert.ToDecimal(TPcp.Text);
            Datos.Tcmc8 = TTcmc8.Text;
            Datos.TPem = TNcam5.Text;
            Datos.TCsist1 = FbcT.Text;
            Datos.AutR = AutR.Text;
            Datos.SCdt = SCdt.Text;

            foreach (Form form in System.Windows.Forms.Application.OpenForms)
            {
                if (form is Form1 form1)
                {
                    form1.GuardarValores();
                    break;
                }
            }

            _localForm1.GuardarValores();
            // Cerrar Form5
            this.Close();
        }

        private void TQevp_SelectedIndexChanged(object sender, EventArgs e)
        {
            CSerie.Items.Clear();

            if (TQevp.Text == "4")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEA", "HED", "HEDP", "HEC" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "MVP" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "6")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEA", "HED", "HEC" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "GRB" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "9")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEA", "HED", "HEDP", "HEC" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "TTL", "PIL", "MRL", "GRL", "FRL" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "4.5")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEB", "HED", "HEC", "HER" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "7")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEB", "HED", "HEBF", "HEC", "HEP" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "TTB", "PIB", "MVB", "MRB", "FRB", "ECB" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "10")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEB", "HED", "HEC", "HEF" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "12")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEB", "HED", "HEC", "HEF" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "TTX", "MRX", "GRX", "TTX", "FRL" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "3.5")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "HEP", "HER" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else if (TQevp.Text == "2.8")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "PIA" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }

            else if (TQevp.Text == "4.2")
            {
                if (FbcT.Text == "Hispania")
                {
                    string[] installs = new string[] { "TTM", "GRM", "TTM", "PIM", "MVM", "FRM", "FCM", "FBV", "ECM" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Frimetal")
                {
                    string[] installs = new string[] { "PIA" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Ser")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Kobol")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "FrigaBohn")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Guntner")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Escofred")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Sereva")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
                if (FbcT.Text == "Eco")
                {
                    string[] installs = new string[] { "" };
                    CSerie.Items.AddRange(installs);
                }
            }
            else
            {
                CSerie.Text = "";
            }
            this.Controls.Add(this.CSerie);
        }

        private void Rfcalx_SelectedIndexChanged(object sender, EventArgs e)
        {
            //Refrigerant Selection
            int Nrfg = 0;
            int Nffg1 = 0;
            int Nfrg2 = 0;

            if (Rfcalx.Text == "R410A") { Nrfg = 1; Nfrg2 = 2; Nfrg2 = 2; }
            if (Rfcalx.Text == "R131a") { Nrfg = 2; Nfrg2 = 3; Nfrg2 = 3; }
            if (Rfcalx.Text == "R22") { Nrfg = 3; Nfrg2 = 4; Nfrg2 = 4; }
            if (Rfcalx.Text == "R404A") { Nrfg = 4; Nfrg2 = 5; Nfrg2 = 6; }
            if (Rfcalx.Text == "R507") { Nrfg = 5; Nfrg2 = 7; Nfrg2 = 7; }
            if (Rfcalx.Text == "R407C") { Nrfg = 6; Nfrg2 = 8; Nfrg2 = 9; }
            if (Rfcalx.Text == "R23") { Nrfg = 7; Nfrg2 = 10; Nfrg2 = 10; }
            if (Rfcalx.Text == "R454A") { Nrfg = 8; Nfrg2 = 13; Nfrg2 = 14; }
            if (Rfcalx.Text == "R454C") { Nrfg = 9; Nfrg2 = 15; Nfrg2 = 16; }
            if (Rfcalx.Text == "RR1234yf") { Nrfg = 10; Nfrg2 = 17; Nfrg2 = 17; }
            if (Rfcalx.Text == "RR1234ze") { Nrfg = 11; Nfrg2 = 18; Nfrg2 = 18; }
            if (Rfcalx.Text == "R32") { Nrfg = 12; Nfrg2 = 21; Nfrg2 = 21; }
            if (Rfcalx.Text == "R452A") { Nrfg = 13; Nfrg2 = 22; Nfrg2 = 23; }
            if (Rfcalx.Text == "R448A") { Nrfg = 14; Nfrg2 = 24; Nfrg2 = 25; }
            if (Rfcalx.Text == "R449A") { Nrfg = 15; Nfrg2 = 26; Nfrg2 = 27; }
            if (Rfcalx.Text == "R450A") { Nrfg = 16; Nfrg2 = 28; Nfrg2 = 29; }
            if (Rfcalx.Text == "R513A") { Nrfg = 17; Nfrg2 = 30; Nfrg2 = 31; }
            if (Rfcalx.Text == "R455A") { Nrfg = 18; Nfrg2 = 34; Nfrg2 = 35; }

            //Press. regulators
            int Press = 0;
            string PressR = "";
            int Press1 = 0;
            int Press2 = 0;
            if (Press == 1) { PressR = "R134a"; Press1 = 3; Press2 = 3; }
            if (Press == 2) { PressR = "R22"; Press1 = 4; Press2 = 4; }
            if (Press == 3) { PressR = "R404A"; Press1 = 5; Press2 = 6; }
            if (Press == 4) { PressR = "R507"; Press1 = 7; Press2 = 7; }
            if (Press == 5) { PressR = "R407C"; Press1 = 8; Press2 = 9; }
        }

        private void OExcel_Click(object sender, EventArgs e)
        {
            //Refrigerant Selection
            int Nrfg = 0;
            int Nffg1 = 0;
            int Nfrg2 = 0;

            if (Rfcalx.Text == "R410A") { Nrfg = 1; Nfrg2 = 2; Nfrg2 = 2; }
            if (Rfcalx.Text == "R131a") { Nrfg = 2; Nfrg2 = 3; Nfrg2 = 3; }
            if (Rfcalx.Text == "R22") { Nrfg = 3; Nfrg2 = 4; Nfrg2 = 4; }
            if (Rfcalx.Text == "R404A") { Nrfg = 4; Nfrg2 = 5; Nfrg2 = 6; }
            if (Rfcalx.Text == "R507") { Nrfg = 5; Nfrg2 = 7; Nfrg2 = 7; }
            if (Rfcalx.Text == "R407C") { Nrfg = 6; Nfrg2 = 8; Nfrg2 = 9; }
            if (Rfcalx.Text == "R23") { Nrfg = 7; Nfrg2 = 10; Nfrg2 = 10; }
            if (Rfcalx.Text == "R454A") { Nrfg = 8; Nfrg2 = 13; Nfrg2 = 14; }
            if (Rfcalx.Text == "R454C") { Nrfg = 9; Nfrg2 = 15; Nfrg2 = 16; }
            if (Rfcalx.Text == "RR1234yf") { Nrfg = 10; Nfrg2 = 17; Nfrg2 = 17; }
            if (Rfcalx.Text == "RR1234ze") { Nrfg = 11; Nfrg2 = 18; Nfrg2 = 18; }
            if (Rfcalx.Text == "R32") { Nrfg = 12; Nfrg2 = 21; Nfrg2 = 21; }
            if (Rfcalx.Text == "R452A") { Nrfg = 13; Nfrg2 = 22; Nfrg2 = 23; }
            if (Rfcalx.Text == "R448A") { Nrfg = 14; Nfrg2 = 24; Nfrg2 = 25; }
            if (Rfcalx.Text == "R449A") { Nrfg = 15; Nfrg2 = 26; Nfrg2 = 27; }
            if (Rfcalx.Text == "R450A") { Nrfg = 16; Nfrg2 = 28; Nfrg2 = 29; }
            if (Rfcalx.Text == "R513A") { Nrfg = 17; Nfrg2 = 30; Nfrg2 = 31; }
            if (Rfcalx.Text == "R455A") { Nrfg = 18; Nfrg2 = 34; Nfrg2 = 35; }

            //Press. regulators
            int Press = 0;
            string PressR = "";
            int Press1 = 0;
            int Press2 = 0;
            if (Press == 1) { PressR = "R134a"; Press1 = 3; Press2 = 3; }
            if (Press == 2) { PressR = "R22"; Press1 = 4; Press2 = 4; }
            if (Press == 3) { PressR = "R404A"; Press1 = 5; Press2 = 6; }
            if (Press == 4) { PressR = "R507"; Press1 = 7; Press2 = 7; }
            if (Press == 5) { PressR = "R407C"; Press1 = 8; Press2 = 9; }

            decimal Tve = 0;//te= ºC
            decimal Kwattp = 0;//Kw DT1=8K
            decimal Am2;//Superficie (m²)
            decimal Dm3;//Volumen Interno (dm3)
            decimal Kgn;//Peso Neto (kg)
            decimal Liq1;//Entrada 1
            decimal Liq2;//Entrada 2
            decimal Dmm;//Diámetro (Ф mm)
            decimal Fase;//Fase
            decimal Pdcw;//Potencia (W)
            decimal InA;//Intensidad (A)
            decimal Tiro;//Tiro de Aire m)"
            decimal TPdcw;//Total (W)
            decimal DimA;//A
            decimal DimB;//B
            decimal DimC;//C
            decimal Palet = 0;//Paso aleta @"\fnstptecuc.tpt"
            decimal Evap;
            decimal HKwatt;
            decimal MaxEp = 0;
            decimal Kwattd = 0;
            decimal Kwattm = 0;
            decimal DKwatt = 0;
            int LPs = 0;//Dt de temperatura calculado
            int Hur1 = 0;//Humedad relativa Maxima
            int Hur2 = 0;//Humedad relativa Minima
            Hur1 = Datos.Hur11;// Leer el valor de la humedad relativa
            Hur2 = Datos.Hur21;// Leer el valor de la humedad relativa

            string connectionString3 = "Data Source=LENOVO;Initial Catalog=Humedad;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString3))
            {
                // Abrir la conexión
                connection.Open();
                int Sum = Hur1 + Hur2;
                string F = "";
                if (Sum == 0) { F = "F150"; } else { F = "F" + Sum.ToString(); }
                // Verificar que Sum no sea 0
                if (Sum == 0)
                {
                    // Crear un formulario para ingresar los valores
                    Form inputForm = new Form();
                    System.Windows.Forms.Label label1 = new System.Windows.Forms.Label() { Left = 50, Top = 30, Text = "Humedad Máx:" };
                    System.Windows.Forms.TextBox textBox1 = new System.Windows.Forms.TextBox() { Left = 50, Top = 50, Width = 100 };
                    System.Windows.Forms.Label label2 = new System.Windows.Forms.Label() { Left = 50, Top = 80, Text = "Humedad Min:" };
                    System.Windows.Forms.TextBox textBox2 = new System.Windows.Forms.TextBox() { Left = 50, Top = 110, Width = 100 };
                    System.Windows.Forms.Button confirmation = new System.Windows.Forms.Button() { Text = "Aceptar", Left = 50, Top = 140, Width = 100, DialogResult = DialogResult.OK };
                    inputForm.Controls.Add(label1);
                    inputForm.Controls.Add(textBox1);
                    inputForm.Controls.Add(label2);
                    inputForm.Controls.Add(textBox2);
                    inputForm.Controls.Add(confirmation);
                    inputForm.AcceptButton = confirmation;

                    if (inputForm.ShowDialog() == DialogResult.OK)
                    {
                        if (int.TryParse(textBox1.Text, out int HumA) && int.TryParse(textBox2.Text, out int HumB))
                        {
                            Hur1 = HumA;
                            Hur2 = HumB;
                            Sum = Hur1 + Hur2;
                        }
                        else
                        {
                            MessageBox.Show("Error: Los valores ingresados no son válidos.");
                        }
                    }
                }
                
                
                string query = $"SELECT {F} FROM [Humedad].[dbo].[Humedad] WHERE F2 = 40";
                SqlCommand command = new SqlCommand(query, connection);

                // Ejecutar la consulta y obtener el resultado
                object result = command.ExecuteScalar();

                // Verificar si el resultado no es null (es decir, se encontró un valor)
                if (result != null)
                {
                    // Convertir el resultado a un entero (asumiendo que es un valor entero válido)
                    try
                    {
                        LPs = Convert.ToInt16(result);
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("Error: El resultado no es un valor entero válido.");
                    }
                }
                else
                {
                    MessageBox.Show("No se encontró ningún valor.");
                }
            }

            int selected = Datos.selected;
            string unidades1 = "";
            string unidades2 = "";
            string unidades3 = "";
            string unidades4 = "";
            if (selected == 1) { unidades1 = "Kw"; unidades2 = "°C"; unidades3 = "bar"; unidades4 = "%-Kw"; }
            if (selected == 2) { unidades1 = "tons"; unidades2 = "°F"; unidades3 = "Psi"; unidades4 = "%-tons"; }
            if (selected == 3) { unidades1 = "MBH"; unidades2 = "°F"; unidades3 = "Psi"; unidades4 = "%-MBH"; }
            int LPs1 = 0;
            decimal Kwatt2 = 0;
            if (DTsl.Text == "") { DTsl.Text = "AUT"; } else { if (DTsl.Text == "AUT") { LPs1 = LPs; } else { LPs1 = Convert.ToInt16(float.Parse(DTsl.Text)); } }

            DTpd.Text = LPs1.ToString();
            Kwatt2 = Datos.TFWe;
            if (selected == 1) { TPcp.Text = Math.Round((Kwatt2), 2, MidpointRounding.ToEven).ToString(); }
            if (selected == 2) { TPcp.Text = Math.Round((Kwatt2 * 0.284M), 2, MidpointRounding.ToEven).ToString(); }
            if (selected == 3) { TPcp.Text = Math.Round((Kwatt2 * 14.5038M), 2, MidpointRounding.ToEven).ToString(); }

            int NEvapc = 0;//Valor numerico del numero de evaporadores
            if (NEvap.Text == "") { NEvap.Text = "1"; NEvapc = Convert.ToInt16(NEvap.Text); } else { NEvapc = Convert.ToInt16(NEvap.Text); }
            Kwattp = Kwatt2 / (NEvapc);//Potencia de evaporacion por Evaporador de proyecto

            Datos.DTpde1 = LPs1.ToString();
            //Datos.TCsist1 = FbcT.Text;
            Datos.NEvap1 = NEvap.Text;
            decimal TCentx1 = 0;
            TCentx1 = Datos.TCentd;
            if (TeVP.Text == "") { TeVP.Text = "AUT"; TCentx.Text = TCentx1.ToString(); } else { if (TeVP.Text != "AUT") { TCentx1 = Convert.ToInt16(float.Parse(TeVP.Text)); } else { TCentx.Text = TCentx1.ToString(); } }
            int TcmF = 0;
            int TcmC = 0;
            decimal FR = 0;

            // Establecer la conexión con la base de datos
            string connectionString1 = "Data Source=LENOVO;Initial Catalog=Humedad;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString1))
            {
                // Abrir la conexión
                connection.Open();
                int Sum = Hur1 + Hur2;
                string F = "";
                if (Sum == 0) { F = "F150"; } else { F = "F" + Sum.ToString(); }

                // Crear un comando SQL para ejecutar la consulta
                string query = $"SELECT {F} FROM [Humedad].[dbo].[Humedad] WHERE F2 = " + TCentx1 + ""; // Renombrar a query1 para evitar duplicados
                SqlCommand command = new SqlCommand(query, connection);

                // Ejecutar la consulta y obtener el resultado
                object result = command.ExecuteScalar();

                // Verificar si el resultado no es null (es decir, se encontró un valor)
                if (result != null)
                {
                    // Convertir el resultado a un entero (asumiendo que es un valor entero válido)
                    try
                    {

                        FR = Convert.ToDecimal(result);
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("Error: El resultado no es un valor entero válido.");
                    }
                }
                else
                {

                }

            }

            // Establecer la conexión con la base de datos
            string connectionString2 = "Data Source=LENOVO;Initial Catalog=Humedad;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString2))

            {
                connection.Open();
                string query = $"SELECT F1 FROM [Humedad].[dbo].[Humedad] WHERE F2 = " + TCentx1 + ""; // Renombrar a query1 para evitar duplicados
                SqlCommand command = new SqlCommand(query, connection);

                // Ejecutar la consulta y obtener el resultado
                object result = command.ExecuteScalar();

                // Verificar si el resultado no es null (es decir, se encontró un valor)
                if (result != null)
                {
                    // Convertir el resultado a un entero (asumiendo que es un valor entero válido)
                    try
                    {

                        SCdt.Text = result.ToString();
                    }
                    catch (FormatException)
                    {
                        MessageBox.Show("Error: El resultado no es un valor entero válido.");
                    }
                }
                else
                {

                }
            }
            decimal FTr = 0;
            // Establecer la conexión con la base de datos
            string connectionString6 = "Data Source=LENOVO;Initial Catalog=Humedad;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString6))
            {
                connection.Open();

                // Verificar que Rfcalx.Text no esté vacío o nulo
                if (string.IsNullOrWhiteSpace(Rfcalx.Text))
                {
                    //MessageBox.Show("El nombre de la columna no puede estar vacío.");
                    return;
                }

                string query = $"SELECT [{Rfcalx.Text}] FROM [Humedad].[dbo].[Humedad] WHERE F2 = @F2Value";
                SqlCommand command = new SqlCommand(query, connection);

                // Añadir el parámetro de la consulta
                command.Parameters.AddWithValue("@F2Value", TCentx1);

                // Ejecutar la consulta y obtener el resultado
                object result1 = command.ExecuteScalar();

                // Verificar si el resultado no es null (es decir, se encontró un valor)
                if (result1 != null)
                {
                    try
                    {
                        // Convertir el resultado a un decimal (asumiendo que es un valor decimal válido)
                        FTr = Convert.ToDecimal(result1);
                        //MessageBox.Show($"El valor obtenido es: {FTr}");
                    }
                    catch (FormatException)
                    {
                        //MessageBox.Show("Error: El resultado no es un valor decimal válido.");
                    }
                }
                else
                {
                    //MessageBox.Show("No se encontró el valor especificado.");
                }
            }


            int TCond = 0;
            if (TTcmc6.Text == "")
            {
                MessageBox.Show("Entrar el valor de Temp. Condensación");
            }
            else { TCond = Convert.ToInt16(float.Parse(TTcmc6.Text)); }
            decimal KwattE = 0;

            int autRValue;
            if (string.IsNullOrEmpty(Datos.AutR) || (int.TryParse(Datos.AutR, out autRValue) && autRValue == 0))
            {
                AutR.Text = "0,054";
            }
            else
            {
                // Maneja el caso en que Datos.AutR no sea nulo y no sea igual a 0
                // Puedes agregar lógica adicional aquí si es necesario
            }
            decimal CpP;
            if (decimal.TryParse(AutR.Text, NumberStyles.Any, CultureInfo.GetCultureInfo("es-ES"), out CpP))
            {
                CpP = Math.Round(CpP, 4); // Redondea a 4 dígitos después de la coma
            }
            else
            {
                // Maneja el caso en que la conversión falle
                Console.WriteLine("El valor de AutR.Text no es un número válido.");
            }

            try
            {
                if (((FR * FTr) + 0.7M - CpP) - (-0.0292M) * (TCond - 55) > Kwattp) { MessageBox.Show("Realizar el cálculo de Carga"); Close(); }
                else if (TCentx1 >= 0) { KwattE = Math.Round(Kwattp / (FR * FTr), 2, MidpointRounding.ToEven); } 
                else { KwattE = Math.Round(Kwattp / (FR * FTr), 2, MidpointRounding.ToEven); }
            }
            catch { }

            decimal Evp1 = 0;//Factor de Máxima valor
            decimal Evp2 = 0;//Factor de Minimo valor.
            decimal Tpks = 0;//Coeficiente temperatura
            Tpks = 25;
            decimal TCentr = 0;//Temperatura transición °C a °F

            try
            {
                TRdin.Text = "";
                PEcalc.Text = Math.Round(KwattE, 2, MidpointRounding.ToEven).ToString();
                if (selected != 1) { TCentr = TCentx1 - 32 / 1.8M; } else { TCentr = TCentx1 * 1; }
                TCentx.Text = Math.Round(TCentr, 0, MidpointRounding.ToEven).ToString();//Temperatura de la camara
                if (selected != 1) { TTcmc8.Text = Math.Round((TCentx1 - (LPs1 - 32 / 1.8M)), 0, MidpointRounding.ToEven).ToString(); } else { TTcmc8.Text = Math.Round((TCentx1 - LPs1), 1, MidpointRounding.ToEven).ToString(); }

            }
            catch { }

            // Convertir de °C a K
            string Tkelvis = "";
            try
            {

                double celsius = Convert.ToInt16(TCentx1 - LPs1);
                double kelvin = celsius + 273;
                Tkelvis = "F" + kelvin.ToString();
            }
            catch (FormatException)
            {
                MessageBox.Show("Por favor, ingresa un número válido.");
            }

            decimal Tdin = 0;
            Tdin = 10;
            decimal KwattS = KwattE;
            decimal Rango = Convert.ToDecimal(PRango.Text) * 20;
            decimal KwattM = (KwattS + (KwattS * Rango / 100));
            decimal KwattN = (KwattS - (KwattS * Rango / 100));
            decimal PTEselec = 0;

            string connectionString = "Data Source=LENOVO;Initial Catalog=Selectorv;Integrated Security=True";

            string query1 = @"
            SELECT 
                Pmax, M3, N3 FROM [Selectorv].[dbo].[cpoia_hisp]
            WHERE 
                Fac = @Fac AND Tp = @Tp AND Pe = @Pe AND Pmax <= @PmaxM AND Pmax >= @PmaxN";

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                SqlCommand command = new SqlCommand(query1, connection);
                command.Parameters.AddWithValue("@Fac", FbcT.Text);
                command.Parameters.AddWithValue("@Tp", CSerie.Text);
                command.Parameters.AddWithValue("@Pe", TQevp.Text);
                command.Parameters.AddWithValue("@Tev", TTcmc8.Text);
                command.Parameters.AddWithValue("@PmaxM", KwattS + (KwattS * Rango / 100));
                command.Parameters.AddWithValue("@PmaxN", KwattS - (KwattS * Rango / 100));

                connection.Open();
                SqlDataReader reader = command.ExecuteReader();

                if (reader.Read())
                {
                    decimal avgPmax = reader["Pmax"] != DBNull.Value ? Convert.ToDecimal(reader["Pmax"]) : 0;
                    decimal avgM3 = reader["M3"] != DBNull.Value ? Convert.ToDecimal(reader["M3"]) : 0;
                    decimal avgN3 = reader["N3"] != DBNull.Value ? Convert.ToDecimal(reader["N3"]) : 0;
                    Datos.avgPmax = avgPmax.ToString();
                    Datos.avgM3 = avgM3.ToString();
                    Datos.avgN3 = avgN3.ToString();
                    Console.WriteLine($"Pmax: {avgPmax}, M3: {avgM3}, N3: {avgN3}");
                }

                reader.Close();
            }


            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                connection.Open();
                if (TQevp.Text == "AUT")
                {
                    string query = "Select * From cpoia_hisp";
                    SqlCommand comando = new SqlCommand(query, connection);
                    SqlDataAdapter data = new SqlDataAdapter(comando);
                    System.Data.DataTable tabla = new System.Data.DataTable();
                    data.Fill(tabla);
                    dgv_consultav.DataSource = tabla;
                }
                else
                {

                    try
                    {
                        decimal autR = Convert.ToDecimal(AutR.Text);
                        decimal avgPmax = Convert.ToDecimal(Datos.avgPmax);
                        decimal avgM3 = Convert.ToDecimal(Datos.avgM3);
                        decimal avgN3 = Convert.ToDecimal(Datos.avgN3);
                        decimal tCentx1 = Convert.ToDecimal(TCentx1);
                        decimal lPs1 = Convert.ToDecimal(LPs1);
                        decimal dTpd = Convert.ToDecimal(DTpd.Text);
                        decimal tTcmc6 = Convert.ToDecimal(TTcmc6.Text);
                        decimal Tdevap = (tCentx1 - lPs1 - dTpd + 28) + 0.7m;
                        decimal Tcondens = (tTcmc6 - 55)* avgN3;
                        decimal kwatt = Math.Round((avgM3 * 100)* (avgN3/tCentx1-) * lPs1 + (-(0.029m) * (tTcmc6 - 55 + 10)), 2, MidpointRounding.ToEven);
                        Datos.tKwatt = Math.Round((avgM3 * 100) * (avgN3/lPs1) * lPs1 + (-(0.029m) * (tTcmc6 - 55 + 10)), 2, MidpointRounding.ToEven).ToString();
                        //decimal kwatt = Math.Round((avgPmax + avgM3 * (tCentx1 - LPs1 + 28) + 0.092m * (Tcondens)), 2, MidpointRounding.ToEven);
                        //Datos.tKwatt = Math.Round((avgPmax + avgM3 * (tCentx1 - LPs1 + 28) + 0.092m * (Tcondens)), 2, MidpointRounding.ToEven).ToString();
                        Console.WriteLine($"Kwatt: {kwatt}");
                    }
                    catch (FormatException ex)
                    {
                        Console.WriteLine("Error de formato: " + ex.Message);
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("Error: " + ex.Message);
                    }
                    decimal PmaxM;
                    decimal PmaxN;
                    PmaxM = Convert.ToDecimal(Datos.tKwatt) + (Convert.ToDecimal(Datos.tKwatt) * Rango / 100);
                    PmaxN = Convert.ToDecimal(Datos.tKwatt) - (Convert.ToDecimal(Datos.tKwatt) * Rango / 100);
                    string query = "Select Fac, Tp, Pe, Kwatt, Non From cpoia_hisp where Fac = '" + FbcT.Text + "' and Tp = '" + CSerie.Text + "' and Pe = '" + TQevp.Text + "' and Kwatt BETWEEN '" + Convert.ToInt32(double.Parse(PmaxN.ToString())) + "' and '" + Convert.ToInt32(double.Parse(PmaxM.ToString())) + "'";
                    SqlCommand comando = new SqlCommand(query, connection);
                    SqlDataAdapter data = new SqlDataAdapter(comando);
                    System.Data.DataTable tabla = new System.Data.DataTable();
                    try
                    {
                        data.Fill(tabla);
                        dgv_consultav.DataSource = tabla;
                        // Asigna el DataSource al DataGridView
                        dgv_consultav.DataSource = tabla;
                        dgv_consultav.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

                        // Oculta la columna 'Id' si existe
                        if (dgv_consultav.Columns.Contains("Id"))
                        {
                            dgv_consultav.Columns["Id"].Visible = false;
                        }
                        if (dgv_consultav.Columns.Contains("Fac"))
                        {
                            dgv_consultav.Columns["Fac"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }
                        if (dgv_consultav.Columns.Contains("Non"))
                        {
                            dgv_consultav.Columns["Non"].AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells;
                        }

                        // Fuerza la actualización del diseño del DataGridView
                        dgv_consultav.AutoResizeColumns();
                        dgv_consultav.Refresh();

                        // Encuentra la fila con el valor de Kwatt más cercano a KwattS
                        //decimal kwattS = Convert.ToDecimal(Datos.tKwatt);
                        decimal diferenciaMinima = decimal.MaxValue;
                        int filaCercana = -1;
                        for (int i = 0; i < tabla.Rows.Count; i++)
                        {
                            decimal kwatt = Convert.ToDecimal(tabla.Rows[i]["Kwatt"]);
                            decimal diferencia = Math.Abs(kwatt - Convert.ToDecimal(Datos.tKwatt));
                            if (diferencia < diferenciaMinima)
                            {
                                diferenciaMinima = diferencia;
                                filaCercana = i;
                            }
                        }
                        if (filaCercana != -1)
                        {
                            dgv_consultav.Rows[filaCercana].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                            TValv.Text = tabla.Rows[filaCercana]["Non"].ToString();
                            PTevap.Text = tabla.Rows[filaCercana]["Kwatt"].ToString();
                            Datos.filacercanav = filaCercana.ToString();
                        }

                        SqlDataReader lector = comando.ExecuteReader();
                        if (lector.Read())
                        {
                            double pTevapS = Convert.ToDouble(PTevap.Text);
                            double kwattE = Convert.ToDouble(Datos.tKwatt);
                            double resultado1 = Math.Round((kwattE / pTevapS / 1), 2, MidpointRounding.ToEven);
                            Console.WriteLine($"Resultado1: {resultado1}");
                            TRdin.Text = resultado1.ToString();

                            // Procesar el resultado aparte
                            ProcesarResultado(pTevapS, Convert.ToDouble(PEcalc.Text));
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error: " + ex.Message);
                    }
                }
            }

            if (selected == 1) { PEcalc.Text = Math.Round(KwattE, 2, MidpointRounding.ToEven).ToString(); }
            if (selected == 2) { PEcalc.Text = Math.Round(KwattE * 0.284M, 2, MidpointRounding.ToEven).ToString(); }
            if (selected == 3) { PEcalc.Text = Math.Round(KwattE * 3.41212M, 2, MidpointRounding.ToEven).ToString(); }
        }

        private void dgv_consultav_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (Datos.filacercanav != null)
                {
                    dgv_consultav.Rows[Convert.ToInt16(Datos.filacercanav)].DefaultCellStyle.BackColor = System.Drawing.Color.White; // o el color original
                    dgv_consultav.Rows[Convert.ToInt16(Datos.filacercanav)].DefaultCellStyle.Font = dgv_consultav.DefaultCellStyle.Font;
                }
                PTevap.Text = dgv_consultav.CurrentRow.Cells[4].Value.ToString();
                TValv.Text = dgv_consultav.CurrentRow.Cells[5].Value.ToString();
                double peCalc = Convert.ToDouble(PEcalc.Text);
                double pTevap = Convert.ToDouble(PTevap.Text);
                double resultado2 = Math.Round(peCalc / pTevap / 1, 2, MidpointRounding.ToEven);
                Console.WriteLine($"Resultado2: {resultado2}");
                TRdin.Text = resultado2.ToString();
                ProcesarResultado(pTevap, peCalc);
                dgv_consultav.Rows[e.RowIndex].DefaultCellStyle.BackColor = System.Drawing.Color.LightGreen;
                Datos.filacercanav = e.RowIndex.ToString();
            }
            catch { }
        }
        private void ProcesarResultado(double valorActual, double valorReferencia)
        {
            // Imprimir valores para depuración
            Console.WriteLine($"Valor Actual: {valorActual}, Valor de Referencia: {valorReferencia}");

            // Calcular el porcentaje de diferencia
            double porcentajeDiferencia = ((valorActual - valorReferencia) / valorReferencia) * 100;
            string resultadoConPorcentaje = (porcentajeDiferencia >= 0 ? "+" : "") + porcentajeDiferencia.ToString("F2") + "%";
            Console.WriteLine($"Resultado con porcentaje: {resultadoConPorcentaje}");

            // Mostrar el resultado en el TextBox TRdin
            TRdin.Text = resultadoConPorcentaje;
            Datos.TRdinE = resultadoConPorcentaje;
        }
        private void BCard_Click(object sender, EventArgs e)
        {
            BCdor.Clear();
            if (BCdor.Text == "") { BCdor.Text = TValv.Text; }
            string connectionString5 = "Data Source=LENOVO;Initial Catalog=Selectorv;Integrated Security=True";
            using (SqlConnection connection = new SqlConnection(connectionString5))
            {
                connection.Open();
                string query = $"Select Fac, Tp, Pe, Tev, Kwatt, Non From cpoia_hisp where Non = '" + BCdor.Text + "'";
                SqlCommand command = new SqlCommand(query, connection);

                // Ejecutar la consulta y obtener el resultado
                object result = command.ExecuteScalar();

                // Verificar si el resultado no es null (es decir, se encontró un valor)
                if (result != null)
                {
                    SqlCommand comando = new SqlCommand(query, connection);
                    SqlDataAdapter data = new SqlDataAdapter(comando);
                    System.Data.DataTable tabla = new System.Data.DataTable();
                    data.Fill(tabla);
                    dgv_consultav.DataSource = tabla;
                }
                else
                {
                    MessageBox.Show("Insertat Evaporador de busqueda");
                }

            }

        }


    }





}

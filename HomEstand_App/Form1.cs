using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// RECURSOS DE La APLICACIÓN
using HomEstand_App.Properties;
using System.Reflection;
// CATALOGOS - CSV
using System.IO;
// BASE DE DATOS - ACCESS
using System.Data.OleDb;
// JAROWINKLER - METODOS DE STRING MATCHING
    // INSTALADO VÍA NUGET
using SimMetricsMetricUtilities;
// TUPLES - 
using System.Collections;
// UI Thread - PROCESOS EN SEGUNDO PLANO  
using System.Threading;

namespace HomEstand_App
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
        }

        // CARGA DEL FORMULARIO
        private void Form1_Load(object sender, EventArgs e)
        {
            // INICIALIZAR DELEGADO PARA COMUNICACIÓN CON PROCESOS EN SEGUNDO PLANO
            this.updateStatusDelegate = new UpdateStatusDelegate(this.UpdateStatus);

            // CARGA DE CATALOGO DE COMPAÑIAS
            //load_CATCias();
        }

        // VARIABLES GLOBALES 
            // Conexion a la Base de Datos
        String MyConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Desktop\\H&E_DB.mdb;";

        // THREADING - PROCESOS EN SEGUNDO PLANO
            // THREAD
        private Thread myThread = null;
            // DELEGADO USADO PARA COMUNICAR EL THREAD CON LA APLICACION PRINCIPAL
        private delegate void UpdateStatusDelegate();
        private UpdateStatusDelegate updateStatusDelegate = null;

            // PROGRESO DEL PROCESO EN SEGUNDO PLANO:
        int progMax = 1;
        int progCount = 0;
        DateTime tStart;
        TimeSpan tExec;
            
            // ACTUALIZAR PROGRESO DEL PROCESO EN SEGUNDO PLANO
        private void UpdateStatus()
        {
            tExec = DateTime.Now - tStart;
            txt_TimeExec.Text = String.Format("{0:D2}:{1:D2}:{2:D2}", tExec.Hours, tExec.Minutes, tExec.Seconds);
            txt_ProgCount.Text = progCount.ToString() + " de " + progMax.ToString();
        }

        // CATALOGOS
            // CATALOGO: COMPAÑIAS (NO FUNCIONAL)
        /*   
        public List<String>[] get_CATCias()
        {
            var currentAssembly = Assembly.GetExecutingAssembly();
            using (var stream = currentAssembly.GetManifestResourceStream("HomEstand_App.Cias_WS.csv"))
            //using (var stream = Resources.Cias_WS)
            {
                using (var readCIAS = new StreamReader(stream))
                {
                    // Arreglo de Listas que almacena: 
                    // ID Compañia, Nombre Compañia, Abreviatura Compañia
                    List<String>[] CATCias = new List<String>[3];
                    for (Int32 i = 0; i < CATCias.Length; i++)
                    {
                        CATCias[i] = new List<String>();
                    }

                    int row = 0;
                    while (!readCIAS.EndOfStream)
                    {
                        var line = readCIAS.ReadLine();
                        var values = line.Split(',');

                        // La primera fila (row) almacena el nombre de la columna.
                        // Las siguientes almacenan los valores.
                        if (row > 0)
                        {
                            // ID
                            CATCias[0].Add((values[0].ToString() != "") ? values[0] : "NA");
                            // Nombre
                            CATCias[1].Add((values[1].ToString() != "") ? values[1] : "NA");
                            // Abreviatura
                            CATCias[2].Add((values[2].ToString() != "") ? values[2] : "NA");
                        }
                        row++;
                    }
                    MessageBox.Show("Carga Completa de Catálogo de Compañías", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return CATCias;
                }
            }
        }

        private void load_CATCias()
        {
            // Obteniendo Catalogo de CIAS
            List<String>[] CATCias = get_CATCias();

            OleDbConnection CONNECT = new OleDbConnection(MyConnString);

            for (int i = 0; i < CATCias[0].Count; i++)
            {
                if (TableExists(CONNECT, "DAT_" + CATCias[2].ElementAt(i).ToString()))
                {
                    cmb_Cia.Items.Add(CATCias[2].ElementAt(i).ToString());
                    CiasID[i] = CATCias[0].ElementAt(i).ToString();
                }
            }
        }

            // Catalogo de Compañías
        String[] CiasID = null;

        */
            // CATALOGO: ACRONIMOS
        public List<String>[] getAcroDB()
        {
            using (var readCSV = new StreamReader(@"C:\Users\Griselda\Documents\AcronymDB.csv"))
            {
                List<String>[] acrDB = new List<String>[12];
                for (Int32 i = 0; i < acrDB.Length; i++)
                {
                    acrDB[i] = new List<String>();
                }

                int reg = 0;
                while (!readCSV.EndOfStream)
                {
                    var line = readCSV.ReadLine();
                    var values = line.Split(',');

                    if (reg > 0)
                    {
                        acrDB[0].Add((values[0].ToString() != "") ? values[0] : "NA");
                        // MessageBox.Show(acrDB[0].ElementAt(acrDB[0].Count-1).ToString(), "Trans");
                        acrDB[1].Add((values[1].ToString() != "") ? values[1] : "NA");
                        acrDB[2].Add((values[2].ToString() != "") ? values[2] : "NA");
                        acrDB[3].Add((values[3].ToString() != "") ? values[3] : "NA");
                        acrDB[4].Add((values[4].ToString() != "") ? values[4] : "NA");
                        acrDB[5].Add((values[5].ToString() != "") ? values[5] : "NA");
                        acrDB[6].Add((values[6].ToString() != "") ? values[6] : "NA");
                        acrDB[7].Add((values[7].ToString() != "") ? values[7] : "NA");
                        acrDB[8].Add((values[8].ToString() != "") ? values[8] : "NA");
                        acrDB[9].Add((values[9].ToString() != "") ? values[9] : "NA");
                        acrDB[10].Add((values[10].ToString() != "") ? values[10] : "NA");
                        acrDB[11].Add((values[11].ToString() != "") ? values[11] : "NA");
                    }
                    reg++;
                }

                MessageBox.Show("Carga Completa del Catálogo de Acrónimos", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                   
                return acrDB;
            }
        }

        // UTILIDADES (MÉTODOS)
            // COMPROBAR SI EXISTE UNA TABLA
        public bool TableExists(OleDbConnection CONN, String table)
        {
            CONN.Open();
            var exists = CONN.GetSchema("Tables", new string[4] { null, null, table, "TABLE" }).Rows.Count > 0;
            CONN.Close();
            return exists;
        }

            // EJECUTAR CONSULTA SQL EN ACCESS
        public void doQuery(String CONS, String CONEX)
        {
            OleDbConnection CONNECT = new OleDbConnection(CONEX
                //"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Documents\\Nueva BD\\Test.mdb;"
                );
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONS, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();
                READER.Close();
                CONNECT.Close();
                // MessageBox.Show("SUCCESSS");
            }
            catch (Exception e)
            {
                MessageBox.Show("QUERY: ___" + CONS + "___ EX:___" + e.ToString(), "Error en Query", MessageBoxButtons.OK, MessageBoxIcon.Error);      
            }

        }

            // DETERMINAR SI UN STRING ES EQUIVALENTE A OTRO, SIN IMPORTAR: ORDEN, ESPACIOS, MAYUSCULAS NI MINUSCULAS
        public Boolean eqDescrip(String a, String b)
        {
            a = a.ToUpper(); b = b.ToUpper();
            String aWord = "", bWord = "";
            List<String> listA = new List<String>();
            List<String> listB = new List<String>();
            int aux = 0;

            // Descomponer el String A en palabras y agregarlo a List
            while (a.Length > 0)
            {
                // Remover ' ' al inicio 
                a = a.Trim();
                // MessageBox.Show("_" + a + "_", "String A sin Esp");

                // String valido
                if (a.Length > 0)
                {
                    // Varias palabras
                    if (a.LastIndexOf(' ') > 0)
                    {
                        aux = a.IndexOf(' ');
                        aWord = a.Substring(0, aux);
                        // MessageBox.Show("_" + aWord + "_", "Word A");
                        a = a.Remove(0, aux);
                        // MessageBox.Show("_" + a + "_", "String restante A");
                    }
                    // Ultima palabra
                    else
                    {
                        aWord = a;
                        // MessageBox.Show("_" + aWord + "_", "Last Word A");
                        a = "";
                    }
                    // Agregando palabra a la lista
                    listA.Add(aWord);
                    // MessageBox.Show(listA.Count.ToString(), "Words in listA");
                    aWord = ""; aux = 0;
                }
                else
                {
                    break;
                }
            }

            // Descomponer el String B en palabras y agregarlo a List
            while (b.Length > 0)
            {
                // Remover ' ' al inicio 
                b = b.Trim();

                // String valido
                if (b.Length > 0)
                {
                    // Varias palabras
                    if (b.LastIndexOf(' ') > 0)
                    {
                        aux = b.IndexOf(' ');
                        bWord = b.Substring(0, aux);
                        b = b.Remove(0, aux);
                    }
                    // Ultima palabra
                    else
                    {
                        bWord = b;
                        b = "";
                    }
                    // Agregando palabra a la lista
                    listB.Add(bWord);
                    bWord = ""; aux = 0;
                }
                else
                {
                    break;
                }
            }

            // Verificar que las listas tienen el mismo tamaño y contienen los mismos elementos
            return (listA.Count == listB.Count) && new HashSet<string>(listA).SetEquals(listB);
        }

            // ORDENAR ALFABETICA E INVERSAMENTE UN STRING
        public String sortDescrip(String a, Boolean Inverse)
        {
            a = a.ToUpper();
            String aWord = "";
            List<String> listA = new List<String>();
            int aux = 0;

            // Descomponer el String A en palabras y agregarlo a List
            while (a.Length > 0)
            {
                // Remover ' ' al inicio 
                a = a.Trim();

                // String valido
                if (a.Length > 0)
                {
                    // Varias palabras
                    if (a.LastIndexOf(' ') > 0)
                    {
                        aux = a.IndexOf(' ');
                        aWord = a.Substring(0, aux);
                        // MessageBox.Show("_" + aWord + "_", "Word A");
                        a = a.Remove(0, aux);
                        // MessageBox.Show("_" + a + "_", "String restante A");
                    }
                    // Ultima palabra
                    else
                    {
                        aWord = a;
                        // MessageBox.Show("_" + aWord + "_", "Last Word A");
                        a = "";
                    }
                    // Agregando palabra a la lista
                    listA.Add(aWord);
                    // MessageBox.Show(listA.Count.ToString(), "Words in listA");
                    aWord = ""; aux = 0;
                }
                else
                {
                    break;
                }
            }

            // Ordenando Lista
            listA.Sort();
            // Lista Inversa
            if (Inverse)
            {
                listA.Reverse();
            }

            // Construyendo nueva String
            aWord = "";
            foreach (String Val in listA)
            {
                aWord += Val + " ";
            }
            return aWord;
        }

            // OBTENER EL INDICE DE UN VALOR EN UNA MATRIZ DE DOBLES
        public Tuple<int, int> getIndex(double[,] jaggedArray, double value)
        {
            int w = jaggedArray.GetLength(0); // width
            int h = jaggedArray.GetLength(1); // height

            for (int x = 0; x < w; ++x)
            {
                for (int y = 0; y < h; ++y)
                {
                    if (jaggedArray[x, y].Equals(value))
                        return Tuple.Create(x, y);
                }
            }

            return Tuple.Create(-1, -1);
        }

        /////////////////////////////////////////////////////////////

        /// INTERFAZ DE USUARIO 

        ///////////////////////////////////////////////////////////

        // PANEL DE NAVEGACIÓN
            // COMBO-BOX: COMPAÑIAS
        private void cmb_Cia_SelectedIndexChanged(object sender, EventArgs e)
        {
            int[] num_Comp = { 7, 21, 5, 26, 20, 2, 12 };
            txt_CiaID.Text = (cmb_Cia.SelectedIndex > 0) ? num_Comp[cmb_Cia.SelectedIndex].ToString(): "";
        
            //txt_CiaID.Text = CiasID[cmb_Cia.SelectedIndex];
            String cMarca = "";
            if (cmb_Marca.SelectedIndex >= 0)
            {
                cMarca = cmb_Marca.SelectedItem.ToString();
            }
            cmb_Marca.SelectedIndex = -1;
            cmb_Marca.Items.Clear();

            if (cmb_Cia.SelectedIndex >= 0)
            {
                OleDbConnection CONNECT = new OleDbConnection(
                    MyConnString);
                String CONSULT = "SELECT DISTINCT Marca as Brand FROM DAT_" + cmb_Cia.SelectedItem.ToString();
                try
                {
                    CONNECT.Open();
                    OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                    OleDbDataReader READER = COMMAND.ExecuteReader();

                    if (READER.HasRows)
                    {
                        while (READER.Read())
                        {
                            if (READER["Brand"].ToString() != "")
                                cmb_Marca.Items.Add(READER["Brand"]);
                        }
                    }
                    READER.Close();
                    CONNECT.Close();
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString(), "ERROR");
                }
                finally
                {
                    if (cmb_Marca.Items.Contains(cMarca))
                    {
                        cmb_Marca.SelectedItem = cMarca;
                    }
                    else
                    {
                        cmb_Marca.SelectedItem = -1;
                        cmb_Tipo0.SelectedItem = -1;
                        cmb_Tipo1.SelectedItem = -1;
                        cmb_Tipo0.Items.Clear();
                        cmb_Tipo1.Items.Clear();
                    }
                }
            }
        }

            // COMBO-BOX: MARCAS
        private void cmb_Marca_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmb_Cia.SelectedIndex >= 0)
            {
                if (cmb_Marca.SelectedIndex >= 0)
                {
                    cmb_Tipo0.SelectedIndex = -1;
                    cmb_Tipo1.SelectedIndex = -1;
                    cmb_Tipo0.Items.Clear();
                    cmb_Tipo1.Items.Clear();

                    try
                    {
                        OleDbConnection CONNECT = new OleDbConnection(
                            MyConnString);
                        String CONSULT = "SELECT DISTINCT Tipo as Type FROM DAT_" + cmb_Cia.SelectedItem.ToString() + " WHERE Marca = '" + cmb_Marca.SelectedItem.ToString() + "'";
                
                        CONNECT.Open();
                        OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                        OleDbDataReader READER = COMMAND.ExecuteReader();

                        if (READER.HasRows)
                        {
                            while (READER.Read())
                            {
                                if (READER["Type"].ToString() != "")
                                {
                                    cmb_Tipo0.Items.Add(READER["Type"]);
                                    cmb_Tipo1.Items.Add(READER["Type"]);
                                }
                            }
                        }
                        READER.Close();
                        CONNECT.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                    }
                }
                else
                {
                    cmb_Marca.SelectedIndex = -1;
                }
            }
            else
            {
                MessageBox.Show("Selecciona una Compañía", "ERROR");
                cmb_Cia.SelectedIndex = -1;
            }

        }

            // REFRESCAR INFORMACIÓN (COMBO-BOXES)
        private void btnRInfo_Click(object sender, EventArgs e)
        {
            cmb_Tipo0.SelectedIndex = -1;
            cmb_Tipo1.SelectedIndex = -1;
            cmb_Marca.SelectedIndex = -1;
            cmb_Cia.SelectedIndex = -1;
            txt_CiaID.Clear();
            cmb_Tipo0.Text = "";
            cmb_Tipo1.Text = "";
            cmb_Marca.Text = "";
            cmb_Cia.Text = "";
        }

        // PANEL 1: MARCA Y TIPO
            // BOTON: CAMBIAR TIPO
        private void btn_CambiarTipo_Click(object sender, EventArgs e)
        {
            if (txt_NTipo.Text.Trim() != "")
            {
                if (cmb_Tipo0.SelectedIndex >= 0)
                {
                    if (chk_TDSimple.Checked)
                    {
                        cambiarTipo(" " + cmb_Tipo0.SelectedItem.ToString().Trim() + " ", txt_NTipo.Text, cmb_Marca.SelectedItem.ToString(), cmb_Cia.SelectedItem.ToString(), true);
                    }
                    else
                    {
                        cambiarTipo(" " + cmb_Tipo0.SelectedItem.ToString().Trim() + " ", txt_NTipo.Text, cmb_Marca.SelectedItem.ToString(), cmb_Cia.SelectedItem.ToString(), false);
                    }
                    txt_NTipo.Clear();
                }
                else
                {
                    MessageBox.Show("Selecciona el Tipo a Sustituir", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else {
                MessageBox.Show("Introduce el Nuevo Tipo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
        }

        private void cambiarTipo(String Tipo, String nTipo, String Marca, String Cia, Boolean tDSimple)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "'";
            try
            {
                CONNECT.Open();
                // Si el Tipo está en la Descripción Simple
                OleDbCommand COMMAND = (tDSimple) ? 
                    new OleDbCommand(CONSULT, CONNECT) : 
                    new OleDbCommand(CONSULT + " AND Tipo = '" + Tipo.Trim() + "'", CONNECT);

                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    String x;
                    
                    if (tDSimple)
                    {
                        x = " " + READER["DescripSimple"].ToString().Trim() + " ";
                        if (x.Contains(Tipo))
                        {
                            //REMOVER Tipo de la Descripcion Simple
                            x = x.Replace(Tipo, " ");

                            doQuery("UPDATE DAT_" + Cia + " SET Tipo = '" + nTipo + "', DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                MyConnString
                                 );
                        }
                    }
                    else
                    {
                        doQuery("UPDATE DAT_" + Cia + " SET Tipo = '" + nTipo + "' WHERE Tipo = '" + Tipo.Trim() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                MyConnString
                                 );
                    }

                }
                MessageBox.Show("Tipo cambiado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information); 
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
        }

            // BOTON: CAMBIAR MARCA
        private void btn_CambiarMarca_Click(object sender, EventArgs e)
        {
            if (cmb_Tipo0.SelectedIndex >= 0)
            {
                if (cmb_NMarca.SelectedIndex >= 0)
                {
                    cambiarMarca(cmb_Tipo0.SelectedItem.ToString(), cmb_Marca.SelectedItem.ToString(), cmb_NMarca.SelectedItem.ToString(), cmb_Cia.SelectedItem.ToString());
                    cmb_NMarca.SelectedIndex = -1;
                }
                else {
                    MessageBox.Show("Introduce la Nueva Marca", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
            }
            else {
                MessageBox.Show("Introduce el Tipo cuya Marca se va a Sustituir", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
            }
        }

        private void cambiarMarca(String Tipo, String Marca, String nMarca, String Cia)
        {
            try
            {
                OleDbConnection CONNECT = new OleDbConnection(
                    MyConnString);
                String CONSULT = "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "' AND Tipo = '" + Tipo + "' ORDER BY Clave";
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);

                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    doQuery("UPDATE DAT_" + Cia + " SET Marca = '" + nMarca + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",
                                MyConnString
                                 );
                }     
                MessageBox.Show("Marca cambiada correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // PANEL 2: ACRONIMOS GENERALES
            // COMBO-BOX: CAMPO (BASE DE DATOS)
        private void cmb_Campo0_SelectedIndexChanged(object sender, EventArgs e)
        {
            // Trans, Cil, Vest, Aire, QC, Equipado, EE, BAire, Sonido, ABS, RA, FN, CodRaro, DH 
            String[,] fieldOptions = { 
                                       { "1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11", "12", 
                                           "13", "14", "15", "16", "17", "18", "19", "20", "21", "22", 
                                           "23", "24", "25", "26", "27","28", "29","30", 
                                           "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "2Ptas", "2y4Ptas", "3Ptas", "3y5Ptas", "4Ptas", "5Ptas", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "Aut", "Std", "7G-DCT", "7G-Tronic", "9G-DCT", "9G-Tronic", "AMG-Speedshift", "ASG", "AutoStick", "CVT", "DCT", 
                                           "Drivelogic", "DSG", "Dualogic", "DuoSelect", "Easytronic", "EDC", "Geartronic", "GETRAG", "G-Tronic", "HSD", 
                                           "Hydramatic", "Lineartronic", "M-DKG", "MCT", "Multitronic", "PDK", "Powershift", "Q-Tronic", "R-Tronic", "Secuencial", "SelectShift",
                                           "Selespeed", "Sentronic", "Shiftmatic", "Shiftronic", "SMG-II", "SMG", "Sportronic", "SportShift", "Steptronic", "S-Tronic",
                                           "TCT", "Tiptronic", "Touchtronic", "X-Tronic"}, 
                                        { "0Cil", "2Cil", "3Cil", "4Cil", "5Cil", "6Cil", "8Cil", "10Cil", "12Cil", 
                                            "", "", "", "", "", "", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "Alcantara", "Gamusina", "Gamuza", "Leatherette", "Napa", "Piel parcial", "Piel", "Tela", "Terciopelo", "Velour", 
                                            "Vinil", "", "", "","", "", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "C/AAcc", "S/AAcc", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "C/Qcc", "S/Qcc", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "6CD", "AFS", "ASIST.EST.", "AUDIO MANEJO", "C/CAMARA", "C/LOCKER", "CAM.TRAS.", "COMAND ONLINE", "COMP.VIAJE.", "CTROL/AUDIO", 
                                            "CTROL/VOZ", "EQUIPADO", "F/BI-XENON", "F/XENON", "GETRONIC", "GMLINK", "FULL LINK", "GPS", "HIELERA", "HILL HOLDER", "JOYBOX", 
                                            "MEDIA NAV", "MULTIMEDIA", "MYGIG", "NAVIGON", "PTA.TRAS.ELEC.", "SEMIEQUIPADO", "RNS-510", "SIST.ENTRET.", "SIST.NAV.", 
                                            "TPM", "TV", "UCONNECT", "WIFI", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "CE", "EE", "SE", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "CB", "CBL", "SB", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "AM", "BD", "BOSE", "BT", "CD", "CT", "DVD", "DYNAUDIO", "FENDER", "FM",
                                            "MP3", "RADIO", "SS", "USB", "","", "", "","", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "ABS", "D/ABS", "D/V", "D/T", "DIS", "NEU", "TAM", "V", "V/DIS", "V/T", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "R-13", "R-14", "R-15", "R-16", "R-17", "R-18", "R-19", "R-20", "R-21", "R-22",
                                            "R-25", "RA", "RA-14", "RA-15", "RA-16", "RA-17", "RA-18", "RA-19", "RA-20",
                                            "RA-21", "", "", "", "", "", "", "", "", "", 
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "FN", "", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "SM", "", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""},
                                        { "DH", "DHS", "", "", "", "", "", "", "", "", "", "", "", "","", "", "", "","", "", "", "", "",
                                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""}
                                     };


            cmb_NAcr.Text = "";
            cmb_NAcr.Items.Clear();
            cmb_NAcr.SelectedIndex = -1;
            if (cmb_Campo0.SelectedIndex >= 0)
            {
                for (int i = 0; i < 45; i++)
                {
                    if (fieldOptions[cmb_Campo0.SelectedIndex, i] != "")
                        cmb_NAcr.Items.Add(fieldOptions[cmb_Campo0.SelectedIndex, i]);
                }
                
            }
            chk_AcDSimple0.Checked = false;
        }

            // CHECK-BOX: ACRONIMO EN DESCRIPCION SIMPLE
        private void chk_AcDSimple_CheckedChanged(object sender, EventArgs e)
        {
            cmb_Campo0.SelectedIndex = -1;
            cmb_NAcr.SelectedIndex = -1;
        }

            // BOTON: CAMBIAR ACRONIMO GENERAL
        private void btn_CambiarAcroGen_Click(object sender, EventArgs e)
        {
            if (txt_Acr0.Text.Trim() != "")
            {
                if (chk_AcDSimple0.Checked)
                {
                    if (txt_NAcr0.Text.Trim() != "")
                    {
                        cambiarAcroGen(" " + txt_Acr0.Text.Trim() + " ", txt_NAcr0.Text.Trim(), "DescripSimple", cmb_Cia.SelectedItem.ToString(), true);
                        txt_Acr0.Clear();
                        txt_NAcr0.Clear();
                    }
                    else {
                        MessageBox.Show("Introduce el Nuevo Acrónimo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else {
                    if (cmb_Campo0.SelectedIndex >= 0)
                    {
                        if (cmb_NAcr.SelectedIndex >= 0)
                        {
                            cambiarAcroGen(" " + txt_Acr0.Text.Trim() + " ", cmb_NAcr.SelectedItem.ToString(), cmb_Campo0.SelectedItem.ToString(), cmb_Cia.SelectedItem.ToString(), false);
                            txt_Acr0.Clear();
                        }
                        else {
                            MessageBox.Show("Introduce el Nuevo Acrónimo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else {
                        MessageBox.Show("Selecciona un Campo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }  
                }
            }
            else {
                MessageBox.Show("Introduce el Acrónimo a Sustituir", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void cambiarAcroGen(String Acro, String nAcro, String Campo, String Cia, Boolean DSimple)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM DAT_" + Cia + " ORDER BY Clave, Modelo";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    String x = " " + READER["DescripSimple"].ToString() + " ";
                    if (DSimple)
                    {
                        if (x.Contains(Acro))
                        {
                            x = x.Replace(Acro, " " + nAcro.Trim() + " ");
                            doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString(),

                                MyConnString
                                 );
                        }
                    }
                    else
                    {
                        String y = READER[Campo].ToString();

                        if (x.Contains(Acro))
                        {
                            if (!(" " + y.Trim() + " ").Contains(" " + nAcro + " "))
                            {
                                x = x.Replace(Acro, " ");
                                y = (y.Length > 0) ? y + " " + nAcro : Acro;
                                doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "', " + Campo + " = '" + y + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                    "' AND Modelo = " + READER["Modelo"].ToString(),

                                    MyConnString
                                     );
                            }
                            else
                            {
                                x = x.Replace(Acro, " ");
                                doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                    "' AND Modelo = " + READER["Modelo"].ToString(),

                                    MyConnString
                                     );
                            }
                        }
                    }
                }
                MessageBox.Show("Acrónimo cambiado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

            // BOTON: ELIMINAR ACRONIMO GENERAL
        private void btn_EliminarAcroGen_Click(object sender, EventArgs e)
        {
            if (txt_Acr0.Text.Trim() != "")
            {
                if (chk_AcDSimple0.Checked)
                {
                    DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (dialogResult == DialogResult.Yes)
                    {
                        eliminarAcroGen(" " + txt_Acr0.Text + " ", cmb_Cia.SelectedItem.ToString(), "DescripSimple"); 
                    }
                }
                else
                {
                    if (cmb_Campo0.SelectedIndex >= 0)
                    {
                        DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (dialogResult == DialogResult.Yes)
                        {
                            eliminarAcroGen(" " + txt_Acr0.Text + " ", cmb_Cia.SelectedItem.ToString(), cmb_Campo0.SelectedItem.ToString());
                        }
                    }
                    else {
                        MessageBox.Show("Selecciona el Campo donde se eliminará el Acrónimo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                txt_Acr0.Clear();
            }
            else
            {
                MessageBox.Show("Introduce el Acrónimo a Eliminar", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void eliminarAcroGen(String Acro, String Cia, String Campo)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULT = "SELECT * FROM DAT_" + Cia + " ORDER BY Clave";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x = " " + READER[Campo].ToString().Trim() + " ";
                        if (x.Contains(Acro))
                        {
                            x = x.Replace(Acro, " ");
                            doQuery("UPDATE DAT_" + Cia + " SET " + Campo + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                "' AND Modelo = " + READER["Modelo"].ToString(),

                                MyConnString
                                );
                        }

                    }
                }
                MessageBox.Show("Acrónimo eliminado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // PANEL 3: ACRONIMOS POR MARCA
            // CHECK-BOX: ACRONIMO EN DESCRIPCION SIMPLE
        private void chk_AcDSimple1_CheckedChanged(object sender, EventArgs e)
        {
            cmb_Campo1.SelectedIndex = -1;
        }

            // BOTON: CAMBIAR ACRONIMO POR MARCA
        private void btn_CambiarAcroMar_Click(object sender, EventArgs e)
        {

            if (cmb_Marca.SelectedIndex >= 0) 
            {
                if (txt_Acr1.Text.Trim() != "")
                {
                    if (txt_NAcr1.Text.Trim() != "")
                    {
                        if (chk_AcTipo.Checked)
                        {
                            if (cmb_Tipo1.SelectedIndex >= 0)
                            {
                                if (chk_AcDSimple1.Checked)
                                {
                                    // MARCA-TIPO Y DSIMPLE
                                    cambiarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    txt_NAcr1.Text.Trim(),
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    cmb_Tipo1.SelectedItem.ToString(),
                                    "DescripSimple"
                                    );
                                }
                                else {
                                    // MARCA-TIPO Y CAMPO
                                    cambiarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    txt_NAcr1.Text.Trim(),
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    cmb_Tipo1.SelectedItem.ToString(),
                                    cmb_Campo1.SelectedItem.ToString()
                                    );
                                }
                                txt_NAcr1.Clear();
                            }
                            else {
                                MessageBox.Show("Selecciona un Tipo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            }
                        }
                        else {
                            if (chk_AcDSimple1.Checked)
                            {
                                // MARCA Y DSIMPLE
                                cambiarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    txt_NAcr1.Text.Trim(),
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    "noTipo",
                                    "DescripSimple"
                                    );
                            }
                            else {
                                // MARCA Y CAMPO
                                cambiarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ", 
                                    txt_NAcr1.Text.Trim(), 
                                    cmb_Cia.SelectedItem.ToString(), 
                                    cmb_Marca.SelectedItem.ToString(), 
                                    "noTipo", 
                                    cmb_Campo1.SelectedItem.ToString()
                                    );
                            }
                            txt_NAcr1.Clear();
                        }
                    }
                    else {
                        MessageBox.Show("Introduce Acrónimo a Insertar", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else {
                    MessageBox.Show("Introduce Acrónimo a Sustituir", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error); 
                }
            }
            else {
                MessageBox.Show("Selecciona una Marca", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                
            }
        }

        private void cambiarAcroMar(String Acro, String nAcro, String Cia, String Marca, String Tipo, String Campo)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);

            String CONSULT = (Tipo == "noTipo") ?  
                "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "' ORDER BY Clave" :
                "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "' AND Tipo = '" + Tipo + "' ORDER BY Clave";

            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x;
                        if (Campo != "DescripSimple") {
                            x = " " + READER["DescripSimple"].ToString().Trim() + " ";
                            String y = READER[Campo].ToString();
                            if (x.Contains(Acro))
                            {
                                x = x.Replace(Acro, " ");
                                y = (!y.Contains(nAcro)) ? y + " " + nAcro : nAcro;

                                if (Tipo == "noTipo")
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "', " + Campo + " = '" + y + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                        "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                        MyConnString
                                        );
                                }
                                else {
                                    doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "', " + Campo + " = '" + y + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                       "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                                       MyConnString
                                       );
                                }
                            }
                        }
                        else
                        {
                            x = " " + READER[Campo].ToString() + " ";
                            if (x.Contains(Acro))
                            {
                                x = x.Replace(Acro, " " + nAcro + " ");
                                if (Tipo == "noTipo")
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET " + Campo + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                        "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                        MyConnString
                                        );
                                }
                                else
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET " + Campo + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                       "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                                       MyConnString
                                       );
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("Acrónimo sustituido correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

            // BOTON: ELIMINAR ACRONIMO POR MARCA
        private void btnEliminarAcroMar_Click(object sender, EventArgs e)
        {
            if (cmb_Marca.SelectedIndex >= 0)
            {
                if (txt_Acr1.Text.Trim() != "")
                {
                    if (chk_AcTipo.Checked)
                    {
                        // ELIMINAR EN TIPO ESPECIFICO
                        if (cmb_Tipo1.SelectedIndex >= 0)
                        {
                            if (chk_AcDSimple1.Checked)
                            {
                                // MARCA-TIPO Y DSIMPLE
                                DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    eliminarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    cmb_Tipo1.SelectedItem.ToString(),
                                    "DescripSimple"
                                    );
                                }
                            }
                            else
                            {
                                // MARCA-TIPO Y CAMPO
                                DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                                if (dialogResult == DialogResult.Yes)
                                {
                                    eliminarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    cmb_Tipo1.SelectedItem.ToString(),
                                    cmb_Campo1.SelectedItem.ToString()
                                    );
                                }
                            }
                            txt_Acr1.Clear();
                        }
                        else
                        {
                            MessageBox.Show("Selecciona un Tipo", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }
                    else { 
                        // NO TIPO
                        if (chk_AcDSimple1.Checked)
                        {
                            // MARCA Y DSIMPLE
                            DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (dialogResult == DialogResult.Yes)
                            {
                                eliminarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    "noTipo",
                                    "DescripSimple"
                                    );
                            }
                        }
                        else {
                            // MARCA Y CAMPO
                            DialogResult dialogResult = MessageBox.Show("¿Borrar contenido especificado?", "Confirmación", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                            if (dialogResult == DialogResult.Yes)
                            {
                                eliminarAcroMar(
                                    " " + txt_Acr1.Text.Trim() + " ",
                                    cmb_Cia.SelectedItem.ToString(),
                                    cmb_Marca.SelectedItem.ToString(),
                                    "noTipo",
                                    cmb_Campo1.SelectedItem.ToString()
                                    );
                            }
                        }
                        txt_Acr1.Clear();
                    }
                }
                else
                {
                    MessageBox.Show("Introduce Acrónimo a Sustituir", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Selecciona una Marca", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void eliminarAcroMar(String Acro, String Cia, String Marca, String Tipo, String Campo)
        {
            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);

            String CONSULT = (Tipo == "noTipo") ?
                "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "' ORDER BY Clave" :
                "SELECT * FROM DAT_" + Cia + " WHERE Marca = '" + Marca + "' AND Tipo = '" + Tipo + "' ORDER BY Clave";

            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULT, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                if (READER.HasRows)
                {
                    while (READER.Read())
                    {
                        String x;
                        if (Campo != "DescripSimple")
                        {
                            x = " " + READER[Campo].ToString().Trim() + " ";
                            if (x.Contains(Acro))
                            {
                                x = x.Replace(Acro, " ");

                                if (Tipo == "noTipo")
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET " + Campo + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                        "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                        MyConnString
                                        );
                                }
                                else
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET " + Campo + " = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                       "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                                       MyConnString
                                       );
                                }
                            }
                        }
                        else
                        {
                            x = " " + READER["DescripSimple"].ToString() + " ";
                            if (x.Contains(Acro))
                            {
                                x = x.Replace(Acro, " ");
                                if (Tipo == "noTipo")
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                        "' AND Marca = '" + READER["Marca"].ToString() + "'",

                                        MyConnString
                                        );
                                }
                                else
                                {
                                    doQuery("UPDATE DAT_" + Cia + " SET DescripSimple = '" + x.Trim() + "' WHERE Clave = '" + READER["Clave"].ToString() +
                                       "' AND Marca = '" + READER["Marca"].ToString() + "' AND Tipo = '" + READER["Tipo"].ToString() + "'",

                                       MyConnString
                                       );
                                }
                            }
                        }
                    }
                }
                MessageBox.Show("Acrónimo eliminado correctamente.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
                READER.Close();
                CONNECT.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // PANEL 4 - ESTANDARIZACIÓN
          // GENERAR TABLA ESTANDARIZADA
        private void btn_GenerarTablaSTD_Click(object sender, EventArgs e)
        {
            if (cmb_Cia.SelectedIndex >= 0)
            {
                try
                {
                    String nomCia = cmb_Cia.SelectedItem.ToString();
                    tStart = DateTime.Now;

                    // VERIFICAR SI LA TABLA ESTANDARIZADA CONTIENE REGISTROS
                    int count = 0;
                    try
                    {
                        OleDbConnection CONNECT_COUNT = new OleDbConnection(MyConnString);
                        CONNECT_COUNT.Open();
                        OleDbCommand COMMAND_COUNT = new OleDbCommand("SELECT COUNT (*) FROM STD_" + nomCia, CONNECT_COUNT);

                        count = Convert.ToInt32(COMMAND_COUNT.ExecuteScalar());

                        CONNECT_COUNT.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show("EX:___" + Ex.ToString(), "Error al obtener registros de STD_"+ nomCia, MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // TABLA SIN REGISTROS, GENERAR NUEVA
                    if (count == 0)
                    {
                        //MessageBox.Show("Compañía no disponible", "Error");
                        int numCia = Convert.ToInt32(txt_CiaID.Text);
                        this.myThread = null;
                        this.myThread = new Thread(
                            () => DBToStand(nomCia, numCia)
                        );
                        this.myThread.Start();
                    }
                    else
                    {
                        MessageBox.Show("La Tabla STD_" + nomCia + " ya contiene registros", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else {
                MessageBox.Show("Selecciona una compañía.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
            // GENERAR TABLA ESTANDARIZADA PARA COMPAÑIA GENERICA
        public void DBToStand(String Company, Int32 numCompany)
        {
            progCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM DAT_" + Company, CONNECTPB);

                progMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }

            OleDbConnection CONNECT = new OleDbConnection(
                MyConnString);
            String CONSULTA = "SELECT * FROM DAT_" + Company + " ORDER BY Clave";
            // var np = "";
            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND = new OleDbCommand(CONSULTA, CONNECT);
                OleDbDataReader READER = COMMAND.ExecuteReader();

                while (READER.Read())
                {
                    if (READER["Tipo"].ToString() != "" && READER["Marca"].ToString() != "ESPECIALES" && READER["Marca"].ToString() != "FRONTERIZO" && READER["Marca"].ToString() != "LEGALIZADO" && READER["Marca"].ToString() != "FRONTERIZO" && READER["Marca"].ToString() != "PLAN PISO" && READER["Marca"].ToString() != "CLASICO")
                    {
                        String desTSM = (READER["DescripSimple"].ToString().Trim().Length > 0) ?
                            READER["DescripSimple"].ToString().Trim() + " " : "";
                        desTSM += (READER["Equipado"].ToString().Trim().Length > 0) ? READER["Equipado"].ToString().Trim() + " " : "";
                        desTSM += (READER["Trans"].ToString().Trim().Length > 0) ? READER["Trans"].ToString().Trim() + " " : "";
                        desTSM += (READER["Puertas"].ToString().Trim().Length > 0) ? READER["Puertas"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vestiduras"].ToString().Trim().Length > 0) ? READER["Vestiduras"].ToString().Trim() + " " : "";
                        desTSM += (READER["ABS"].ToString().Trim().Length > 0) ? READER["ABS"].ToString().Trim() + " " : "";
                        desTSM += (READER["Aire"].ToString().Trim().Length > 0) ? READER["Aire"].ToString().Trim() + " " : "";
                        desTSM += (READER["QC"].ToString().Trim().Length > 0) ? READER["QC"].ToString().Trim() + " " : "";
                        // np = READER["Clave"].ToString().Trim() + " : " + READER["Modelo"].ToString().Trim();
                        desTSM += (Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj " : "";
                        desTSM += (READER["EE"].ToString().Trim().Length > 0) ? READER["EE"].ToString().Trim() + " " : "";
                        desTSM += (READER["Vidrios"].ToString().Trim().Length > 0) ? READER["Vidrios"].ToString().Trim() + " " : "";
                        desTSM += (READER["BAire"].ToString().Trim().Length > 0) ? READER["BAire"].ToString().Trim() + " " : "";
                        desTSM += (READER["Sonido"].ToString().Trim().Length > 0) ? READER["Sonido"].ToString().Trim() + " " : ""; ;
                        desTSM += (READER["FN"].ToString().Trim().Length > 0) ? READER["FN"].ToString().Trim() + " " : "";
                        desTSM += (READER["DH"].ToString().Trim().Length > 0) ? READER["DH"].ToString().Trim() + " " : "";
                        desTSM += (READER["DT"].ToString().Trim().Length > 0) ? READER["DT"].ToString().Trim() + " " : "";
                        desTSM += (READER["RA"].ToString().Trim().Length > 0) ? READER["RA"].ToString().Trim() + " " : "";
                        desTSM += (READER["Cilindros"].ToString().Trim().Length > 0) ? READER["Cilindros"].ToString().Trim() + " " : "";
                        desTSM += (READER["FL"].ToString().Trim().Length > 0) ? READER["FL"].ToString().Trim() + " " : "";
                        desTSM += (READER["BF"].ToString().Trim().Length > 0) ? READER["BF"].ToString().Trim() + " " : "";
                        desTSM += (READER["PE"].ToString().Trim().Length > 0) ? READER["PE"].ToString().Trim() + " " : "";
                        desTSM += (READER["TP"].ToString().Trim().Length > 0) ? READER["TP"].ToString().Trim() + " " : "";
                        desTSM += (READER["TC"].ToString().Trim().Length > 0) ? READER["TC"].ToString().Trim() + " " : "";
                        desTSM += (READER["CodRaro"].ToString().Trim().Length > 0) ? READER["CodRaro"].ToString().Trim() + " " : "";

                        doQuery("INSERT INTO STD_" + Company +
                               "(Cia, TipoTar, Clave, Marca, Tipo, Modelo, DescripCia, DescripSimple, " +
                               "Equipado, Trans, Puertas, Vestiduras, ABS, Aire, QC, NPass, EE, Vidrios, BAire, Sonido, " +
                               "FN, DH, DT, RA, Cilindros, FL, BF, PE, TP, TC, CodRaro, "
                               + "DescripTSM)" +
                                "VALUES (" +
                               numCompany.ToString() +
                                     ", 0, '" +
                               READER["Clave"].ToString().Trim() + "', '" +
                               READER["Marca"].ToString().Trim() + "', '" +
                               READER["Tipo"].ToString().Trim() + "', " +
                               READER["Modelo"].ToString() + ", '" +
                               READER["DescripCia"].ToString().Trim() + "', '" +
                               READER["DescripSimple"].ToString().Trim() + "', '" +
                               READER["Equipado"].ToString().Trim() + "', '" +
                               READER["Trans"].ToString().Trim() + "', '" +
                               READER["Puertas"].ToString().Trim() + "', '" +
                               READER["Vestiduras"].ToString().Trim() + "', '" +
                               READER["ABS"].ToString().Trim() + "', '" +
                               READER["Aire"].ToString().Trim() + "', '" +
                               READER["QC"].ToString().Trim() + "', '" +
                                    ((Convert.ToInt32(READER["NPass"].ToString()) > 0) ? READER["NPass"].ToString().Trim() + "Pasaj" : "") + "', '" +
                               READER["EE"].ToString().Trim() + "', '" +
                               READER["Vidrios"].ToString().Trim() + "', '" +
                               READER["BAire"].ToString().Trim() + "', '" +
                               READER["Sonido"].ToString().Trim() + "', '" +
                               READER["FN"].ToString().Trim() + "', '" +
                               READER["DH"].ToString().Trim() + "', '" +
                               READER["DT"].ToString().Trim() + "', '" +
                               READER["RA"].ToString().Trim() + "', '" +
                               READER["Cilindros"].ToString().Trim() + "', '" +
                               READER["FL"].ToString().Trim() + "', '" +
                               READER["BF"].ToString().Trim() + "', '" +
                               READER["PE"].ToString().Trim() + "', '" +
                               READER["TP"].ToString().Trim() + "', '" +
                               READER["TC"].ToString().Trim() + "', '" +
                               READER["CodRaro"].ToString().Trim() + "', '" +
                               desTSM.Trim() +
                               "')"
                               ,
                               MyConnString
                           );
                    }
                    progCount++;
                    this.Invoke(this.updateStatusDelegate);
                }
                READER.Close();
                CONNECT.Close();
                MessageBox.Show("Tabla STD_" + Company + " Generada Correctamente", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error al generar STD_" + Company, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                this.myThread.Abort();
            }
        }     

          // AGREGAR TABLA ESTANDARIZADA A HOMOLOGACIÓN
        private void btn_AgregarCiaHom_Click(object sender, EventArgs e)
        {
            if (cmb_Cia.SelectedIndex >= 0)
            {
                tStart = DateTime.Now;
                String nomCia = cmb_Cia.SelectedItem.ToString();
                int numCia = Convert.ToInt32(txt_CiaID.Text);

                // VERIFICAR SI LA TABLA ESTANDARIZADA YA FUE AGREGADA A DATOS ESTANDARIZADOS
                int count = 0;
                try
                {
                    OleDbConnection CONNECT_COUNT = new OleDbConnection(MyConnString);
                    CONNECT_COUNT.Open();
                    OleDbCommand COMMAND_COUNT = new OleDbCommand("SELECT COUNT (Cia_" + nomCia.ToString() + ") FROM DatosEstandarizados WHERE (Cia_" + nomCia.ToString() + " <> '')", CONNECT_COUNT);

                    count = Convert.ToInt32(COMMAND_COUNT.ExecuteScalar());

                    CONNECT_COUNT.Close();
                }
                catch (Exception Ex)
                {
                    MessageBox.Show("EX:___" + Ex.ToString(), "Error al obtener registros estandarizados en Cia_" + numCia.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                }

                if (count == 0)
                {
                    if (nu_Precision.Value > 69)
                    {
                        this.myThread = null;
                        this.myThread = new Thread(
                                () => CIAToStand(numCia, nomCia, Convert.ToDouble(nu_Precision.Value / 100) - 0.001)
                            );
                        this.myThread.Start();
                    }
                    else {
                        MessageBox.Show("Favor de evaluar con una mayor precisión." , "Error" + numCia.ToString(), MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
                else
                {
                    MessageBox.Show("La compañía " + nomCia + " ya está agregada a la homologación", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            else
            {
                MessageBox.Show("Selecciona una compañía.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public void CIAToStand(Int32 numCia, String nomCia, Double rAccuracy)
        {
            // Acronym DataBase
            // Fields: Trans, Gear, Pts, Pass, Brakes, Vest, Sound, Equip, Air, AirBag, QC, Descrip
            List<String>[] AcroDB = getAcroDB();

            // DataTable from DatosEstandarizados
            DataTable DT_Std = new DataTable();
            DT_Std.Columns.Add("sTrans", typeof(String));
            DT_Std.Columns.Add("sGear", typeof(String));
            DT_Std.Columns.Add("sCyl", typeof(String));
            DT_Std.Columns.Add("sPts", typeof(String));
            DT_Std.Columns.Add("sPass", typeof(String));
            DT_Std.Columns.Add("sBrakes", typeof(String));
            DT_Std.Columns.Add("sVest", typeof(String));
            DT_Std.Columns.Add("sSound", typeof(String));
            DT_Std.Columns.Add("sEquip", typeof(String));
            DT_Std.Columns.Add("sAir", typeof(String));
            DT_Std.Columns.Add("sAirBag", typeof(String));
            DT_Std.Columns.Add("sQC", typeof(String));
            DT_Std.Columns.Add("sDescrip", typeof(String));
            // DataTable from New Company
            DataTable DT_New = new DataTable();
            DT_New.Columns.Add("sTrans", typeof(String));
            DT_New.Columns.Add("sGear", typeof(String));
            DT_New.Columns.Add("sCyl", typeof(String));
            DT_New.Columns.Add("sPts", typeof(String));
            DT_New.Columns.Add("sPass", typeof(String));
            DT_New.Columns.Add("sBrakes", typeof(String));
            DT_New.Columns.Add("sVest", typeof(String));
            DT_New.Columns.Add("sSound", typeof(String));
            DT_New.Columns.Add("sEquip", typeof(String));
            DT_New.Columns.Add("sAir", typeof(String));
            DT_New.Columns.Add("sAirBag", typeof(String));
            DT_New.Columns.Add("sQC", typeof(String));
            DT_New.Columns.Add("sDescrip", typeof(String));

            // CEVIC List from DatosEstandarizados
            List<String> CEVList = new List<String>();
            // Key (Clave) List from EstandarizadosCia
            List<String> NEWKList = new List<String>();
            // Key (Clave) List from EstandarizadosCia
            List<String> NEWTSMList = new List<String>();

            // Campos de CEVIC
            String cveCEVIC, Mar, Typ, Mod, cveCo;
            int nMod = 0;
            Boolean NMod = false;

            progCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PB = new OleDbCommand("SELECT COUNT (*) FROM STD_" + nomCia, CONNECTPB);
                
                progMax = Convert.ToInt32(COMMAND_PB.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex) {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error al obtener registros de STD_" + nomCia, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            try
            {
                OleDbConnection CONNECT_SD = new OleDbConnection(MyConnString);
                CONNECT_SD.Open();
                OleDbCommand COMMAND_SD = new OleDbCommand("SELECT DISTINCT Marca, Tipo, Modelo FROM STD_" + nomCia,
                    CONNECT_SD);
                OleDbDataReader READER_SD = COMMAND_SD.ExecuteReader();

                while (READER_SD.Read())
                {
                    NMod = false;
                    // Getting from NewCompany
                    try
                    {
                        OleDbConnection CONNECT_NEWC = new OleDbConnection(MyConnString);
                        CONNECT_NEWC.Open();
                        OleDbCommand COMMAND_NEWC = new OleDbCommand("SELECT * FROM STD_" + nomCia + " WHERE " +
                                "Marca = '" + READER_SD["Marca"].ToString() +
                                "' AND Tipo = '" + READER_SD["Tipo"].ToString() +
                                "' AND Modelo = " + READER_SD["Modelo"].ToString() + "",
                            CONNECT_NEWC);
                        OleDbDataReader READER_NEWC = COMMAND_NEWC.ExecuteReader();

                        NEWKList.Clear();
                        NEWTSMList.Clear();

                        DT_New.Rows.Clear();

                        while (READER_NEWC.Read())
                        {
                            NEWKList.Add(READER_NEWC["Clave"].ToString());
                            NEWTSMList.Add(READER_NEWC["DescripTSM"].ToString());

                            DT_New.Rows.Add(
                                READER_NEWC["Trans"].ToString().Trim(),
                                READER_NEWC["Trans"].ToString().Trim(),
                                READER_NEWC["Cilindros"].ToString().Trim(),
                                READER_NEWC["Puertas"].ToString().Trim(),
                                READER_NEWC["NPass"].ToString().Trim(),
                                READER_NEWC["ABS"].ToString().Trim(),
                                READER_NEWC["Vestiduras"].ToString().Trim(),
                                READER_NEWC["Sonido"].ToString().Trim(),
                                READER_NEWC["Equipado"].ToString().Trim(),
                                READER_NEWC["Aire"].ToString().Trim(),
                                READER_NEWC["BAire"].ToString().Trim(),
                                READER_NEWC["QC"].ToString().Trim(),
                                READER_NEWC["DescripSimple"].ToString().Trim()
                            );
                        }

                        READER_NEWC.Close();
                        CONNECT_NEWC.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show("EX:___" + Ex.ToString(), "Error en SELECT_FROM_NEW_COMP", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // Getting from DatosEstandarizados
                    try
                    {
                        OleDbConnection CONNECT_STD = new OleDbConnection(MyConnString);
                        CONNECT_STD.Open();
                        OleDbCommand COMMAND_STD = new OleDbCommand("SELECT * FROM DatosEstandarizados WHERE CEVIC IN (" +
                            "SELECT MyCEVICPool FROM (" +
                                "SELECT IIF([Cia_" + numCia.ToString() + "] Is Null, 'unavailable', [Cia_" + numCia.ToString() + "]) AS MyCIA, CEVIC AS MyCEVICPool " +
                                "FROM DatosEstandarizados WHERE " +
                                "Marca = '" + READER_SD["Marca"].ToString() +
                                "' AND Tipo = '" + READER_SD["Tipo"].ToString() +
                                "' AND Modelo = '" + READER_SD["Modelo"].ToString() + "') " +
                            "WHERE MyCIA = 'unavailable')",
                            CONNECT_STD);
                        OleDbDataReader READER_STD = COMMAND_STD.ExecuteReader();

                        // Modelo Disponible en DatosEstandarizados
                        if (READER_STD.HasRows)
                        {
                            //NMod = false;

                            CEVList.Clear();
                            DT_Std.Rows.Clear();

                            while (READER_STD.Read())
                            {
                                // Getting CEVIC
                                CEVList.Add(READER_STD["CEVIC"].ToString());

                                // Getting INFO from Model
                                String nTSM = " " + READER_STD["Descripcion"].ToString().Trim() + " ";
                                // Transmission
                                String nTrans = "";
                                for (int i = 0; i < AcroDB[0].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[0].ElementAt(i) + " ") && AcroDB[0].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[0].ElementAt(i) + " ", " ");
                                        nTrans += AcroDB[0].ElementAt(i) + " ";
                                    }
                                }
                                // GearBox
                                String nGear = "";
                                for (int i = 0; i < AcroDB[1].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[1].ElementAt(i) + " ") && AcroDB[1].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[1].ElementAt(i) + " ", " ");
                                        nGear += AcroDB[1].ElementAt(i) + " ";
                                    }
                                }
                                // Cylinders
                                String nCyl = "";
                                for (int i = 0; i < AcroDB[2].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[2].ElementAt(i) + " ") && AcroDB[2].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[2].ElementAt(i) + " ", " ");
                                        nCyl += AcroDB[2].ElementAt(i) + " ";
                                    }
                                }
                                // Passengers
                                String nPass = "";
                                for (int i = 0; i < AcroDB[3].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[3].ElementAt(i) + " ") && AcroDB[3].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[3].ElementAt(i) + " ", " ");
                                        nPass += AcroDB[3].ElementAt(i) + " ";
                                    }
                                }
                                // Doors
                                String nPts = "";
                                for (int i = 0; i < AcroDB[4].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[4].ElementAt(i) + " ") && AcroDB[4].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[4].ElementAt(i) + " ", " ");
                                        nPts += AcroDB[4].ElementAt(i) + " ";
                                    }
                                }
                                // Brakes
                                String nBrakes = "";
                                for (int i = 0; i < AcroDB[5].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[5].ElementAt(i) + " ") && AcroDB[5].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[5].ElementAt(i) + " ", " ");
                                        nBrakes += AcroDB[5].ElementAt(i) + " ";
                                    }
                                }
                                // Vest
                                String nVest = "";
                                for (int i = 0; i < AcroDB[6].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[6].ElementAt(i) + " ") && AcroDB[6].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[6].ElementAt(i) + " ", " ");
                                        nVest += AcroDB[6].ElementAt(i) + " ";
                                    }
                                }
                                // Sound
                                String nSound = "";
                                for (int i = 0; i < AcroDB[7].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[7].ElementAt(i) + " ") && AcroDB[7].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[7].ElementAt(i) + " ", " ");
                                        nSound += AcroDB[7].ElementAt(i) + " ";
                                    }
                                }
                                // Equipment
                                String nEquip = "";
                                for (int i = 0; i < AcroDB[8].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[8].ElementAt(i) + " ") && AcroDB[8].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[8].ElementAt(i) + " ", " ");
                                        nEquip += AcroDB[8].ElementAt(i) + " ";
                                    }
                                }
                                // AC Air
                                String nAir = "";
                                for (int i = 0; i < AcroDB[9].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[9].ElementAt(i) + " ") && AcroDB[9].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[9].ElementAt(i) + " ", " ");
                                        nAir += AcroDB[9].ElementAt(i) + " ";
                                    }
                                }
                                // AirBag
                                String nAirBag = "";
                                for (int i = 0; i < AcroDB[10].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[10].ElementAt(i) + " ") && AcroDB[10].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[10].ElementAt(i) + " ", " ");
                                        nAirBag += AcroDB[10].ElementAt(i) + " ";
                                    }
                                }
                                // QC
                                String nQC = "";
                                for (int i = 0; i < AcroDB[11].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[11].ElementAt(i) + " ") && AcroDB[11].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[11].ElementAt(i) + " ", " ");
                                        nQC += AcroDB[11].ElementAt(i) + " ";
                                    }

                                }

                                // Adding Model to DataTable
                                DT_Std.Rows.Add(
                                    nTrans.Trim(),
                                    nGear.Trim(),
                                    nCyl.Trim(),
                                    nPts.Trim(),
                                    nPass.Trim(),
                                    nBrakes.Trim(),
                                    nVest.Trim(),
                                    nSound.Trim(),
                                    nEquip.Trim(),
                                    nAir.Trim(),
                                    nAirBag.Trim(),
                                    nQC.Trim(),
                                    nTSM.Trim()
                                );

                            }
                        }
                        else
                        {
                            NMod = true;
                            // Modelo No Disponible en DatosEstandarizados
                            // Add All NEWKList to DATOSESTAND

                            Mar = (READER_SD["Marca"].ToString().Length > 3) ?
                                    (READER_SD["Marca"].ToString()).Substring(0, 3) :
                                    (READER_SD["Marca"].ToString());
                            Typ = (READER_SD["Tipo"].ToString().Length > 2) ?
                                (READER_SD["Tipo"].ToString()).Substring(0, 2) :
                                (READER_SD["Tipo"].ToString());
                            Mod = READER_SD["Modelo"].ToString();

                            for (int i = 0; i < NEWKList.Count; i++)
                            {
                                // Generando CEVIC
                                cveCEVIC = Mar + Typ + Mod + NEWKList.ElementAt(i) + "_X00";

                                String myQuery = "INSERT INTO DatosEstandarizados " +
                                            "(Cia_" + numCia.ToString() + ", Cia_Disponible, CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                            "VALUES ('" +
                                            NEWKList.ElementAt(i) + "', '" +
                                            numCia.ToString() + "| ', '" +
                                            cveCEVIC + "', '" +
                                            Mod +
                                            "', '', '', '', '', '" +
                                            READER_SD["Marca"].ToString() + "', '" +
                                            READER_SD["Tipo"].ToString() + "', '" +
                                            NEWTSMList.ElementAt(i) + "')";

                                doQuery(myQuery
                                ,
                                MyConnString
                                );
                            }

                            NEWKList.Clear();
                            NEWTSMList.Clear();
                        }

                        READER_STD.Close();
                        CONNECT_STD.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show("EX:___" + Ex.ToString(), "Error en SELECT_FROM_DATOS_STD", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }

                    // Evaluate Similarty
                    // Si el Modelo esta en DatosEstandarizados
                    if (NMod == false)
                    {

                        Int32[] MResult = new Int32[DT_New.Rows.Count];
                        MResult = evMatModels(DT_Std, DT_New, rAccuracy);

                        for (Int32 i = 0; i < MResult.Length; i++)
                        {
                            switch (MResult[i])
                            {
                                case -2:
                                    MessageBox.Show("Registro no evaluado.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                // No es Match, Insertar Nuevo Registro
                                case -1:
                                    // Generando CEVIC
                                    Mar = (READER_SD["Marca"].ToString().Length > 3) ?
                                        (READER_SD["Marca"].ToString()).Substring(0, 3) :
                                        (READER_SD["Marca"].ToString());
                                    Typ = (READER_SD["Tipo"].ToString().Length > 2) ?
                                        (READER_SD["Tipo"].ToString()).Substring(0, 2) :
                                        (READER_SD["Tipo"].ToString());
                                    Mod = READER_SD["Modelo"].ToString();
                                    cveCo = NEWKList.ElementAt(i);
                                    cveCEVIC = Mar + Typ + Mod + cveCo + "_X";

                                    try
                                    {
                                        OleDbConnection CONNECT_NR = new OleDbConnection(MyConnString);
                                        CONNECT_NR.Open();
                                        // SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '8514_X??'
                                        OleDbCommand COMMAND_CEVIC = new OleDbCommand("SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '" + cveCEVIC + "__'", CONNECT_NR);
                                        nMod = Convert.ToInt32(COMMAND_CEVIC.ExecuteScalar());
                                        CONNECT_NR.Close();

                                        cveCEVIC += nMod.ToString("D2");
                                        String myQuery = "INSERT INTO DatosEstandarizados " +
                                                "(Cia_" + numCia.ToString() + ", Cia_Disponible, CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                                "VALUES ('" +
                                                cveCo + "', '" +
                                                numCia.ToString() + "| ', '" +
                                                cveCEVIC + "', '" +
                                                Mod +
                                                "', '', '', '', '', '" +
                                                READER_SD["Marca"].ToString() + "', '" +
                                                READER_SD["Tipo"].ToString() + "', '" +
                                                NEWTSMList.ElementAt(i) + "')";

                                        doQuery(myQuery
                                            ,
                                            MyConnString
                                        );

                                    }
                                    catch (Exception Ex)
                                    {
                                        MessageBox.Show("EX:___" + Ex.ToString(), "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    break;
                                // Match Exitoso, Insercion de Referencia
                                default:
                                    if (MResult[i] >= 0 && MResult[i] <= DT_Std.Rows.Count)
                                    {
                                        try
                                        {
                                            OleDbConnection CONNECT_NR = new OleDbConnection(MyConnString);
                                            CONNECT_NR.Open();
                                            // SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '8514_X??'
                                            OleDbCommand COMMAND_CEVIC = new OleDbCommand("SELECT Cia_Disponible FROM DatosEstandarizados WHERE CEVIC = '" + CEVList.ElementAt(MResult[i]) + "'", CONNECT_NR);
                                            String cDisp = COMMAND_CEVIC.ExecuteScalar().ToString();
                                            CONNECT_NR.Close();

                                            String myQuery = "UPDATE DatosEstandarizados " +
                                                "SET Cia_" + numCia.ToString() + " = '" + NEWKList.ElementAt(i) + "'" +
                                                ", Cia_Disponible = '" + sortDescrip(cDisp.Trim() + " " + numCia.ToString() + "| ", false).Trim() + "' " +
                                                "WHERE CEVIC = '" + CEVList.ElementAt(MResult[i]) + "' " +
                                                "AND Modelo = '" + READER_SD["Modelo"].ToString() + "'";

                                            doQuery(myQuery
                                                ,
                                                MyConnString
                                            );
                                        }
                                        catch (Exception ex)
                                        {
                                            MessageBox.Show(ex.ToString());
                                        }

                                    }
                                    else
                                    {
                                        MessageBox.Show("Registro no válido.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    }
                                    break;
                            }
                        }
                    }
                }
                READER_SD.Close();
                CONNECT_SD.Close();

                MessageBox.Show("Compañía " + nomCia + "agregada a la tabla estandarizada.", "Éxito", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception Ex)
            {
                MessageBox.Show("EX:___" + Ex.ToString(), "Error en SELECT_DISTINCT_MAR/TIP/MOD", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

            // EVALUAR MATRIZ DE DESCRIPCIONES
        public int[] evMatModels(DataTable DT_Std, DataTable DT_New, Double evAccuracy)
        {
            // JaroWinkler Object
            var Jw = new JaroWinkler();

            // Matriz que almacena los coeficientes de similaridad de la descripcion simple
            double[,] MatSimD = new double[DT_New.Rows.Count, DT_Std.Rows.Count];
            // Matriz que almacena el número de campos que los modelos tienen en común
            int[,] MatSimF = new int[DT_New.Rows.Count, DT_Std.Rows.Count];

            int NewCount = 0;
            foreach (DataRow RowNModel in DT_New.Rows)
            {
                int StdCount = 0;
                foreach (DataRow RowSModel in DT_Std.Rows)
                {
                    // Transmision
                    if (eqDescrip(RowNModel.Field<String>(0), RowSModel.Field<String>(0)) || // Mismo valor para el campo
                        RowSModel.Field<String>(0).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(0), RowSModel.Field<String>(0)) > evAccuracy) // Alta similaridad en el campo
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Gearbox | Caja de Cambios
                    if (eqDescrip(RowNModel.Field<String>(1), RowSModel.Field<String>(1)) ||
                        RowNModel.Field<String>(1).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(1), RowSModel.Field<String>(1)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Cilindros
                    if (eqDescrip(RowNModel.Field<String>(2), RowSModel.Field<String>(2)) ||
                        RowSModel.Field<String>(2).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(2), RowSModel.Field<String>(2)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Pasajeros
                    if (eqDescrip(RowNModel.Field<String>(3), RowSModel.Field<String>(3)) ||
                        RowSModel.Field<String>(3).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(3), RowSModel.Field<String>(3)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Puertas
                    if (eqDescrip(RowNModel.Field<String>(4), RowSModel.Field<String>(4)) ||
                        RowSModel.Field<String>(4).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(4), RowSModel.Field<String>(4)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Frenos
                    if (eqDescrip(RowNModel.Field<String>(5), RowSModel.Field<String>(5)) ||
                        RowSModel.Field<String>(5).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(5), RowSModel.Field<String>(5)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Vestiduras
                    if (eqDescrip(RowNModel.Field<String>(6), RowSModel.Field<String>(6)) ||
                        RowSModel.Field<String>(6).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(6), RowSModel.Field<String>(6)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Sonido
                    if (eqDescrip(RowNModel.Field<String>(7), RowSModel.Field<String>(7)) ||
                        RowSModel.Field<String>(7).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(7), RowSModel.Field<String>(7)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Equipamiento
                    if (eqDescrip(RowNModel.Field<String>(8), RowSModel.Field<String>(8)) ||
                        RowSModel.Field<String>(8).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(8), RowSModel.Field<String>(8)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Aire
                    if (eqDescrip(RowNModel.Field<String>(9), RowSModel.Field<String>(9)) ||
                        RowSModel.Field<String>(9).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(9), RowSModel.Field<String>(9)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Bolsa de Aire
                    if (eqDescrip(RowNModel.Field<String>(10), RowSModel.Field<String>(10)) ||
                        RowSModel.Field<String>(10).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(10), RowSModel.Field<String>(10)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // QC
                    if (eqDescrip(RowNModel.Field<String>(11), RowSModel.Field<String>(11)) ||
                        RowSModel.Field<String>(11).Length == 0 || // El registro estandarizado carece de valor para el campo
                        Jw.GetSimilarity(RowNModel.Field<String>(11), RowSModel.Field<String>(11)) > evAccuracy)
                    {
                        MatSimF[NewCount, StdCount]++;
                    }
                    // Descripcion Simple
                    MatSimD[NewCount, StdCount] =
                        // Verificando si el modelo estandarizado contiene Descripcion Simple
                        (RowSModel.Field<String>(12).Length > 0) ?
                            Math.Max(
                                Math.Max(
                        // Descripcion ordenada alfabeticamente
                                    Jw.GetSimilarity(sortDescrip(RowNModel.Field<String>(12), false),
                                        sortDescrip(RowSModel.Field<String>(12), false)),
                        // Descripcion ordenada inversamente
                                    Jw.GetSimilarity(sortDescrip(RowNModel.Field<String>(12), true),
                                        sortDescrip(RowSModel.Field<String>(12), true))
                                     ),
                        // Descripcion por defecto
                                Jw.GetSimilarity(RowNModel.Field<String>(12),
                                        RowSModel.Field<String>(12))
                                )
                            :
                        // Si la descripcion estandarizada está vacía, son compatibles
                            0.7;
                    StdCount++;
                }
                NewCount++;
            }

            // Vector de resultados
            // -2: El modelo no ha sido evaluado
            // -1: El modelo no tiene equivalencia
            // N: Match con N (N >= 0)
            int[] Result = new int[DT_New.Rows.Count];
            for (int i = 0; i < DT_New.Rows.Count; i++)
            {
                Result[i] = -2;
            }

            // Valor Maximo obtenido para cada modelo
            Double Max;
            Tuple<int, int> posMax;

            // Matches
            Int32 maxMatches = (DT_New.Rows.Count > DT_Std.Rows.Count) ? DT_Std.Rows.Count : DT_Std.Rows.Count;
            Int32 nMatches = 0;

            do
            {
                // Obteniendo coeficiente Maximo
                Max = MatSimD.Cast<Double>().Max();
                // Obteniendo Posicion del coeficiente Maximo
                posMax = getIndex(MatSimD, Max);
                //MessageBox.Show(MatSimD.GetValue(posMax.Item1, posMax.Item2).ToString(), posMax.ToString() + Max.ToString());

                if (Max > evAccuracy)
                {
                    // Campos compatibles, Match resuelto
                    if (MatSimF[posMax.Item1, posMax.Item2] >= Convert.ToInt32(evAccuracy * 12))
                    {
                        for (int j = 0; j < DT_Std.Rows.Count; j++)
                        {
                            MatSimD[posMax.Item1, j] = 0;
                        }
                        for (int i = 0; i < DT_New.Rows.Count; i++)
                        {
                            MatSimD[i, posMax.Item2] = 0;
                        }

                        Result[posMax.Item1] = posMax.Item2;
                        nMatches++;
                    }
                    // Campos incompatibles, Match no completado
                    else
                    {
                        MatSimF[posMax.Item1, posMax.Item2] = 0;
                    }
                }
                // Descripcion Simple no compatible
                else
                {
                    // Forzando Match en caso de no cumplir
                    if (nMatches < maxMatches)
                    {
                        for (int j = 0; j < DT_Std.Rows.Count; j++)
                        {
                            MatSimD[posMax.Item1, j] = 0;
                        }
                        for (int i = 0; i < DT_New.Rows.Count; i++)
                        {
                            MatSimD[i, posMax.Item2] = 0;
                        }

                        Result[posMax.Item1] = posMax.Item2;
                        nMatches++;
                        continue;
                    }
                    else
                    {
                        for (int i = 0; i < DT_New.Rows.Count; i++)
                        {
                            if (Result[i] == -2)
                                Result[i] = -1;
                        }
                        break;
                    }
                }
            } while (nMatches < maxMatches);


            for (int i = 0; i < DT_New.Rows.Count; i++)
            {
                if (Result[i] == -2)
                    Result[i] = -1;
            }
            return Result;
        }

        // PANEL 5 - CAMPO DE PRUEBAS
        private void btnTest1_Click(object sender, EventArgs e)
        {
            double[,] Arry = { { 1.0, 7.25, 4.64 }, { 4.23, 7.251, 4.65 }, { 1.5, 4.7, 4.78 }, { 1.2, 5.6513, 5.6529 } };
            double Max = Arry.Cast<double>().Max();
            Tuple<int, int> position = getIndex(Arry, Max);
            MessageBox.Show(position.ToString(), Max.ToString());
        }

    }
}

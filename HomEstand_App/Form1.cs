using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
// RECURSOS
using HomEstand_App.Properties;
using System.Reflection;
// CSV Read - CATALOGOS
using System.IO;
// Access DB - BASE DE DATOS
using System.Data.OleDb;
// 
using SimMetricsMetricUtilities;
// TUPLES
using System.Collections;

namespace HomEstand_App
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }


        // VARIABLES GLOBALES 
        // Conexion a la Base de Datos
        String MyConnString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=C:\\Users\\Griselda\\Desktop\\H&E_DB.mdb;";

        // Catalogo de Compañías
        String[] CiasID = null;

        // CARGA DEL FORMULARIO
        private void Form1_Load(object sender, EventArgs e)
        {
            //
            load_CATCias();
        }

        // COMPROBAR SI EXISTE UNA TABLA
        public bool TableExists(OleDbConnection CONN, String table)
        {
            CONN.Open();
            var exists = CONN.GetSchema("Tables", new string[4] { null, null, table, "TABLE" }).Rows.Count > 0;
            CONN.Close();
            return exists;
        }

        // FUNCIONES PARA CATALOGOS
        // COMPAÑIAS
        public List<String>[] get_CATCias()
        {
            var currentAssembly = Assembly.GetExecutingAssembly();
            using (var stream = currentAssembly.GetManifestResourceStream("HomEstand_App.Cias_WS.csv"))
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

        private void load_CATCias()
        {
            // Obteniendo Catalogo de CIAS
            List<String>[] CATCias = get_CATCias();

            OleDbConnection CONNECT = new OleDbConnection(MyConnString);

            for (int i = 0; i < CATCias[0].Count; i++)
            {
                if (TableExists(CONNECT, "D_" + CATCias[2].ElementAt(i).ToString()))
                {
                    cmb_Cia.Items.Add(CATCias[2].ElementAt(i).ToString());
                    CiasID[i] = CATCias[0].ElementAt(i).ToString();
                }
            }
        }

        // Get Acro DB
        // Get Acronym Database from .CSV
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

                MessageBox.Show("Acronym Database Loaded", "SUCCESS");
                return acrDB;
            }
        }


        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            txt_CiaID.Text = CiasID[cmb_Cia.SelectedIndex];
        }


        private void btnTest1_Click(object sender, EventArgs e)
        {
            // OleDbConnection CONNECT = new OleDbConnection(MyConnString);
            //ssageBox.Show(TableExists(CONNECT, "D_" + txtTest1.Text).ToString());
            //MessageBox.Show(Resources.Cias_WS);
        }

        // EJECUTAR CONSULTA
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
                MessageBox.Show("QUERY: ___" + CONS + "___ EX:___" + e.ToString(), "ERROR IN QUERY");
            }

        }

        public Boolean equalDescrip(String a, String b)
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

        public String sortDescrip(String a)
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

            listA.Sort();

            aWord = "";
            foreach (String Val in listA)
            {
                aWord += Val + " ";
            }

            return aWord;
        }

        public Int32 evalReader(DataTable myStandData)
        {
            // JaroWinkler Object
            var Jw = new JaroWinkler();

            int RowC = 0;

            double[] MatchD = new double[myStandData.Rows.Count];
            Int32[] MatchF = new Int32[myStandData.Rows.Count];

            DataRow MyModel = null;

            // Evaluando Elementos
            foreach (DataRow Row in myStandData.Rows)
            {
                if (RowC == 0)
                {
                    MyModel = Row;
                }
                else
                {
                    // Transmision
                    if (equalDescrip(MyModel.Field<String>(0), Row.Field<String>(0)) || Row.Field<String>(0).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // GearBox
                    if (equalDescrip(MyModel.Field<String>(1), Row.Field<String>(1)) || Row.Field<String>(1).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Cilindros
                    if (equalDescrip(MyModel.Field<String>(2), Row.Field<String>(2)) || Row.Field<String>(2).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Pasajeros
                    if (equalDescrip(MyModel.Field<String>(3), Row.Field<String>(3)) || Row.Field<String>(3).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Puertas
                    if (equalDescrip(MyModel.Field<String>(4), Row.Field<String>(4)) || Row.Field<String>(4).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // ABS
                    if (equalDescrip(MyModel.Field<String>(5), Row.Field<String>(5)) || Row.Field<String>(5).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Vestidura
                    if (equalDescrip(MyModel.Field<String>(6), Row.Field<String>(6)) || Row.Field<String>(6).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Sonido
                    if (equalDescrip(MyModel.Field<String>(7), Row.Field<String>(7)) || Row.Field<String>(7).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Equipado
                    if (equalDescrip(MyModel.Field<String>(8), Row.Field<String>(8)) || Row.Field<String>(8).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // Aire
                    if (equalDescrip(MyModel.Field<String>(9), Row.Field<String>(9)) || Row.Field<String>(9).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // BAire
                    if (equalDescrip(MyModel.Field<String>(10), Row.Field<String>(10)) || Row.Field<String>(10).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // QC
                    if (equalDescrip(MyModel.Field<String>(11), Row.Field<String>(11)) || Row.Field<String>(11).Length == 0)
                    {
                        MatchF[RowC]++;
                    }
                    // DescripSimple
                    MatchD[RowC] = (Row.Field<String>(12).Length > 0) ?
                        Jw.GetSimilarity(sortDescrip(MyModel.Field<String>(12)), sortDescrip(Row.Field<String>(12)))
                        : 0.70;
                    /*
                    if (Levenshtein.CalculateDistance(
                            sortDescrip(MyModel.Field<String>(12)), sortDescrip(Row.Field<String>(12)), 1)
                            < Convert.ToInt32(Math.Max(Row.Field<String>(12).Length, MyModel.Field<String>(12).Length) * 0.7))
                    {
                             MatchD[RowC] = true;
                    }  */
                }
                RowC++;
            }

            int bestMatch = 0;
            for (int i = 0; i < MatchF.Length; i++)
            {
                if (i == 0)
                {
                    MatchD[i] = 0.699;
                    MatchF[i] = 6;
                }
                else
                {
                    if (MatchF[i] > MatchF[bestMatch] && MatchD[i] > MatchD[bestMatch])
                        bestMatch = i;
                }
            }

            return bestMatch = (MatchF[bestMatch] > 7 && MatchD[bestMatch] > 0.7) ? bestMatch : 0;
        }


        public void insertNewComp(int numCia, String nomCia)
        {
            OleDbConnection CONNECT = new OleDbConnection(MyConnString);
            OleDbConnection CONNECT2 = new OleDbConnection(MyConnString);
            // Cambiar nombre de acuerdo con la Tabla de la Cia
            String CONSULT_COMP = "SELECT * FROM Estandarizados_" + nomCia; // +"2";
            String CONSULT_NEW = "SELECT * FROM DatosEstandarizados WHERE ";

            // Campos de CEVIC
            String cveCEVIC, Mar, Typ, Mod, cveCo;
            int nMod = 0;

            // Acronym DataBase
            // Fields: Trans, Gear, Pts, Pass, Brakes, Vest, Sound, Equip, Air, AirBag, QC, Descrip
            List<String>[] AcroDB = getAcroDB();

            // DataTable from DatosEstandarizados
            DataTable dTStand = new DataTable();
            dTStand.Columns.Add("sTrans", typeof(String));
            dTStand.Columns.Add("sGear", typeof(String));
            dTStand.Columns.Add("sCyl", typeof(String));
            dTStand.Columns.Add("sPts", typeof(String));
            dTStand.Columns.Add("sPass", typeof(String));
            dTStand.Columns.Add("sBrakes", typeof(String));
            dTStand.Columns.Add("sVest", typeof(String));
            dTStand.Columns.Add("sSound", typeof(String));
            dTStand.Columns.Add("sEquip", typeof(String));
            dTStand.Columns.Add("sAir", typeof(String));
            dTStand.Columns.Add("sAirBag", typeof(String));
            dTStand.Columns.Add("sQC", typeof(String));
            dTStand.Columns.Add("sDescrip", typeof(String));

            List<String> CEVList = new List<String>();

            // Progress Count
            // pBCount = 0;
            /* try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PBAR = new OleDbCommand("SELECT COUNT (*) FROM Estandarizados_" + nomCia, CONNECTPB);

                pBMax = Convert.ToInt32(COMMAND_PBAR.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "Error en Count de Registros");
            }
            //this.Invoke(this.setBarDelegate);
             * */

            try
            {
                CONNECT.Open();
                OleDbCommand COMMAND_COMP = new OleDbCommand(CONSULT_COMP, CONNECT);
                OleDbDataReader READER_COMP = COMMAND_COMP.ExecuteReader();

                while (READER_COMP.Read())
                {
                    try
                    {
                        CONNECT2.Open();
                        OleDbCommand COMMAND_NEW = new OleDbCommand(CONSULT_NEW +
                            "CEVIC IN (" +
                                "SELECT MyCEVICPool FROM (" +
                                    "SELECT IIF([Cia_" + numCia.ToString() + "] Is Null, 'unavailable', [Cia_" + numCia.ToString() + "]) AS MyCIA, CEVIC AS MyCEVICPool FROM DatosEstandarizados " +
                                    "WHERE Marca = '" + READER_COMP["Marca"].ToString() + "' AND Tipo = '" + READER_COMP["Tipo"].ToString() + "' AND Modelo = '" + READER_COMP["Modelo"].ToString() + "')" +
                                "WHERE MyCIA = 'unavailable')",
                            CONNECT2);
                        OleDbDataReader READER_NEW = COMMAND_NEW.ExecuteReader();

                        String myQuery = "";
                        // POSIBLE MATCH
                        if (READER_NEW.HasRows)
                        {
                            CEVList.Clear();
                            dTStand.Rows.Clear();
                            // Adding New Model as First Row
                            dTStand.Rows.Add(
                                    READER_COMP["Trans"].ToString(),
                                    READER_COMP["Trans"].ToString(),
                                    READER_COMP["Cilindros"].ToString(),
                                    READER_COMP["Puertas"].ToString(),
                                    READER_COMP["NPass"].ToString(),
                                    READER_COMP["ABS"].ToString(),
                                    READER_COMP["Vestiduras"].ToString(),
                                    READER_COMP["Sonido"].ToString(),
                                    READER_COMP["Equipado"].ToString(),
                                    READER_COMP["Aire"].ToString(),
                                    READER_COMP["BAire"].ToString(),
                                    READER_COMP["QC"].ToString(),
                                    READER_COMP["DescripSimple"].ToString()
                                );

                            while (READER_NEW.Read())
                            {
                                CEVList.Add(READER_NEW["CEVIC"].ToString());

                                String dsTrans = "", dsGear = "", dsCyl = "", dsPts = "", dsPass = "", dsBrakes = "",
                                       dsVest = "", dsSound = "", dsEquip = "", dsAir = "", dsAirBag = "", dsQC = "";
                                String nTSM = " " + READER_NEW["Descripcion"].ToString().Trim() + " ";

                                // Getting Info. from Models
                                for (int i = 0; i < AcroDB[0].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[0].ElementAt(i) + " ") && AcroDB[0].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[0].ElementAt(i) + " ", " ");
                                        dsTrans += AcroDB[0].ElementAt(i) + " ";
                                        //MessageBox.Show("Contains Trans: " + dsTrans);
                                    }
                                }
                                for (int i = 0; i < AcroDB[1].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[1].ElementAt(i) + " ") && AcroDB[1].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[1].ElementAt(i) + " ", " ");
                                        dsGear += AcroDB[1].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[2].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[2].ElementAt(i) + " ") && AcroDB[2].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[2].ElementAt(i) + " ", " ");
                                        dsCyl += AcroDB[2].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[3].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[3].ElementAt(i) + " ") && AcroDB[3].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[3].ElementAt(i) + " ", " ");
                                        dsPass += AcroDB[3].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[4].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[4].ElementAt(i) + " ") && AcroDB[4].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[4].ElementAt(i) + " ", " ");
                                        dsPts += AcroDB[4].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[5].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[5].ElementAt(i) + " ") && AcroDB[5].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[5].ElementAt(i) + " ", " ");
                                        dsBrakes += AcroDB[5].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[6].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[6].ElementAt(i) + " ") && AcroDB[6].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[6].ElementAt(i) + " ", " ");
                                        dsVest += AcroDB[6].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[7].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[7].ElementAt(i) + " ") && AcroDB[7].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[7].ElementAt(i) + " ", " ");
                                        dsSound += AcroDB[7].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[8].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[8].ElementAt(i) + " ") && AcroDB[8].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[8].ElementAt(i) + " ", " ");
                                        dsEquip += AcroDB[8].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[9].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[9].ElementAt(i) + " ") && AcroDB[9].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[9].ElementAt(i) + " ", " ");
                                        dsAir += AcroDB[9].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[10].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[10].ElementAt(i) + " ") && AcroDB[10].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[10].ElementAt(i) + " ", " ");
                                        dsAirBag += AcroDB[10].ElementAt(i) + " ";
                                    }
                                }
                                for (int i = 0; i < AcroDB[11].Count; i++)
                                {
                                    if (nTSM.Contains(" " + AcroDB[11].ElementAt(i) + " ") && AcroDB[11].ElementAt(i) != "NA")
                                    {
                                        nTSM = nTSM.Replace(" " + AcroDB[11].ElementAt(i) + " ", " ");
                                        dsQC += AcroDB[11].ElementAt(i) + " ";
                                    }

                                }

                                //MessageBox.Show("Descrip: _" + nTSM.Trim() + "_");
                                dTStand.Rows.Add(
                                    dsTrans.Trim(),
                                    dsGear.Trim(),
                                    dsCyl.Trim(),
                                    dsPts.Trim(),
                                    dsPass.Trim(),
                                    dsBrakes.Trim(),
                                    dsVest.Trim(),
                                    dsSound.Trim(),
                                    dsEquip.Trim(),
                                    dsAir.Trim(),
                                    dsAirBag.Trim(),
                                    dsQC.Trim(),
                                    nTSM.Trim()
                                );
                            }

                            // EVALUATE SIMILARTY
                            int standRef = evalReader(dTStand);
                            // Si es mayor a 0, insertar Referencia (Modelo Equivalente)
                            if (standRef > 0)
                            {
                                myQuery = "UPDATE DatosEstandarizados SET Cia_" + numCia.ToString() + " = '" + READER_COMP["Clave"].ToString() +
                                "' WHERE CEVIC = '" + CEVList.ElementAt(standRef - 1) + "' AND Modelo = '" + READER_COMP["Modelo"].ToString() + "'";
                                doQuery(myQuery
                                    ,
                                    MyConnString
                                );

                                myQuery = "";
                                //MessageBox.Show("MATCH EXITOSO EN REGISTRO", "INFO");
                                CEVList.Clear();
                            }
                            // Sin Equivalencia, Nuevo Registro
                            else
                            {
                                // CALCULAR CEVIC
                                Mar = (READER_COMP["Marca"].ToString().Length > 3) ? (READER_COMP["Marca"].ToString()).Substring(0, 3) : (READER_COMP["Marca"].ToString());
                                Typ = (READER_COMP["Tipo"].ToString().Length > 2) ? (READER_COMP["Tipo"].ToString()).Substring(0, 2) : (READER_COMP["Tipo"].ToString());
                                Mod = READER_COMP["Modelo"].ToString();
                                cveCo = READER_COMP["Clave"].ToString();
                                cveCEVIC = Mar + Typ + Mod + cveCo + "_X";

                                // COMPROBAR SI CEVIC ESTÁ EN LA TABLA
                                try
                                {
                                    OleDbConnection CONNECT3 = new OleDbConnection(MyConnString);
                                    CONNECT3.Open();
                                    // SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '8514_X??'
                                    OleDbCommand COMMAND_CEVIC = new OleDbCommand("SELECT COUNT (*) FROM DatosEstandarizados WHERE CEVIC LIKE '" + cveCEVIC + "__'", CONNECT);
                                    nMod = (Int32)COMMAND_CEVIC.ExecuteScalar();
                                    CONNECT3.Close();

                                    cveCEVIC += nMod.ToString("D2");
                                    myQuery = "INSERT INTO DatosEstandarizados" +
                                        // Cambiar Nº de Compañia
                                            "(Cia_" + numCia.ToString() + ", CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                            "VALUES ('" +
                                            cveCo + "','" +
                                            cveCEVIC + "','" +
                                            Mod +
                                            "','','','','','" +
                                            READER_COMP["Marca"].ToString() + "','" +
                                            READER_COMP["Tipo"].ToString() + "','" +
                                            READER_COMP["DescripTSM"].ToString() + "')";

                                    doQuery(myQuery
                                    ,
                                    MyConnString
                                    );

                                }
                                catch (Exception ex)
                                {
                                    MessageBox.Show(ex.ToString());
                                }
                                myQuery = "";
                                cveCEVIC = "";

                                //break;
                                //MessageBox.Show("New");
                            }
                        }
                        else
                        {
                            Mar = (READER_COMP["Marca"].ToString().Length > 3) ? (READER_COMP["Marca"].ToString()).Substring(0, 3) : (READER_COMP["Marca"].ToString());
                            Typ = (READER_COMP["Tipo"].ToString().Length > 2) ? (READER_COMP["Tipo"].ToString()).Substring(0, 2) : (READER_COMP["Tipo"].ToString());
                            Mod = READER_COMP["Modelo"].ToString();
                            nMod = 0;
                            cveCo = READER_COMP["Clave"].ToString();
                            cveCEVIC = Mar + Typ + Mod + cveCo + "_X" + nMod.ToString("D2");

                            myQuery = "INSERT INTO DatosEstandarizados" +
                                // Cambiar Nº de Compañia
                                "(Cia_" + numCia.ToString() + ", CEVIC, Modelo, CveMarca_Cia, CveTipo_Cia, CveVersion_Cia, CveTrans_Cia, Marca, Tipo, Descripcion)" +
                                "VALUES ('" +

                                cveCo + "','" +
                                cveCEVIC + "','" +
                                Mod +
                                "','','','','','" +
                                READER_COMP["Marca"].ToString() + "','" +
                                READER_COMP["Tipo"].ToString() + "','" +
                                READER_COMP["DescripTSM"].ToString() + "')";

                            doQuery(myQuery
                                ,
                                MyConnString
                            );

                            myQuery = "";
                            cveCEVIC = "";
                            //MessageBox.Show("NUEVO REGISTRO, NUEVO AUTO", "INFO");
                        }
                        // Progress Update
                        //pBCount++;
                        //this.Invoke(this.updateStatusDelegate);

                        READER_NEW.Close();
                        CONNECT2.Close();
                    }
                    catch (Exception e)
                    {
                        MessageBox.Show(e.ToString(), "ERROR ON INSERT ON DATOS ESTANDARIZADOS");
                    }
                }
                READER_COMP.Close();
                CONNECT.Close();
                MessageBox.Show("SUCCESS");
            }
            catch (Exception e)
            {
                MessageBox.Show(e.ToString(), "ERROR ON GETTING FROM COMPANY");
            }
            finally
            {
                //this.workerThread.Abort();
            }
        }

        public void CIAToStand(int numCia, String nomCia)
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
            // Key (Clave) List from DatosEstandarizados
            List<String> NEWKList = new List<String>();

            // Campos de CEVIC
            String cveCEVIC, Mar, Typ, Mod, cveCo;
            int nMod = 0;

            /* Progress Count
            pBCount = 0;
            try
            {
                OleDbConnection CONNECTPB = new OleDbConnection(MyConnString);
                CONNECTPB.Open();
                OleDbCommand COMMAND_PB = new OleDbCommand("SELECT COUNT (*) FROM Estandarizados_" + nomCia, CONNECTPB);
                
                pBMax = Convert.ToInt32(COMMAND_PB.ExecuteScalar());

                CONNECTPB.Close();
            }
            catch (Exception Ex) {
                MessageBox.Show(Ex.ToString(), "ERROR EN PB_COUNT");
            }
            //this.Invoke(this.setBarDelegate);
            */
            try
            {
                OleDbConnection CONNECT_SD = new OleDbConnection(MyConnString);
                CONNECT_SD.Open();
                OleDbCommand COMMAND_SD = new OleDbCommand("SELECT DISTINCT Marca, Tipo, Modelo FROM Estandarizados_" + nomCia,
                    CONNECT_SD);
                OleDbDataReader READER_SD = COMMAND_SD.ExecuteReader();

                while (READER_SD.Read())
                {
                    // Getting from NewCompany
                    try
                    {
                        OleDbConnection CONNECT_NEWC = new OleDbConnection(MyConnString);
                        CONNECT_NEWC.Open();
                        OleDbCommand COMMAND_NEWC = new OleDbCommand("SELECT * FROM Estandarizados_" + nomCia + " WHERE " +
                                "Marca = '" + READER_SD["Marca"].ToString() +
                                "' AND Tipo = '" + READER_SD["Tipo"].ToString() +
                                "' AND Modelo = '" + READER_SD["Modelo"].ToString() + "'",
                            CONNECT_NEWC);
                        OleDbDataReader READER_NEWC = COMMAND_NEWC.ExecuteReader();

                        NEWKList.Clear();
                        DT_New.Rows.Clear();

                        while (READER_NEWC.Read())
                        {
                            NEWKList.Add(READER_NEWC["Clave"].ToString());

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
                        MessageBox.Show(Ex.ToString(), "ERROR EN SELECT_FROM_NEW_COMP");
                    }

                    // Getting from DatosEstandarizados
                    try
                    {
                        OleDbConnection CONNECT_STD = new OleDbConnection(MyConnString);
                        CONNECT_STD.Open();
                        OleDbCommand COMMAND_STD = new OleDbCommand("SELECT * FROM DatosEstandarizados WHERE CEVIC IN (" +
                            "SELECT MyCEVICPool FROM (" +
                                "SELECT IIF([Cia_" + numCia.ToString() + "] Is Null, 'unavailable', Cia_" + numCia.ToString() + "]) AS MyCIA, CEVIC AS MyCEVICPool " +
                                "FROM DatosEstandarizados WHERE " +
                                "Marca = '" + READER_SD["Marca"].ToString() +
                                "' AND Tipo = '" + READER_SD["Tipo"].ToString() +
                                "' AND Modelo = '" + READER_SD["Modelo"].ToString() + "') " +
                            "WHERE MyCIA = 'unavailable')",
                            CONNECT_STD);
                        OleDbDataReader READER_STD = COMMAND_STD.ExecuteReader();

                        if (READER_STD.HasRows)
                        {
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
                            // Add All NEWKList to DATOSESTAND
                            foreach (String CEV in NEWKList) {
                                
                            }
                            NEWKList.Clear();
                        }

                        READER_STD.Close();
                        CONNECT_STD.Close();
                    }
                    catch (Exception Ex)
                    {
                        MessageBox.Show(Ex.ToString(), "ERROR EN SELECT_FROM_DATOS_STD");
                    }
                }

                // Evaluate Similarty
                if (DT_Std.Rows.Count > 0)
                {
                    Int32[] MResult = new Int32[DT_New.Rows.Count];

                    MResult = evMatModels(DT_Std, DT_New);


                }
                else { 
                
                }

                READER_SD.Close();
                CONNECT_SD.Close();
            }
            catch (Exception Ex)
            {
                MessageBox.Show(Ex.ToString(), "ERROR EN SELECT_DISTINCT_MAR/TIP/MOD");
            }
        }

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

        public int[] evMatModels(DataTable DT_Std, DataTable DT_New) 
        {
            // JaroWinkler Object
            var Jw = new JaroWinkler();

            int[] Result = new int[DT_New.Rows.Count];

            double [,] MatSimDes =  new double[DT_New.Rows.Count, DT_Std.Rows.Count];
            int [,] MatSimF =  new int[DT_New.Rows.Count, DT_Std.Rows.Count];

            int StdCount = 0;
            foreach(DataRow RowSModel in DT_Std.Rows) {
                int NewCount = 0;
                foreach(DataRow RowNModel in DT_New.Rows) {
                    // Transmission
                    if(equalDescrip(RowNModel.Field<String>(0), RowSModel.Field<String>(0)) || 
                        RowNModel.Field<String>(0).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Gearbox
                    if(equalDescrip(RowNModel.Field<String>(1), RowSModel.Field<String>(1)) || 
                        RowNModel.Field<String>(1).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Cylinders
                    if(equalDescrip(RowNModel.Field<String>(2), RowSModel.Field<String>(2)) || 
                        RowNModel.Field<String>(2).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Passengers
                    if(equalDescrip(RowNModel.Field<String>(3), RowSModel.Field<String>(3)) || 
                        RowNModel.Field<String>(3).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Doors
                    if(equalDescrip(RowNModel.Field<String>(4), RowSModel.Field<String>(4)) || 
                        RowNModel.Field<String>(4).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Brakes
                    if(equalDescrip(RowNModel.Field<String>(5), RowSModel.Field<String>(5)) || 
                        RowNModel.Field<String>(5).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Vest
                    if(equalDescrip(RowNModel.Field<String>(6), RowSModel.Field<String>(6)) || 
                        RowNModel.Field<String>(6).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Sound
                    if(equalDescrip(RowNModel.Field<String>(7), RowSModel.Field<String>(7)) || 
                        RowNModel.Field<String>(7).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Equipment
                    if(equalDescrip(RowNModel.Field<String>(8), RowSModel.Field<String>(8)) || 
                        RowNModel.Field<String>(8).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // Air
                    if(equalDescrip(RowNModel.Field<String>(9), RowSModel.Field<String>(9)) || 
                        RowNModel.Field<String>(9).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // AirBag
                    if(equalDescrip(RowNModel.Field<String>(10), RowSModel.Field<String>(10)) || 
                        RowNModel.Field<String>(10).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }
                    // QC
                    if(equalDescrip(RowNModel.Field<String>(11), RowSModel.Field<String>(11)) || 
                        RowNModel.Field<String>(11).Length == 0) {
                        MatSimF[StdCount, NewCount]++;
                    }

                    MatSimDes[StdCount, NewCount] = 
                        (RowSModel.Field<String>(12).Length > 0 ) ?  
                            Jw.GetSimilarity(sortDescrip(RowNModel.Field<String>(12)), 
                                sortDescrip(RowSModel.Field<String>(12))) 
                            :
                            0.7;

                    NewCount++;
                }
                StdCount++;
            }
                
            for (int i = 0; i < DT_Std.Rows.Count; i++) {
                // Get MAX AND EVALUATE 
                for(int j = 0; j < DT_Std.Rows.Count; i++) {
                    
                }

            }

            double Max;
            do
            {
                Max = MatSimDes.Cast<double>().Max();


            } while (Max > 0);

            // PROVI
            int [] X = { 1, 2 };
            return X;

            ////////////////////////////////////////////////////////////
            
            /*
            int bestMatch = 0;
            for (int i = 0; i < MatchF.Length; i++) {
                if (i == 0) {
                    MatchD[i] = 0.7;
                    MatchF[i] = 6;
                }
                else
                {
                    if (MatchF[i] > MatchF[bestMatch] && MatchD[i] > MatchD[bestMatch])
                        bestMatch = i;
                }
            }

            return bestMatch = (MatchF[bestMatch] > 7 && MatchD[bestMatch] > 0.7) ? bestMatch : 0;
        */
           }



    }
}

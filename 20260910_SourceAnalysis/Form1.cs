using ICSharpCode.SharpZipLib.Zip;
using MathNet.Numerics;
using MathNet.Numerics.Statistics;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Text;
using Org.BouncyCastle.Asn1.X509; // Clipboard 用
using Org.BouncyCastle.Crypto.Agreement.Srp;
using Org.BouncyCastle.Ocsp;
using ScottPlot;
using ScottPlot.Control;
using ScottPlot.Drawing.Colormaps;
using ScottPlot.Plottable;
using ScottPlot.Renderable;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.WindowsRuntime;
using System.Runtime.Remoting.Contexts;
using System.Runtime.Versioning;
using System.Security.Cryptography;
using System.Security.Principal;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Interop;
using static EMS_Small_Chb.Form1;
using static EMS_Small_Chb.Json;
using static System.Net.WebRequestMethods;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using Excel = Microsoft.Office.Interop.Excel;
//using SixLabors.ImageSharp.Metadata.Profiles.Exif;

namespace EMS_Small_Chb
{
    public partial class Form1 : Form
    {
        #region 定数定義
        public const int BRACKET_CSIZE = 12;
        public const int BRACKET_LSIZE = 5;
        public const int RECDATAMAXLENGTH = 2000;
        public const int TEMPMAX = 200;
        public const int TEMPMIN = -100;
        public const int HUMIMAX = 110;
        public const int HUMIMIN = 0;

        public const int HUMISETSTEPS = 4;
        public const int TEMPSETSTEPS = 6;

        public const int HUMIENDCNT = 500;

        public const int STABLE_STATE = 0;
        public const int UP_STATE = 1;
        public const int DOWN_STATE = 2;
        public const int NON_STATE = 3;
        public const int JUDGE_TIMES = 30;
        public const int JUDGE_AFTER_TIME = 10;
        // ゾンタークの式の係数
        public const double S1 = -6096.9385;
        public const double S2 = 21.2409642;
        public const double S3 = -0.02711193;
        public const double S4 = 0.00001673952;
        public const double S5 = 2.433502;

        #endregion
        #region 構造体定義

        public class RecData
        {
            public string[] Date = new string[RECDATAMAXLENGTH];
            public string[] Time = new string[RECDATAMAXLENGTH];
            public double[] Times = new double[RECDATAMAXLENGTH];
            public double[] TempSV = new double[RECDATAMAXLENGTH];
            public double[] HumiSV = new double[RECDATAMAXLENGTH];
            public double[] TempPV = new double[RECDATAMAXLENGTH];
            public double[] WetTempPV = new double[RECDATAMAXLENGTH];
            public double[] RoomTemp = new double[RECDATAMAXLENGTH];
            public double[] HumiPV = new double[RECDATAMAXLENGTH];
        }
        public class ChkSpec 
        {
            public String InspectPrgName = "";
            public String CheckSheetName = "";
            public bool HumiUse = false;
            public bool Temp180Use = false;
            public bool WindowUse = false;
            public bool IndoorUse = false;
            public bool OpeWindowUse = false;
            public bool RefLeakTestUse = false;

            public int MashineNo = 0;
            public int PoleUpSpec = 60;
            public float PoleUpStartTemp = -40;
            public float PoleUpEndTemp = 150;
            public int PoleDownSpec = 50;
            public float PoleDownStartTemp = 20;
            public float PoleDownEndTemp = -40;

            public float ChgUpSpec = 3.2f;
            public float ChgUpStartTemp = -21;
            public float ChgUpEndTemp = 131;
            public float ChgDownSpec = 2.1f;
            public float ChgDownStartTemp = 131;
            public float ChgDownEndTemp = -21;

            public float AccuracyTemp1Spec = 1.5f;
            public float AccuracyTemp1StartTemp = -40.0f;
            public float AccuracyTemp1EndTemp = 100.0f;

            public float AccuracyTemp2Spec = 1.9f;
            public float AccuracyTemp2StartTemp = 100.1f;
            public float AccuracyTemp2EndTemp = 150.0f;

            public float AccuracyTemp3Spec = 6.1f;
            public float AccuracyTemp3StartTemp = 150.1f;
            public float AccuracyTemp3EndTemp = 180.0f;

            public float VarTemp1Spec = 0.3f;
            public float VarTemp1StartTemp = -40;
            public float VarTemp1EndTemp = 100;
            public float VarTemp2Spec = 0.5f;
            public float VarTemp2StartTemp = 100.1f;
            public float VarTemp2EndTemp = 150;
            public float VarTemp3Spec = 0.5f;
            public float VarTemp3StartTemp = 150.1f;
            public float VarTemp3EndTemp = 180;
            public float VarHumiSpec = 0.3f;
            public float SpecMaxTemp = 150f;
            public float SpecMinTemp = -40f;
            public float SpecMaxUpTemp = 157f;
            public float SpecMinDownTemp = -43f;

        }
        public class ChgRate
        {
            public float StartTemp = float.MinValue;
            public float EndTemp = float.MinValue;
            public int StartXPos = int.MinValue;
            public int EndXPos = int.MinValue;
            public float StartYValue = float.MinValue;
            public float EndYValue = float.MinValue;
            public float ResultValue = float.MinValue;
            public float ChangeRate = float.MinValue;
            public bool StartExist = false;
            public bool EndExist = false;
        }
        public class JdgCond 
        {
            public float LimitWidth = (float)1.0;
            public float ChgJudgeTemp = (float)2.0;
            public float MaxMinWidth = (float)0.3;
            public float StableCount = (float)10.0;
            public int ChgToleranceLimit = 5;
            public int JdgTimes = JUDGE_TIMES;
            public int JdgAfterTime = JUDGE_AFTER_TIME;
        }
        public class WIDTH_FLUCT 
        {
            public int StartPosition = 0;
            public int EndPosition = 0;
            public float SetValue = float.MinValue;         // 確認設定値
            public float StartValue = float.MinValue;
            public float EndValue = float.MinValue;
            public float TempSvAve = float.MaxValue;        // 温度制御値平均
            public float HumiSvAve = float.MaxValue;        // 湿度制御値平均
            public float TempPvAve = float.MaxValue;        // 温度モニタ乾球温度値平均
            public float WetPvAve = float.MaxValue;         // 温度モニタ湿球温度値平均
            public float HumiPvAve = float.MaxValue;        // 相対湿度値平均
            public float TempSvStd = float.MaxValue;        // 温度制御値　標準偏差
            public float HumiSvStd = float.MaxValue;        // 湿度制御値　標準偏差
            public float TempPvStd = float.MaxValue;        // 温度モニタ乾球温度値　標準偏差
            public float WetPvStd = float.MaxValue;         // 温度モニタ湿球温度値　標準偏差
            public float HumiPvStd = float.MaxValue;        // 相対湿度値　標準偏差
            public float Std = float.MaxValue;              // 
            public bool Exist = false;                      // 検出記憶
            public bool Judge = false;                      // 判定結果
        }
        public class STATE_INFO 
        {
            public bool Temp1 = false;
            public bool Temp2 = false;
            public bool PoleUpChgStart = false;
            public bool PoleUpChgEnd = false;
        }
        public class HUMI_SET 
        {
            public float SetTemp = float.MinValue;
            public float SetHumi = float.MinValue;
            public float HumiAccuracy = float.MinValue;
            public float SetTime = float.MinValue;
        }
        public class TEMP_SET
        {
            public float SetTemp = float.MinValue;
            public float SetTime = float.MinValue;
        }
        public class Option
        {
            public String OptionCode { get; set; } = "";
            public String OptionName { get; set; } = "";
            public String OptionKind { get; set; } = "";
            public String OptionValue { get; set; } = "";
        }
        public class Header
        {
            public int MashinCode = 0;
            public string CreateNo = "";
            public String PuroductModel = "";
            public String PuroductName = "";
            public String PuroductOrder = "";
            public int VoltSpec = 200;
            public int FreqSpec = 60;
            public int TempSpec = 100;
            public int WVHoleSpec = 0;
            public String UserName = "";
            public String[] device_name = new string[4];
            public String den_serial_no = "";
            public String rec_serial_no1 = "";
            public String rec_serial_no2 = "";
            public List<Option> Option = new List<Option>();
        }

        public class ExcelData 
        {
            public int EndRecord = 0;
            public int HumiEndRecord = 0;
            public int TempEndRecord = 0;
            public string NowCSVFilePath = "";
            public string NowExcelFileName = "";
            public string ModelName = "";
            public string SerialNo="";
            public string UserName = "";
            public bool CheckWindow = false;
            public bool InnerDoor = false;
            public bool OperationHole = false;
            public float RoomTempAverage = 0f;
            public float PoleDownRoomTempAverage = 0f;
            public string AdditionalSpecifications = "";
            public int[] CableHole = new int[]{0,0,0,0};
            public string TestingPerson = "";
            public string TestingPersonEn = "";
            public string TestingDate = "";
            public string TestingVolt = "";
            public string TestingFreq = "";
            public string TestingRecorderNo = "";
            public string TestingRecorderModel = "";
            public string TestingRecorderSerialNo = "";
            public FormsPlot formsPlot1=null;
            public string MasterPath = "";

            public string lblTempReachMemL = "";
            public string lblTempReachMemH = "";
            public string lblTempReachJdgeL = "";
            public string lblTempReachJdgeH = "";

            public string lblTempOpeMemL = "";
            public string lblTempOpeMemH = "";
            public string lblTempOpeJdgeL = "";
            public string lblTempOpeJdgeH = "";

            public string lblHumiOpeHumi3 = "";
            public string lblHumiOpeHumi4 = "";

            public string lblChgUpStart = "";
            public string lblChgUpEnd = "";
            public string lblChgUpSpec = "";
            public string lblChgUpStr = "";
            public string lblChgUpValue = "";
            public string lblChgUpTime = "";

            public string lblPoleUpStart = "";
            public string lblPoleUpEnd = "";
            public string lblPoleUpSpec = "";
            public string lblPoleUpStr = "";
            public string lblPoleUpValue = "";

            public string lblChgDownStart = "";
            public string lblChgDownEnd = "";
            public string lblChgDownSpec = "";
            public string lblChgDownStr = "";
            public string lblChgDownValue = "";
            public string lblChgDownTime = "";

            public string lblPoleDownStart = "";
            public string lblPoleDownEnd = "";
            public string lblPoleDownSpec = "";
            public string lblPoleDownStr = "";
            public string lblPoleDownValue = "";

            public string lblVarTemp1Start = "";
            public string lblVarTemp1End = "";
            public string lblVarTemp1Spec = "";

            public string lblTempFluctMemM = "";
            public string lblTempFluctMemL = "";

            public string lblTempFluctJdgeL = "";
            public string lblTempFluctJdgeM = "";

            public string lblVarTemp2Start = "";
            public string lblVarTemp2End = "";
            public string lblVarTemp2Spec = "";
            public string lblTempFluctMemH = "";
            public string lblTempFluctJdgeH = "";

            public string lblHumiFluctStd1 = "";
            public string lblHumiFluctStd2 = "";
            public string lblHumiFluctStd3 = "";
            public string lblHumiFluctStd4 = "";

            public string lblVarHumiSpec = "";

            public string lblHumiFluctJdge1 = "";
            public string lblHumiFluctJdge2 = "";
            public string lblHumiFluctJdge3 = "";
            public string lblHumiFluctJdge4 = "";

            public string lblCsvFileName = "";

            private static CURRENTS Cv = new CURRENTS();

        }
        private static ExcelData ExcelValue= new ExcelData();
        private static int ChkDataMaxlength = 1200;
        #region コントロール配列
        //ボタンコントロール配列のフィールドを作成
        private System.Windows.Forms.TextBox[] MemberBox;
        private System.Windows.Forms.TextBox[] MemberNoBox;
        private System.Windows.Forms.TextBox[] MemberENBox;

        private System.Windows.Forms.TextBox[] VoltKind;
        private System.Windows.Forms.TextBox[] RecorderNoBox;
        private System.Windows.Forms.TextBox[] RecorderModelBox;
        private System.Windows.Forms.TextBox[] RecorderSerialNoBox;
        private System.Windows.Forms.RadioButton[] OpeTemp;
        private System.Windows.Forms.RadioButton[] OpeHumi;
        private System.Windows.Forms.RadioButton[] CentTemp;
        private System.Windows.Forms.RadioButton[] CentHumi;
        private System.Windows.Forms.RadioButton[] Humi;
        private System.Windows.Forms.RadioButton[] RoomTemp;
        private System.Windows.Forms.Label[] HumiOpeSet;
        private System.Windows.Forms.Label[] HumiOpeTemp;
        private System.Windows.Forms.Label[] HumiOpeHumi;
        private System.Windows.Forms.Label[] HumiOpeJdge;
        private System.Windows.Forms.Label[] HumiOpeMonTemp;
        private System.Windows.Forms.Label[] HumiOpeMonHumi;
        private System.Windows.Forms.Label[] HumiOpeHumiPV;
        private System.Windows.Forms.Label[] HumiFluctStd;
        private System.Windows.Forms.Label[] HumiFluctJdge;
        #endregion コントロール配列
        #endregion 構造体定義
        #region クラスインスタンス定義
        // Class 登録
        //private static Access Db;
        #endregion

        #region Public変数
        // スレッドセーフなキュー
        private static ConcurrentQueue<string> DataQueue = new ConcurrentQueue<string>();

        public static bool bExcelWriteBusy = false;
        public static bool bExcelWriteEnd = false;
        public static bool bExcelWriteFail= false;
        public static bool DebugMode = false;
        public static bool FIRST_STEP = false;
        public static bool FLUCT_USE = false;
        public static ChgRate ChgUpValue = new ChgRate();
        public static ChgRate ChgDownValue = new ChgRate();
        public static ChgRate ExtUpValue = new ChgRate();
        public static ChgRate ExtDownValue = new ChgRate();
        public static JdgCond JdgCnd = new JdgCond();
        public static ChkSpec ChkNowSpec = new ChkSpec();
        public static List<HUMI_SET> Humiset = new List<HUMI_SET>();
        public static List<TEMP_SET> Tempset = new List<TEMP_SET>();
        public static String sTestCheckMsg1 = "";
        public static String sTestCheckMsg2 = "";
        public static List<WIDTH_FLUCT> TempFluct = new List<WIDTH_FLUCT>();
        public static List<WIDTH_FLUCT> HumiFluct = new List<WIDTH_FLUCT>();
        public static WIDTH_FLUCT MaxTemp = new WIDTH_FLUCT();
        public static WIDTH_FLUCT MinTemp = new WIDTH_FLUCT();
        public static string NowCSVFilePath = "";
        public static string NowExcelFileName = "";
        public static ulong Data_id { get; set; }
        public static Header HeaderData = new Header();
        #endregion Public変数

        #region Private変数
        private static int iFirstChgRecord = 0; 
        private List<CURRENT> CurrentInfo = new List<CURRENT>();

        private static string user = "";
        private static string username = "";
        private static System.Collections.ArrayList devList = new System.Collections.ArrayList();
        private static bool FLUCT20_Exist = false;
        #region Jsonファイル読込用
        private Json.MASHINE Mashine = new Json.MASHINE();
        private Json.MEMBER Member = new Json.MEMBER();
        private Json.RECORDER Recorders = new Json.RECORDER();
        private static Json.SYSTEM SystemSet = new Json.SYSTEM();
        private static Json.DATEBASE DatabaseInfo = new Json.DATEBASE();
        private Json.JUDGEMENT Judgment = new Json.JUDGEMENT();
        private Json.OPTION OptionSet = new Json.OPTION();
        #endregion Jsonファイル読込用
        private static RecData RecValue = new RecData();
        private STATE_INFO StateInfo = new STATE_INFO();
        private String[] MemberName = new string[10];
        private String ModelName = "";
        private String UserName = "";
        private String SerialNo = "";
        private bool CheckWindow = false;
        private bool InnerDoor = false;
        private bool OperationHole = false;
        private int[] CableHole = { 0, 1, 0, 0 };
        private String AdditionalSpecifications = "";
        private String TestingPerson = "";
        private String TestingPersonEn = "";
        private static String TestingDate = "";
        private static String TestingVolt = "100V";
        private static String TestingFreq = "50Hz";
        private String TestingRecorderNo = "";
        private String TestingRecorderModel = "";
        private String TestingRecorderSerialNo = "";
        #region CSVファイル読込制御
        private static int ChkLength = 0;
        private static int RecLength = 0;
        private static int RecLengthbak = 0;
        private static int ChgCheckToleranceCount = 0;
        private static int PoleCheckToleranceCount = 0;
        private static int HumiEndRecord = 0;
        private static int TempEndRecord = 0;
        private static bool tCSVReadRun = false;
        private static bool tCSVReadEnd = false;
        private static bool tCSVReadStop = false;
        private static String tCSVMsg = "";
        private static ScatterPlot[] scat = new ScatterPlot[6];
        private static ScatterPlot[] bracketFree = new ScatterPlot[8];
        private static int vMin = TEMPMIN;
        private static int vMax = TEMPMAX;
        private static float RoomTempAverage = 0;
        private static float PoleDownRoomTempAverage = 0;

        #endregion CSVファイル読込制御
        private static String AppVer = "";
        private static string MasterPath = Directory.GetCurrentDirectory() + "\\Document\\";
        #region バーコード関連変数
        private static string BarCodeData = "";
        private static string ReturnBarCodeData = "";
        private static bool BarCodeStart = false;
        private static char[] modulus43CharList =
        {
            '0','1','2','3','4','5','6','7','8','9',
            'A','B','C','D','E','F','G','H','I','J',
            'K','L','M','N','O','P','Q','R','S','T',
            'U','V','W','X','Y','Z','-','.',' ','$',
            '/','+','%'
        };
        #endregion バーコード関連変数
        #region DataBase 関連変数
        private string strConnString = "";      // Oracle
        private static bool bFirstFlg = true;
        private static bool bDBReadyFlg = false;
        private static bool bDBReadyEnd = false;

        // データベース有効フラグ
        private static bool bDBOracleFlg = false;
        private static int bDBReadyCnt = 0;

        #endregion DataBase 関連変数
        #endregion

        public Form1()
        {
            InitializeComponent();

            //コマンドライン引数を表示する
            Console.WriteLine(System.Environment.CommandLine);

            //コマンドライン引数を配列で取得する
            string[] cmds = System.Environment.GetCommandLineArgs();
            //コマンドライン引数を列挙する
            foreach (string cmd in cmds)
            {
                Console.WriteLine(cmd);
            }
            
            if (cmds.Length == 2 ) 
            {
                if (cmds[1] == "admin") { this.tabControl1.TabPages.Insert(0, this.tabPage3); DebugMode = true; }
                else this.tabControl1.TabPages.Remove(this.tabPage3);
            }
            else this.tabControl1.TabPages.Remove(this.tabPage3);
           
        }

        /// モジュラス43　計算
        /// CODE39
        /// 
        private static string GetModulus43(string Value)
        {
            try
            {
                if (!Regex.IsMatch(Value, @"^[A-Z|0-9|\-|.| |$|/|+|%]+$"))
                {
                    throw new FormatException();
                }

                long x = 0;
                for (int i = 0; i < Value.Length; i++)
                {
                    x += Array.IndexOf(modulus43CharList, Value[i]);
                }

                return modulus43CharList[x % 43].ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return "";
            }
        }
        private void ControlArayCreate() 
        {
            //メンバーテキストボックスコントロール配列の作成
            this.MemberBox = new System.Windows.Forms.TextBox[10];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.MemberBox = new System.Windows.Forms.TextBox[]
                {this.txtMember01, this.txtMember02, this.txtMember03, this.txtMember04, this.txtMember05,
                 this.txtMember06, this.txtMember07, this.txtMember08, this.txtMember09, this.txtMember10 };

            //製造番号テキストボックスコントロール配列の作成
            this.MemberNoBox = new System.Windows.Forms.TextBox[10];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.MemberNoBox = new System.Windows.Forms.TextBox[]
                {this.txtNo01, this.txtNo02, this.txtNo03, this.txtNo04, this.txtNo05,
                 this.txtNo06, this.txtNo07, this.txtNo08, this.txtNo09, this.txtNo10 };

            //英文名テキストボックスコントロール配列の作成
            this.MemberENBox = new System.Windows.Forms.TextBox[10];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.MemberENBox = new System.Windows.Forms.TextBox[]
                {this.txtEN01, this.txtEN02, this.txtEN03, this.txtEN04, this.txtEN05,
                 this.txtEN06, this.txtEN07, this.txtEN08, this.txtEN09, this.txtEN10 };

            //記録計品番テキストボックスコントロール配列の作成
            this.RecorderNoBox = new System.Windows.Forms.TextBox[16];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.RecorderNoBox = new System.Windows.Forms.TextBox[]
                {this.txtRecNo01, this.txtRecNo02, this.txtRecNo03, this.txtRecNo04, this.txtRecNo05,
                 this.txtRecNo06, this.txtRecNo07, this.txtRecNo08, this.txtRecNo09, this.txtRecNo10,
                 this.txtRecNo11, this.txtRecNo12, this.txtRecNo13, this.txtRecNo14, this.txtRecNo15, this.txtRecNo16};

            //記録計型式テキストボックスコントロール配列の作成
            this.RecorderModelBox = new System.Windows.Forms.TextBox[16];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.RecorderModelBox = new System.Windows.Forms.TextBox[]
                {this.txtRecModel01, this.txtRecModel02, this.txtRecModel03, this.txtRecModel04, this.txtRecModel05,
                 this.txtRecModel06, this.txtRecModel07, this.txtRecModel08, this.txtRecModel09, this.txtRecModel10,
                 this.txtRecModel11, this.txtRecModel12, this.txtRecModel13, this.txtRecModel14, this.txtRecModel15, this.txtRecModel16};

            //記録計製造番号テキストボックスコントロール配列の作成
            this.RecorderSerialNoBox = new System.Windows.Forms.TextBox[16];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.RecorderSerialNoBox = new System.Windows.Forms.TextBox[]
                {this.txtRecSerialNo01, this.txtRecSerialNo02, this.txtRecSerialNo03, this.txtRecSerialNo04, this.txtRecSerialNo05,
                 this.txtRecSerialNo06, this.txtRecSerialNo07, this.txtRecSerialNo08, this.txtRecSerialNo09, this.txtRecSerialNo10,
                 this.txtRecSerialNo11, this.txtRecSerialNo12, this.txtRecSerialNo13, this.txtRecSerialNo14, this.txtRecSerialNo15, this.txtRecSerialNo16};

            //電源電圧テキストボックスコントロール配列の作成
            this.VoltKind = new System.Windows.Forms.TextBox[8];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.VoltKind = new System.Windows.Forms.TextBox[]
                {this.txtVoltKind01, this.txtVoltKind02, this.txtVoltKind03, this.txtVoltKind04, this.txtVoltKind05 , this.txtVoltKind06 , this.txtVoltKind07 , this.txtVoltKind08};

            //制御温度ラジオボタンコントロール配列の作成
            this.OpeTemp = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.OpeTemp = new System.Windows.Forms.RadioButton[]
                {this.rdbOpeTempColor1, this.rdbOpeTempColor2, this.rdbOpeTempColor3, this.rdbOpeTempColor4, this.rdbOpeTempColor5 , this.rdbOpeTempColor6 };

            //制御湿度ラジオボタンコントロール配列の作成
            this.OpeHumi = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.OpeHumi = new System.Windows.Forms.RadioButton[]
                {this.rdbOpeHumiColor1, this.rdbOpeHumiColor2, this.rdbOpeHumiColor3, this.rdbOpeHumiColor4, this.rdbOpeHumiColor5 , this.rdbOpeHumiColor6 };

            //中央乾球温度ラジオボタンコントロール配列の作成
            this.CentTemp = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.CentTemp = new System.Windows.Forms.RadioButton[]
                {this.rdbCentTempColor1, this.rdbCentTempColor2, this.rdbCentTempColor3, this.rdbCentTempColor4, this.rdbCentTempColor5 , this.rdbCentTempColor6 };

            //中央湿球温度ラジオボタンコントロール配列の作成
            this.CentHumi = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.CentHumi = new System.Windows.Forms.RadioButton[]
                {this.rdbCentHumiColor1, this.rdbCentHumiColor2, this.rdbCentHumiColor3, this.rdbCentHumiColor4, this.rdbCentHumiColor5 , this.rdbCentHumiColor6 };

            //相対湿度ラジオボタンコントロール配列の作成
            this.Humi = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.Humi = new System.Windows.Forms.RadioButton[]
                {this.rdbHumiColor1, this.rdbHumiColor2, this.rdbHumiColor3, this.rdbHumiColor4, this.rdbHumiColor5 , this.rdbHumiColor6 };

            //室内温度ラジオボタンコントロール配列の作成
            this.RoomTemp = new System.Windows.Forms.RadioButton[6];
            //テキストボックスコントロールの配列にすでに作成されているインスタンスを代入
            this.RoomTemp = new System.Windows.Forms.RadioButton[]
                {this.rdbRTColor1, this.rdbRTColor2, this.rdbRTColor3, this.rdbRTColor4, this.rdbRTColor5 , this.rdbRTColor6 };

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeSet = new System.Windows.Forms.Label[4];
            this.HumiOpeSet = new System.Windows.Forms.Label[]
            { this.lblHumiOpeSet1,this.lblHumiOpeSet2,this.lblHumiOpeSet3,this.lblHumiOpeSet4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeTemp = new System.Windows.Forms.Label[4];
            this.HumiOpeTemp = new System.Windows.Forms.Label[]
            { this.lblHumiOpeTemp1,this.lblHumiOpeTemp2,this.lblHumiOpeTemp3,this.lblHumiOpeTemp4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeHumi = new System.Windows.Forms.Label[4];
            this.HumiOpeHumi = new System.Windows.Forms.Label[]
            { this.lblHumiOpeHumi1,this.lblHumiOpeHumi2,this.lblHumiOpeHumi3,this.lblHumiOpeHumi4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeJdge = new System.Windows.Forms.Label[4];
            this.HumiOpeJdge = new System.Windows.Forms.Label[]
            { this.lblHumiOpeJdge1,this.lblHumiOpeJdge2,this.lblHumiOpeJdge3,this.lblHumiOpeJdge4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeMonTemp = new System.Windows.Forms.Label[4];
            this.HumiOpeMonTemp = new System.Windows.Forms.Label[]
            { this.lblHumiOpeMonTemp1,this.lblHumiOpeMonTemp2,this.lblHumiOpeMonTemp3,this.lblHumiOpeMonTemp4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeMonHumi = new System.Windows.Forms.Label[4];
            this.HumiOpeMonHumi = new System.Windows.Forms.Label[]
            { this.lblHumiOpeMonHumi1,this.lblHumiOpeMonHumi2,this.lblHumiOpeMonHumi3,this.lblHumiOpeMonHumi4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiOpeHumiPV = new System.Windows.Forms.Label[4];
            this.HumiOpeHumiPV = new System.Windows.Forms.Label[]
            { this.lblHumiOpeHumiPV1,this.lblHumiOpeHumiPV2,this.lblHumiOpeHumiPV3,this.lblHumiOpeHumiPV4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiFluctStd = new System.Windows.Forms.Label[4];
            this.HumiFluctStd = new System.Windows.Forms.Label[]
            { this.lblHumiFluctStd1,this.lblHumiFluctStd2,this.lblHumiFluctStd3,this.lblHumiFluctStd4};

            // 温湿度制御ラベルコントロール配列の作成
            this.HumiFluctJdge = new System.Windows.Forms.Label[4];
            this.HumiFluctJdge = new System.Windows.Forms.Label[]
            { this.lblHumiFluctJdge1,this.lblHumiFluctJdge2,this.lblHumiFluctJdge3,this.lblHumiFluctJdge4};

        }
        private void MemberNameGet() 
        {
            //設定されているコンボボックスの初期化
            string User = Environment.UserName;
            int uNo = 0;
            cmbMemberName.Items.Clear();
            //設定されているメンバーの名前を取り込む
            for (int i = 0; i < 10; i++)
            {
                MemberName[i] = MemberBox[i].Text;
                cmbMemberName.Items.Add(MemberName[i]);
                if (Member.Member[i].No == User) uNo = i;
            }
            cmbMemberName.SelectedIndex = uNo;
        }
        private  void DbOpen() 
        {
            //String cPath = Directory.GetCurrentDirectory() + "\\Document\\";
            //cPath += "EMS_Small_Chb.accdb";
            //Db.DbPathSet(cPath);
            //Db.DbOpen();
        }
        private void DbRead() 
        {
            // 担当者
            System.Data.DataTable dataTable = new System.Data.DataTable();
            string sql = "SELECT * FROM 担当者";
            // Db.DbAllSelect(sql,ref dataTable);
            // 取得結果表示
            //foreach (DataRow dr in dataTable.Rows)
            //{
            //    Console.WriteLine((String)dr.ItemArray[0]);
            //}

            // 記録計
            // 製品仕様
        }
        private void DbWrite()
        {
            // 担当者
            // 記録計
            // 製品仕様
        }
        private void DbClose() 
        {
            //Db.DbClose();
        }
        private String DataGet() 
        {
            String Msg = "";
            // コントロールに設定されている内容を設定ファイルへ書き込むデータに設定

            //
            CableHole[0]=int.Parse(cmbCableHole25.Text);
            CableHole[1] = int.Parse(cmbCableHole50.Text);
            CableHole[2] = int.Parse(cmbCableHole100.Text);
            CableHole[3] = int.Parse(cmbCableHoleFlat.Text);
            ExcelValue.CableHole[0] = CableHole[0];
            ExcelValue.CableHole[1] = CableHole[1];
            ExcelValue.CableHole[2] = CableHole[2];
            ExcelValue.CableHole[3] = CableHole[3];

            // 担当者
            for (int i = 0; i < 10; ++i)
            {
                Member.Member[i].No = this.MemberNoBox[i].Text;
                Member.Member[i].Name = this.MemberBox[i].Text;
                Member.Member[i].EnName = this.MemberENBox[i].Text;
            }
            // 記録計
            for (int i = 0; i < Recorders.Recorder.Count; ++i)
            {
                Recorders.Recorder[i].No  = this.RecorderNoBox[i].Text;
                Recorders.Recorder[i].Model = this.RecorderModelBox[i].Text;
                Recorders.Recorder[i].SerialNo = this.RecorderSerialNoBox[i].Text;
            }
            // システム設定
            SystemSet.Volt01 = this.VoltKind[0].Text;
            SystemSet.Volt02 = this.VoltKind[1].Text;
            SystemSet.Volt03 = this.VoltKind[2].Text;
            SystemSet.Volt04 = this.VoltKind[3].Text;
            SystemSet.Volt05 = this.VoltKind[4].Text;
            SystemSet.Volt06 = this.VoltKind[5].Text;
            SystemSet.Volt07 = this.VoltKind[6].Text;
            SystemSet.Volt08 = this.VoltKind[7].Text;

            SystemSet.ExcelPath = lblExcelFolderPath.Text;
            string CfDrv = txtReadDrive.Text;
            if(CfDrv!="D" && CfDrv != "E" && CfDrv != "F" && CfDrv != "G") 
            {
                // ＣＦカードドライブ指定間違い
                Msg += "ＣＦカードドライブ指定が誤っています。\r\n";
                Msg += "(D/E/F/G)の何れかを設定してください。。\r\n";
            }
            else SystemSet.CfDrive = txtReadDrive.Text;

            double dRecLength = ValueCheck(txtRecLength.Text);
            if (dRecLength > 1500) dRecLength = 1500;
            SystemSet.Rec_Length = (int)dRecLength;

            for (int i = 0; i < 6; ++i)
            {
                if (this.OpeTemp[i].Checked) SystemSet.OpeTempColor = i;
                if (this.OpeHumi[i].Checked) SystemSet.OpeHumiColor = i;
                if (this.CentTemp[i].Checked) SystemSet.CentTempColor = i;
                if (this.CentHumi[i].Checked) SystemSet.CentHumiColor = i;
                if (this.Humi[i].Checked) SystemSet.HumiColor = i;
                if (this.RoomTemp[i].Checked) SystemSet.RoomTempColor = i;
            }

            // 判定条件
            double ChkValue = ValueCheck(txtlimitWidth.Text);
            if (ChkValue > 20 || ChkValue < 1) 
            {
                // 温度判定リミット範囲異常
                Msg += "目標安定確認温度指定が誤っています。\r\n";
                Msg += "1.0℃ 以上 20.0℃ 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.CheckTemp = (float)ChkValue;

            ChkValue = ValueCheck(txtChgTemp.Text);
            if (ChkValue > 5 || ChkValue < 1)
            {
                // 極値変化判定温度範囲異常
                Msg += "極値変化判定温度指定が誤っています。\r\n";
                Msg += "1.0℃ 以上 5.0℃ 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.ChgJudgeTemp = (float)ChkValue;

            ChkValue = ValueCheck(txtmaxminWidth.Text);
            if (ChkValue > 1 || ChkValue < 0)
            {
                // 温度判定最大最小範囲異常
                Msg += "最大最小安定確認温度指定が誤っています。\r\n";
                Msg += "0.0℃ 以上 1.0℃ 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.MaxMinWidth = (float)ChkValue;

            ChkValue = ValueCheck(txtstabCount.Text);
            if (ChkValue > 20 || ChkValue < 4)
            {
                // 温度判定安定回数範囲異常
                Msg += "安定確認回数指定が誤っています。\r\n";
                Msg += "4 回 以上 20 回 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.StabCheckCount = (int)ChkValue;

            ChkValue = ValueCheck(txtFluctJdgeTime.Text);
            if (ChkValue > 50 || ChkValue < 2)
            {
                // 温度変動判定確認回数範囲異常
                Msg += "安定確認回数指定が誤っています。\r\n";
                Msg += "2 回 以上 50 回 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.FluctJdgTimes = (int)ChkValue;

            ChkValue = ValueCheck(txtFluctJdgeAfterTime.Text);
            if (ChkValue > 30 || ChkValue < 2)
            {
                // 温度変動判定確認回数範囲異常
                Msg += "安定確認後回数指定が誤っています。\r\n";
                Msg += "2 回 以上 30 回 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.FluctJdgAfterTimes = (int)ChkValue;

            ChkValue = ValueCheck(txtToleranceLimit.Text);
            if (ChkValue > 10 || ChkValue < 1)
            {
                // 変化許容回数範囲異常
                Msg += "変化許容回数指定が誤っています。\r\n";
                Msg += "1 回 以上 10 回 以下で設定してください。。\r\n\r\n";
            }
            else Judgment.ToleranceLimit = (int)ChkValue;

            return Msg;

        }
        private void DataSet() 
        {
            // 設定ファイルから読み込んだデータをコントロールに設定
            // 担当者
            for(int i = 0 ; i < 10; ++i ) 
            {
                this.MemberBox[i].Text = ""; this.MemberNoBox[i].Text = ""; this.MemberENBox[i].Text = "";
                if (i < Member.Member.Count) 
                { 
                    this.MemberBox[i].Text = Member.Member[i].Name;
                    this.MemberNoBox[i].Text = Member.Member[i].No;
                    this.MemberENBox[i].Text = Member.Member[i].EnName; 
                }
            }
            // 記録計
            for (int i = 0; i < 16; ++i)
            {
                if (i >= Recorders.Recorder.Count)
                {
                    this.RecorderNoBox[i].Text = "";
                    this.RecorderModelBox[i].Text = "";
                    this.RecorderSerialNoBox[i].Text = "";
                }
                else 
                {
                    this.RecorderNoBox[i].Text = Recorders.Recorder[i].No;
                    this.RecorderModelBox[i].Text = Recorders.Recorder[i].Model;
                    this.RecorderSerialNoBox[i].Text = Recorders.Recorder[i].SerialNo;
                }
            }
            // システム設定
            // 温湿度特性判定の有無
            FLUCT_USE = SystemSet.TempSpecJdge;
            FIRST_STEP = SystemSet.First_Step;
            gbHumiVar.Visible = false;
            gbTempVar.Visible = false;
            gpTempReach1.Visible = false;
            gpTempReach2.Visible = false;
            gpTempControl.Visible = false;
            gpTempFluct.Visible = false;
            gpHumiCharact.Visible = false;

            if (FLUCT_USE && FIRST_STEP)
            {
                gbHumiVar.Visible = true;
                gbTempVar.Visible = true;
            }
            else if(!FIRST_STEP)
            {
                gpTempReach1.Visible = true;
                gpTempReach2.Visible = true;
                gpTempControl.Visible = true;
                gpTempFluct.Visible = true;

                gpHumiCharact.Visible = true;
            }

            cmbTestVolt.Items.Clear();
            this.VoltKind[0].Text = SystemSet.Volt01;
            cmbTestVolt.Items.Add(SystemSet.Volt01);

            this.VoltKind[1].Text = SystemSet.Volt02;
            cmbTestVolt.Items.Add(SystemSet.Volt02);

            this.VoltKind[2].Text = SystemSet.Volt03;
            cmbTestVolt.Items.Add(SystemSet.Volt03);

            this.VoltKind[3].Text = SystemSet.Volt04;
            cmbTestVolt.Items.Add(SystemSet.Volt04);

            this.VoltKind[4].Text = SystemSet.Volt05;
            cmbTestVolt.Items.Add(SystemSet.Volt05);

            this.VoltKind[5].Text = SystemSet.Volt06;
            cmbTestVolt.Items.Add(SystemSet.Volt06);

            this.VoltKind[6].Text = SystemSet.Volt07;
            cmbTestVolt.Items.Add(SystemSet.Volt07);

            this.VoltKind[7].Text = SystemSet.Volt08;
            cmbTestVolt.Items.Add(SystemSet.Volt08);

            cmbTestVolt.SelectedIndex = 0;

            lblExcelFolderPath.Text = SystemSet.ExcelPath;
            txtReadDrive.Text = SystemSet.CfDrive;

            cmbTestFreq.Items.Clear();
            cmbTestFreq.Items.Add("50");
            cmbTestFreq.Items.Add("60");
            cmbTestFreq.SelectedIndex = 0;

            this.OpeTemp[SystemSet.OpeTempColor].Checked = true;
            this.OpeHumi[SystemSet.OpeHumiColor].Checked = true;
            this.CentTemp[SystemSet.CentTempColor].Checked = true;
            this.CentHumi[SystemSet.CentHumiColor].Checked = true;
            this.Humi[SystemSet.HumiColor].Checked = true;
            this.RoomTemp[SystemSet.RoomTempColor].Checked = true;
            // 判定条件.Json
            txtlimitWidth.Text = Judgment.CheckTemp.ToString("F1");
            txtmaxminWidth.Text = Judgment.MaxMinWidth.ToString("F1");
            txtstabCount.Text = Judgment.StabCheckCount.ToString("");

            txtFluctJdgeTime.Text = Judgment.FluctJdgTimes.ToString("");
            txtFluctJdgeAfterTime.Text = Judgment.FluctJdgAfterTimes.ToString("");
        }
        private string CheckJudge(POLELIMIT Model,String sMsg ) 
        {
            String cMsg = "";
            if (Model == null) cMsg = cMsg + sMsg + "が設定されていません。\r\n";
            else
            {
                if (Model.StartTemp == float.MinValue) cMsg = cMsg + sMsg + "の開始温度が設定されていません。\r\n";
                if (Model.EndTemp == float.MinValue) cMsg = cMsg + sMsg + "の終了温度が設定されていません。\r\n";
                if (Model.Judgment == float.MinValue) cMsg = cMsg + sMsg + "の判定温度が設定されていません。\r\n";
            }
            return cMsg;
        }
        private String CheckHumiJudge(HUMIACCURACY Model, String sMsg)
        {
            String cMsg = "";
            if (Model == null) cMsg = cMsg + sMsg + "が設定されていません。\r\n";
            else
            {
                if (Model.SetTemp == float.MinValue) cMsg = cMsg + sMsg + "の設定温度が設定されていません。\r\n";
                if (Model.SetHumi == float.MinValue) cMsg = cMsg + sMsg + "の設定湿度が設定されていません。\r\n";
                if (Model.Judgment == float.MinValue) cMsg = cMsg + sMsg + "の判定湿度が設定されていません。\r\n";
            }
            return cMsg;
        }
        private void JsonFileRead() 
        {
            String cPath = Directory.GetCurrentDirectory() + "\\Document\\";

            // システム設定.Json
            // システム設定をJsonファイルから取り込み
            WriteLog("システム設定ファイル読込み" + "\r\n");
            TextMsgWrite("システム設定ファイル読込み" + "\r\n");
            SystemSet = Json.SystemRead(cPath);
            lblExcelFolderPath.Text = SystemSet.ExcelPath;
            txtRecLength.Text = SystemSet.Rec_Length.ToString();

            MasterPath = SystemSet.MasterPath;
            ExcelValue.MasterPath = MasterPath;

            // 製品仕様.Json
            // 製品仕様をJsonファイルから取り込み
            WriteLog("製品仕様ファイル読込み" + "\r\n");
            TextMsgWrite("製品仕様ファイル読込み" + "\r\n");
            if(SystemSet.ClosedXMLUsed) Mashine = Excel_Operation.MashineSpecRead(MasterPath);
            else Mashine = Json.MashineSpecRead(MasterPath);

            // 設定情報確認（必須項目が設定されているか？nullデータが無にする。）
            String cMsg = "";
            int ModelCnt = 0;
            foreach (Json.MODELS Model in Mashine.Models)
            {
                ++ModelCnt; String cMsgNow = "";
                // キーコード確認
                if (Model.BarCode == "" || Model.Model == "" )
                {
                    cMsgNow += "バーコード・型式の何れかが設定されていません。\r\n"; 
                }

                // 温度検査プログラム項目確認
                if (Model.ProgramName == "")
                {
                    cMsgNow += "検査プログラム名が設定されていません。\r\n";
                }

                // 温度検査プログラム項目確認
                if (Model.TempSteps == null )
                {
                    cMsgNow += "温度検査プログラム項目が設定されていません。\r\n";
                }

                // 湿度設定項目確認
                if ( Model.bHumi)
                {
                    // 湿度検査プログラム項目確認
                    if (Model.HumiSteps == null)
                    {
                        cMsgNow += "湿度検査プログラム項目が設定されていません。\r\n";
                    }
                    // 湿度変動
                    if (Model.VarLimitHumi == null)
                    {
                        cMsgNow += "湿度変動設定項目が設定されていません。\r\n";
                    }
                    else if(Model.VarLimitHumi.Judgment == float.MinValue)
                    {
                        cMsgNow += "湿度変動判定基準設定項目が設定されていません。\r\n";
                    }
                    //cMsgNow += CheckHumiJudge(Model.AccuracyHumi1, "湿度精度１");
                    //cMsgNow += CheckHumiJudge(Model.AccuracyHumi2, "湿度精度２");
                    //cMsgNow += CheckHumiJudge(Model.AccuracyHumi3, "湿度精度３");
                    //cMsgNow += CheckHumiJudge(Model.AccuracyHumi4, "湿度精度４");
                }

                // 温度判定条件設定確認
                // 温度極値到達上昇設定
                cMsgNow += CheckJudge(Model.PolelLmitUp, "温度極値到達上昇");
                // 温度極値到達降下設定
                cMsgNow += CheckJudge(Model.PolelLmitDown, "温度極値到達降下");
                // 温度変化速度上昇設定
                cMsgNow += CheckJudge(Model.ChgLimitUp, "温度変化速度上昇");
                // 温度変化速度降下設定
                cMsgNow += CheckJudge(Model.ChgLimitDown, "温度変化速度降下");

                // 温度精度設定確認
                cMsgNow += CheckJudge(Model.AccuracyLowTemp, "温度精度低温");
                cMsgNow += CheckJudge(Model.AccuracyMiddleTemp, "温度精度中温");

                // 温度変動設定確認
                cMsgNow += CheckJudge(Model.VarLimitLowTemp, "温度変動低温");
                cMsgNow += CheckJudge(Model.VarLimitMiddleTemp, "温度変動中温");

                // 180℃仕様の必須項目確認
                if (Model.b180Temp) 
                {
                    // 高温精度設定
                    cMsgNow += CheckJudge(Model.AccuracyHighTemp, "温度精度高温");
                    // 高温変動設定
                    cMsgNow += CheckJudge(Model.VarLimitHighTemp, "温度変動高温");
                }

                if (cMsgNow != "")
                {
                    cMsg += "設定番号:" + ModelCnt .ToString("D3") + "\r\n" + cMsgNow;
                }

            }
            if (cMsg != "") { Console.WriteLine(cMsg); }

            // 担当者.Json
            // 担当者をJsonファイルから取り込み
            WriteLog("担当者ファイル読込み" + "\r\n");
            TextMsgWrite("担当者ファイル読込み" + "\r\n");
            Member = Json.MemberRead(MasterPath);

            // 記録計.Json
            // 記録計をJsonファイルから取り込み
            WriteLog("記録計ファイル読込み" + "\r\n");
            TextMsgWrite("記録計ファイル読込み" + "\r\n");
            Recorders = Json.RecorderRead(MasterPath);

            // データベース情報.Json
            // データベース設定をJsonファイルから取り込み
            WriteLog("データベース情報ファイル読込み" + "\r\n");
            TextMsgWrite("データベース情報ファイル読込み" + "\r\n");
            DatabaseInfo = Json.DateBaseRead(MasterPath);

            // 判定条件.Json
            // 判定条件をJsonファイルから取り込み
            WriteLog("判定条件ファイル読込み" + "\r\n");
            TextMsgWrite("判定条件ファイル読込み" + "\r\n");
            Judgment = Json.JudgmentRead(MasterPath);

            // Option.Json
            WriteLog("Optionファイル読込み" + "\r\n");
            TextMsgWrite("Optionファイル読込み" + "\r\n");
            OptionSet = Json.OptionRead(MasterPath);

            JdgCnd.LimitWidth = Judgment.CheckTemp;
            JdgCnd.ChgJudgeTemp = Judgment.ChgJudgeTemp;
            JdgCnd.MaxMinWidth = Judgment.MaxMinWidth;

            JdgCnd.StableCount = Judgment.StabCheckCount;

            JdgCnd.ChgToleranceLimit = Judgment.ToleranceLimit;

            JdgCnd.JdgTimes = Judgment.FluctJdgTimes;
            JdgCnd.JdgAfterTime = Judgment.FluctJdgAfterTimes;

        }
        private void JsonFileWrite()
        {
            String cPath = Directory.GetCurrentDirectory() + "\\Document\\";
            // 製品仕様.Json
            // 製品仕様をJsonファイルへ書込み
            //WriteLog("製品仕様ファイル書込み" + "\r\n");
            //TextMsgWrite("製品仕様ファイル書込み" + "\r\n");
            //Json.MashineSpecWrite(cPath, Mashine);
            // 担当者.Json
            // 担当者をJsonファイルへ書込み
            WriteLog("担当者ファイル書込み" + "\r\n");
            TextMsgWrite("担当者ファイル書込み" + "\r\n");
            Json.MemberWrite(cPath, Member);
            // 記録計.Json
            // 記録計をJsonファイルへ書込み
            WriteLog("記録計ファイル書込み" + "\r\n");
            TextMsgWrite("記録計ファイル書込み" + "\r\n");
            Json.RecorderWrite(cPath, Recorders);
            // システム設定.Json
            // システム設定をJsonファイルへ書込み
            WriteLog("システム設定ファイル書込み" + "\r\n");
            TextMsgWrite("システム設定ファイル書込み" + "\r\n");
            Json.SystemWrite(cPath, SystemSet);
            // 判定条件.Json
            // 判定条件をJsonファイルへ書込み
            WriteLog("判定条件ファイル書込み" + "\r\n");
            TextMsgWrite("判定条件ファイル書込み" + "\r\n");
            Json.JudgmentWrite(cPath, Judgment);
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            //二重起動をチェックする
            if (System.Diagnostics.Process.GetProcessesByName(
                System.Diagnostics.Process.GetCurrentProcess().ProcessName).Length > 1)
            {
                //すでに起動していると判断して終了
                MessageBox.Show("多重起動はできません。");
                return;
            }

            //自分自身のバージョン情報を取得する
            System.Diagnostics.FileVersionInfo ver =
                System.Diagnostics.FileVersionInfo.GetVersionInfo(
                System.Reflection.Assembly.GetExecutingAssembly().Location);
            //結果を表示
            Console.WriteLine(ver);
            AppVer = ver.FileVersion;

            String Path = Directory.GetCurrentDirectory() + "\\Document\\";
            //var path = @"C:\temp\hidden2\";

            Directory.CreateDirectory(Path);

            var directoryInfo = new DirectoryInfo(Path);
            directoryInfo.Attributes |= System.IO.FileAttributes.Hidden;
            directoryInfo.Attributes |= System.IO.FileAttributes.System;

            // コントロールの配列化
            ControlArayCreate();

            // データベースファイル読込
            // Db = new Access();
            // DbOpen();
            // DbRead();

            user = Environment.UserName;

            // Jsonファイル読込
            JsonFileRead();
            DataSet();

            if (SystemSet.OracleRead)
            {
                //接続文字列(Oracle用)
                strConnString = "User Id = " + DatabaseInfo.MashineDb.UserID
                + "; Password =" + DatabaseInfo.MashineDb.Password
                + "; Data Source = (DESCRIPTION = (ADDRESS_LIST = (ADDRESS = (PROTOCOL = TCP)(HOST = " + DatabaseInfo.MashineDb.SeverAdr
                + ")(PORT = " + DatabaseInfo.MashineDb.PortNo
                + "))) (CONNECT_DATA = (SERVICE_NAME = " + DatabaseInfo.MashineDb.DataBaseName
                + ")))";
                DBStart();

                if (!bDBOracleFlg)
                {
                    MessageBox.Show("基幹システムと接続されておりません。");
                    return;
                }
            }
            // USBシリアルデバイス検索
            string dL = SystemSet.Com_Port;
            ShowSerialPortName.ShowSerialPortName.GetDeviceNames();
            devList = ShowSerialPortName.ShowSerialPortName.deviceNameList;
            String[] ConPortList = new string[devList.Count];
            int j = 0;
            for (int i = 0; i < devList.Count; ++i) 
            {
                if (devList[i].ToString().Contains("USB シリアル デバイス (COM"))
                {
                    dL = devList[i].ToString();
                    dL = dL.Replace("USB シリアル デバイス (", "");
                    dL = dL.Replace(")", "");
                    // 検出
                    ConPortList[j] = dL;
                    ++j;
                }
            }

            for (int k = 0; k < j; ++k)
            {
                // バーコードリーダーＣＯＭポートオープン
                dL = ConPortList[k];
                BarCodeReader.PortName = dL;
                try
                {
                    BarCodeReader.Open();
                    if (BarCodeReader.IsOpen) break;
                }
                catch (Exception ex)
                {
                    string cMsg = ConPortList[k] + ":オープン確認異常 : ";
                    WriteLog(cMsg + ex.ToString() + "\r\n");
                }
            }

            // バーコードリーダーポートオープン異常
            if (!BarCodeReader.IsOpen)
            {
                string wMsg = "バーコードリーダーが接続できておりません。\r\n";
                wMsg += "製造番号を現品移動票から読み込みできませんので、\r\n";
                wMsg += "アプリケーションを終了してＵＳＢポートを確認して下さい。\r\n";
                MessageBox.Show(wMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            // 新しいCOM番号を登録
            SystemSet.Com_Port = dL;
            // 製品型式コンボボックス初期化 
            foreach (Json.MODELS Model in Mashine.Models)
            {
                Console.WriteLine(Model.Model);
                cmbModel.Items.Add(Model.ProgramName);
            }
            lblCsvFileName.Text = "";
            // 製品仕様を初期化
            int SelModel = 0;
            cmbModel.SelectedIndex = SelModel;
            ChkNowSpec.MashineNo = SelModel;
            ModelSpecSet(SelModel);
            // グラフエリア初期表示
            ViewGraphinit();
            // メンバーの名前を取り込みコンボボックスに設定する
            MemberNameGet();
            this.Top = 0; this.Left = 0;
            //自分自身のフォームを最大化
            this.WindowState = FormWindowState.Maximized;
        }
        public static void WriteLog(string cMsg)
        {
            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("utf-8");
            // 日時ファイル名生成
            DateTime dt = DateTime.Now;
            //string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\";
            //string path = "D:\\";
            path += "Log";
            if (!Directory.Exists(path))
            {
                // ディレクトリ作成
                Directory.CreateDirectory(path);
            }
            String FName = path + "\\" + dt.ToString("yyyyMMddHH") + ".Log";
            string nowstr = dt.ToString("yyyy/MM/dd,HH:mm:ss,fff,");
            int count = 0;
            while (true)
            {
                try
                {
                    // ファイルを開く(Append)
                    StreamWriter writer = new StreamWriter(@FName, true, enc);
                    // テキストを書き込む
                    writer.Write(nowstr + cMsg);
                    // ファイルを閉じる
                    writer.Close();
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    count++;
                    if (count > 10) break;
                }
            }
            //if (tCSVReadRun) tCSVMsg += cMsg + "\r\n";
            Console.WriteLine(cMsg);
        }
        private static void WriteCSVLog(string cMsg)
        {
            //if (!DebugMode) return;

            // 文字コードを指定
            Encoding enc = Encoding.GetEncoding("utf-8");
            // 日時ファイル名生成
            DateTime dt = DateTime.Now;

            //string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\";
            //string path = "D:\\";

            path += "CSV";
            if (!Directory.Exists(path))
            {
                // ディレクトリ作成
                Directory.CreateDirectory(path);
            }
            String FName = path + "\\" + NowExcelFileName + "_自動判定ログ.CSV";
            string nowstr = dt.ToString("yyyy/MM/dd,HH:mm:ss,fff,");
            int count = 0;
            while (true)
            {
                try
                {
                    // ファイルを開く(Append)
                    StreamWriter writer = new StreamWriter(@FName, true, enc);
                    // テキストを書き込む
                    writer.Write(nowstr + cMsg);
                    // ファイルを閉じる
                    writer.Close();
                    break;
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString());
                    count++;
                    if (count > 10) break;
                }
            }
        }
        private void TextMsgWrite(String cMsg) 
        {
            if (txtMessage.TextLength > 65535) txtMessage.Text = "";
            txtMessage.Text += cMsg;
            txtMessage.SelectionStart = txtMessage.TextLength;
            txtMessage.ScrollToCaret();
            txtMessage.Update();
            Console.WriteLine(cMsg);
        }
        private String PoleSet(float value) 
        {
            string sRet = "";

            if (value > 0) sRet += "+";
            sRet += value.ToString();

            return sRet;
        }
        private void ModelSpecSet(int i) 
        {
            WriteLog("製品仕様設定" + "\r\n");
            TextMsgWrite("製品仕様設定" + "\r\n");

            // 機種名表示変更
            ModelName = Mashine.Models[i].Model;
            ExcelValue.ModelName = ModelName;

            lblModel.Text = ModelName;

            //Model.ProgramName
            ChkNowSpec.InspectPrgName = Mashine.Models[i].ProgramName;
            lblInspection_program_name.Text = ChkNowSpec.InspectPrgName;

            //Model.CheckSheetName
            ChkNowSpec.CheckSheetName = Mashine.Models[i].CheckSheetName;
            lblCheckSheet_name.Text = ChkNowSpec.CheckSheetName;

            // 180℃仕様設定
            if (Mashine.Models[i].b180Temp) { ChkNowSpec.Temp180Use = true; gB180.Visible = true; lbl180Spec.Text = "180℃仕様有"; }
            else { ChkNowSpec.Temp180Use = false; gB180.Visible = false; lbl180Spec.Text = "180℃仕様無"; }

            // 観測窓仕様設定
            if (Mashine.Models[i].bChkWindow) { ChkNowSpec.WindowUse = true; lblObservationWindow.Text = "観測窓有"; gpTempReach2.Visible = true; }
            else { ChkNowSpec.WindowUse = false; lblObservationWindow.Text = "観測窓無"; gpTempReach2.Visible = false; }

            // 内扉仕様設定
            if (Mashine.Models[i].bInWindow) { ChkNowSpec.IndoorUse = true; lblInnerDoor.Text = "内扉有"; }
            else { ChkNowSpec.IndoorUse = false; lblInnerDoor.Text = "内扉無"; }

            // 電圧・周波数基幹システムより設定
            // 基幹システムデータベースアクセス
            if (SystemSet.OracleRead && bDBOracleFlg)
            {
                // 電圧検索更新
                for (int iNo = 0; iNo < cmbTestVolt.Items.Count; iNo++)
                {
                    if (HeaderData.VoltSpec.ToString() == cmbTestVolt.Items[iNo].ToString())
                    {
                        cmbTestVolt.SelectedIndex = iNo;
                        break;
                    }
                }
                // 周波数検索更新
                for (int iNo = 0; iNo < cmbTestFreq.Items.Count; iNo++)
                {
                    if (HeaderData.FreqSpec.ToString() == cmbTestFreq.Items[iNo].ToString())
                    {
                        cmbTestFreq.SelectedIndex = iNo;
                        break;
                    }
                }
            }

            // 製品最大最小温度設定
            ChkNowSpec.SpecMaxTemp = Mashine.Models[i].HighTemp;
            ChkNowSpec.SpecMinTemp = Mashine.Models[i].LowTemp;

            // 性能確認エリア初期化
            TempFluct.Clear();
            HumiFluct.Clear();

            MaxTemp = new WIDTH_FLUCT();
            MinTemp = new WIDTH_FLUCT();

            // 湿度タイプ設定
            if (Mashine.Models[i].bHumi) ChkNowSpec.HumiUse = true;
            else ChkNowSpec.HumiUse = false;

            if(!FIRST_STEP) gpHumiCharact.Visible = false;

            if (ChkNowSpec.HumiUse) {
                if (!FIRST_STEP) gpHumiCharact.Visible = true;
                // 温湿度設定
                Humiset.Clear();
                for (int j = 0; j < HUMISETSTEPS; ++j)
                {
                    //WIDTH_FLUCT TempCheckSet = new WIDTH_FLUCT();
                    WIDTH_FLUCT HumiCheckSet = new WIDTH_FLUCT();
                    HUMI_SET HSet = new HUMI_SET();
                    HSet.SetHumi = Mashine.Models[i].HumiSteps[j].SetHumi;
                    HSet.SetTemp = Mashine.Models[i].HumiSteps[j].SetTemp;
                    HSet.HumiAccuracy = Mashine.Models[i].HumiSteps[j].HumiAccuracy;
                    HSet.SetTime = Mashine.Models[i].HumiSteps[j].SetTime;
                    Humiset.Add(HSet);
                    HumiCheckSet.SetValue = HSet.SetHumi;
                    //TempCheckSet.SetValue = HSet.SetTemp;
                    HumiFluct.Add(HumiCheckSet);
                    //TempFluct.Add(TempCheckSet);
                }
            }
            // 温度設定
            float fMaxTemp = float.MinValue;
            float fMinTemp = float.MaxValue;

            // 20℃温度安定検出フラグクリア(使用時は、false 未使用時は、true)
            FLUCT20_Exist = false;

            Tempset.Clear();
            for (int j = 0; j < TEMPSETSTEPS; ++j)
            {
                WIDTH_FLUCT TempCheckSet = new WIDTH_FLUCT();
                TEMP_SET TSet = new TEMP_SET();
                TSet.SetTemp = Mashine.Models[i].TempSteps[j].SetTemp;
                TSet.SetTime = Mashine.Models[i].TempSteps[j].SetTime;
                Tempset.Add(TSet);
                if (fMaxTemp < TSet.SetTemp) fMaxTemp = TSet.SetTemp;
                if (fMinTemp > TSet.SetTemp) fMinTemp = TSet.SetTemp;
                TempCheckSet.SetValue = TSet.SetTemp;
                if (TSet.SetTemp != 157 && TSet.SetTemp != 185 &&
                   TSet.SetTemp != -23 && TSet.SetTemp != -43 && TSet.SetTemp != -63) TempFluct.Add(TempCheckSet);
            }

            // 製品最大上最小下温度設定
            ChkNowSpec.SpecMaxUpTemp = fMaxTemp;
            ChkNowSpec.SpecMinDownTemp = fMinTemp;

            // 温度極値上昇
            ChkNowSpec.PoleUpSpec = (int)Mashine.Models[i].PolelLmitUp.Judgment;
            lblPoleUpSpec.Text = ChkNowSpec.PoleUpSpec.ToString();

            ChkNowSpec.PoleUpStartTemp = Mashine.Models[i].PolelLmitUp.StartTemp;
            lblPoleUpStart.Text = PoleSet(ChkNowSpec.PoleUpStartTemp);

            ChkNowSpec.PoleUpEndTemp = Mashine.Models[i].PolelLmitUp.EndTemp;
            lblPoleUpEnd.Text = PoleSet(ChkNowSpec.PoleUpEndTemp);

            lblPoleUpValue.Text = "";
            lblPoleUpStr.Text = "";
            lblPoleUpStr.BackColor = Color.LightGray;

            // 温度極値降下
            ChkNowSpec.PoleDownSpec = (int)Mashine.Models[i].PolelLmitDown.Judgment;
            lblPoleDownSpec.Text = ChkNowSpec.PoleDownSpec.ToString();

            ChkNowSpec.PoleDownStartTemp = Mashine.Models[i].PolelLmitDown.StartTemp;
            lblPoleDownStart.Text = PoleSet(ChkNowSpec.PoleDownStartTemp);

            ChkNowSpec.PoleDownEndTemp = Mashine.Models[i].PolelLmitDown.EndTemp;
            lblPoleDownEnd.Text = PoleSet(ChkNowSpec.PoleDownEndTemp);

            lblPoleDownValue.Text = "";
            lblPoleDownStr.Text = "";
            lblPoleDownStr.BackColor = Color.LightGray;

            // 温度精度低温
            ChkNowSpec.AccuracyTemp1Spec = Mashine.Models[i].AccuracyLowTemp.Judgment;
            ChkNowSpec.AccuracyTemp1StartTemp = Mashine.Models[i].AccuracyLowTemp.StartTemp;
            ChkNowSpec.AccuracyTemp1EndTemp = Mashine.Models[i].AccuracyLowTemp.EndTemp;

            // 温度精度中温
            ChkNowSpec.AccuracyTemp2Spec = Mashine.Models[i].AccuracyMiddleTemp.Judgment;
            ChkNowSpec.AccuracyTemp2StartTemp = Mashine.Models[i].AccuracyMiddleTemp.StartTemp;
            ChkNowSpec.AccuracyTemp2EndTemp = Mashine.Models[i].AccuracyMiddleTemp.EndTemp;

            // 温度精度高温
            if (ChkNowSpec.Temp180Use) 
            {
                ChkNowSpec.AccuracyTemp3Spec = Mashine.Models[i].AccuracyHighTemp.Judgment;
                ChkNowSpec.AccuracyTemp3StartTemp = Mashine.Models[i].AccuracyHighTemp.StartTemp;
                ChkNowSpec.AccuracyTemp3EndTemp = Mashine.Models[i].AccuracyHighTemp.EndTemp;
            }

            // 温度変化率上昇
            ChkNowSpec.ChgUpSpec = Mashine.Models[i].ChgLimitUp.Judgment;
            lblChgUpSpec.Text = ChkNowSpec.ChgUpSpec.ToString("F1");

            ChkNowSpec.ChgUpStartTemp = Mashine.Models[i].ChgLimitUp.StartTemp;
            lblChgUpStart.Text = PoleSet(ChkNowSpec.ChgUpStartTemp);

            ChkNowSpec.ChgUpEndTemp = Mashine.Models[i].ChgLimitUp.EndTemp;
            lblChgUpEnd.Text = PoleSet(ChkNowSpec.ChgUpEndTemp);

            lblChgUpValue.Text = "";
            lblChgUpTime.Text = "";
            lblChgUpStr.Text = "";
            lblChgUpStr.BackColor = Color.LightGray;

            // 温度変化率降下
            ChkNowSpec.ChgDownSpec = Mashine.Models[i].ChgLimitDown.Judgment;
            lblChgDownSpec.Text = ChkNowSpec.ChgDownSpec.ToString("F1");

            ChkNowSpec.ChgDownStartTemp = Mashine.Models[i].ChgLimitDown.StartTemp;
            lblChgDownStart.Text = PoleSet(ChkNowSpec.ChgDownStartTemp);

            ChkNowSpec.ChgDownEndTemp = Mashine.Models[i].ChgLimitDown.EndTemp;
            lblChgDownEnd.Text = Mashine.Models[i].ChgLimitDown.EndTemp.ToString();

            lblChgDownValue.Text = "";
            lblChgDownTime.Text = "";
            lblChgDownStr.Text = "";
            lblChgDownStr.BackColor = Color.LightGray;

            if (FLUCT_USE)
            //    if (FLUCT_USE && FIRST_STEP)
            {
                // 温度変動１
                ChkNowSpec.VarTemp1Spec = Mashine.Models[i].VarLimitLowTemp.Judgment;
                lblVarTemp1Spec.Text = ChkNowSpec.VarTemp1Spec.ToString();

                ChkNowSpec.VarTemp1StartTemp = Mashine.Models[i].VarLimitLowTemp.StartTemp;
                lblVarTemp1Start.Text = ChkNowSpec.VarTemp1StartTemp.ToString();

                ChkNowSpec.VarTemp1EndTemp = Mashine.Models[i].VarLimitLowTemp.EndTemp;
                lblVarTemp1End.Text = ChkNowSpec.VarTemp1EndTemp.ToString();

                lblVarTemp1Str.Text = "";
                lblVarTemp1Str.BackColor = Color.LightGray;

                // 温度変動２
                ChkNowSpec.VarTemp2Spec = Mashine.Models[i].VarLimitMiddleTemp.Judgment;
                lblVarTemp2Spec.Text = ChkNowSpec.VarTemp2Spec.ToString();

                ChkNowSpec.VarTemp2StartTemp = Mashine.Models[i].VarLimitMiddleTemp.StartTemp;
                lblVarTemp2Start.Text = ChkNowSpec.VarTemp2StartTemp.ToString();

                ChkNowSpec.VarTemp2EndTemp = Mashine.Models[i].VarLimitMiddleTemp.EndTemp;
                lblVarTemp2End.Text = ChkNowSpec.VarTemp2EndTemp.ToString();

                lblVarTemp2Str.Text = "";
                lblVarTemp2Str.BackColor = Color.LightGray;

                // 温度変動３
                if (ChkNowSpec.Temp180Use)
                {

                    ChkNowSpec.VarTemp3Spec = Mashine.Models[i].VarLimitHighTemp.Judgment;
                    lblVarTemp3Spec.Text = ChkNowSpec.VarTemp3Spec.ToString();

                    ChkNowSpec.VarTemp3StartTemp = Mashine.Models[i].VarLimitHighTemp.StartTemp;
                    lblVarTemp3Start.Text = ChkNowSpec.VarTemp3StartTemp.ToString();

                    ChkNowSpec.VarTemp3EndTemp = Mashine.Models[i].VarLimitHighTemp.EndTemp;
                    lblVarTemp3End.Text = ChkNowSpec.VarTemp3EndTemp.ToString();

                    lblVarTemp3Str.Text = "";
                    lblVarTemp3Str.BackColor = Color.LightGray;
                }
                if (Mashine.Models[i].bHumi) gbHumiVar.Visible = true;
                else gbHumiVar.Visible = false;

                if (ChkNowSpec.HumiUse)
                {
                    // 湿度変動
                    ChkNowSpec.VarHumiSpec = Mashine.Models[i].VarLimitHumi.Judgment;
                    lblVarHumiSpec.Text = ChkNowSpec.VarHumiSpec.ToString();
                    lblVarHumiStr.Text = "";
                    lblVarHumiStr.BackColor = Color.LightGray;

                }
            }

            // 温度到達（制御温度値判定）最低温度
            //lblTempReachSetL.Text = PoleSet((Mashine.Models[i].LowTemp - 3 ));
            lblTempReachSetL.Text = PoleSet(ChkNowSpec.SpecMinDownTemp);
            lblTempReachSpecL.Text = Mashine.Models[i].AccuracyLowTemp.Judgment.ToString();
            lblTempReachMemL.Text = "";
            lblTempReachJdgeL.Text = "";
            lblTempReachJdgeL.BackColor = Color.LightGray;

            // 温度到達（制御温度値判定）最高温度
            //lblTempReachSetH.Text = PoleSet((Mashine.Models[i].HighTemp + 3));
            lblTempReachSetH.Text = PoleSet(ChkNowSpec.SpecMaxUpTemp);
            float fValue = Mashine.Models[i].AccuracyMiddleTemp.Judgment;
            if (Mashine.Models[i].b180Temp) fValue = Mashine.Models[i].AccuracyHighTemp.Judgment;
            lblTempReachSpecH.Text = fValue.ToString();
            lblTempReachMemH.Text = "";
            lblTempReachJdgeH.Text = "";
            lblTempReachJdgeH.BackColor = Color.LightGray;

            // 温度到達（槽中央乾球温度値判定）仕様最低温度
            lblTempReachSetM.Text = PoleSet(Mashine.Models[i].LowTemp);
            lblTempReachSpecM.Text = Mashine.Models[i].AccuracyLowTemp.Judgment.ToString();
            lblTempReachMemM.Text = "";
            lblTempReachJdgeM.Text = "";
            lblTempReachJdgeM.BackColor = Color.LightGray;

            // 温度制御（制御温度値判定）最高温度
            lblTempOpeSetH.Text = PoleSet(Mashine.Models[i].HighTemp);
            lblTempOpeSpecH.Text = "±1.2";
            lblTempOpeMemH.Text = "";
            lblTempOpeJdgeH.Text = "";
            lblTempOpeJdgeH.BackColor = Color.LightGray;

            // 温度制御（制御温度値判定）中間温度
            lblTempOpeSetM.Text = "+20";
            lblTempOpeSpecM.Text = "±1.2";
            lblTempOpeMemM.Text = "";
            lblTempOpeJdgeM.Text = "";
            lblTempOpeJdgeM.BackColor = Color.LightGray;

            // 温度制御（制御温度値判定）最低温度
            lblTempOpeSetL.Text = PoleSet(Mashine.Models[i].LowTemp);
            lblTempOpeSpecL.Text = "±1.2";
            lblTempOpeMemL.Text = "";
            lblTempOpeJdgeL.Text = "";
            lblTempOpeJdgeL.BackColor = Color.LightGray;

            // 温度変動（槽中央温度値判定）最高温度
            lblTempFluctSetH.Text = PoleSet(Mashine.Models[i].HighTemp);
            lblTempFluctSpecH.Text = "±0.5";
            lblTempFluctMemH.Text = "";
            lblTempFluctJdgeH.Text = "";
            lblTempFluctJdgeH.BackColor = Color.LightGray;

            // 温度変動（槽中央温度値判定）中間温度
            lblTempFluctSetM.Text = "+20";
            lblTempFluctSpecM.Text = "±0.3";
            lblTempFluctMemM.Text = "";
            lblTempFluctJdgeM.Text = "";
            lblTempFluctJdgeM.BackColor = Color.LightGray;

            // 温度変動（槽中央温度値判定）最低温度
            lblTempFluctSetL.Text = PoleSet(Mashine.Models[i].LowTemp);
            lblTempFluctSpecL.Text = "±0.3";
            lblTempFluctMemL.Text = "";
            lblTempFluctJdgeL.Text = "";
            lblTempFluctJdgeL.BackColor = Color.LightGray;


            // 温湿度制御・変動判定

            for (int j= 0; j < 4; ++j)
            {
                HumiOpeSet[j].Text = "+" + Humiset[j].SetTemp.ToString() + "℃/" + Humiset[j].SetHumi.ToString() + "%RH";
                HumiOpeTemp[j].Text = "";
                HumiOpeHumi[j].Text = "";
                HumiOpeJdge[j].Text = "";
                HumiOpeJdge[j].BackColor = Color.LightGray;
                HumiOpeMonTemp[j].Text = "";
                HumiOpeMonHumi[j].Text = "";
                HumiOpeHumiPV[j].Text = "";
                HumiFluctStd[j].Text = "";
                HumiFluctJdge[j].Text = "";
                HumiFluctJdge[j].BackColor = Color.LightGray;
            }
            // 電流情報登録
            CurrentInfo = new List<CURRENT>();
            foreach (VOLTCURRENTS cur in Mashine.Models[i].Volts)
            {
                if(cur.Volt == HeaderData.VoltSpec.ToString() + "V") 
                {
                    // 電圧仕様の電流値情報を取得
                    foreach(CURRENT VoltCurrent in cur.VoltCurrents.Currents) 
                    {
                        CurrentInfo.Add(VoltCurrent);
                    }
                    break;
                }
            }
            // 自動判定条件
            //txtlimitWidth.Text= JdgCnd.LimitWidth.ToString("F1");
            txtlimitWidth.Text = Judgment.CheckTemp.ToString("F1");

            txtChgTemp.Text = JdgCnd.ChgJudgeTemp.ToString("F1");

            txtmaxminWidth.Text = JdgCnd.MaxMinWidth.ToString("F1");

            int cnt = (int)JdgCnd.StableCount;
            txtstabCount.Text = cnt.ToString("");

            txtEndRecord.Text = ChkDataMaxlength.ToString("");

            SetExcellFileName();

            ViewGraphinit();
        }
        private void ViewGraphinit()
        {
            WriteLog("グラフ描画初期化" + "\r\n");
            TextMsgWrite("グラフ描画初期化" + "\r\n");
            ChkLength = ChkDataMaxlength;
            // グラフ更新
            formsPlot1.Plot.Clear();
            String msg = "温度";
            if (ChkNowSpec.HumiUse) msg = "温湿度";
            else msg = "温度";
            // グラフにタイトルを付けます
            formsPlot1.Plot.Title("<    " + ModelName + "               " + msg + "確認試験               No.   " + SerialNo + "     >");

            // X軸, Y軸に名前を付けます
            // fontNameでフォントを指定することができます。
            formsPlot1.Plot.XLabel("経過時間(分)");
            formsPlot1.Plot.YLabel("温度(℃)");

            formsPlot1.Plot.YAxis2.Label("");
            formsPlot1.Plot.YAxis2.Dims.SetAxis(HUMIMIN, HUMIMAX);
            // 目盛り線の表示を設定します(trueにしないと表示されません)
            if (ChkNowSpec.HumiUse) { formsPlot1.Plot.YAxis2.Ticks(true); formsPlot1.Plot.YAxis2.Label("湿度(%RH)"); }
            else formsPlot1.Plot.YAxis2.Ticks(false);

            // Ｘ軸の軸範囲設定
            formsPlot1.Plot.XAxis.Dims.SetAxis(0, ChkLength);
            // Ｙ軸の軸範囲設定
            formsPlot1.Plot.YAxis.Dims.SetAxis(TEMPMIN, TEMPMAX);

            // Legend(凡例)の表示を指定します(指定しないと表示されません)
            formsPlot1.Plot.Legend();
            // ScottPlotのコントロールに描画(表示)します。
            formsPlot1.Render();
            formsPlot1.Visible = true;
        }
        private void ScatSet(ref ScatterPlot scat, double[] ValueX, double[] ValueY, Color DrawColor, int LineWidth, LineStyle dashDot, String Label, bool Second)
        {
            scat = formsPlot1.Plot.AddScatterLines(ValueX, ValueY, DrawColor, LineWidth, dashDot);
            scat.Label = Label;
            scat.OnNaN = ScatterPlot.NanBehavior.Gap;
            if (Second) scat.YAxisIndex = 1;
        }
        private void BracketSet(Bracket Bracket,double SPos,float SValue, double EPos, float EValue ,String Msg,
            float FontSize,bool LabelCounterClockwise,float EdgeLength,int YAxisIndex) 
        {
            Bracket = formsPlot1.Plot.AddBracket(SPos, SValue, EPos, EValue, Msg);
            Bracket.Font.Size = FontSize;
            Bracket.LabelCounterClockwise = LabelCounterClockwise;
            Bracket.EdgeLength = EdgeLength;
            Bracket.YAxisIndex = YAxisIndex;
        }
        private void BracketFreeSet(ScatterPlot[] bracketFree, int iIndex, double SPos, float SValue, double EPos, float EValue)
        {
            // 開始ライン
            double[] timed = new double[2];
            double[] plotd = new double[2];
            timed[0] = SPos; timed[1] = SPos; plotd[0] = SValue; plotd[1] = vMin + 35;
            ScatSet(ref bracketFree[iIndex], timed, plotd, Color.Black, 1, LineStyle.DashDotDot, "", false);
            // 終了ライン
            timed = new double[2];
            plotd = new double[2];
            timed[0] = EPos; timed[1] = EPos; plotd[0] = vMin + 35; plotd[1] = EValue;
            ScatSet(ref bracketFree[iIndex + 1], timed, plotd, Color.Black, 1, LineStyle.DashDotDot, "", false);

        }
        private void BracketJudge(ref Bracket BR, ref ChgRate CR, ref double[] timed, String cMsg)
        {
            if (CR.EndExist)
            {
                if (SystemSet.GraphRequired)
                {
                    BracketFreeSet(bracketFree, 0, timed[CR.StartXPos], CR.StartYValue, timed[CR.EndXPos], CR.EndYValue);
                    if (cMsg.Contains("極値"))
                    {
                        BracketSet(BR, timed[CR.StartXPos], vMin + 35, timed[CR.EndXPos], vMin + 35, cMsg, BRACKET_CSIZE, false, BRACKET_LSIZE, 0);
                    }
                    else 
                    {
                        BracketSet(BR, timed[CR.StartXPos], vMin + 20, timed[CR.EndXPos], vMin + 20, cMsg, BRACKET_CSIZE, false, BRACKET_LSIZE, 0);
                    }
                }
                else
                {
                    BracketSet(BR, timed[CR.StartXPos], CR.StartYValue, timed[CR.EndXPos], CR.EndYValue, cMsg, BRACKET_CSIZE, false, BRACKET_LSIZE, 0);
                }
            }
        }
        private void ViewGraph()
        {
            WriteLog("グラフ描画更新" + "\r\n");
            TextMsgWrite("グラフ描画更新" + "\r\n");
            ChkLength = ChkDataMaxlength;
            try
            {
                RecLength = ChkLength;
                #region// データの切り出し
                double[] timed = new double[RecLength];
                double[] TempSV = new double[RecLength];
                double[] HumiSV = new double[RecLength];
                double[] TempPV = new double[RecLength];
                double[] WetTempPV = new double[RecLength];
                double[] RoomTemp = new double[RecLength];
                double[] HumiPV = new double[RecLength];
                Array.Copy(RecValue.Times, 0, timed, 0, RecLength);
                Array.Copy(RecValue.TempSV, 0, TempSV, 0, RecLength);
                Array.Copy(RecValue.HumiSV, 0, HumiSV, 0, RecLength);
                Array.Copy(RecValue.TempPV, 0, TempPV, 0, RecLength);
                Array.Copy(RecValue.WetTempPV, 0, WetTempPV, 0, RecLength);
                Array.Copy(RecValue.RoomTemp, 0, RoomTemp, 0, RecLength);
                Array.Copy(RecValue.HumiPV, 0, HumiPV, 0, RecLength);
                #endregion

                // 結果表示
                // グラフ更新
                // グラフにタイトルを付けます
                //if(!ChkNowSpec.HumiUse) formsPlot1.Plot.Title("小型環境試験器　標準検査データ（温度制御）");
                //else formsPlot1.Plot.Title("小型環境試験器　標準検査データ（湿度制御）");
                // X軸, Y軸に名前を付けます
                formsPlot1.Plot.XLabel("経過時間(分)");
                formsPlot1.Plot.YLabel("温度(℃)");
                // Ｘ軸の軸範囲設定
                formsPlot1.Plot.XAxis.Dims.SetAxis(0, ChkLength);
                string cMsg = "";
                if (SystemSet.GraphRequired)
                {
                    int yValue = vMin + 75;
                    // ヘッダー表示
                    cMsg = "試験実施日\t:" + TestingDate + "\r\n";
                    cMsg += "周囲温度\t:" + "+" + RoomTempAverage.ToString("F1") + " ℃\r\n";
                    cMsg += "試験電圧\t:" + "AC" + HeaderData.VoltSpec.ToString("") + "V / " + HeaderData.FreqSpec.ToString("") + "Hz\r\n";
                    cMsg += "記録計品番\t:" + TestingRecorderNo + "\r\n";
                    cMsg += "記録計型式\t:" + TestingRecorderModel + "\r\n";
                    cMsg += "製造番号\t:" + TestingRecorderSerialNo + "\r\n";
                    var txt = formsPlot1.Plot.AddText(cMsg, 3, yValue);
                    txt.Font.Color = Color.Black; ;
                    txt.BackgroundColor = Color.White;
                    txt.BackgroundFill = true;
                    txt.Rotation = 0;
                    txt.BorderSize = 1;
                    txt.BorderColor = Color.Black;
                    txt.DragEnabled = false;
                }

                var axisLimits = formsPlot1.Plot.GetAxisLimits();

                // プロットデータ準備
                //ScatterPlot[] scat = new ScatterPlot[6];
                ScatSet(ref scat[0], timed, TempSV, Color.Brown, 1, LineStyle.Solid, "制御温度",false);
                ScatSet(ref scat[2], timed, TempPV, Color.Blue, 1, LineStyle.Solid, "中央乾球温度",false);
                ScatSet(ref scat[4], timed, RoomTemp, Color.Yellow, 1, LineStyle.Solid, "周囲温度",false);
                if (ChkNowSpec.HumiUse)
                {
                    ScatSet(ref scat[1], timed, HumiSV, Color.Green, 1, LineStyle.Solid, "制御湿度", false);
                    ScatSet(ref scat[3], timed, WetTempPV, Color.Red, 1, LineStyle.Solid, "中央湿球温度", false);
                    ScatSet(ref scat[5], timed, HumiPV, Color.Cyan, 1, LineStyle.Solid, "相対湿度", true);
                }
                // 温度確認のブラケット表示
                Bracket[] TempBracket = new Bracket[4];

                cMsg = "温度変化速度(降下):\r\n" + ChgDownValue.ResultValue.ToString("F1") + "℃/分(" + (ChgDownValue.EndXPos-ChgDownValue.StartXPos).ToString() + "分)";
                BracketJudge(ref TempBracket[0], ref ChgDownValue, ref timed, cMsg);

                cMsg = "極値到達降下時間:\r\n" + ExtDownValue.ResultValue.ToString() + "分";
                BracketJudge(ref TempBracket[1], ref ExtDownValue, ref timed, cMsg);

                cMsg = "極値到達上昇時間:\r\n" + ExtUpValue.ResultValue.ToString() + "分";
                BracketJudge(ref TempBracket[2], ref ExtUpValue, ref timed, cMsg);

                cMsg = "温度変化速度(上昇):\r\n" + ChgUpValue.ResultValue.ToString("F1") + "℃/分(" + (ChgUpValue.EndXPos - ChgUpValue.StartXPos).ToString() + "分)";
                BracketJudge(ref TempBracket[3], ref ChgUpValue, ref timed, cMsg);

                if ((FLUCT_USE && FIRST_STEP) || !FIRST_STEP)
                {
                    // 温度幅確認のブラケット表示
                    if (TempFluct.Count != 0)
                    {
                        Bracket[] TempWidthBracket = new Bracket[TempFluct.Count];
                        for (int i = 0; i < TempFluct.Count; i++)
                        {
                            if (TempFluct[i].StartPosition != 0)
                            {
                                if (TempFluct[i].Exist) BracketSet(TempWidthBracket[i], TempFluct[i].StartPosition, TempFluct[i].StartValue, TempFluct[i].EndPosition, TempFluct[i].StartValue,
                                    "温度幅:" + TempFluct[i].TempPvStd.ToString("F1") + "℃", BRACKET_CSIZE, true, BRACKET_LSIZE, 0);
                            }
                        }
                    }

                    // 湿度幅確認のブラケット表示
                    if (HumiFluct.Count != 0 && ChkNowSpec.HumiUse)
                    {
                        Bracket[] HumiWidthBracket = new Bracket[HumiFluct.Count];
                        for (int i = 0; i < HumiFluct.Count; i++)
                        {
                            bool bCow = false;
                            if (HumiFluct[i].StartPosition != 0)
                            {
                                if (i % 2 == 1) bCow = true;
                                if (HumiFluct[i].Exist) BracketSet(HumiWidthBracket[i], HumiFluct[i].StartPosition, HumiFluct[i].StartValue, HumiFluct[i].EndPosition, HumiFluct[i].StartValue,
                                    "湿度幅:" + HumiFluct[i].HumiPvStd.ToString("F1") + "%RH", BRACKET_CSIZE, bCow, BRACKET_LSIZE, 1);
                            }
                        }
                    }

                    // 最小温度確認のブラケット表示
                    Bracket[] MaxMinBracket = new Bracket[2];
                    if(MinTemp.Exist) BracketSet(MaxMinBracket[0], MinTemp.StartPosition, MinTemp.StartValue, MinTemp.EndPosition, MinTemp.StartValue,
                        "最低温度:" + MinTemp.TempSvAve.ToString("F1") + "℃", BRACKET_CSIZE, false, BRACKET_LSIZE, 0);

                    // 最大温度確認のブラケット表示
                    if (MaxTemp.Exist) BracketSet(MaxMinBracket[1], MaxTemp.StartPosition, MaxTemp.StartValue, MaxTemp.EndPosition, MaxTemp.StartValue,
                        "最高温度:" + MaxTemp.TempSvAve.ToString("F1") + "℃", BRACKET_CSIZE, true, BRACKET_LSIZE, 0);


                }
                // 温湿度ステップ書き込み
                if (SystemSet.GraphRequired)
                {
                    //int iNo = ProgStep.Count;
                    int iNo = 4;bool bFirst = true; 
                    string cMsg3 = "";
                    Bracket[] bracketStep = new Bracket[iNo];
                    double sPos = 0; double nPos = 0; string cMsg2 = "";
                    if (ChkNowSpec.HumiUse)
                    {
                        iNo = Mashine.Models[ChkNowSpec.MashineNo].HumiSteps.Count;
                        bracketStep = new Bracket[iNo];
                        for (int i = 0; i < iNo; ++i)
                        {
                            if(bFirst && i == 0 && false) 
                            {
                                sPos = ((double)iFirstChgRecord / 60) - (double)Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetTime;
                                bFirst = false;
                            }
                            cMsg3 = "";
                            if (Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetTemp > 0) cMsg3 = "+";
                            cMsg2 = cMsg3 + Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetTemp.ToString("F0") + "℃\r\n";
                            float fValue = Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetHumi;
                            if (ChkNowSpec.HumiUse && fValue != 0) cMsg2 += fValue.ToString("F0") + "%RH\r\n";
                            cMsg2 += Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetTime.ToString("F1") + "Hr";
                            nPos = sPos + Mashine.Models[ChkNowSpec.MashineNo].HumiSteps[i].SetTime;
                            BracketSet(bracketStep[i], (sPos + 0.1) * 60, 50, (nPos - 0.1) * 60, 50,
                            cMsg2, BRACKET_CSIZE, true, BRACKET_LSIZE, 0);
                            sPos = nPos;
                        }
                    }
                    iNo = Mashine.Models[ChkNowSpec.MashineNo].TempSteps.Count;
                    //int iNo = ProgStep.Count;
                    bracketStep = new Bracket[iNo];
                    nPos = 0; cMsg2 = "";
                    for (int i = 0; i < iNo; ++i)
                    {
                        if (bFirst && i == 0 && false)
                        {
                            sPos = ((double)iFirstChgRecord / 60) - (double)Mashine.Models[ChkNowSpec.MashineNo].TempSteps[i].SetTime;
                            bFirst = false;
                        }
                        cMsg3 = "";
                        if (Mashine.Models[ChkNowSpec.MashineNo].TempSteps[i].SetTemp > 0) cMsg3 = "+";
                        cMsg2 = cMsg3 + Mashine.Models[ChkNowSpec.MashineNo].TempSteps[i].SetTemp.ToString("F0") + "℃\r\n";
                        cMsg2 += Mashine.Models[ChkNowSpec.MashineNo].TempSteps[i].SetTime.ToString("F1") + "Hr";
                        nPos = sPos + Mashine.Models[ChkNowSpec.MashineNo].TempSteps[i].SetTime;
                        BracketSet(bracketStep[i], (sPos + 0.1) * 60, 50, (nPos - 0.1) * 60, 50,
                        cMsg2, BRACKET_CSIZE, true, BRACKET_LSIZE, 0);
                        sPos = nPos;
                    }
                }
                // Y軸０にLineを引く
                formsPlot1.Plot.AddHorizontalLine(0);
                // Legend(凡例)の表示を指定します(指定しないと表示されません)
                formsPlot1.Plot.Legend();

                // 凡例の位置を左下に設定
                formsPlot1.Plot.Legend(location: ScottPlot.Alignment.LowerLeft);

                formsPlot1.Plot.SetAxisLimits(axisLimits);

                // ScottPlotのコントロールに描画(表示)します。
                formsPlot1.Render();
                formsPlot1.Visible = true;

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
        }
        private static double ValueCheck( string StrValue) 
        {
            double retValue = double.NaN;
            bool isNumber = double.TryParse(StrValue, out retValue);
            if(!isNumber) retValue = double.NaN;
            return retValue;
        }
        private static string GetHumiPV(double DryTemp, double WetTemp, double WindSpeed) 
        {
            string ReturnValue = "error";
            double T1 = 0; double T2 = 0; double esd = 0; double esw = 0;
            double e = double.MinValue; double ans = 0;

            if(DryTemp <= 100 && DryTemp >= 0 &&
                WetTemp <=100 && WetTemp >= 0) 
            {
                if (WetTemp > DryTemp) WetTemp = DryTemp;
                // 乾球温度DryTempでの飽和水蒸気分圧：esd
                T1 = DryTemp + 273.15;
                // ゾンタークの式
                esd = Math.Exp(S1 *  Math.Pow(T1 , -1) + S2 + S3 * T1 + S4 * Math.Pow(T1 , 2) + S5 * Math.Log(T1));
                // 湿球温度WetTempでの飽和水蒸気分圧：esw
                T2 = WetTemp + 273.15;
                esw = Math.Exp(S1 * Math.Pow(T2, -1) + S2 + S3 * T2 + S4 * Math.Pow(T2, 2) + S5 * Math.Log(T2));
                if(WindSpeed > 0 && WindSpeed <= 0.5) 
                {
                    // ペルンターの式（風速：静穏）
                    e = esw - 0.0012 * 101325 * (DryTemp - WetTemp) * (1 + WetTemp / 610);
                    
                }
                else if (WindSpeed > 0.5 && WindSpeed <= 1.5) 
                {
                    // ペルンターの式（風速：弱風）
                    e = esw - 0.0008 * 101325 * (DryTemp - WetTemp) * (1 + WetTemp / 610);
                }
                else if (WindSpeed > 1.5) 
                {
                    // スプルングの式
                    e = esw - 0.000662 * 101325 * (DryTemp - WetTemp);
                }
                if (e != double.MinValue) 
                {
                    double v = e / esd * 100;
                    ReturnValue = v.ToString(); 
                }
            }
            return ReturnValue;
        }
        private bool PoleChgTempChk(double[] ChkV , float StartTemp,int cnt)
        {
            bool Ret = false; double[] Chk3V = new double[cnt-1];
            if (StartTemp != float.MinValue)
            {
                // 目標温度近傍温度確認
                for (int i = 0; i < cnt - 1; i++)
                {
                    if (ChkV[i] > StartTemp + JdgCnd.LimitWidth || ChkV[i] < StartTemp - JdgCnd.LimitWidth) return Ret;
                }
            }
            // 安定幅内にあるか？
            Array.Copy(ChkV, 0, Chk3V, 0, cnt-1);
            double cValue = Math.Abs(Chk3V.Max() - Chk3V.Min());
            if(Math.Abs(ChkV[cnt-1]- ChkV[cnt-2]) <= cValue ) return Ret;
            Ret = true;
            return Ret;
        }
        private bool PoleChgTempWidthChk(double[] ChkV,int cnt )
        {
            bool Ret = false; double[] Chk3V = new double[cnt - 1];
            // 安定幅内にあるか？
            Array.Copy(ChkV, 0, Chk3V, 0, cnt - 1);
            double cValue = double.Parse(Math.Abs(Chk3V.Max() - Chk3V.Min()).ToString("F1"));
            if (cValue > JdgCnd.LimitWidth * 2) return Ret;
            if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < JdgCnd.LimitWidth) return Ret;
            Ret = true;
            return Ret;
        }
        private bool PoleChgHumiWidthChk(double[] ChkV, int cnt)
        {
            bool Ret = false; double[] Chk3V = new double[cnt - 1];
            // 安定幅内にあるか？
            Array.Copy(ChkV, 0, Chk3V, 0, cnt - 1);
            double cValue = double.Parse(Math.Abs(Chk3V.Max() - Chk3V.Min()).ToString("F1"));
            if (cValue > JdgCnd.MaxMinWidth) return Ret;
            if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < JdgCnd.LimitWidth) return Ret;
            Ret = true;
            return Ret;
        }
        private static bool PoleChgWidthChk(double[] ChkV, int cnt)
        {
            bool Ret = false; double[] Chk3V = new double[cnt - 1];
            // 安定幅内にあるか？
            Array.Copy(ChkV, 0, Chk3V, 0, cnt - 1);
            double cValue = double.Parse(Math.Abs(Chk3V.Max() - Chk3V.Min()).ToString("F1"));
            if (cValue > JdgCnd.MaxMinWidth * 2) return Ret;
            //if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < JdgCnd.ChgJudgeTemp) return Ret;
            if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < 1.0f) return Ret;
            Ret = true;
            return Ret;
        }
        private static bool MaxMinTempChk(double[] ChkV, float dTemp ,int cnt )
        {
            bool Ret = false; double[] Chk3V = new double[cnt]; double[] ChkAve = new double[cnt];

            // 安定幅内にあるか？
            Array.Copy(ChkV, ChkV.Count() - cnt - 2, Chk3V, 0, cnt);
            for (int i = 0 ; i < cnt ; ++i) 
            {
                if (Chk3V[i] < dTemp - 2 || Chk3V[i] > dTemp + 2) return Ret;
            }

            if (Math.Abs(ChkV[ChkV.Count() - 1] - ChkV[ChkV.Count() - 3]) < 2.0f) return Ret;
            Ret = true;
            return Ret;
        }
        private static bool PoleChgWidthChk2(double[] ChkV, int cnt,float Width,float Chg)
        {
            bool Ret = false; double[] Chk3V = new double[cnt - JdgCnd.JdgAfterTime];

            // 安定幅内にあるか？
            Array.Copy(ChkV, 0, Chk3V, 0, cnt - JdgCnd.JdgAfterTime);
            double cValue = double.Parse(Math.Abs(Chk3V.Max() - Chk3V.Min()).ToString("F1"));
            if (cValue > Width) return Ret;
            if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < Chg) return Ret;
            Ret = true;

            return Ret;
        }
        private static bool PoleChgCheckTemp(double[] ChkV, int cnt)
        {
            bool Ret = false; 
            if (Math.Abs(ChkV[cnt - 1] - ChkV[cnt - 2]) < JdgCnd.ChgJudgeTemp) return Ret;
            Ret = true;
            return Ret;
        }
        private static int UpDownCheck(double[] ChkValue, double[] jdgStatus) 
        {
            //  0:安定    1:上昇    2:降下    3:不安定
            int Status = NON_STATE; bool fUp = false; bool fDown = false;
            double cValue = double.Parse(Math.Abs(ChkValue.Max() - ChkValue.Min()).ToString("F1"));
            for (int i = 1 ; i < ChkValue.Count() ; ++i) 
            {
                if (!fUp && !fDown && ChkValue[i - 1] < ChkValue[i]) fUp = true;
                if (!fUp && !fDown && ChkValue[i - 1] > ChkValue[i]) fDown = true;
                if (fUp && ChkValue[i - 1] > ChkValue[i]) { fUp = false; break; }
                if (fDown && ChkValue[i - 1] < ChkValue[i]) { fDown = false; break; }
            }

            if(!fUp && !fDown) 
            {
                double fValue = ChkValue[ChkValue.Count() - 1] - ChkValue[ChkValue.Count() - 2];
                if (jdgStatus[jdgStatus.Count() - 1] == UP_STATE && fValue >= 0) fUp = true;
                else if (jdgStatus[jdgStatus.Count() - 1] == DOWN_STATE && fValue <= 0) fDown = true;
                else if (fValue > JdgCnd.ChgJudgeTemp) fUp = true;
                else if (fValue < JdgCnd.ChgJudgeTemp * -1) fDown = true;
            }

            if (cValue < JdgCnd.MaxMinWidth) Status = STABLE_STATE;
            else if (fUp) Status = UP_STATE;
            else if (fDown) Status = DOWN_STATE;
            return Status;
        }
        static async Task DelaySample() 
        {
            await Task.Delay(10);
        }
        private static void DataShift(double[] ValueAlay) 
        {
            for (int i = 1; i < ValueAlay.Count(); ++i) 
            {
                ValueAlay[i - 1] = ValueAlay[i];
            }
        }
        private static void FluctSet(ref WIDTH_FLUCT Fluct, int cnt,int FluctTime, double[] Value, String wMsg, String Msg) 
        {
            Fluct.Exist = true;
            Fluct.StartPosition = cnt - FluctTime;
            Fluct.EndPosition = cnt - (int)JdgCnd.JdgAfterTime;
            Fluct.StartValue = (float)Value[0];
            Fluct.EndValue = (float)Value[(int)JdgCnd.JdgTimes - 1];
            
            double[] Values = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(Value, 0, Values, 0, (int)JdgCnd.JdgTimes);
            Fluct.Std = (float)Values.StandardDeviation() * 2;

            double[] TempSvs = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(RecValue.TempSV, cnt - (int)JdgCnd.JdgTimes - (int)JdgCnd.JdgAfterTime, TempSvs, 0, (int)JdgCnd.JdgTimes);
            Fluct.TempSvAve = (float)TempSvs.Average();

            double[] HumiSvs = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(RecValue.HumiSV, cnt - (int)JdgCnd.JdgTimes - (int)JdgCnd.JdgAfterTime, HumiSvs, 0, (int)JdgCnd.JdgTimes);
            Fluct.HumiSvAve = (float)HumiSvs.Average();

            double[] TempPvs = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(RecValue.TempPV, cnt - (int)JdgCnd.JdgTimes - (int)JdgCnd.JdgAfterTime, TempPvs, 0, (int)JdgCnd.JdgTimes);
            Fluct.TempPvAve = (float)TempPvs.Average();

            double[] WetPvs = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(RecValue.WetTempPV, cnt - (int)JdgCnd.JdgTimes - (int)JdgCnd.JdgAfterTime, WetPvs, 0, (int)JdgCnd.JdgTimes);
            Fluct.WetPvAve = (float)WetPvs.Average();

            double[] HumiPvs = new double[(int)JdgCnd.JdgTimes];
            Array.Copy(RecValue.HumiPV, cnt - (int)JdgCnd.JdgTimes - (int)JdgCnd.JdgAfterTime, HumiPvs, 0, (int)JdgCnd.JdgTimes);
            Fluct.HumiPvAve = (float)HumiPvs.Average();

            String cMsg = wMsg + "," + Msg;
            cMsg += "," + Fluct.StartPosition.ToString();
            cMsg += "," + Fluct.EndPosition.ToString();
            cMsg += "," + Fluct.StartValue.ToString("F1");
            cMsg += "," + Fluct.EndValue.ToString("F1");
            cMsg += "," + Fluct.TempSvAve.ToString("F1");
            cMsg += "," + Fluct.HumiSvAve.ToString("F1");
            cMsg += "," + Fluct.Std.ToString("F1");

            WriteLog(cMsg + "\r\n");
            WriteCSVLog(cMsg + "\r\n");
        }
        private static void ChgChk(int  cnt,ref  ChgRate ChgValue,double[] jdgStatus,int chkCnt,ref bool bChgStatus, double[] ChkValue, String wMsg,String sMsg ,int State,float StartTemp, float EndTemp) 
        {
            // 上昇・降下判定データ入替
            double SCValue = ChkValue[chkCnt]; double STemp = StartTemp; double ECValue = ChkValue[chkCnt]; double ETemp = EndTemp;
            if (State == DOWN_STATE) { SCValue = StartTemp; STemp = ChkValue[chkCnt]; ECValue = EndTemp; ETemp = ChkValue[chkCnt]; }

            // 条件不成立の確認処理
            if (ChgValue.StartExist && !ChgValue.EndExist && jdgStatus[chkCnt] != State && jdgStatus[chkCnt] != NON_STATE)
            //if (ChgValue.StartExist && !ChgValue.EndExist && (jdgStatus[chkCnt] != State || jdgStatus[chkCnt] == NON_STATE))
            {
                ++ChgCheckToleranceCount;
                if (ChgCheckToleranceCount > JdgCnd.ChgToleranceLimit) 
                {
                    ChgValue.StartExist = false; bChgStatus = false;
                    ChgCheckToleranceCount = 0;
                } 
            }

            // 変化開始記憶
            //if (!bChgStatus && SCValue < STemp && jdgStatus[chkCnt] == STABLE_STATE) bChgStatus = true;
            if (!bChgStatus && SCValue < STemp) { bChgStatus = true; ChgCheckToleranceCount = 0; }

            // 変化開始温度通過記憶
            if (!ChgValue.StartExist && !ChgValue.EndExist && jdgStatus[chkCnt] == State && SCValue > STemp && bChgStatus)
            {
                ChgValue.StartExist = true;
                ChgValue.StartTemp = (float)ChkValue[chkCnt - 1];
                ChgValue.StartXPos = cnt - 1;
                ChgValue.StartYValue = ChgValue.StartTemp;
            }
            // 変化終了温度通過記憶(判定条件取得完了)
            else if (ECValue >= ETemp && ChgValue.StartExist && bChgStatus && !ChgValue.EndExist)
            {
                // 変化速度検出
                ChgValue.EndExist = true;
                ChgValue.EndTemp = (float)ChkValue[chkCnt];
                ChgValue.EndXPos = cnt;
                ChgValue.EndYValue = ChgValue.EndTemp;

                float fValue = 0;

                // 実測値で変化率を演算
                fValue = Math.Abs((ChgValue.EndYValue - ChgValue.StartYValue) / (ChgValue.EndXPos - ChgValue.StartXPos));

                // 設定値で変化率を演算
                double dValue = Math.Round(Math.Abs((EndTemp - StartTemp) / (ChgValue.EndXPos - ChgValue.StartXPos)),1, MidpointRounding.AwayFromZero);
                //ChgValue.ResultValue = Math.Abs((EndTemp - StartTemp) / (ChgValue.EndXPos - ChgValue.StartXPos));
                ChgValue.ResultValue = (float)(dValue);
                // 設定値で変化率を演算
                ChgValue.ChangeRate = fValue;

                String cMsg = wMsg + "," + "変化速度" + sMsg + "検出";
                cMsg += "," + ChgValue.StartXPos.ToString();
                cMsg += "," + ChgValue.EndXPos.ToString();
                cMsg += "," + ChgValue.StartYValue.ToString("F1");
                cMsg += "," + ChgValue.EndYValue.ToString("F1");
                cMsg += "," + ChgValue.ResultValue.ToString("F1");
                cMsg += "," + fValue.ToString("F1");
                WriteLog(cMsg + "\r\n");
                WriteCSVLog(cMsg + "\r\n");

            }
            //else if (jdgStatus[chkCnt] != State && !ChgValue.EndExist)
            //{
            //    // 変化速度未検出
            //    ChgValue.StartExist = false;
            //    ChgValue.EndExist = false;
            //    // bChgStatus = false;
            //}
        }
        private static void PoleChk(int cnt, ref ChgRate PoleValue, double[] jdgStatus, int chkCnt, ref bool bPoleStatus, double[] ChkValue, String wMsg, String sMsg, int State, float StartTemp, float EndTemp)
        {
            // 上昇・降下判定データ入替
            double ECValue = ChkValue[chkCnt]; double ETemp = EndTemp;
            if (State == DOWN_STATE) { ECValue = EndTemp; ETemp = ChkValue[chkCnt]; }

            // 条件不成立の確認処理
            if (PoleValue.StartExist && !PoleValue.EndExist && jdgStatus[chkCnt] != State && jdgStatus[chkCnt] != NON_STATE)
            {
                ++PoleCheckToleranceCount;
                if (PoleCheckToleranceCount > JdgCnd.ChgToleranceLimit ) 
                {
                    PoleValue.StartExist = false; bPoleStatus = false;
                    PoleCheckToleranceCount = 0;
                }
            }

            // 極値開始記憶
            if (!bPoleStatus && StartTemp - JdgCnd.LimitWidth < ChkValue[chkCnt] && StartTemp + JdgCnd.LimitWidth > ChkValue[chkCnt] && jdgStatus[chkCnt] == STABLE_STATE) 
            { 
                bPoleStatus = true;
                PoleCheckToleranceCount = 0;
            }

            // 極値開始温度通過記憶
            if (PoleChgCheckTemp(ChkValue, (int)JdgCnd.StableCount) && bPoleStatus && !PoleValue.StartExist && !PoleValue.EndExist)
            {
                PoleValue.StartExist = true;
                PoleValue.StartTemp = (float)ChkValue[chkCnt - 1];
                PoleValue.StartXPos = cnt - 1;
                PoleValue.StartYValue = PoleValue.StartTemp;
            }
            // 極値終了温度通過記憶(判定条件取得完了)
            else if (ECValue >= ETemp && PoleValue.StartExist && bPoleStatus && !PoleValue.EndExist)
            {
                // 極値速度検出
                PoleValue.EndExist = true;
                PoleValue.EndTemp = (float)ChkValue[chkCnt];
                PoleValue.EndXPos = cnt;
                PoleValue.EndYValue = PoleValue.EndTemp;
                PoleValue.ResultValue = PoleValue.EndXPos - PoleValue.StartXPos;
                String cMsg = wMsg + "," + "極値到達" + sMsg + "検出";
                cMsg += "," + PoleValue.StartXPos.ToString();
                cMsg += "," + PoleValue.EndXPos.ToString();
                cMsg += "," + PoleValue.StartYValue.ToString("F1");
                cMsg += "," + PoleValue.EndYValue.ToString("F1");
                cMsg += "," + PoleValue.ResultValue.ToString("F1");
                WriteLog(cMsg + "\r\n");
                WriteCSVLog(cMsg + "\r\n");
            }
            //else if (jdgStatus[chkCnt] != State && !PoleValue.EndExist)
            //{
            //    // 極値到達未検出
            //    PoleValue.StartExist = false;
            //    PoleValue.EndExist = false;
            //    // bPoleStatus = false;
            //}
        }
        private static float ChgRound(double ChgV,int digit) 
        {
            float fRet = 0;
            fRet = (float)(Math.Round(ChgV, digit, MidpointRounding.AwayFromZero));
            return fRet;
        }
        private static void SetFluctDataSet(WIDTH_FLUCT Fluct ,int cnt ,int FluctTime ) 
        {

            // 演算用配列準備
            double[] TempSV = new double[FluctTime - JdgCnd.JdgAfterTime];
            double[] TempPV = new double[FluctTime - JdgCnd.JdgAfterTime];
            double[] HumiSV = new double[FluctTime - JdgCnd.JdgAfterTime];
            double[] WetPV = new double[FluctTime - JdgCnd.JdgAfterTime];
            double[] HumiPV = new double[FluctTime - JdgCnd.JdgAfterTime];
            // 演算用配列にﾃﾞｰﾀ複写
            Array.Copy(RecValue.TempSV,Fluct.StartPosition, TempSV, 0, FluctTime - JdgCnd.JdgAfterTime);
            Array.Copy(RecValue.TempPV,Fluct.StartPosition, TempPV, 0, FluctTime - JdgCnd.JdgAfterTime);
            Array.Copy(RecValue.HumiSV,Fluct.StartPosition, HumiSV, 0, FluctTime - JdgCnd.JdgAfterTime);
            Array.Copy(RecValue.HumiPV,Fluct.StartPosition, HumiPV, 0, FluctTime - JdgCnd.JdgAfterTime);
            Array.Copy(RecValue.WetTempPV,Fluct.StartPosition, WetPV, 0, FluctTime - JdgCnd.JdgAfterTime);

           Fluct.TempSvAve = ChgRound(TempSV.Average(), 1);
           Fluct.TempPvAve = ChgRound(TempPV.Average(), 1);
           Fluct.HumiSvAve = ChgRound(HumiSV.Average(), 1);
           Fluct.HumiPvAve = ChgRound(HumiPV.Average(), 1);
           Fluct.WetPvAve = ChgRound(WetPV.Average(), 1);

           Fluct.TempSvStd = ChgRound(TempSV.StandardDeviation() * 2, 1);
           Fluct.TempPvStd = ChgRound(TempPV.StandardDeviation() * 2, 1);
           Fluct.HumiSvStd = ChgRound(HumiSV.StandardDeviation() * 2, 1);
           Fluct.HumiPvStd = ChgRound(HumiPV.StandardDeviation() * 2, 1);
           Fluct.WetPvStd = ChgRound(WetPV.StandardDeviation() * 2, 1);
        }
        private static void CsvRead() 
        {
            bool bOK = false;int cnt = 0; bool bHumiEnd = false;bool EndCheckStable = false;bool bNG = true;bool bFirstEnd = false;
            double[] jdgStatus = new double[(int)JdgCnd.StableCount];   //  0:安定    1:上昇    2:降下    3:不安定

            //int FluctTime = (int)JdgCnd.JdgTimes + (int)JdgCnd.JdgAfterTime;
            int FluctTime = (int)JdgCnd.JdgAfterTime + (int)JdgCnd.JdgTimes;
            int TempDeadZone = 10;int HumiDeadZone = 10;double RoomTempSum = 0; int RoomTempSumNonCnt = 0;

            double[] ChkValue = new double[(int)JdgCnd.StableCount]; int chkCnt = 0; string wMsg = "";
            double[] FluctTempValue = new double[FluctTime]; int chkFluctTempCnt = 0;float BeforeTemp = float.MinValue;
            double[] FluctHumiValue = new double[FluctTime]; int chkFluctHumiCnt = 0;float BeforeHumi = float.MinValue;
            double[] FluctHumiPVValue = new double[FluctTime];
            bool bChgUpStatus = false; bool bChgDownStatus = false; bool bPoleUpStatus = false; bool bPoleDownStatus = false;
            bool bEndDownState = false;int iEndDownCount = 0; 

            // CSVデータファイル項目名
            int DATEMON = 0; int TIMEMON = 1; int TEMPSV = 2; int HUMISV = 3; int TEMPPV = 4; int HUMIWET = 5; int AMBTEMP = 6;

            // レコード長の初期化
            HumiEndRecord = 0;TempEndRecord = 0; ChkDataMaxlength = int.MaxValue;
            String fName = NowCSVFilePath;
            bool bHeaderNG = false;
            iFirstChgRecord = 0;
            try
            {
                WriteLog("CSVファイル読込み開始" + "\r\n");
                WriteCSVLog("CSVファイル読込み開始" + "\r\n");
                WriteLog("状態判定実施" + "\r\n");
                WriteCSVLog("状態判定実施" + "\r\n");
                WriteLog("判定ポイント検索実施,TempSV,HumiSV,TempPV,WetTemp,AmbTemp,HumiPV,槽内状態" + "\r\n");
                WriteCSVLog("判定ポイント検索実施,TempSV,HumiSV,TempPV,WetTemp,AmbTemp,HumiPV,槽内状態" + "\r\n");
                var task = DelaySample();
                // CSVファイル読込み
                // 読み込みたいCSVファイルのパスを指定して開く
                using (StreamReader sr = new StreamReader(@fName, Encoding.GetEncoding("shift_jis")))
                {
                    // 末尾まで繰り返す
                    while (!sr.EndOfStream)
                    {
                        // CSVファイルの一行を読み込む
                        string line = sr.ReadLine();
                        // 読み込んだ一行をカンマ毎に分けて配列に格納する
                        if (line == "") break;
                        string[] values = line.Split(',');
                        if (values.Length == 0) break;
                        if (bOK)
                        {
                            if (!bFirstEnd)
                            {
                                bFirstEnd = true;
                                #region 温湿度確認
                                wMsg = "";
                                double chkValue = ValueCheck(values[HUMISV]);
                                if (ChkNowSpec.HumiUse && Double.IsNaN(chkValue))
                                {
                                    // 湿度タイプ設定にもかかわらず、最初の湿度制御設定値が設定されていない
                                    wMsg += "湿度タイプ設定であるが、湿度制御設定されていない\r\n";
                                }
                                else if (!ChkNowSpec.HumiUse && !Double.IsNaN(chkValue))
                                {
                                    // 温度タイプ設定にもかかわらず、最初の湿度制御設定値が設定されている
                                    wMsg += "温度タイプ設定であるが、湿度制御設定されている\r\n";
                                }
                                if (wMsg != "")
                                {
                                    wMsg = "温湿度設定タイプ確認\r\n" + wMsg + "ＣＳＶファイル選択を見直してください。\r\n";
                                    MessageBox.Show(wMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                    break;
                                }
                                #endregion 温湿度確認
                                TestingDate = values[DATEMON];
                                ExcelValue.TestingDate = TestingDate;
                            }
                            //CSV 日付・時刻記憶
                            if (!DateTime.TryParse(values[0], out DateTime dateresult)) break;
                            if (!TimeSpan.TryParse(values[1], out TimeSpan timeresult)) break;
                            RecValue.Date[cnt] = values[0];
                            RecValue.Time[cnt] = values[1];

                            #region グラフデータ設定
                            RecValue.Times[cnt] = (double)cnt;
                            RecValue.TempSV[cnt] = ValueCheck(values[TEMPSV]);
                            RecValue.HumiSV[cnt] = ValueCheck(values[HUMISV]);
                            RecValue.TempPV[cnt] = ValueCheck(values[TEMPPV]);
                            RecValue.WetTempPV[cnt] = ValueCheck(values[HUMIWET]);

                            RecValue.RoomTemp[cnt] = ValueCheck(values[AMBTEMP]);

                            //RoomTempSum += RecValue.RoomTemp[cnt];
                            if (!Double.IsNaN(RecValue.RoomTemp[cnt])) RoomTempSum += RecValue.RoomTemp[cnt];
                            else RoomTempSumNonCnt += 1;
                            RecValue.HumiPV[cnt] = ValueCheck(GetHumiPV(RecValue.TempPV[cnt], RecValue.WetTempPV[cnt], 1));

                            if (bHumiEnd)
                            {
                                RecValue.HumiSV[cnt] = double.NaN;
                                RecValue.WetTempPV[cnt] = double.NaN;
                                RecValue.HumiPV[cnt] = double.NaN;
                            }

                            #endregion グラフデータ設定
                            #region グラフデータLog設定
                            wMsg = RecValue.Times[cnt].ToString();
                            wMsg += "," + RecValue.TempSV[cnt].ToString("F1");
                            wMsg += "," + RecValue.HumiSV[cnt].ToString("F1");
                            wMsg += "," + RecValue.TempPV[cnt].ToString("F1");
                            wMsg += "," + RecValue.WetTempPV[cnt].ToString("F1");
                            wMsg += "," + RecValue.RoomTemp[cnt].ToString("F1");
                            wMsg += "," + RecValue.HumiPV[cnt].ToString("F1");
                            #endregion グラフデータLog設定

                            // 温度精度設定（装置毎の温度範囲毎の温度精度を設定する）
                            if (RecValue.TempPV[cnt] <= ChkNowSpec.AccuracyTemp1EndTemp + ChkNowSpec.AccuracyTemp2Spec)
                            {
                                JdgCnd.LimitWidth = ChkNowSpec.AccuracyTemp1Spec;
                            } else if (RecValue.TempPV[cnt] <= ChkNowSpec.AccuracyTemp2EndTemp + ChkNowSpec.AccuracyTemp3Spec)
                            {
                                JdgCnd.LimitWidth = ChkNowSpec.AccuracyTemp2Spec;
                            }
                            else
                            {
                                JdgCnd.LimitWidth = ChkNowSpec.AccuracyTemp3Spec;
                            }

                            // 湿度レコード完了確認
                            double dCheck = ValueCheck(values[HUMISV]);
                            if (!bHumiEnd && ((ChkNowSpec.HumiUse && cnt > HUMIENDCNT) || dCheck < 1 || Double.IsNaN(dCheck)))
                            {
                                HumiEndRecord = cnt;
                                ExcelValue.HumiEndRecord = cnt;
                                bHumiEnd = true;
                            }
                            // 判定データ準備
                            ChkValue[chkCnt] = RecValue.TempPV[cnt];
                            FluctTempValue[chkFluctTempCnt] = RecValue.TempSV[cnt];
                            FluctHumiValue[chkFluctHumiCnt] = RecValue.HumiSV[cnt];
                            FluctHumiPVValue[chkFluctHumiCnt] = RecValue.HumiPV[cnt];
                            if (cnt == 600)
                            {
                                Console.WriteLine("");
                            }

                            // 温度判定処理
                            String sMsg = "不安定領域";
                            if (chkCnt == JdgCnd.StableCount - 1)
                            {
                                // 温度判定処理
                                jdgStatus[chkCnt] = UpDownCheck(ChkValue, jdgStatus);
                                if (jdgStatus[chkCnt] == STABLE_STATE) sMsg = "安定領域";
                                else if (jdgStatus[chkCnt] == UP_STATE) sMsg = "上昇領域";
                                else if (jdgStatus[chkCnt] == DOWN_STATE) sMsg = "降下領域";

                                // 最初の変化を確認する
                                if ((jdgStatus[chkCnt] == UP_STATE || jdgStatus[chkCnt] == DOWN_STATE) && jdgStatus[chkCnt - 3] == STABLE_STATE)
                                {
                                    if (iFirstChgRecord == 0 && cnt > JdgCnd.StableCount * 2)
                                    {
                                        iFirstChgRecord = cnt;
                                        Console.WriteLine("最初に安定から変化したレコード　= " + iFirstChgRecord.ToString());
                                    }
                                }

                                // 温度特性確認
                                if (bHumiEnd)
                                {
                                    #region 変化速度上昇判定：上昇領域で開始温度を越えた所から終了温度を越えた所まで
                                    ChgChk(cnt, ref ChgUpValue, jdgStatus, chkCnt, ref bChgUpStatus, ChkValue, wMsg, "上昇", UP_STATE, ChkNowSpec.ChgUpStartTemp, ChkNowSpec.ChgUpEndTemp);
                                    #endregion 変化速度上昇判定：上昇領域で開始温度を越えた所から終了温度を越えた所まで

                                    #region 変化速度降下判定：降下領域で開始温度を越えた所から終了温度を越えた所まで
                                    ChgChk(cnt, ref ChgDownValue, jdgStatus, chkCnt, ref bChgDownStatus, ChkValue, wMsg, "降下", DOWN_STATE, ChkNowSpec.ChgDownStartTemp, ChkNowSpec.ChgDownEndTemp);
                                    #endregion 変化速度降下判定：降下領域で開始温度を越えた所から終了温度を越えた所まで

                                    #region 極値到達上昇判定：開始温度近傍の安定領域から終了温度を越えた所まで
                                    PoleChk(cnt, ref ExtUpValue, jdgStatus, chkCnt, ref bPoleUpStatus, ChkValue, wMsg, "上昇", UP_STATE, ChkNowSpec.PoleUpStartTemp, ChkNowSpec.PoleUpEndTemp);
                                    #endregion 極値到達上昇判定：開始温度近傍の安定領域から終了温度を越えた所まで

                                    #region 極値到達降下判定：開始温度近傍の安定領域から終了温度を越えた所まで
                                    PoleChk(cnt, ref ExtDownValue, jdgStatus, chkCnt, ref bPoleDownStatus, ChkValue, wMsg, "降下", DOWN_STATE, ChkNowSpec.PoleDownStartTemp, ChkNowSpec.PoleDownEndTemp);
                                    #endregion 極値到達降下判定：開始温度近傍の安定領域から終了温度を越えた所まで
                                }
                                // 温度判定エリアリング処理（1データシフト）
                                DataShift(ChkValue);
                                DataShift(jdgStatus);
                            }
                            else
                            {
                                ++chkCnt;
                            }
                            WriteLog(wMsg + "," + sMsg + "\r\n");
                            WriteCSVLog(wMsg + "," + sMsg + "\r\n");
                            //DataQueue.Enqueue(wMsg + "," + sMsg + "\r\n");

                            #region 温度変動判定処理・温度到達判定処理
                            if (bHumiEnd)
                            {
                                TempDeadZone += 1;
                                if (chkFluctTempCnt == FluctTime - 1)
                                {
                                    // MaxUp/MinDown確認
                                    int CheckDataCount = (int)JdgCnd.JdgTimes;
                                    float uTemp = ChkNowSpec.SpecMaxUpTemp;
                                    float dTemp = ChkNowSpec.SpecMinDownTemp;
                                    float AveValue = float.MinValue;
                                    if (MaxMinTempChk(FluctTempValue, uTemp, CheckDataCount) && !MaxTemp.Exist)
                                    {
                                        double[] ChkAve = new double[CheckDataCount];
                                        Array.Copy(RecValue.TempSV, cnt - CheckDataCount - 2, ChkAve, 0, CheckDataCount);
                                        AveValue = (float)ChkAve.Average();

                                        // 最大最大検出
                                        MaxTemp.Exist = true;
                                        MaxTemp.StartPosition = cnt - CheckDataCount;
                                        MaxTemp.EndPosition = cnt - 2;
                                        MaxTemp.StartValue = (float)RecValue.TempPV[MaxTemp.StartPosition];
                                        MaxTemp.TempSvAve = AveValue;
                                    }
                                    AveValue = float.MinValue;
                                    if (MaxMinTempChk(FluctTempValue, dTemp, CheckDataCount) && !MinTemp.Exist)
                                    {
                                        double[] ChkAve = new double[CheckDataCount];
                                        Array.Copy(RecValue.TempSV, cnt - CheckDataCount - 2, ChkAve, 0, CheckDataCount);
                                        AveValue = (float)ChkAve.Average();

                                        // 最小最小検出
                                        MinTemp.Exist = true;
                                        MinTemp.StartPosition = cnt - CheckDataCount;
                                        MinTemp.EndPosition = cnt - 2;
                                        MinTemp.StartValue = (float)RecValue.TempPV[MinTemp.StartPosition];
                                        MinTemp.TempSvAve = AveValue;
                                    }

                                    if (TempDeadZone > 10)
                                    {
                                        // 温度変動判定処理
                                        // 温度精度確認
                                        float ChgBeforeTemp = (float)FluctTempValue[FluctTime - JdgCnd.JdgAfterTime - 2];
                                        float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                                        if (ChgBeforeTemp > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                                        if (ChgBeforeTemp > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                                        if (PoleChgWidthChk2(FluctTempValue, FluctTime, JdgCnd.MaxMinWidth * 2, JdgCnd.ChgJudgeTemp))
                                        {
                                            // 安定幅以内で安定していて最後に2℃変化した
                                            // 該当する設定の温度を検索
                                            for (int i = 0; i < TempFluct.Count; ++i)
                                            {
                                                // 最低温度判定は20℃の温度安定検出後のみ判定する
                                                if ((TempFluct[i].SetValue == ChkNowSpec.SpecMinTemp && FLUCT20_Exist) || !(TempFluct[i].SetValue == ChkNowSpec.SpecMinTemp))
                                                {
                                                    // 直前温度が最大越え温度・最小越え温度の除外処理
                                                    if (!((uTemp - 1 <= ChgBeforeTemp) && (uTemp + 1 >= ChgBeforeTemp) || (dTemp - 1 <= ChgBeforeTemp) && (dTemp + 1 >= ChgBeforeTemp)))
                                                    {
                                                        //　温度精度以内である事を確認する
                                                        if (TempFluct[i].SetValue + TempWidth >= ChgBeforeTemp && TempFluct[i].SetValue - TempWidth <= ChgBeforeTemp && !TempFluct[i].Exist)
                                                        {
                                                            // 対象温度であると判断。合わせて未検出だった
                                                            TempFluct[i].Exist = true;
                                                            // 開始・終了ポイント・平均・標準偏差演算設定
                                                            TempFluct[i].StartPosition = cnt - FluctTime;
                                                            TempFluct[i].EndPosition = cnt - JdgCnd.JdgAfterTime;
                                                            TempFluct[i].StartValue = (float)RecValue.TempPV[TempFluct[i].StartPosition];
                                                            TempFluct[i].EndValue = (float)RecValue.TempPV[TempFluct[i].EndPosition];
                                                            SetFluctDataSet(TempFluct[i], cnt, FluctTime);
                                                            // 乾球温度で判定
                                                            float CheckSpec = 0.3f;
                                                            if (TempFluct[i].SetValue > 100) CheckSpec = 0.5f;
                                                            if (TempFluct[i].TempPvStd <= CheckSpec) TempFluct[i].Judge = true;
                                                            String cMsg = wMsg + ",温度変動";
                                                            cMsg += "," + TempFluct[i].StartPosition.ToString();
                                                            cMsg += "," + TempFluct[i].EndPosition.ToString();
                                                            cMsg += "," + TempFluct[i].StartValue.ToString("F1");
                                                            cMsg += "," + TempFluct[i].EndValue.ToString("F1");
                                                            cMsg += "," + TempFluct[i].TempPvStd.ToString("F1");
                                                            WriteLog(cMsg + "\r\n");
                                                            WriteCSVLog(cMsg + "\r\n");
                                                            //DataQueue.Enqueue(cMsg + "\r\n");
                                                            TempDeadZone = 0;
                                                            // 20℃の温度安定を検出した事を記憶する
                                                            if (TempFluct[i].SetValue == 20) FLUCT20_Exist = true;
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    // 温度変動判定エリアリング処理（1データシフト）
                                    DataShift(FluctTempValue);
                                }
                                else
                                {
                                    ++chkFluctTempCnt;
                                }
                            } 
                            #endregion 温度変動判定処理・温度到達判定処理
                            #region 湿度変動判定処理
                            if (!bHumiEnd)
                            {
                                HumiDeadZone += 1;
                                if (chkFluctHumiCnt == FluctTime - 1)
                                {
                                    if (HumiDeadZone > 10)
                                    {
                                        // 湿度変動判定処理
                                        if (PoleChgWidthChk2(FluctHumiValue, FluctTime, 5f, 5f))
                                        {
                                            // 3%以内で安定していて最後に5%変化した

                                            // 湿度精度確認
                                            float ChgBeforeTemp = (float)FluctTempValue[FluctTime - JdgCnd.JdgAfterTime - 2];
                                            float ChgBeforeHumi = (float)FluctHumiValue[FluctTime - JdgCnd.JdgAfterTime - 2];

                                            for (int i = 0; i < HumiFluct.Count; ++i)
                                            {
                                                float HumiWidth = Humiset[i].HumiAccuracy;
                                                // 湿度精度以内であることを確認する。
                                                if (HumiFluct[i].SetValue + HumiWidth >= ChgBeforeHumi && HumiFluct[i].SetValue - HumiWidth <= ChgBeforeHumi && !HumiFluct[i].Exist)
                                                {
                                                    // 対象湿度であると判断。合わせて未検出だった
                                                    HumiFluct[i].Exist = true;
                                                    // 開始・終了ポイント・平均・標準偏差演算設定
                                                    HumiFluct[i].StartPosition = cnt - FluctTime;
                                                    HumiFluct[i].EndPosition = cnt - JdgCnd.JdgAfterTime;
                                                    HumiFluct[i].StartValue = (float)RecValue.HumiPV[HumiFluct[i].StartPosition];
                                                    HumiFluct[i].EndValue = (float)RecValue.HumiPV[HumiFluct[i].EndPosition];
                                                    SetFluctDataSet(HumiFluct[i], cnt, FluctTime);

                                                    // 相対湿度で判定
                                                    if (HumiFluct[i].HumiPvStd <= 3f) HumiFluct[i].Judge = true;

                                                    String cMsg = wMsg + ",湿度変動";
                                                    cMsg += "," + HumiFluct[i].StartPosition.ToString();
                                                    cMsg += "," + HumiFluct[i].EndPosition.ToString();
                                                    cMsg += "," + HumiFluct[i].StartValue.ToString("F1");
                                                    cMsg += "," + HumiFluct[i].EndValue.ToString("F1");
                                                    cMsg += "," + HumiFluct[i].HumiPvStd.ToString("F1");

                                                    WriteLog(cMsg + "\r\n");
                                                    WriteCSVLog(cMsg + "\r\n");
                                                    //DataQueue.Enqueue(cMsg + "\r\n");
                                                    HumiDeadZone = 0;
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    // 湿度変動判定エリアリング処理（1データシフト）
                                    DataShift(FluctHumiValue);
                                    DataShift(FluctHumiPVValue);
                                }
                                else
                                {
                                    ++chkFluctHumiCnt;
                                }
                            }
                            #endregion 湿度変動判定処理
                            // 次のレコード確認
                            ++cnt;
                            ChkLength = cnt;
                        }
                        // 不要なヘッダーを飛ばす
                        if ((values[0] == "日付" || values[0] == "Date") && !bOK )
                        {
                            if (!values[2].Contains("制御温度") && !values[2].Contains("TEMP")) bHeaderNG = true;
                            if (!values[3].Contains("制御湿度") && !values[3].Contains("HUMI")) bHeaderNG = true;
                            if (!values[4].Contains("中央乾球") && !values[4].Contains("CENT")) bHeaderNG = true;
                            if (!values[5].Contains("中央湿球") && !values[5].Contains("WET")) bHeaderNG = true;
                            if (!values[6].Contains("周囲温度") && !values[6].Contains("AMB")) bHeaderNG = true;
                            bOK = true;
                        }

                        // 温度判定４種類取得して安定領域になったら取り込み終了
                        if (ChgUpValue.EndExist && ChgDownValue.EndExist
                            && ExtUpValue.EndExist && ExtDownValue.EndExist
                            && jdgStatus[chkCnt] == STABLE_STATE
                            && !EndCheckStable) EndCheckStable = true;
                        if (EndCheckStable && jdgStatus[chkCnt] == DOWN_STATE && !bEndDownState) bEndDownState = true;
                        if (bEndDownState) ++iEndDownCount;
                        if (iEndDownCount > 30) break;
                        if (cnt >= SystemSet.Rec_Length) break;
                        if (bHeaderNG) break;
                    }
                    // 検索結果確認
                    ChkDataMaxlength = cnt - 1;
                    ExcelValue.EndRecord = ChkDataMaxlength;
                    if (cnt != 0) 
                    { 
                        RoomTempAverage = ((float)RoomTempSum / (cnt - RoomTempSumNonCnt));
                        ExcelValue.RoomTempAverage = RoomTempAverage;

                    }
                    if (EndCheckStable)
                    {
                        TempEndRecord = ChkDataMaxlength - HumiEndRecord + 1;
                        ExcelValue.TempEndRecord = TempEndRecord;
                        string cMsg = "";
                        if (ChkNowSpec.HumiUse) cMsg = "湿度タイプ：";
                        else cMsg = "温度タイプ：";
                        cMsg += "湿度レコード長 = " + HumiEndRecord.ToString() + " 温度レコード長 = " + TempEndRecord.ToString();
                        WriteLog(cMsg + "\r\n");
                        WriteCSVLog(cMsg + "\r\n");
                        //DataQueue.Enqueue(cMsg + "\r\n");
                        StateOut("温度変化値降下\r\n", ChgDownValue);
                        StateOut("温度極値到達降下\r\n", ExtDownValue);
                        StateOut("温度変化値上昇\r\n", ChgUpValue);
                        StateOut("温度極値到達上昇\r\n", ExtUpValue);
                    }
                    else
                    {
                        string cMsg = "";
                        TempEndRecord = ChkDataMaxlength - HumiEndRecord + 1;
                        ExcelValue.TempEndRecord = TempEndRecord;
                        if (bHeaderNG)
                        {
                            // ヘッダー異常だった
                            cMsg = "\r\nデータの並びが違います。\r\n";
                            cMsg += "日付,時刻,制御温度,制御湿度,槽中央乾球温度,槽中央湿球温度,周囲温度の\r\n";
                            cMsg += "順番である必要があります。\r\n";
                            cnt = int.MinValue;
                        }
                        else 
                        { 
                            // 判定条件が見つからなかった
                            cMsg = "\r\n判定条件に一致するデータを検出出来ませんでした。\r\n";
                            if (!ChgUpValue.EndExist) cMsg += "温度上昇変化速度未検出\r\n";
                            if (!ChgDownValue.EndExist) cMsg += "温度降下変化速度未検出\r\n";
                            if (!ExtUpValue.EndExist) cMsg += "温度上昇極値到達未検出\r\n";
                            if (!ExtDownValue.EndExist) cMsg += "温度降下極値到達未検出\r\n";
                            cnt = int.MaxValue;
                        }
                        WriteCSVLog(cMsg);
                        //DataQueue.Enqueue(cMsg);
                    }
                    // 2024/11/06　改善の為、解除
                    // 2024/08/22　改善の為、一時保留
                    WriteLog("平均室温：" + RoomTempAverage.ToString("F1") + "℃\r\n");
                    // 2024/08/22　改善の為、一時保留
                    // 2024/11/06　改善の為、解除
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            WriteCSVLog("CSVファイル読込み終了" + "\r\n\r\n");
            //DataQueue.Enqueue("CSVファイル読込み終了" + "\r\n\r\n");
            tCSVReadEnd = true;
            RecLength = cnt;
        }
        private static void StateOut(string msg ,ChgRate disp) 
        {
            msg += ",,,開始POS," + disp.StartXPos.ToString() + "\r\n";
            msg += ",,,終了POS," + disp.EndXPos.ToString() + "\r\n";
            msg += ",,,開始温度," + disp.StartYValue.ToString() + "\r\n";
            msg += ",,,終了温度," + disp.EndYValue.ToString() + "\r\n";
            WriteLog(msg);
            WriteCSVLog(msg);
        }
        private void StateInit() 
        {
            ChgUpValue.StartExist = false; ChgUpValue.EndExist = false;
            ChgDownValue.StartExist = false; ChgDownValue.EndExist = false;
            ExtUpValue.StartExist = false; ExtUpValue.EndExist = false;
            ExtDownValue.StartExist = false; ExtDownValue.EndExist = false;
        }
        private string ChkInput() 
        {
            string returnMsg = "";
            // 製造番号
            //SerialNo = lblSerial_number.Text;
            if (SerialNo == "") returnMsg = "製造番号が登録されていません\r\n";
            // ユーザー名
            UserName = txtUserName.Text;
            if (UserName == "") returnMsg = "ユーザー名が設定されていません\r\n";

            // 判定条件確認と設定
            JdgCnd.LimitWidth = float.Parse(txtlimitWidth.Text);
            JdgCnd.ChgJudgeTemp = float.Parse(txtChgTemp.Text);
            JdgCnd.MaxMinWidth = float.Parse(txtmaxminWidth.Text);
            JdgCnd.StableCount = float.Parse(txtstabCount.Text);

            JdgCnd.ChgToleranceLimit = int.Parse(txtToleranceLimit.Text);

            ChkDataMaxlength = int.Parse(txtEndRecord.Text);

            returnMsg += DataGet();

            return returnMsg;
        }
        private string JudgeValueStr(float value, int digit, int dec, bool exist)
        {
            string sRet = "";
            for (int i = 0; i < digit; ++i) sRet += "-";
            if (exist)
            {
                if (digit == 2 && Math.Abs(value) < 100) sRet = value.ToString("F" + dec.ToString());
                else if (digit == 3 && Math.Abs(value) < 1000) sRet = value.ToString("F" + dec.ToString());
            }
            return sRet;
        }
        private void TestResultDisply() 
        {
            int elapsedtime = 0;
            // 温度極値上昇

            elapsedtime = (int)ExtUpValue.ResultValue;
            if (!ExtUpValue.EndExist) { elapsedtime = int.MaxValue; ExtUpValue.ResultValue = int.MaxValue; }
            //lblPoleUpValue.Text = elapsedtime.ToString();
            lblPoleUpValue.Text = JudgeValueStr(elapsedtime, 2, 0, ExtUpValue.EndExist);
            if (elapsedtime <= ChkNowSpec.PoleUpSpec && ExtUpValue.EndExist) { lblPoleUpStr.Text = "GOOD"; lblPoleUpStr.BackColor = Color.Lime; }
            else { lblPoleUpStr.Text = "NG"; lblPoleUpStr.BackColor = Color.Red; }

            // 温度極値降下
            elapsedtime = (int)ExtDownValue.ResultValue;
            if (!ExtDownValue.EndExist) { elapsedtime = int.MaxValue; ExtDownValue.ResultValue = int.MaxValue; }
            //lblPoleDownValue.Text = elapsedtime.ToString();
            lblPoleDownValue.Text = JudgeValueStr(elapsedtime, 2, 0, ExtDownValue.EndExist);

            if (elapsedtime <= ChkNowSpec.PoleDownSpec && ExtDownValue.EndExist) { lblPoleDownStr.Text = "GOOD"; lblPoleDownStr.BackColor = Color.Lime; }
            else { lblPoleDownStr.Text = "NG"; lblPoleDownStr.BackColor = Color.Red; }

            // 温度極値降下時の室温平均温度算出
            PoleDownRoomTempAverage = RoomTempAverage;
            if (ExtDownValue.EndExist)
            {
                int sPos = ExtDownValue.StartXPos;
                int ePos = ExtDownValue.EndXPos;
                double[] RoomTemp = new double[ePos - sPos];
                Array.Copy(RecValue.RoomTemp, sPos, RoomTemp, 0, ePos - sPos);
                PoleDownRoomTempAverage= ChgRound(RoomTemp.Average(), 1);
            }
            ExcelValue.PoleDownRoomTempAverage = PoleDownRoomTempAverage;

            // 温度変化速度上昇
            float chgtime = 0f;
            chgtime = ChgUpValue.ResultValue;
            if (!ChgUpValue.EndExist) { chgtime = float.MinValue; ChgUpValue.ResultValue = float.MinValue; }
            //lblChgUpValue.Text = chgtime.ToString("F1");
            lblChgUpValue.Text = JudgeValueStr(chgtime, 2, 1, ChgUpValue.EndExist);
            lblChgUpTime.Text = JudgeValueStr(ChgUpValue.EndXPos- ChgUpValue.StartXPos, 3, 0, ChgUpValue.EndExist) + "分";
            if (chgtime >= ChkNowSpec.ChgUpSpec && ChgUpValue.EndExist) { lblChgUpStr.Text = "GOOD"; lblChgUpStr.BackColor = Color.Lime; }
            else { lblChgUpStr.Text = "NG"; lblChgUpStr.BackColor = Color.Red; }

            // 温度変化速度降下
            chgtime = ChgDownValue.ResultValue;
            if (!ChgDownValue.EndExist) { chgtime = float.MinValue; ChgDownValue.ResultValue = float.MinValue; }
            //lblChgDownValue.Text = chgtime.ToString("F1");
            lblChgDownValue.Text = JudgeValueStr(chgtime, 2, 1, ChgDownValue.EndExist);
            lblChgDownTime.Text = JudgeValueStr(ChgDownValue.EndXPos - ChgDownValue.StartXPos, 3, 0, ChgDownValue.EndExist) + "分";
            if (chgtime >= ChkNowSpec.ChgDownSpec && ChgDownValue.EndExist) { lblChgDownStr.Text = "GOOD"; lblChgDownStr.BackColor = Color.Lime; }
            else { lblChgDownStr.Text = "NG"; lblChgDownStr.BackColor = Color.Red; }

            if (FLUCT_USE && FIRST_STEP)
            {
                bool bNgFlg = false; bool bNg1Flg = false; bool bNg2Flg = false; bool bNg3Flg = false;
                WIDTH_FLUCT TFluct = new WIDTH_FLUCT();
                WIDTH_FLUCT HFluct = new WIDTH_FLUCT();
                // 温度変動幅
                foreach (WIDTH_FLUCT tf in TempFluct)
                {
                    if (ChkNowSpec.VarTemp1EndTemp >= tf.StartValue)
                    {
                        // 低温度域での判定確認
                        if (!tf.Judge) bNg1Flg = true;
                    }
                    else if (ChkNowSpec.VarTemp2EndTemp >= tf.StartValue)
                    {
                        // 中温度域での判定確認
                        if (!tf.Judge) bNg2Flg = true;
                    }
                    else
                    {
                        // 高温度域での判定確認
                        if (!tf.Judge) bNg3Flg = true;
                    }
                }
                if (bNg1Flg) { lblVarTemp1Str.Text = "NG"; lblVarTemp1Str.BackColor = Color.Red; }
                else { lblVarTemp1Str.Text = "GOOD"; lblVarTemp1Str.BackColor = Color.Lime; }
                if (bNg2Flg) { lblVarTemp2Str.Text = "NG"; lblVarTemp2Str.BackColor = Color.Red; }
                else { lblVarTemp2Str.Text = "GOOD"; lblVarTemp2Str.BackColor = Color.Lime; }
                if (bNg3Flg) { lblVarTemp3Str.Text = "NG"; lblVarTemp3Str.BackColor = Color.Red; }
                else { lblVarTemp3Str.Text = "GOOD"; lblVarTemp3Str.BackColor = Color.Lime; }

                // 湿度変動幅
                foreach (WIDTH_FLUCT hf in HumiFluct)
                {
                    // 判定確認
                    if (!hf.Judge) bNgFlg = true;
                }
                if (bNgFlg) { lblVarHumiStr.Text = "NG"; lblVarHumiStr.BackColor = Color.Red; }
                else { lblVarHumiStr.Text = "GOOD"; lblVarHumiStr.BackColor = Color.Lime; }
            }
            else
            {
                // 第二ステップ
                // 温度到達　最低温度
                lblTempReachJdgeL.Text = "NG"; lblTempReachJdgeL.BackColor = Color.Red;
                if (MinTemp.Exist)
                {
                    //if (Math.Abs(MinTemp.TempSvAve) < 100) lblTempReachMemL.Text = MinTemp.TempSvAve.ToString("F1");
                    lblTempReachMemL.Text = JudgeValueStr(MinTemp.TempSvAve, 3, 1, true);
                    lblTempReachJdgeL.Text = "GOOD"; lblTempReachJdgeL.BackColor = Color.Lime;
                }
                else lblTempReachMemL.Text = "-.-";
                // 温度到達　最高温度
                lblTempReachJdgeH.Text = "NG"; lblTempReachJdgeH.BackColor = Color.Red;
                if (MaxTemp.Exist)
                {
                    //if(Math.Abs(MaxTemp.TempSvAve) < 100) lblTempReachMemH.Text = MaxTemp.TempSvAve.ToString("F1");
                    lblTempReachMemH.Text = JudgeValueStr(MaxTemp.TempSvAve, 3, 1, true);
                    lblTempReachJdgeH.Text = "GOOD"; lblTempReachJdgeH.BackColor = Color.Lime;
                }
                else lblTempReachMemH.Text = "-.-";
                // 検査結果（温度変動）
                if (TempFluct.Count != 0)
                {
                    // 温度到達　仕様値最低温度
                    lblTempReachJdgeM.Text = "NG"; lblTempReachJdgeM.BackColor = Color.Red; lblTempReachMemM.Text = "---.-";
                    lblTempOpeJdgeH.Text = "NG"; lblTempOpeJdgeH.BackColor = Color.Red; lblTempOpeMemH.Text = "---.-";
                    lblTempOpeJdgeM.Text = "NG"; lblTempOpeJdgeM.BackColor = Color.Red; lblTempOpeMemM.Text = "---.-";
                    lblTempOpeJdgeL.Text = "NG"; lblTempOpeJdgeL.BackColor = Color.Red; lblTempOpeMemL.Text = "---.-";
                    lblTempFluctJdgeH.Text = "NG"; lblTempFluctJdgeH.BackColor = Color.Red; lblTempFluctMemH.Text = "-.-";
                    lblTempFluctJdgeM.Text = "NG"; lblTempFluctJdgeM.BackColor = Color.Red; lblTempFluctMemM.Text = "-.-";
                    lblTempFluctJdgeL.Text = "NG"; lblTempFluctJdgeL.BackColor = Color.Red; lblTempFluctMemL.Text = "-.-";

                    float MaxValue = float.MinValue;
                    float MinSVValue = float.MaxValue; float MainDownSVValue = float.MaxValue;
                    float MaxSVValue = float.MinValue; float MaxUpSVValue = float.MinValue;

                    foreach (WIDTH_FLUCT fluct in TempFluct)
                    {
                        float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                        if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                        if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                        if (fluct.TempSvAve >= ChkNowSpec.SpecMaxTemp - TempWidth / 2 && fluct.TempSvAve <= ChkNowSpec.SpecMaxTemp + TempWidth / 2)
                        {
                            // 最大温度の安定があった。
                            //lblTempOpeMemH.Text = fluct.TempSvAve.ToString("F1");
                            lblTempOpeMemH.Text = JudgeValueStr(fluct.TempSvAve, 3, 1, true);
                            MaxSVValue = ChkNowSpec.SpecMaxTemp;
                            if (Math.Abs(fluct.TempSvAve - MaxSVValue) < 1.2) { lblTempOpeJdgeH.Text = "GOOD"; lblTempOpeJdgeH.BackColor = Color.Lime; }

                            //lblTempFluctMemH.Text = fluct.TempPvStd.ToString("F2");
                            lblTempFluctMemH.Text = JudgeValueStr(fluct.TempPvStd, 3, 1, true);
                            if (fluct.TempPvStd <= 0.5f) { lblTempFluctJdgeH.Text = "GOOD"; lblTempFluctJdgeH.BackColor = Color.Lime; }

                        }
                        else if (fluct.TempSvAve >= ChkNowSpec.SpecMinTemp - TempWidth && fluct.TempSvAve <= ChkNowSpec.SpecMinTemp + TempWidth)
                        {
                            MinSVValue = ChkNowSpec.SpecMinTemp;
                            // 最小温度の安定があった。
                            //lblTempOpeMemL.Text = fluct.TempSvAve.ToString("F1");
                            lblTempOpeMemL.Text = JudgeValueStr(fluct.TempSvAve, 3, 1, true);
                            if (Math.Abs(fluct.TempSvAve - MinSVValue) < 1.2) { lblTempOpeJdgeL.Text = "GOOD"; lblTempOpeJdgeL.BackColor = Color.Lime; }

                            //lblTempFluctMemL.Text = fluct.TempPvStd.ToString("F2");
                            lblTempFluctMemL.Text = JudgeValueStr(fluct.TempPvStd, 3, 1, true);
                            if (fluct.TempPvStd <= 0.3f) { lblTempFluctJdgeL.Text = "GOOD"; lblTempFluctJdgeL.BackColor = Color.Lime; }

                            // 槽中央温度で判定
                            //lblTempReachMemM.Text = fluct.TempPvAve.ToString("F1");
                            lblTempReachMemM.Text = JudgeValueStr(fluct.TempPvAve, 3, 1, true);
                            lblTempReachJdgeM.Text = "GOOD"; lblTempReachJdgeM.BackColor = Color.Lime;
                        }
                        else if (fluct.TempSvAve > 20 - TempWidth && fluct.TempSvAve < 20 + TempWidth)
                        {
                            // 20℃の安定があった。
                            //lblTempOpeMemM.Text = fluct.TempSvAve.ToString("F1");
                            lblTempOpeMemM.Text = JudgeValueStr(fluct.TempSvAve, 3, 1, true);
                            if (Math.Abs(fluct.TempSvAve - 20) < 1.2) { lblTempOpeJdgeM.Text = "GOOD"; lblTempOpeJdgeM.BackColor = Color.Lime; }

                            //lblTempFluctMemM.Text = fluct.TempPvStd.ToString("F2");
                            lblTempFluctMemM.Text = JudgeValueStr(fluct.TempPvStd, 3, 1, true);
                            Console.WriteLine(fluct.TempPvStd);
                            if (fluct.TempPvStd <= 0.3f) { lblTempFluctJdgeM.Text = "GOOD"; lblTempFluctJdgeM.BackColor = Color.Lime; }

                        }
                    }
                }
                // 検査結果（湿度変動）
                if (HumiFluct.Count != 0)
                {
                    float MaxValue = float.MinValue;
                    for (int i = 0; i < 4; ++i)
                    {
                        HumiOpeJdge[i].Text = "NG"; HumiOpeJdge[i].BackColor = Color.Red;
                        HumiFluctJdge[i].Text = "NG"; HumiFluctJdge[i].BackColor = Color.Red;
                        HumiOpeTemp[i].Text = "--.-"; HumiOpeHumi[i].Text = "--.-";
                        HumiOpeMonTemp[i].Text = "--.-"; HumiOpeMonHumi[i].Text = "--.-";
                        HumiOpeHumiPV[i].Text = "--.-"; HumiFluctStd[i].Text = "-.-";
                    }
                    foreach (WIDTH_FLUCT fluct in HumiFluct)
                    {
                        float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                        if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                        if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                        for (int i = 0; i < 4; ++i)
                        {
                            if (fluct.TempSvAve >= Humiset[i].SetTemp - TempWidth && fluct.TempSvAve <= Humiset[i].SetTemp + TempWidth &&
                                fluct.HumiSvAve >= Humiset[i].SetHumi - TempWidth && fluct.HumiSvAve <= Humiset[i].SetHumi + TempWidth)
                            {
                                // 第iステップの設定安定が合った
                                //HumiOpeTemp[i].Text = fluct.TempSvAve.ToString("F1");
                                HumiOpeTemp[i].Text = JudgeValueStr(fluct.TempSvAve, 3, 1, true);
                                //HumiOpeHumi[i].Text = fluct.HumiSvAve.ToString("F1");
                                HumiOpeHumi[i].Text = JudgeValueStr(fluct.HumiSvAve, 3, 1, true);
                                if (Math.Abs(fluct.HumiSvAve - Humiset[i].SetHumi) <= 3) { HumiOpeJdge[i].Text = "GOOD"; HumiOpeJdge[i].BackColor = Color.Lime; }
                                //HumiOpeMonTemp[i].Text = fluct.TempPvAve.ToString("F1");
                                HumiOpeMonTemp[i].Text = JudgeValueStr(fluct.TempPvAve, 3, 1, true);
                                //HumiOpeMonHumi[i].Text = fluct.WetPvAve.ToString("F1");
                                HumiOpeMonHumi[i].Text = JudgeValueStr(fluct.WetPvAve, 3, 1, true);
                                //HumiOpeHumiPV[i].Text = fluct.HumiPvAve.ToString("F1");
                                HumiOpeHumiPV[i].Text = JudgeValueStr(fluct.HumiPvAve, 3, 1, true);
                                //HumiFluctStd[i].Text = fluct.HumiPvStd.ToString("F2");
                                HumiFluctStd[i].Text = JudgeValueStr(fluct.HumiPvStd, 3, 1, true);
                                if (fluct.HumiPvStd <= 3) { HumiFluctJdge[i].Text = "GOOD"; HumiFluctJdge[i].BackColor = Color.Lime; }
                            }

                        }
                    }
                }
                // 表示内容をメモリに保存
                ExcelValue.lblTempReachMemL = lblTempReachMemL.Text;
                ExcelValue.lblTempReachMemH = lblTempReachMemH.Text;
                ExcelValue.lblTempReachJdgeL = lblTempReachJdgeL.Text;
                ExcelValue.lblTempReachJdgeH = lblTempReachJdgeH.Text;

                ExcelValue.lblTempOpeMemL = lblTempOpeMemL.Text;
                ExcelValue.lblTempOpeMemH = lblTempOpeMemH.Text;
                ExcelValue.lblTempOpeJdgeL = lblTempOpeJdgeL.Text;
                ExcelValue.lblTempOpeJdgeH = lblTempOpeJdgeH.Text;

                ExcelValue.lblHumiOpeHumi3 = HumiOpeHumi[2].Text;
                ExcelValue.lblHumiOpeHumi4 = HumiOpeHumi[3].Text;

                ExcelValue.lblChgUpStart = lblChgUpStart.Text;
                ExcelValue.lblChgUpEnd = lblChgUpEnd.Text;
                ExcelValue.lblChgUpSpec = lblChgUpSpec.Text;
                ExcelValue.lblChgUpStr = lblChgUpStr.Text;
                ExcelValue.lblChgUpValue = lblChgUpValue.Text;
                ExcelValue.lblChgUpTime = lblChgUpTime.Text;

                ExcelValue.lblPoleUpStart = lblPoleUpStart.Text;
                ExcelValue.lblPoleUpEnd = lblPoleUpEnd.Text;
                ExcelValue.lblPoleUpSpec = lblPoleUpSpec.Text;
                ExcelValue.lblPoleUpStr = lblPoleUpStr.Text;
                ExcelValue.lblPoleUpValue = lblPoleUpValue.Text;

                ExcelValue.lblChgDownStart = lblChgDownStart.Text;
                ExcelValue.lblChgDownEnd = lblChgDownEnd.Text;
                ExcelValue.lblChgDownSpec = lblChgDownSpec.Text;
                ExcelValue.lblChgDownStr = lblChgDownStr.Text;
                ExcelValue.lblChgDownValue = lblChgDownValue.Text;
                ExcelValue.lblChgDownTime = lblChgDownTime.Text;

                ExcelValue.lblPoleDownStart = lblPoleDownStart.Text;
                ExcelValue.lblPoleDownEnd = lblPoleDownEnd.Text;
                ExcelValue.lblPoleDownSpec = lblPoleDownSpec.Text;
                ExcelValue.lblPoleDownStr = lblPoleDownStr.Text;
                ExcelValue.lblPoleDownValue = lblPoleDownValue.Text;

                ExcelValue.lblVarTemp1Start = lblVarTemp1Start.Text;
                ExcelValue.lblVarTemp1End = lblVarTemp1End.Text;
                ExcelValue.lblVarTemp1Spec = lblVarTemp1Spec.Text;

                ExcelValue.lblTempFluctMemM = lblTempFluctMemM.Text;
                ExcelValue.lblTempFluctMemL = lblTempFluctMemL.Text;

                ExcelValue.lblTempFluctJdgeL = lblTempFluctJdgeL.Text;
                ExcelValue.lblTempFluctJdgeM = lblTempFluctJdgeM.Text;

                ExcelValue.lblVarTemp2Start = lblVarTemp2Start.Text;
                ExcelValue.lblVarTemp2End = lblVarTemp2End.Text;
                ExcelValue.lblVarTemp2Spec = lblVarTemp2Spec.Text;
                ExcelValue.lblTempFluctMemH = lblTempFluctMemH.Text;
                ExcelValue.lblTempFluctJdgeH = lblTempFluctJdgeH.Text;

                ExcelValue.lblHumiFluctStd1 = lblHumiFluctStd1.Text;
                ExcelValue.lblHumiFluctStd2 = lblHumiFluctStd2.Text;
                ExcelValue.lblHumiFluctStd3 = lblHumiFluctStd3.Text;
                ExcelValue.lblHumiFluctStd4 = lblHumiFluctStd4.Text;

                ExcelValue.lblVarHumiSpec = lblVarHumiSpec.Text;

                ExcelValue.lblHumiFluctJdge1 = lblHumiFluctJdge1.Text;
                ExcelValue.lblHumiFluctJdge2 = lblHumiFluctJdge2.Text;
                ExcelValue.lblHumiFluctJdge3 = lblHumiFluctJdge3.Text;
                ExcelValue.lblHumiFluctJdge4 = lblHumiFluctJdge4.Text;
            }
        }
        private void TestDataCheck() 
        {
            // 温度変化速度降下
            //ChgDownValue.StartXPos = 96;
            //ChgDownValue.StartYValue = 131.3f;
            //ChgDownValue.EndXPos = 149;
            //ChgDownValue.EndYValue = -21.6f;
            ChgDownValue.ResultValue = Math.Abs((ChgDownValue.EndYValue - ChgDownValue.StartYValue) / (ChgDownValue.EndXPos - ChgDownValue.StartXPos));

            // 温度極値到達時間降下
            //ExtDownValue.StartXPos = 315;
            //ExtDownValue.StartYValue = 20.1f;
            //ExtDownValue.EndXPos = 346;
            //ExtDownValue.EndYValue = -40.5f;
            ExtDownValue.ResultValue = ExtDownValue.EndXPos - ExtDownValue.StartXPos;

            // 温度変化速度上昇
            //ChgUpValue.StartXPos = 466;
            //ChgUpValue.StartYValue = -25.7f;
            //ChgUpValue.EndXPos = 496;
            //ChgUpValue.EndYValue = 132.1f;
            ChgUpValue.ResultValue = Math.Abs((ChgUpValue.EndYValue - ChgUpValue.StartYValue) / (ChgUpValue.EndXPos - ChgUpValue.StartXPos));

            // 温度極値到達時間上昇
            //ExtUpValue.StartXPos = 464;
            //ExtUpValue.StartYValue = -39.9f;
            //ExtUpValue.EndXPos = 502;
            //ExtUpValue.EndYValue = 152.4f;
            ExtUpValue.ResultValue = ExtUpValue.EndXPos - ExtUpValue.StartXPos;

        }
        private void ModelSelInitial(int iNo) 
        {
            // 検査実施日・記録計情報・ＣＳＶファイル・エクセルデータファイル名クリア
            lblTestImplementationDate.Text = "";
            lblRecorderNo.Text = "";
            lblRecorderModel.Text = "";
            lblRecorderSerialNo.Text = "";
            lblCsvFileName.Text = "";
            lblCreateExcelFileName.Text = "";
            // 製品仕様を初期化
            ChkNowSpec.MashineNo = iNo;
            ModelSpecSet(iNo);

            // フォームタイトル変更
            string Title = "小型環境試験器　標準検査データ";
            this.Text = Title;
            if (ChkNowSpec.HumiUse) this.Text += "（湿度タイプ）";
            else this.Text += "（温度タイプ）";

            this.Text += " Ver" + AppVer;

            // 異常時の背景色をクリア
            txtMessage.BackColor = Color.White;

            // ボタンを有効/無効設定
            btnCSVFileSelect.Enabled = true;
            btnStart.Enabled = false;
            btnExcelSave.Enabled = false;
        }
        private void cmbModel_SelectedIndexChanged(object sender, EventArgs e)
        {
            ModelSelInitial(cmbModel.SelectedIndex);
        }
        private void Form1_Resize(object sender, EventArgs e)
        {
            if(this.Width<1920 || this.Height < 1050) 
            { 
                this.Width = 1920; this.Height = 1050;
                tabControl1.Width = 1891; tabControl1.Height = 1000;
            }
            else 
            {
                // Graph 拡大
                formsPlot1.Width = (this.Width - 1920) + 1871;
                formsPlot1.Height = (this.Height - 1050) + 956;

                tabControl1.Width = (this.Width - 1920) + 1891;
                tabControl1.Height = (this.Height - 1050) + 1000;

            }

        }
        private void btnStart_Click_1(object sender, EventArgs e)
        {
            String cMsg = ""; String aMsg = "";
            // 入力値確認
            cMsg = ChkInput();
            if (cMsg == "")
            {
                SetExcellFileName();

                DialogResult result = DialogResult.No;
                aMsg += "ユーザー:" + UserName + "\r\n";
                aMsg += "製造番号:" + SerialNo + "\r\n";
                aMsg += "装置型式:" + ModelName + "\r\n";
                aMsg += "検査プログラム名:" + ChkNowSpec.InspectPrgName + "\r\n";
                aMsg += "ＣＳＶファイル名:" + lblCsvFileName.Text + "\r\n\r\n";

                aMsg += "付加仕様など、黄色背景色部分の設定は正しいですか？\r\n\r\n";
                aMsg += "上記のデータ解析を開始します。宜しいですか？";
                result = MessageBox.Show(aMsg, "解析開始確認警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                if (result == DialogResult.Yes)
                {
                    // 製品仕様を初期化
                    // ModelSpecSet(cmbModel.SelectedIndex);

                    // 担当者名登録
                    TestingPerson = cmbMemberName.Text;
                    ExcelValue.TestingPerson = TestingPerson;

                    // 英文登録
                    foreach (MEMBERS EnName in Member.Member) 
                    {
                        if(EnName.Name== TestingPerson)
                        {
                            TestingPersonEn = EnName.EnName;
                            ExcelValue.TestingPersonEn = TestingPersonEn;
                            break;
                        }
                    }

                    // 検査電源情報
                    TestingVolt = cmbTestVolt.Text;
                    ExcelValue.TestingVolt = TestingVolt;
                    TestingFreq = cmbTestFreq.Text;
                    ExcelValue.TestingFreq = TestingFreq;

                    // 観測窓設定
                    CheckWindow = false;
                    if (ChkNowSpec.WindowUse) CheckWindow = true;
                    ExcelValue.CheckWindow = CheckWindow;

                    // 内扉設定
                    InnerDoor = false;
                    if (ChkNowSpec.IndoorUse) InnerDoor = true;
                    ExcelValue.InnerDoor = InnerDoor;

                    // 操作孔（側面）設定
                    //OperationHole = false;
                    //if (cmbOperationHole.Text == "操作孔有") OperationHole = true;
                    //ExcelValue.OperationHole = OperationHole;

                    // 付加仕様記載内容読み込み
                    AdditionalSpecifications = txtAdditionalSpec.Text;
                    ExcelValue.AdditionalSpecifications = AdditionalSpecifications;

                    // データを初期化
                    //if (TempFluct != null) TempFluct.Clear();
                    //if (HumiFluct != null) HumiFluct.Clear();

                    // 初期化
                    txtMessage.BackColor = Color.White;

                    // CSVファイル読込み
                    if (NowCSVFilePath != "")
                    {
                        //RecLength = CsvRead(NowCSVFilePath);
                        // CsvRead();

                        ChgRate ChgUpValue = new ChgRate();
                        ChgRate ChgDownValue = new ChgRate();
                        ChgRate ExtUpValue = new ChgRate();
                        ChgRate ExtDownValue = new ChgRate();

                        tCSVReadRun = true;
                        tCSVReadEnd = false;
                        tCSVReadStop = false;
                        tCSVMsg = "";
                        RecLength = 0;

                        // 自動判定CSVファイル削除
                        string path = System.AppDomain.CurrentDomain.BaseDirectory;
                        path += "CSV";
                        String FName = path + "\\" + NowExcelFileName + "_自動判定ログ.CSV";
                        if (System.IO.File.Exists(FName)) System.IO.File.Delete(FName);

                        // ボタンを有効/無効設定
                        btnCSVFileSelect.Enabled = false;
                        btnStart.Enabled = false;
                        btnExcelSave.Enabled = false;

                        Thread thread = new Thread(new ThreadStart(CsvRead));
                        thread.Start();
                    }
                    else
                    {
                        // CSVファイルが選択されておりません。
                        cMsg = "CSVファイルが選択されていません。\r\n";
                        cMsg += "CSVファイルを選択してください。\r\n";
                    }
                }
            }
            else
            {
                // 入力異常をメッセージウィンドウで表示する
                //cMsg = "設定項目に入力異常がありました。\r\n";
                cMsg += "設定項目を確認してください。\r\n";
            }
            if (cMsg != "")
            {
                // 異常処理
                // 異常時の背景色を赤色
                //txtMessage.BackColor = Color.Red;
                TextMsgWrite(cMsg);

                WriteLog(cMsg);
                MessageBox.Show(cMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // ボタンを有効/無効設定
                //btnCSVFileSelect.Enabled = true;
                //btnStart.Enabled = false;
                //btnExcelSave.Enabled = false;
            }
        }
        private void SetExcellFileName() 
        {
            // エクセルファイル名設定
            //SerialNo = lblSerial_number.Text;
            UserName = txtUserName.Text;
            ExcelValue.UserName = UserName;

            NowExcelFileName = "検査データ_" + SerialNo + "_" + UserName + "_" + ModelName;
            ExcelValue.NowExcelFileName = NowExcelFileName;

            lblCreateExcelFileName.Text = NowExcelFileName;
            lblCreateExcelFileName.Refresh();
        }
        private void btnCSVFileSelect_Click(object sender, EventArgs e)
        {
            SetExcellFileName();
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            string drv = txtReadDrive.Text;
            if (drv != "D" && drv != "E" && drv != "f" && drv != "G")
            {
                // ドライブ指定エラー
                WriteLog("ＣＦカード読み取りドライブ指定エラー" + "\r\n");
                TextMsgWrite("ＣＦカード読み取りドライブ指定エラー" + "\r\n");
            }
            else
            {
                lblCsvFileName.Text = ""; NowCSVFilePath = "";
                drv += ":\\";
                //OpenFileDialogクラスのインスタンスを作成
                OpenFileDialog ofd = new OpenFileDialog();

                //はじめのファイル名を指定する
                //はじめに「ファイル名」で表示される文字列を指定する
                ofd.FileName = "default.csv";
                //はじめに表示されるフォルダを指定する
                //指定しない（空の文字列）の時は、現在のディレクトリが表示される
                ofd.InitialDirectory = @drv;
                //[ファイルの種類]に表示される選択肢を指定する
                //指定しないとすべてのファイルが表示される
                ofd.Filter = "CSVファイル(*.CSV;*.csv)|*.csv;*.csv|すべてのファイル(*.*)|*.*";
                //[ファイルの種類]ではじめに選択されるものを指定する
                //2番目の「すべてのファイル」が選択されているようにする
                ofd.FilterIndex = 1;
                //タイトルを設定する
                ofd.Title = "開くファイルを選択してください";
                //ダイアログボックスを閉じる前に現在のディレクトリを復元するようにする
                ofd.RestoreDirectory = true;
                //存在しないファイルの名前が指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                ofd.CheckFileExists = true;
                //存在しないパスが指定されたとき警告を表示する
                //デフォルトでTrueなので指定する必要はない
                ofd.CheckPathExists = true;

                //ダイアログを表示する
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    //OKボタンがクリックされたとき、選択されたファイル名を表示する
                    Console.WriteLine(ofd.FileName);
                    NowCSVFilePath = ofd.FileName;
                    ExcelValue.NowCSVFilePath = NowCSVFilePath;

                    WriteLog(ofd.FileName + "\r\n");
                    TextMsgWrite(ofd.FileName + "\r\n");

                    string fName = Path.GetFileName(ofd.FileName);
                    lblCsvFileName.Text = fName.Substring(0, fName.Length - 4);
                    ExcelValue.lblCsvFileName = lblCsvFileName.Text;

                    // ファイル名から記録計情報を取得
                    // 記録計情報
                    bool bExist = false; string wMsg = "";
                    // 記録計品番の後半7文字が記録されている
                    string IdNo = fName.Substring(0, 7);
                    // 製品の製造番号の後半8文字が記録されている
                    string Rec_SerialNo = fName.Substring(8, 8);
                    TestingRecorderNo = "";
                    TestingRecorderModel = "";
                    TestingRecorderSerialNo = "";
                    foreach (Json.RECORDERS Rec in Recorders.Recorder)
                    {
                        if (Rec.No.Contains(IdNo))
                        {
                            TestingRecorderNo = Rec.No;
                            ExcelValue.TestingRecorderNo = TestingRecorderNo;
                            TestingRecorderModel = Rec.Model;
                            ExcelValue.TestingRecorderModel = TestingRecorderModel;
                            TestingRecorderSerialNo = Rec.SerialNo;
                            ExcelValue.TestingRecorderSerialNo = TestingRecorderSerialNo;
                            bExist = true;
                            break;
                        }
                    }
                    //]

                    if (bExist)
                    {
                        // 記録計品番が存在した。
                        lblRecorderNo.Text = TestingRecorderNo;
                        lblRecorderModel.Text = TestingRecorderModel;
                        lblRecorderSerialNo.Text = TestingRecorderSerialNo;

                        if (Rec_SerialNo == SerialNo || !SystemSet.SerialNoCheck) 
                        { 
                            // BarCodeReader読み込みの製造番号と一致した。

                            // ボタンを有効/無効設定
                            btnCSVFileSelect.Enabled = false;
                            btnStart.Enabled = true;
                            btnExcelSave.Enabled = false;
                        }
                        else 
                        {
#if true
                            wMsg = "ＣＳＶファイル名に製造番号が含まれていません。\r\n";
                            wMsg += "記録計の品番の後に製造番号が必要です。\r\n";
                            wMsg += "ＣＳＶファイル名を確認して下さい。\r\n";
                            MessageBox.Show(wMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            // ボタンを有効/無効設定
                            btnCSVFileSelect.Enabled = true;
                            btnStart.Enabled = false;
                            btnExcelSave.Enabled = false;
#else
                            // ボタンを有効/無効設定
                            btnCSVFileSelect.Enabled = false;
                            btnStart.Enabled = true;
                            btnExcelSave.Enabled = false;

#endif
                        }
                    }
                    else 
                    {
                        wMsg = "ＣＳＶファイル名に記録計の品番が含まれていません。\r\n";
                        wMsg += "記録計の品番が、ＣＳＶファイル名の先頭に必要です。\r\n";
                        wMsg += "ＣＳＶファイル名を確認して下さい。\r\n";
                        MessageBox.Show(wMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    if (wMsg!="") 
                    {
                        // 記録計品番が存在しなかった。
                        lblRecorderNo.Text = "";
                        lblRecorderModel.Text = "";
                        lblRecorderSerialNo.Text = "";
                        // ボタンを有効/無効設定
                        //btnCSVFileSelect.Enabled = false;
                        btnStart.Enabled = false;
                        btnExcelSave.Enabled = false;
                    }
                }

            }
        }
        private void btnExcelSave_Click(object sender, EventArgs e)
        {
            bool bRet = false; double dValue = 0;

            // ボタンを有効/無効設定
            btnCSVFileSelect.Enabled = false;
            btnStart.Enabled = false;
            btnExcelSave.Enabled = false;
            btnExcelSave.Refresh();
            ExcelValue.formsPlot1 = formsPlot1;
            if (SystemSet.ClosedXMLUsed)
            {
                if (!bExcelWriteBusy)
                {
                    bExcelWriteBusy = true; bExcelWriteEnd = false; bExcelWriteFail = false;
                    //Task<bool> bReturn = Excel_Operation.ExcelFileWrite(SystemSet, ExcelValue, RecValue, CurrentInfo);
                    //if(bReturn.Result)
                    //{
                    //    // エクセル処理異常
                    //    txtMessage.BackColor = Color.Red;
                    //}
                    //Thread.CurrentThread.SetApartmentState(ApartmentState.Unknown);
                    //Thread.CurrentThread.SetApartmentState(ApartmentState.STA);
                    Thread thread1 = new Thread(() => Excel_Operation.ExcelFileWrite(SystemSet, ExcelValue, RecValue, CurrentInfo));
                    thread1.Start();
                }
                //var response = await Excel_Operation.ExcelFileWrite(SystemSet, ExcelValue, RecValue);

                //if (!response) 
                //{
                    // エクセル処理異常
                //    txtMessage.BackColor = Color.Red;
                //}
            }
            else
            {
                // 起動されているExcelのプロセスID記憶
                int[] ExcelPid = new int[20]; int iID = 1;
                System.Diagnostics.Process[] prcs = System.Diagnostics.Process.GetProcessesByName("excel");

                if (prcs != null)
                {
                    foreach (System.Diagnostics.Process p in prcs)
                    {
                        ExcelPid[iID] = p.Id;
                        Console.WriteLine("Excel Process ID:" + iID.ToString() + " PID = " + ExcelPid[iID].ToString());
                        ++iID;
                    }
                    Array.Resize(ref ExcelPid, iID);
                    // エクセル起動数設定
                    ExcelPid[0] = iID - 1;
                }
                else
                {
                    Array.Resize(ref ExcelPid, 1);
                    // 起動無し設定
                    ExcelPid[0] = 0;
                }

                // Excel.Application の新しいインスタンスを生成する
                var xlApp = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbooks xlBooks = xlApp.Workbooks; ;
                Microsoft.Office.Interop.Excel.Workbook ExcelWB;
                Microsoft.Office.Interop.Excel.Worksheet ExcelWS;
                string cMsg = "";

                String DocPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);

                //String DistFileName = Directory.GetCurrentDirectory() + "\\Document\\" + "一時ファイル.xlsx";
                //String DistFileName = "D:\\一時ファイル.xlsx";
                //String DistFileName = MasterFilePath + "\\" + "一時ファイル.xlsx";
                String DistFileName = @DocPath + "\\" + "一時ファイル.xls";

                if (System.IO.File.Exists(DistFileName)) System.IO.File.Delete(DistFileName);

                //String cPath = Directory.GetCurrentDirectory() + "\\Document\\";
                String cPath = SystemSet.MasterPath;

                String SourceFileName = cPath;

                bool bGrapfCopyErr = false;

                if (ChkNowSpec.HumiUse) SourceFileName += "検査データ_湿度用.xls";
                else SourceFileName += "検査データ_温度用.xls";
                try
                {
                    if (!System.IO.File.Exists(DistFileName))
                    {
                        // ファイル複写
                        System.IO.File.Copy(SourceFileName, DistFileName);
                        //FileInfoオブジェクトを作成
                        System.IO.FileInfo fi = new System.IO.FileInfo(DistFileName);
                        //隠し属性があるか調べる
                        if ((fi.Attributes & System.IO.FileAttributes.Hidden) ==
                            System.IO.FileAttributes.Hidden)
                        {
                            Console.WriteLine("隠し属性があります。");
                            //隠し属性を削除する
                            fi.Attributes &= ~System.IO.FileAttributes.Hidden;
                        }
                    }
                    // 解析結果データ準備
                    cMsg = "エクセル起動\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);
                    // xlApplication から WorkBooks を取得する
                    // 既存の Excel ブックを開く
                    ExcelWB = xlBooks.Open(DistFileName);
                    ExcelWS = (Worksheet)ExcelWB.Sheets["測定データ"];
                    ExcelWS.Activate();
                    // Excel を表示する
                    //xlApp.Visible = true;
                    #region 検査結果書込み

                    // 解析結果データ準備
                    cMsg = "解析結果データ準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    // ファイル情報
                    ExcelWS.Range["B27"].Value = HumiEndRecord;
                    ExcelWS.Range["B28"].Value = TempEndRecord;
                    ExcelWS.Range["B29"].Value = NowCSVFilePath;
                    ExcelWS.Range["J28"].Value = SystemSet.ExcelPath;
                    ExcelWS.Range["J29"].Value = NowExcelFileName;

                    // 装置情報
                    ExcelWS.Range["B3"].Value = ModelName;
                    ExcelWS.Range["B4"].Value = SerialNo;
                    ExcelWS.Range["B5"].Value = UserName;

                    ExcelWS.Range["G27"].Value = ChkNowSpec.SpecMinTemp;
                    ExcelWS.Range["G28"].Value = ChkNowSpec.SpecMaxTemp;

                    // オプション情報
                    ExcelWS.Range["B6"].Value = "なし";
                    if (CheckWindow) ExcelWS.Range["B6"].Value = "あり";
                    ExcelWS.Range["B7"].Value = "なし";
                    if (InnerDoor) ExcelWS.Range["B7"].Value = "あり";
                    ExcelWS.Range["B8"].Value = "なし";
                    if (OperationHole) ExcelWS.Range["B8"].Value = "あり";
                    ExcelWS.Range["B21"].Value = "なし";
                    if (ChkNowSpec.Temp180Use) ExcelWS.Range["B21"].Value = "あり";

                    // 検査プログラム名
                    ExcelWS.Range["B22"].Value = ChkNowSpec.InspectPrgName;

                    // 検査チェックシート名
                    ExcelWS.Range["B24"].Value = ChkNowSpec.CheckSheetName;

                    // 2024/11/06　改善の為、解除
                    // 2024/08/22　改善の為、一時保留
                    // 平均室温書出し
                    ExcelWS.Range["B23"].Value = RoomTempAverage.ToString("F1");
                    // 2024/08/22　改善の為、一時保留
                    // 2024/11/06　改善の為、解除

                    // 付加仕様情報
                    ExcelWS.Range["B9"].Value = AdditionalSpecifications;
                    // ケーブル孔仕様情報
                    ExcelWS.Range["B10"].Value = CableHole[0];
                    ExcelWS.Range["B11"].Value = CableHole[1];
                    ExcelWS.Range["B12"].Value = CableHole[2];
                    ExcelWS.Range["B13"].Value = CableHole[3];
                    // 検査状況情報
                    ExcelWS.Range["B14"].Value = TestingPerson;
                    ExcelWS.Range["B25"].Value = TestingPersonEn;

                    ExcelWS.Range["B15"].Value = TestingDate;
                    ExcelWS.Range["B16"].Value = TestingVolt;
                    ExcelWS.Range["B17"].Value = TestingFreq;
                    // 記録計情報
                    ExcelWS.Range["B18"].Value = TestingRecorderNo;
                    ExcelWS.Range["B19"].Value = TestingRecorderModel;
                    ExcelWS.Range["B20"].Value = TestingRecorderSerialNo;

                    // 解析結果データ準備
                    cMsg = "自動判定条件準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    // 自動判定条件

                    ExcelWS.Range["W2"].Value = JdgCnd.ChgJudgeTemp.ToString("F1");

                    ExcelWS.Range["W3"].Value = JdgCnd.MaxMinWidth.ToString("F1");
                    ExcelWS.Range["W4"].Value = JdgCnd.StableCount.ToString("");
                    ExcelWS.Range["W5"].Value = JdgCnd.ChgToleranceLimit.ToString("");
                    ExcelWS.Range["W6"].Value = JdgCnd.JdgTimes.ToString("");
                    ExcelWS.Range["W7"].Value = JdgCnd.JdgAfterTime.ToString("");

                    // 製品仕様（極値到達時間上昇）
                    ExcelWS.Range["E3"].Value = ChkNowSpec.PoleUpSpec.ToString("");
                    ExcelWS.Range["F3"].Value = ChkNowSpec.PoleUpStartTemp.ToString("F1");
                    ExcelWS.Range["G3"].Value = ChkNowSpec.PoleUpEndTemp.ToString("F1");
                    // 製品仕様（極値到達時間降下）
                    ExcelWS.Range["E4"].Value = ChkNowSpec.PoleDownSpec.ToString("");
                    ExcelWS.Range["F4"].Value = ChkNowSpec.PoleDownStartTemp.ToString("F1");
                    ExcelWS.Range["G4"].Value = ChkNowSpec.PoleDownEndTemp.ToString("F1");
                    // 製品仕様（変化速度上昇）
                    ExcelWS.Range["E5"].Value = ChkNowSpec.ChgUpSpec.ToString("F1");
                    ExcelWS.Range["F5"].Value = ChkNowSpec.ChgUpStartTemp.ToString("F1");
                    ExcelWS.Range["G5"].Value = ChkNowSpec.ChgUpEndTemp.ToString("F1");
                    // 製品仕様（変化速度降下）
                    ExcelWS.Range["E6"].Value = ChkNowSpec.ChgDownSpec.ToString("F1");
                    ExcelWS.Range["F6"].Value = ChkNowSpec.ChgDownStartTemp.ToString("F1");
                    ExcelWS.Range["G6"].Value = ChkNowSpec.ChgDownEndTemp.ToString("F1");
                    // 温度変動低温製品仕様
                    ExcelWS.Range["E7"].Value = ChkNowSpec.VarTemp1Spec.ToString("F1");
                    ExcelWS.Range["F7"].Value = ChkNowSpec.VarTemp1StartTemp.ToString("F1");
                    ExcelWS.Range["G7"].Value = ChkNowSpec.VarTemp1EndTemp.ToString("F1");
                    // 温度変動高温製品仕様
                    ExcelWS.Range["E8"].Value = ChkNowSpec.VarTemp2Spec.ToString("F1");
                    ExcelWS.Range["F8"].Value = ChkNowSpec.VarTemp2StartTemp.ToString("F1");
                    ExcelWS.Range["G8"].Value = ChkNowSpec.VarTemp2EndTemp.ToString("F1");
                    // 湿度変動製品仕様
                    ExcelWS.Range["E9"].Value = ChkNowSpec.VarHumiSpec.ToString("F1");

                    // 検査結果（極値到達時間上昇）
                    ExcelWS.Range["E12"].Value = ExtUpValue.ResultValue.ToString("");
                    ExcelWS.Range["G12"].Value = "NG";
                    if (ChkNowSpec.PoleUpSpec >= ExtUpValue.ResultValue) ExcelWS.Range["G12"].Value = "GOOD";
                    else { if (Math.Abs(ExtUpValue.ResultValue) > 1000) ExcelWS.Range["E12"].Value = "--"; }

                    // 検査結果（極値到達時間降下）
                    ExcelWS.Range["E13"].Value = ExtDownValue.ResultValue.ToString("");
                    ExcelWS.Range["G13"].Value = "NG";
                    if (ChkNowSpec.PoleDownSpec >= ExtDownValue.ResultValue) ExcelWS.Range["G13"].Value = "GOOD";
                    else { if (Math.Abs(ExtDownValue.ResultValue) > 1000) ExcelWS.Range["E13"].Value = "--"; }

                    // 検査結果（変化速度上昇）
                    ExcelWS.Range["E14"].Value = (ChgUpValue.EndXPos - ChgUpValue.StartXPos).ToString("");
                    ExcelWS.Range["F14"].Value = ChgUpValue.ResultValue.ToString("F1");
                    ExcelWS.Range["G14"].Value = "NG";
                    if (ChkNowSpec.ChgUpSpec <= ChgUpValue.ResultValue) ExcelWS.Range["G14"].Value = "GOOD";
                    else { if (Math.Abs(ChgUpValue.EndXPos - ChgUpValue.StartXPos) > 1000) ExcelWS.Range["E14"].Value = "---"; }

                    // 検査結果（変化速度降下）
                    ExcelWS.Range["E15"].Value = (ChgDownValue.EndXPos - ChgDownValue.StartXPos).ToString("");
                    ExcelWS.Range["F15"].Value = ChgDownValue.ResultValue.ToString("F1");
                    ExcelWS.Range["G15"].Value = "NG";
                    if (ChkNowSpec.ChgDownSpec <= ChgDownValue.ResultValue) ExcelWS.Range["G15"].Value = "GOOD";
                    else { if (Math.Abs(ChgDownValue.EndXPos - ChgDownValue.StartXPos) > 1000) ExcelWS.Range["E15"].Value = "---"; }

                    // 解析結果データ準備
                    cMsg = "検査パターン準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    // 検査パターン
                    for (int i = 0; i < TEMPSETSTEPS; ++i)
                    {
                        ExcelWS.Range["J" + (i + 4).ToString()].Value = Tempset[i].SetTemp.ToString("F1");
                        ExcelWS.Range["K" + (i + 4).ToString()].Value = Tempset[i].SetTime.ToString("F1");
                    }
                    if (ChkNowSpec.HumiUse)
                    {
                        for (int i = 0; i < HUMISETSTEPS; ++i)
                        {
                            ExcelWS.Range["J" + (i + 12).ToString()].Value = Humiset[i].SetTemp.ToString("F1");
                            ExcelWS.Range["k" + (i + 12).ToString()].Value = Humiset[i].SetHumi.ToString("");
                            ExcelWS.Range["L" + (i + 12).ToString()].Value = Humiset[i].SetTime.ToString("F1");
                        }
                    }

                    // 温度精度・変動幅
                    ExcelWS.Range["N4"].Value = ChkNowSpec.AccuracyTemp1StartTemp.ToString("F1");
                    ExcelWS.Range["O4"].Value = ChkNowSpec.AccuracyTemp1EndTemp.ToString("F1");
                    ExcelWS.Range["P4"].Value = "±" + ChkNowSpec.AccuracyTemp1Spec.ToString("F1");

                    ExcelWS.Range["R4"].Value = ChkNowSpec.VarTemp1StartTemp.ToString("F1");
                    ExcelWS.Range["S4"].Value = ChkNowSpec.VarTemp1EndTemp.ToString("F1");
                    ExcelWS.Range["T4"].Value = ChkNowSpec.VarTemp1Spec.ToString("F1");

                    ExcelWS.Range["N5"].Value = ChkNowSpec.AccuracyTemp2StartTemp.ToString("F1");
                    ExcelWS.Range["O5"].Value = ChkNowSpec.AccuracyTemp2EndTemp.ToString("F1");
                    ExcelWS.Range["P5"].Value = "±" + ChkNowSpec.AccuracyTemp2Spec.ToString("F1");

                    ExcelWS.Range["R5"].Value = ChkNowSpec.VarTemp2StartTemp.ToString("F1");
                    ExcelWS.Range["S5"].Value = ChkNowSpec.VarTemp2EndTemp.ToString("F1");
                    ExcelWS.Range["T5"].Value = ChkNowSpec.VarTemp2Spec.ToString("F1");

                    if (ChkNowSpec.Temp180Use)
                    {
                        ExcelWS.Range["N6"].Value = ChkNowSpec.AccuracyTemp3StartTemp.ToString("F1");
                        ExcelWS.Range["O6"].Value = ChkNowSpec.AccuracyTemp3EndTemp.ToString("F1");
                        ExcelWS.Range["P6"].Value = "±" + ChkNowSpec.AccuracyTemp3Spec.ToString("F1");

                        ExcelWS.Range["R6"].Value = ChkNowSpec.VarTemp3StartTemp.ToString("F1");
                        ExcelWS.Range["S6"].Value = ChkNowSpec.VarTemp3EndTemp.ToString("F1");
                        ExcelWS.Range["T6"].Value = ChkNowSpec.VarTemp3Spec.ToString("F1");
                    }

                    // 解析結果データ準備
                    cMsg = "検査結果準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    // 湿度精度・変動幅
                    if (ChkNowSpec.HumiUse)
                    {
                        for (int i = 0; i < HumiFluct.Count; ++i)
                        {
                            ExcelWS.Range["P" + (i + 12).ToString()].Value = "±" + Humiset[i].HumiAccuracy.ToString("F1");
                        }
                        ExcelWS.Range["Q12"].Value = "3.0 %RH";
                    }
                    ExcelWS.Range["I20"].Value = "+" + ChkNowSpec.SpecMaxTemp.ToString();
                    ExcelWS.Range["I22"].Value = ChkNowSpec.SpecMinTemp.ToString();
                    ExcelWS.Range["I25"].Value = (ChkNowSpec.SpecMinTemp - 3).ToString() + " ℃に到達している事";
                    ExcelWS.Range["I26"].Value = "+157 ℃に到達している事";
                    if (ChkNowSpec.Temp180Use) ExcelWS.Range["I26"].Value = "+185 ℃に到達している事";
                    ExcelWS.Range["N25"].Value = ChkNowSpec.SpecMinTemp.ToString() + " ℃に到達している事";

                    if ((FLUCT_USE && FIRST_STEP) || !FIRST_STEP)
                    {
                        // 検査結果（温度変動）
                        if (TempFluct.Count != 0)
                        {
                            float MaxValue = float.MinValue;
                            float MinSVValue = float.MaxValue; float MainDownSVValue = float.MaxValue;
                            float MaxSVValue = float.MinValue; float MaxUpSVValue = float.MinValue;
                            bool fNG = false; int i = 0;


                            float MaxUpTemp = 157; float MinDownTemp = ChkNowSpec.SpecMinTemp - 3;
                            if (ChkNowSpec.Temp180Use) MaxUpTemp = 185;

                            ExcelWS.Range["L20"].Value = "NG";
                            ExcelWS.Range["L21"].Value = "NG";
                            ExcelWS.Range["L22"].Value = "NG";
                            ExcelWS.Range["M25"].Value = "NG";
                            ExcelWS.Range["M26"].Value = "NG";
                            ExcelWS.Range["R25"].Value = "NG";

                            foreach (WIDTH_FLUCT fluct in TempFluct)
                            {
                                float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                                if (fluct.TempSvAve >= ChkNowSpec.SpecMaxTemp - TempWidth / 2 && fluct.TempSvAve <= ChkNowSpec.SpecMaxTemp + TempWidth / 2)
                                {
                                    // 最大温度の安定があった。
                                    MaxSVValue = ChkNowSpec.SpecMaxTemp;
                                    ExcelWS.Range["J20"].Value = fluct.TempSvAve.ToString("F1");
                                    ExcelWS.Range["K20"].Value = (fluct.TempSvAve - MaxSVValue).ToString("F1");
                                    if (Math.Abs(fluct.TempSvAve - MaxSVValue) < 1.2) ExcelWS.Range["L20"].Value = "GOOD";
                                }

                                else if (fluct.TempSvAve >= ChkNowSpec.SpecMinTemp - TempWidth && fluct.TempSvAve <= ChkNowSpec.SpecMinTemp + TempWidth)
                                {
                                    MinSVValue = ChkNowSpec.SpecMinTemp;
                                    // 最小温度の安定があった。
                                    ExcelWS.Range["J22"].Value = fluct.TempSvAve.ToString("F1");
                                    ExcelWS.Range["K22"].Value = (fluct.TempSvAve - MinSVValue).ToString("F1") + " ℃";
                                    if (Math.Abs(fluct.TempSvAve - MinSVValue) < 1.2) ExcelWS.Range["L22"].Value = "GOOD";

                                    // 槽中央温度で判定
                                    ExcelWS.Range["Q25"].Value = fluct.TempPvAve.ToString("F1");
                                    if (Math.Abs(fluct.TempPvAve - ChkNowSpec.SpecMinTemp) < TempWidth) ExcelWS.Range["R25"].Value = "GOOD";
                                }
                                else if (fluct.TempSvAve > 20 - TempWidth && fluct.TempSvAve < 20 + TempWidth)
                                {
                                    // 20℃の安定があった。
                                    ExcelWS.Range["J21"].Value = fluct.TempSvAve.ToString("F1");
                                    ExcelWS.Range["K21"].Value = (fluct.TempSvAve - 20).ToString("F1");
                                    if (Math.Abs(fluct.TempSvAve - 20) < 1.2) ExcelWS.Range["L21"].Value = "GOOD";
                                }

                                if (fluct.TempPvStd > MaxValue && fluct.Exist) MaxValue = fluct.TempPvStd;
                                if (!fluct.Judge && fluct.Exist) fNG = true;

                                // 取得データ書出し
                                if (fluct.Exist)
                                {
                                    int chkTemp = FluctSetTemp(fluct.EndValue);
                                    if (chkTemp == int.MinValue) chkTemp = (int)fluct.EndValue;
                                    ExcelWS.Range["D" + (19 + i).ToString()].Value = chkTemp.ToString("");
                                    ExcelWS.Range["E" + (19 + i).ToString()].Value = fluct.TempPvStd.ToString("F1");
                                }
                                ++i;
                            }
                            ExcelWS.Range["E16"].Value = MaxValue.ToString("F1");
                            ExcelWS.Range["F16"].Value = "NG";
                            if (!fNG) ExcelWS.Range["F16"].Value = "GOOD";

                            // 最高最高温度確認
                            ExcelWS.Range["M26"].Value = "NG";
                            if (MaxTemp.Exist)
                            {
                                ExcelWS.Range["L26"].Value = MaxTemp.TempSvAve.ToString("F1");
                                ExcelWS.Range["M26"].Value = "GOOD";
                            }

                            // 最低最低温度確認
                            ExcelWS.Range["M25"].Value = "NG";
                            if (MinTemp.Exist)
                            {
                                ExcelWS.Range["L25"].Value = MinTemp.TempSvAve.ToString("F1");
                                ExcelWS.Range["M25"].Value = "GOOD";
                            }
                        }

                        // 検査結果（湿度変動）
                        if (HumiFluct.Count != 0)
                        {
                            float MaxValue = float.MinValue;
                            bool fNG = false; int i = 0;
                            ExcelWS.Range["V20"].Value = "NG";
                            ExcelWS.Range["V22"].Value = "NG";
                            ExcelWS.Range["V24"].Value = "NG";
                            ExcelWS.Range["V26"].Value = "NG";
                            foreach (WIDTH_FLUCT fluct in HumiFluct)
                            {

                                float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                                for (int j = 0; j < 4; ++j)
                                {
                                    if (fluct.TempSvAve >= Humiset[j].SetTemp - TempWidth && fluct.TempSvAve <= Humiset[j].SetTemp + TempWidth &&
                                        fluct.HumiSvAve >= Humiset[j].SetHumi - TempWidth && fluct.HumiSvAve <= Humiset[j].SetHumi + TempWidth)
                                    {
                                        // 第Ｊステップの設定安定が合った
                                        ExcelWS.Range["U" + (j * 2 + 20).ToString()].Value = fluct.TempSvAve.ToString("F1");
                                        ExcelWS.Range["U" + (j * 2 + 21).ToString()].Value = fluct.HumiSvAve.ToString("F1");

                                        ExcelWS.Range["W" + (j * 2 + 20).ToString()].Value = fluct.TempPvAve.ToString("F1");
                                        ExcelWS.Range["W" + (j * 2 + 21).ToString()].Value = fluct.WetPvAve.ToString("F1");
                                        ExcelWS.Range["X" + (j * 2 + 21).ToString()].Value = fluct.HumiPvAve.ToString("F1");
                                        if (fluct.HumiSvAve > Humiset[j].SetHumi - 3 && fluct.HumiSvAve < Humiset[j].SetHumi + 3) ExcelWS.Range["V" + (j * 2 + 20).ToString()].Value = "GOOD";
                                        break;
                                    }
                                }
                                // 取得データ書出し
                                if (fluct.Exist)
                                {
                                    if (fluct.HumiPvStd > MaxValue && fluct.Exist) MaxValue = fluct.HumiPvStd;
                                    if (!fluct.Judge && fluct.Exist) fNG = true;
                                    // 取得データ書出し
                                    int chkHumi = (int)(fluct.EndValue);
                                    ExcelWS.Range["F" + (19 + i).ToString()].Value = chkHumi.ToString("");
                                    ExcelWS.Range["G" + (19 + i).ToString()].Value = fluct.HumiPvStd.ToString("F1");
                                }
                                ++i;
                            }
                            ExcelWS.Range["E17"].Value = MaxValue.ToString("F1");
                            ExcelWS.Range["F17"].Value = "NG";
                            if (!fNG) ExcelWS.Range["F17"].Value = "GOOD";
                        }
                    }

                    // 解析結果データ準備
                    cMsg = "確認グラフ画像複写準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    int iErrCnt = 0;
                    while (true)
                    {
                        try
                        {
                            // 確認グラフデータ複写
                            Bitmap bmpImage = formsPlot1.Plot.GetBitmap();
                            Clipboard.SetImage(bmpImage);
                            ExcelWS.Range["I31"].Select();
                            ExcelWS.Paste();
                            break;
                        }
                        catch (Exception ex)
                        {
                            ++iErrCnt;
                            string sMsg = "確認グラフデータ複写異常\r\n" + ex.Message;
                            WriteLog(cMsg);
                            TextMsgWrite(sMsg);
                            Console.WriteLine(iErrCnt.ToString() + ":" + sMsg);
                            if (iErrCnt > 3)
                            {
                                bGrapfCopyErr = true;

                                WriteLog(sMsg);
                                // 異常時の背景色を赤色
                                txtMessage.BackColor = Color.Red;
                                TextMsgWrite(sMsg);
                                Console.WriteLine(sMsg);
                                break;
                            }
                            Thread.Sleep(10);
                        }
                    }
                    #region グラフデータ
                    string SheetName = "";
                    // グラフデータ準備
                    cMsg = "グラフデータ準備\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);
                    xlApp.Run("CSVファイル読込");
                    #endregion グラフデータ

                    xlApp.DisplayAlerts = false;
                    cMsg = "検査データファイル書込み開始\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);
                    ExcelWB.Save();
                    //ExcelWB.SaveCopyAs(SystemSet.ExcelPath + NowExcelFileName + ".xls");
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWS);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWB);
                    #endregion 検査結果書込み

                    #region グラフファイル書込み
                    // グラフファイル準備
                    string GraphDistFileName = "";
                    string GraphDistPDFFileName = "";
                    string GraphSourceFileName = "";
                    string GraphSourcePDFFileName = "";
                    if (SystemSet.GraphRequired)
                    {
                        GraphDistFileName = @DocPath + "\\" + "Graph一時ファイル.xlsx";
                        GraphDistPDFFileName = @DocPath + "\\" + "Graph一時ファイル.pdf";

                        if (System.IO.File.Exists(GraphDistFileName)) System.IO.File.Delete(GraphDistFileName);
                        if (System.IO.File.Exists(GraphDistPDFFileName)) System.IO.File.Delete(GraphDistPDFFileName);
                        cPath = MasterPath + "\\グラフファイル.xlsx";

                        GraphSourceFileName = cPath;
                        SheetName = "グラフファイル";

                        // グラフファイル複写
                        System.IO.File.Copy(GraphSourceFileName, GraphDistFileName);

                        //FileInfoオブジェクトを作成
                        System.IO.FileInfo fi = new System.IO.FileInfo(GraphDistFileName);
                        //隠し属性があるか調べる
                        if ((fi.Attributes & System.IO.FileAttributes.Hidden) ==
                            System.IO.FileAttributes.Hidden)
                        {
                            Console.WriteLine("隠し属性があります。");
                            //隠し属性を削除する
                            fi.Attributes &= ~System.IO.FileAttributes.Hidden;
                        }

                        // xlApplication から WorkBooks を取得する
                        // 既存の Excel ブックを開く
                        ExcelWB = xlBooks.Open(GraphDistFileName);
                        ExcelWS = (Worksheet)ExcelWB.Sheets[SheetName];
                        ExcelWS.Activate();
                        // Excel を表示する

                        // ヘッダー情報
                        //ExcelWS.Range["H1"].Value = HeaderData.PuroductModel;
                        //ExcelWS.Range["J1"].Value = "温湿度確認試験";
                        //if(!ChkNowSpec.HumiUse) ExcelWS.Range["J1"].Value = "温度確認試験";
                        //ExcelWS.Range["Q1"].Value = HeaderData.CreateNo;
                        ExcelWS.Range["W29"].Value = ChkNowSpec.CheckSheetName;

                        // グラフ複写
                        iErrCnt = 0;
                        while (true)
                        {
                            try
                            {
                                // 確認グラフデータ複写
                                Bitmap bmpImage = formsPlot1.Plot.GetBitmap();
                                Clipboard.SetImage(bmpImage);
                                ExcelWS.Range["B2"].Select();
                                ExcelWS.Paste();

                                // 貼り付けられた画像を取得（Shapes の最後の要素が直近貼り付けた図形）
                                Excel.Shape pastedShape = ExcelWS.Shapes.Item(ExcelWS.Shapes.Count);

                                // 幅・高さを66%にスケーリング
                                pastedShape.ScaleWidth(0.67f, Microsoft.Office.Core.MsoTriState.msoTrue);
                                pastedShape.ScaleHeight(0.67f, Microsoft.Office.Core.MsoTriState.msoTrue);
                                break;
                            }
                            catch (Exception ex)
                            {
                                ++iErrCnt;
                                string sMsg = "確認グラフデータ複写異常：" + ex.Message + "\r\n";
                                WriteLog(sMsg);
                                TextMsgWrite(sMsg);
                                Console.WriteLine(iErrCnt.ToString() + ":" + sMsg);
                                if (iErrCnt > 3)
                                {
                                    bGrapfCopyErr = true;

                                    WriteLog(sMsg);
                                    // 異常時の背景色を赤色
                                    txtMessage.BackColor = Color.Red;
                                    TextMsgWrite(sMsg);
                                    Console.WriteLine(sMsg);
                                    break;
                                }
                                Thread.Sleep(10);
                            }
                        }
                        // Top表示
                        ExcelWS.Range["A1"].Activate();
                        cMsg = "グラフファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        TextMsgWrite(cMsg);
                        ExcelWB.Save();

                        // PDFファイル出力
                        ExcelWB.ExportAsFixedFormat(
                        Type: XlFixedFormatType.xlTypePDF,
                        Filename: GraphDistPDFFileName,
                        Quality: XlFixedFormatQuality.xlQualityStandard);

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWS);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWB);
                    }
                    #endregion グラフファイル書込み

                    #region 試験成績書処理
                    string TestDistFileName = "";
                    string TestSourceFileName = "";
                    if (SystemSet.TestReportRequired)
                    {
                        TestDistFileName = @DocPath + "\\" + "Test一時ファイル.xls";

                        if (System.IO.File.Exists(TestDistFileName)) System.IO.File.Delete(TestDistFileName);
                        cPath = SystemSet.MasterPath + "試験検査報告書.xls";

                        TestSourceFileName = cPath;
                        SheetName = "試験検査報告書";

                        // 試験検査報告書ファイル複写
                        System.IO.File.Copy(TestSourceFileName, TestDistFileName);

                        //FileInfoオブジェクトを作成
                        System.IO.FileInfo fi = new System.IO.FileInfo(TestDistFileName);
                        //隠し属性があるか調べる
                        if ((fi.Attributes & System.IO.FileAttributes.Hidden) ==
                            System.IO.FileAttributes.Hidden)
                        {
                            Console.WriteLine("隠し属性があります。");
                            //隠し属性を削除する
                            fi.Attributes &= ~System.IO.FileAttributes.Hidden;
                        }

                        // xlApplication から WorkBooks を取得する
                        // 既存の Excel ブックを開く
                        ExcelWB = xlBooks.Open(TestDistFileName);
                        ExcelWS = (Worksheet)ExcelWB.Sheets[SheetName];
                        ExcelWS.Activate();
                        // Excel を表示する

                        // ヘッダー情報
                        ExcelWS.Range["AA3"].Value = HeaderData.PuroductModel;
                        ExcelWS.Range["AA4"].Value = HeaderData.CreateNo;

                        string strTime1 = TestingDate + " 00:00:00";
                        string strTime = "2024/5/8 00:00:00";

                        DateTime dTime = DateTime.Parse(strTime);
                        ExcelWS.Range["AA5"].Value = dTime;

                        ExcelWS.Range["AA6"].Value = RoomTempAverage.ToString("F1");
                        ExcelWS.Range["AA7"].Value = HeaderData.VoltSpec.ToString();
                        ExcelWS.Range["AA8"].Value = HeaderData.FreqSpec.ToString();
                        ExcelWS.Range["AA9"].Value = TestingPersonEn;
                        ExcelWS.Range["AA10"].Value = TestingPerson;

                        // 湿度仕様有無
                        ExcelWS.Range["AE19"].Value = "FALSE";
                        if (ChkNowSpec.HumiUse) ExcelWS.Range["AE19"].Value = "TRUE";

                        // 観測窓仕様有無
                        ExcelWS.Range["AE7"].Value = "FALSE";
                        if (ChkNowSpec.WindowUse) ExcelWS.Range["AE7"].Value = "TRUE";

                        // 内扉仕様有無
                        ExcelWS.Range["AG7"].Value = "FALSE";
                        if (ChkNowSpec.IndoorUse) ExcelWS.Range["AG7"].Value = "TRUE";

                        // 側面操作孔仕様有無
                        ExcelWS.Range["AI7"].Value = "FALSE";
                        if (ChkNowSpec.OpeWindowUse) ExcelWS.Range["AI7"].Value = "TRUE";

                        // 温度範囲
                        ExcelWS.Range["AE21"].Value = ChkNowSpec.SpecMinTemp.ToString();
                        ExcelWS.Range["AF21"].Value = ChkNowSpec.SpecMaxTemp.ToString();

                        bRet = double.TryParse(lblTempReachMemL.Text, out dValue);
                        if (bRet) ExcelWS.Range["AG21"].Value = dValue.ToString("F1");

                        bRet = double.TryParse(lblTempReachMemH.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH21"].Value = dValue.ToString("F1");

                        ExcelWS.Range["AI21"].Value = "NG";
                        if (lblTempReachJdgeL.Text == "GOOD" && lblTempReachJdgeH.Text == "GOOD") ExcelWS.Range["AI21"].Value = "GOOD";

                        if (ChkNowSpec.HumiUse)
                        {
                            // 湿度範囲  ？？？？
                            //ExcelWS.Range["AE22"].Value = ChkNowSpec.SpecMinHumi.ToString();
                            //ExcelWS.Range["AF22"].Value = ChkNowSpec.SpecMaxHumi.ToString();
                            //ExcelWS.Range["AG22"].Value = lblHumiLowValue.Text;
                            //ExcelWS.Range["AH22"].Value = lblHumiHighValue.Text;
                            // 上限湿度：95％　下限湿度：30%　の測定値で判断か？
                            ExcelWS.Range["AE22"].Value = "30";
                            ExcelWS.Range["AF22"].Value = "95";

                            bRet = double.TryParse(lblHumiOpeHumi3.Text, out dValue);
                            if (bRet) ExcelWS.Range["AG22"].Value = dValue.ToString("F1");

                            bRet = double.TryParse(lblHumiOpeHumi4.Text, out dValue);
                            if (bRet) ExcelWS.Range["AH22"].Value = dValue.ToString("F1");

                            ExcelWS.Range["AI22"].Value = "GOOD";
                            float Humi30 = float.Parse(lblHumiOpeHumi3.Text);
                            float Humi95 = float.Parse(lblHumiOpeHumi4.Text);
                            if (Humi30 > 33 || Humi30 < 27 || Humi95 > 98 || Humi95 < 92) ExcelWS.Range["AI22"].Value = "NG";
                        }
                        // 温度変化上昇
                        //if (ChkNowSpec.ChgUpUse)
                        //{
                        ExcelWS.Range["AE29"].Value = lblChgUpStart.Text;
                        ExcelWS.Range["AF29"].Value = lblChgUpEnd.Text;
                        ExcelWS.Range["AG29"].Value = lblChgUpSpec.Text;
                        bRet = double.TryParse(lblChgUpValue.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH29"].Value = dValue.ToString("F1");
                        ExcelWS.Range["AI29"].Value = lblChgUpStr.Text;
                        //}
                        //else ExcelWS.Range["AJ29"].Value = "FALSE";

                        // 温度到達上昇
                        //if (ChkNowSpec.PoleUpUse)
                        //{
                        ExcelWS.Range["AE30"].Value = lblPoleUpStart.Text;
                        ExcelWS.Range["AF30"].Value = lblPoleUpEnd.Text;
                        ExcelWS.Range["AG30"].Value = lblPoleUpSpec.Text;
                        bRet = double.TryParse(lblPoleUpValue.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH30"].Value = dValue.ToString("F1");
                        ExcelWS.Range["AI30"].Value = lblPoleUpStr.Text;
                        //}
                        //else ExcelWS.Range["AJ30"].Value = "FALSE";

                        // 温度変化降下
                        //if (ChkNowSpec.ChgDownUse)
                        //{
                        ExcelWS.Range["AE31"].Value = lblChgDownStart.Text;
                        ExcelWS.Range["AF31"].Value = lblChgDownEnd.Text;
                        ExcelWS.Range["AG31"].Value = lblChgDownSpec.Text;
                        bRet = double.TryParse(lblChgDownValue.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH31"].Value = dValue.ToString("F1");
                        ExcelWS.Range["AI31"].Value = lblChgDownStr.Text;
                        //}
                        //else ExcelWS.Range["AJ31"].Value = "FALSE";

                        // 温度到達降下
                        //if (ChkNowSpec.PoleDownUse)
                        //{
                        ExcelWS.Range["AE32"].Value = lblPoleDownStart.Text;
                        ExcelWS.Range["AF32"].Value = lblPoleDownEnd.Text;
                        ExcelWS.Range["AG32"].Value = lblPoleDownSpec.Text;
                        bRet = double.TryParse(lblPoleDownValue.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH32"].Value = dValue.ToString("F1");
                        ExcelWS.Range["AI32"].Value = lblPoleDownStr.Text;
                        //}
                        //else ExcelWS.Range["AJ32"].Value = "FALSE";

                        // 温度変動１
                        ExcelWS.Range["AE46"].Value = lblVarTemp1Start.Text;
                        ExcelWS.Range["AF46"].Value = lblVarTemp1End.Text;
                        ExcelWS.Range["AG46"].Value = lblVarTemp1Spec.Text;

                        bRet = double.TryParse(lblTempFluctMemM.Text, out dValue);
                        if (bRet) bRet = double.TryParse(lblTempFluctMemL.Text, out dValue);
                        if (bRet)
                        {
                            float MTemp = float.Parse(lblTempFluctMemM.Text);
                            float LTemp = float.Parse(lblTempFluctMemL.Text);

                            ExcelWS.Range["AH46"].Value = LTemp;
                            if (MTemp > LTemp) ExcelWS.Range["AH46"].Value = MTemp;
                            ExcelWS.Range["AI46"].Value = "NG";
                            if (lblTempFluctJdgeL.Text == "GOOD" && lblTempFluctJdgeM.Text == "GOOD") ExcelWS.Range["AI46"].Value = "GOOD";
                        }
                        else ExcelWS.Range["AI46"].Value = "NG";

                        // 温度変動２
                        ExcelWS.Range["AJ47"].Value = "FALSE";
                        ExcelWS.Range["AE47"].Value = lblVarTemp2Start.Text;
                        ExcelWS.Range["AF47"].Value = lblVarTemp2End.Text;
                        ExcelWS.Range["AG47"].Value = lblVarTemp2Spec.Text;
                        bRet = double.TryParse(lblTempFluctMemH.Text, out dValue);
                        if (bRet) ExcelWS.Range["AH47"].Value = dValue.ToString("F1");
                        ExcelWS.Range["AI47"].Value = lblTempFluctJdgeH.Text;
                        if (HeaderData.TempSpec > 100) ExcelWS.Range["AJ47"].Value = "TRUE";

                        if (ChkNowSpec.HumiUse)
                        {
                            // 湿度変動
                            float fValue1 = float.Parse(lblHumiFluctStd1.Text);
                            float fValue2 = float.Parse(lblHumiFluctStd2.Text);
                            float fValue3 = float.Parse(lblHumiFluctStd3.Text);
                            float fValue4 = float.Parse(lblHumiFluctStd4.Text);
                            float max_value = Math.Max(Math.Max(fValue1, fValue2), Math.Max(fValue3, fValue4));

                            ExcelWS.Range["AG48"].Value = lblVarHumiSpec.Text;
                            ExcelWS.Range["AH48"].Value = max_value.ToString("F1");

                            ExcelWS.Range["AI48"].Value = "NG";
                            if (lblHumiFluctJdge1.Text == "GOOD" && lblHumiFluctJdge2.Text == "GOOD" &&
                                lblHumiFluctJdge3.Text == "GOOD" && lblHumiFluctJdge4.Text == "GOOD") ExcelWS.Range["AI48"].Value = "GOOD";
                        }
#if Flase

                    // 温度設定差
                    ExcelWS.Range["AI52"].Value = "GOOD";
                    ExcelWS.Range["AH52"].Value = TempSettingDiffValue.ToString("F1");
                    if (TempSettingDiffValue > ChkNowSpec.AccuracyTemp1Spec)
                    {
                        ExcelWS.Range["AI52"].Value = "NG";
                    }
                    ExcelWS.Range["AG52"].Value = "1.3";
                    if (!ChkNowSpec.HumiUse)
                    {
                        ExcelWS.Range["AE52"].Value = ChkNowSpec.SpecMaxTemp.ToString("F0");
                        ExcelWS.Range["AE53"].Value = "---";
                        ExcelWS.Range["AG53"].Value = "---";
                        ExcelWS.Range["AH53"].Value = "---";
                        ExcelWS.Range["AI53"].Value = "---";
                    }
                    else
                    {
                        ExcelWS.Range["AE52"].Value = lblSettingTempSetValue.Text;
                        ExcelWS.Range["AE53"].Value = lblSettingHumiSetValue.Text;
                        ExcelWS.Range["AG53"].Value = "3.0";
                        ExcelWS.Range["AI53"].Value = "GOOD";
                        ExcelWS.Range["AH53"].Value = HumiSettingDiffValue.ToString("F1");
                        if (HumiSettingDiffValue > ChkNowSpec.AccuracyHumiSpec)
                        {
                            ExcelWS.Range["AI53"].Value = "NG";
                        }

                    }
                    if (SystemSet.CurrentRequired)
                    {
                        // 積算電力量表示
                        ExcelWS.Range["AD89"].Value = "---";
                        if (PowerInfo.EndExist)
                        {
                            ExcelWS.Range["AD89"].Value = PowerInfo.HourPower.ToString("F1");
                        }
                        // 電流判定結果
                        int j = 0;
                        for (int i = 0; i < 6; ++i)
                        {
                            if (i > 2) j = 6;
                            if (CurInfos[i].Use)
                            {
                                ExcelWS.Range["AD" + (70 + i + j).ToString()].Value = CurInfos[i].name;
                                ExcelWS.Range["AE" + (70 + i + j).ToString()].Value = "TRUE";
                                ExcelWS.Range["AF" + (70 + i + j).ToString()].Value = CurInfos[i].Power.ToString("F1");
                                ExcelWS.Range["AG" + (70 + i + j).ToString()].Value = CurInfos[i].Low.ToString("F1");
                                ExcelWS.Range["AH" + (70 + i + j).ToString()].Value = CurInfos[i].High.ToString("F1");
                                ExcelWS.Range["AI" + (70 + i + j).ToString()].Value = CurInfos[i].Value.ToString("F1");
                                ExcelWS.Range["AJ" + (70 + i + j).ToString()].Value = "NG";
                                if (CurInfos[i].Judge) ExcelWS.Range["AJ" + (70 + i + j).ToString()].Value = "GOOD";
                            }
                            else ExcelWS.Range["AE" + (70 + i + j).ToString()].Value = "FALSE";
                        }
                    }
#endif
                        // Top表示
                        ExcelWS.Range["A1"].Activate();

                        cMsg = "試験検査報告書ファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        TextMsgWrite(cMsg);
                        //警告メッセージを非表示にする
                        ExcelWB.Save();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWS);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWB);
                    }
                    #endregion 試験成績書処理

                    #region アース導通保存
                    string EarthDistFileName = "";
                    string EarthSourceFileName = "";
                    SheetName = "";
                    if (SystemSet.TestReportRequired)
                    {
                        EarthDistFileName = @DocPath + "\\" + "Earth一時ファイル.xls";

                        if (System.IO.File.Exists(EarthDistFileName)) System.IO.File.Delete(EarthDistFileName);
                        cPath = SystemSet.MasterPath + "アース導通試験報告書.xls";

                        EarthSourceFileName = cPath;
                        SheetName = "アース導通";

                        // 試験検査報告書ファイル複写
                        System.IO.File.Copy(EarthSourceFileName, EarthDistFileName);

                        //FileInfoオブジェクトを作成
                        System.IO.FileInfo fi = new System.IO.FileInfo(EarthDistFileName);
                        //隠し属性があるか調べる
                        if ((fi.Attributes & System.IO.FileAttributes.Hidden) ==
                            System.IO.FileAttributes.Hidden)
                        {
                            Console.WriteLine("隠し属性があります。");
                            //隠し属性を削除する
                            fi.Attributes &= ~System.IO.FileAttributes.Hidden;
                        }

                        // xlApplication から WorkBooks を取得する
                        // 既存の Excel ブックを開く
                        ExcelWB = xlBooks.Open(EarthDistFileName);
                        ExcelWS = (Worksheet)ExcelWB.Sheets[SheetName];
                        ExcelWS.Activate();
                        // Excel を表示する

                        // ヘッダー情報
                        ExcelWS.Range["Z5"].Value = HeaderData.PuroductModel;
                        ExcelWS.Range["Z6"].Value = HeaderData.CreateNo;

                        string strTime1 = TestingDate + " 00:00:00";
                        string strTime = "2024/5/8 00:00:00";

                        DateTime dTime = DateTime.Parse(strTime);
                        ExcelWS.Range["Z7"].Value = dTime;
                        ExcelWS.Range["Z8"].Value = TestingPersonEn;
                        ExcelWS.Range["Z9"].Value = TestingPerson;

                        // Top表示
                        ExcelWS.Range["A1"].Activate();

                        xlApp.Run("HideBlankRows");

                        cMsg = "アース導通試験報告書ファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        TextMsgWrite(cMsg);
                        //警告メッセージを非表示にする
                        ExcelWB.Save();

                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWS);
                        System.Runtime.InteropServices.Marshal.ReleaseComObject(ExcelWB);
                    }

                    #endregion アース導通保存

                    // Excel を終了する
                    cMsg = "エクセル終了\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    xlApp.DisplayAlerts = false;
                    xlBooks.Close();
                    xlApp.Quit();
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlBooks);
                    System.Runtime.InteropServices.Marshal.FinalReleaseComObject(xlApp);
                    #region 自動判定ログ保存

                    string path = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\";
                    path += "CSV";
                    String CheckDataSourceFileName = @path + "\\" + NowExcelFileName + "_自動判定ログ.CSV";

                    // 検査データ保存フォルダ作成
                    string DistPath = @SystemSet.ResultPath + NowExcelFileName;
                    if (!Directory.Exists(DistPath))
                    {
                        // ディレクトリ作成
                        Directory.CreateDirectory(DistPath);
                    }

                    // 検査データ移動
                    String CheckDataDistFileName = @DistPath + "\\" + NowExcelFileName + "_自動判定ログ.csv";

                    TestSourceFileName = TestDistFileName;

                    string sFileName = NowExcelFileName.Replace("検査データ_", "");
                    TestDistFileName = @DistPath + "\\" + "試験検査報告書_" + sFileName + ".xls";
                    EarthSourceFileName = EarthDistFileName;
                    EarthDistFileName = @DistPath + "\\" + "アース試験報告書_" + sFileName + ".xls";

                    GraphSourceFileName = GraphDistFileName;
                    GraphDistFileName = @DistPath + "\\" + NowExcelFileName + "_グラフ.xlsx";

                    GraphSourcePDFFileName = GraphDistPDFFileName;
                    GraphDistPDFFileName = @DistPath + "\\" + NowExcelFileName + "_グラフ.pdf";

                    SourceFileName = DistFileName;
                    DistFileName = @DistPath + "\\" + NowExcelFileName + ".xls";

                    DialogResult result = DialogResult.Yes;
                    if (System.IO.File.Exists(DistFileName))
                    {
                        // 保存先にファイルが存在していた場合
                        cMsg = NowExcelFileName + ".xls" + "ファイルが存在します。上書きしますか？";
                        result = MessageBox.Show(cMsg, "上書き警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    }
                    cMsg = "検査データファイル書込み中\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);
                    if (result == DialogResult.Yes)
                    {
                        // ファイル複写（上書き保存）
                        System.IO.File.Copy(SourceFileName, DistFileName, true);

                        if (SystemSet.TestReportRequired)
                        {
                            System.IO.File.Copy(TestSourceFileName, TestDistFileName, true);
                            System.IO.File.Copy(EarthSourceFileName, EarthDistFileName, true);
                        }
                        else
                        {
                            if (System.IO.File.Exists(TestDistFileName)) System.IO.File.Delete(TestDistFileName);
                        }
                        // グラフファイル保存
                        if (SystemSet.GraphRequired)
                        {
                            System.IO.File.Copy(GraphSourceFileName, GraphDistFileName, true);
                            System.IO.File.Copy(GraphSourcePDFFileName, GraphDistPDFFileName, true);
                        }
                        else
                        {
                            if (System.IO.File.Exists(GraphDistFileName)) System.IO.File.Delete(GraphSourceFileName);
                            if (System.IO.File.Exists(GraphDistPDFFileName)) System.IO.File.Delete(GraphSourcePDFFileName);
                        }

                        if (DebugMode)
                        {
                            System.IO.File.Copy(CheckDataSourceFileName, CheckDataDistFileName, true);
                        }
                        // CSV ファイル複写
                        string cFileName = DistPath + "\\" + lblCsvFileName.Text + ".CSV";
                        System.IO.File.Copy(NowCSVFilePath, @cFileName, true);
                    }
                    #endregion 自動判定ログ保存

                    // 一時ファイル削除
                    System.IO.File.Delete(SourceFileName);
                    //if (SystemSet.TestReportRequired) File.Delete(TestSourceFileName);
                    if (SystemSet.GraphRequired)
                    {
                        System.IO.File.Delete(GraphSourceFileName);
                        System.IO.File.Delete(GraphSourcePDFFileName);
                    }
                    if (DebugMode)
                    {
                        System.IO.File.Delete(CheckDataSourceFileName);
                    }
                    cMsg = "検査データファイル書込み完了\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);

                    // 検査保存フォルダをエクスプローラで開く
                    cMsg = "検査データ保存フォルダをエクスプローラで開きます。\r\n";
                    WriteLog(cMsg);
                    TextMsgWrite(cMsg);
                    System.Diagnostics.Process.Start(@DistPath);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    // データ読込エラー
                    cMsg = "エクセルファイル書出しに失敗しました。\r\n";
                    cMsg += "設定されている内容を再度確認して下さい。\r\n";
                    // 異常時の背景色を赤色
                    txtMessage.BackColor = Color.Red;
                    TextMsgWrite(cMsg);

                    MessageBox.Show(cMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
                {
                    System.Diagnostics.Process[] prcs2 = System.Diagnostics.Process.GetProcessesByName("excel");

                    if (prcs2 != null)
                    {
                        foreach (System.Diagnostics.Process p in prcs2)
                        {

                            if (ExcelPid[0] != 0)
                            {
                                bool bExist = false;
                                for (int iCnt = 1; iCnt < ExcelPid[0] + 1; ++iCnt)
                                {
                                    if (ExcelPid[iCnt] == p.Id) bExist = true;
                                }
                                if (!bExist) p.Kill();
                            }
                            else p.Kill();
                        }
                    }
                }
            }
        }
        private void DispInit()
        {
            // 製造番号初期化
            SerialNo = ""; UserName = "";
            lblSerial_number.Text = "";
            txtUserName.Text = "";
            txtSerial_number.Text = "";
            cmbOperationHole.SelectedIndex = 0;
            cmbCableHole25.SelectedIndex = 0;
            cmbCableHole100.SelectedIndex = 0;
            cmbCableHoleFlat.SelectedIndex = 0;
            cmbCableHole50.SelectedIndex = 1;
            cmbTestVolt.SelectedIndex = 0;
            cmbTestFreq.SelectedIndex = 0;
            cmbModel.SelectedIndex = 0;
            txtAdditionalSpec.Text = "";
            lblRoomTemp.Text = "";
        }
        public static  int FluctSetTemp(float cTemp) 
        {
            int iRet = 0;
            float fValue = ChkNowSpec.AccuracyTemp3Spec;

            if (ChkNowSpec.Temp180Use) fValue = ChkNowSpec.AccuracyTemp3Spec;
            else fValue = ChkNowSpec.AccuracyTemp2Spec;

            if (ChkNowSpec.SpecMaxTemp - ChkNowSpec.AccuracyTemp2Spec < cTemp &&
                ChkNowSpec.SpecMaxTemp + ChkNowSpec.AccuracyTemp2Spec > cTemp) iRet = (int)ChkNowSpec.SpecMaxTemp;
            else if (ChkNowSpec.SpecMinTemp - ChkNowSpec.AccuracyTemp1Spec < cTemp &&
                ChkNowSpec.SpecMinTemp + ChkNowSpec.AccuracyTemp1Spec > cTemp) iRet = (int)ChkNowSpec.SpecMinTemp;
            else if (ChkNowSpec.SpecMinDownTemp - ChkNowSpec.AccuracyTemp1Spec < cTemp &&
                ChkNowSpec.SpecMinDownTemp + ChkNowSpec.AccuracyTemp1Spec > cTemp) iRet = (int)ChkNowSpec.SpecMinDownTemp;
            else if (ChkNowSpec.SpecMaxUpTemp - fValue < cTemp &&
                ChkNowSpec.SpecMaxUpTemp + fValue > cTemp) iRet = (int)ChkNowSpec.SpecMaxUpTemp;
            else iRet = int.MinValue;

            return iRet;
        }
        private void btnFolderSelect_Click(object sender, EventArgs e)
        {
            //FolderBrowserDialogクラスのインスタンスを作成
            FolderBrowserDialog fbd = new FolderBrowserDialog();

            //上部に表示する説明テキストを指定する
            fbd.Description = "フォルダを指定してください。";
            //ルートフォルダを指定する
            //デフォルトでDesktop
            fbd.RootFolder = Environment.SpecialFolder.Desktop;
            //最初に選択するフォルダを指定する
            //RootFolder以下にあるフォルダである必要がある
            fbd.SelectedPath = @"D:\";
            //ユーザーが新しいフォルダを作成できるようにする
            //デフォルトでTrue
            fbd.ShowNewFolderButton = true;

            //ダイアログを表示する
            if (fbd.ShowDialog(this) == DialogResult.OK)
            {
                //選択されたフォルダを表示する
                SystemSet.ExcelPath = fbd.SelectedPath;
                SystemSet.ExcelPath= fbd.SelectedPath;
                lblExcelFolderPath.Text = SystemSet.ExcelPath;
                Console.WriteLine(SystemSet.ExcelPath);
            }
        }
        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            String cMsg = "";
            cMsg = DataGet();
            if (cMsg != "")
            {
                cMsg += "上記設定が誤っています。修正しないと終了できません。 ";
                MessageBox.Show(cMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                e.Cancel = true;
            }
            else
            {
                //JsonFileWrite();
                // データベースファイル終了
                DbWrite();
                DbClose();
                // バーコードリーダーＣＯＭポートオープン
                if(BarCodeReader.IsOpen) BarCodeReader.Close();
            }
        }
        private void DBStart()
        {
            // Oracle接続確認
            if (!Oracle.db_Exist(strConnString)) { WriteLog("基幹システム未接続" + "\r\n"); bDBReadyEnd = true; return; }
            bDBOracleFlg = true;
            bDBReadyEnd = true;
            bDBReadyFlg = true;
        }
        private void timer100msec_Tick(object sender, EventArgs e)
        {
            #region Logファイル削除処理 
            // 7日以上前のLogファイルを１分毎に１ファイル削除する
            DateTime dt = DateTime.Now;
            //string path = System.AppDomain.CurrentDomain.BaseDirectory;
            string path = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\";

            path += "Log";
            if (dt.ToString("yyyyMMddhhmmss").Substring(12, 2) == "00")
            {
                // 一週間以上前のファイルは削除
                DateTime DtNow = DateTime.Now;
                DateTime DtNowBefore7 = DtNow.AddDays(-7);
                int CheckDate = int.Parse(DtNowBefore7.ToString("yyyyMMdd"));

                string[] files = Directory.GetFiles(path);
                Array.Sort(files);
                foreach (string file in files)
                {
                    string CheckFileName = file.Substring(file.Length - 14, 8);
                    int FilekDate = int.Parse(CheckFileName);
                    if (FilekDate < CheckDate) { System.IO.File.Delete(file); break; }
                }
            }
            #endregion Logファイル削除処理 

            // CSVReadスレッド終了確認
            if (tCSVReadRun)
            {
                if (tCSVReadEnd)
                {
                    if (RecLength != int.MaxValue && RecLength != int.MinValue)
                    {
                        // CSVReadスレッド処理
                        //TextMsgWrite(tCSVMsg);

                        // 検査日登録
                        lblTestImplementationDate.Text = TestingDate;
                        lblRoomTemp.Text = RoomTempAverage.ToString("F1");

                        tCSVReadRun = false;
                        tCSVReadEnd = false;

                        // データ確認
                        //TestDataCheck();
                        // 結果表示
                        TestResultDisply();
                        // グラフプロット
                        ViewGraphinit();
                        ViewGraph();
                        StateInit();

                        // ボタンを有効/無効設定
                        btnCSVFileSelect.Enabled = false;
                        btnStart.Enabled = false;
                        btnExcelSave.Enabled = true;
                    }
                    else
                    {
                        String cMsg = "";
                        if (RecLength == int.MaxValue)
                        {
                            // データ判定エラー
                            cMsg = "製品仕様判定条件を検出できませんでした。\r\n";
                            cMsg += "設定内容およびＣＦデータを確認して下さい。\r\n";
                            // データ確認
                            //TestDataCheck();
                            // 結果表示
                            TestResultDisply();
                            // グラフプロット
                            ViewGraphinit();
                            ViewGraph();
                        }
                        else
                        {
                            // データ読込エラー
                            cMsg = "データの並びが違います。\r\n";
                            cMsg += "日付,時刻,制御温度,制御湿度,槽中央乾球温度,槽中央湿球温度,周囲温度の\r\n";
                            cMsg += "順番である必要があります。\r\n";
                        }
                        TextMsgWrite(cMsg);
                        txtMessage.BackColor = Color.Red;
                        tCSVReadRun = false;
                        tCSVReadEnd = false;

                        StateInit();

                        // ボタンを有効/無効設定
                        btnCSVFileSelect.Enabled = false;
                        btnStart.Enabled = false;
                        btnExcelSave.Enabled = true;
                    }
                }
                else
                {
                    // CSV読込中
                    string CSVMsg = "読込中レコード： " + ChkLength.ToString() + "\r\n";
                    TextMsgWrite(CSVMsg);
                    //while (DataQueue.TryDequeue(out string item))
                    //{
                    //    // ここでフォームのコントロールを安全に操作できる
                    //    TextMsgWrite(item);
                    //    // 必要に応じてスクロールなどのUI更新も行う
                    //}
                }
            }

            // バーコード取り込みデータ確認
            if (ReturnBarCodeData != "")
            {
                if (ReturnBarCodeData.Substring(0, 1) == "S")
                {
                    // 検査プログラムを受信した
                    for (int i = 0; i < Mashine.Models.Count; ++i)
                    {
                        Json.MODELS model = Mashine.Models[i];
                        if (model.BarCode == ReturnBarCodeData)
                        {
                            ModelSelInitial(i);
                            break;
                        }
                    }
                }
                else
                {
                    // 製造番号を受信した。
                    lblSerial_number.Text = ReturnBarCodeData;
                    //SetExcellFileName();
                }
                SerialNo = ReturnBarCodeData;
                ExcelValue.SerialNo = SerialNo;
                ReturnBarCodeData = "";
                // 基幹システムデータベースアクセス
                if (SystemSet.OracleRead && bDBOracleFlg)
                {
                    HeaderData = new Header();
                    HeaderData = Oracle.headerRead(strConnString, DatabaseInfo, SerialNo);
                    // 製造番号の頭２桁の”00”を撤去する
                    string sNo = HeaderData.CreateNo;
                    if (sNo != "")
                    {
                        if (sNo.Substring(0, 2) == "00")
                        {
                            sNo = sNo.Substring(2, sNo.Length - 2);
                            HeaderData.CreateNo = sNo;
                            SerialNo = sNo;
                            ExcelValue.SerialNo = sNo;
                        }
                        string ChkModel = HeaderData.PuroductModel;
                        // スペック初期化
                        HeaderData.TempSpec = 0;
                        HeaderData.VoltSpec = 0;
                        HeaderData.WVHoleSpec = 0;
                        //小型環境標準設定を実施して置く
                        HeaderData.TempSpec = 150;
                        HeaderData.VoltSpec = 100;

                        ChkNowSpec.RefLeakTestUse = false;
                        ChkNowSpec.WindowUse = false;
                        ChkNowSpec.IndoorUse = false;
                        ChkNowSpec.OpeWindowUse = false;
                        CableHole[0] = 0; CableHole[1] = 1; CableHole[2] = 0; CableHole[3] = 0;

                        if (HeaderData.Option.Count != 0)
                        {
                            // Option 情報解析
                            String OpCode = HeaderData.Option[0].OptionCode;
                            string ModelCode = OpCode.Substring(0, 4);
                            int Model = 0; int MaxCnt = 1;
                            foreach (OPTIONCDS ops in OptionSet.OptionCD)
                            {
                                if (ops.ModelCD == ModelCode)
                                {
                                    MaxCnt = ops.Option.Count;
                                    break;
                                }
                                ++Model;
                            }

                            foreach (Option op in HeaderData.Option)
                            {
                                Console.WriteLine(" オプションコード " + op.OptionCode + " オプション内容 : " + op.OptionName + " 種類 : " + op.OptionKind + " 値 : " + op.OptionValue);
                                for (int i = 0; i < MaxCnt; ++i)
                                {

                                    if (OptionSet.OptionCD[Model].Option[i].OpCode == op.OptionCode)
                                    {
                                        string OpCone = OptionSet.OptionCD[Model].Option[i].OpCode;
                                        string OpName = OptionSet.OptionCD[Model].Option[i].OpName;
                                        // 対象オプションコード確認
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "Volt")
                                        {
                                            HeaderData.VoltSpec = int.Parse(OptionSet.OptionCD[Model].Option[i].Value);
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "Temp")
                                        {
                                            HeaderData.TempSpec = int.Parse(OptionSet.OptionCD[Model].Option[i].Value);
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "Window")
                                        {
                                            ChkNowSpec.WindowUse = false;
                                            if (OptionSet.OptionCD[Model].Option[i].Value == "Use") ChkNowSpec.WindowUse = true;
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "InDoor")
                                        {
                                            ChkNowSpec.IndoorUse = false;
                                            if (OptionSet.OptionCD[Model].Option[i].Value == "Use") ChkNowSpec.IndoorUse = true;
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "SideOpeHole")
                                        {
                                            ChkNowSpec.OpeWindowUse = false;
                                            if (OptionSet.OptionCD[Model].Option[i].Value == "Use") ChkNowSpec.OpeWindowUse = true;
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "WideView")
                                        {
                                            if (OptionSet.OptionCD[Model].Option[i].Value == "操作孔1対") HeaderData.WVHoleSpec = 1;
                                            else if (OptionSet.OptionCD[Model].Option[i].Value == "操作孔2対") HeaderData.WVHoleSpec = 2;
                                        }
                                        if (OptionSet.OptionCD[Model].Option[i].Kind == "CableHole" && true)
                                        {
                                            if (OpName.Contains("削除")) --CableHole[1];
                                            else if (OpName.Contains("φ２５")) ++CableHole[0];
                                            else if (OpName.Contains("φ５０")) ++CableHole[1];
                                            else if (OpName.Contains("φ１００")) ++CableHole[2];
                                            else if (OpName.Contains("フラット")) ++CableHole[3];
                                        }
                                        else if (OptionSet.OptionCD[Model].Option[i].Kind == "Exclusion")
                                        {
                                            // 特殊仕様対応("耐食性冷却器"など)
                                            if (OptionSet.OptionCD[Model].Option[i].OpName == "冷凍機リークテスト" &&
                                                op.OptionName.Contains("冷凍機リークテスト"))
                                            {
                                                // 冷凍機リークテスト使用
                                                ChkNowSpec.RefLeakTestUse = true;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        // オプションコード登録無い場合のオプション名称確認
                                        if (op.OptionName.Contains("観測窓"))
                                        {
                                            ChkNowSpec.WindowUse = true;
                                        }
                                        if (op.OptionName.Contains("内扉"))
                                        {
                                            ChkNowSpec.IndoorUse = true;
                                        }

                                    }
                                }

                            }
                        }
                        Console.WriteLine(" ユーザー名 " + HeaderData.UserName + " オーダー番号 : " + HeaderData.PuroductOrder + " 製品型式 : " + HeaderData.PuroductModel);
                        Console.WriteLine(" 電圧仕様 : " + HeaderData.VoltSpec.ToString("") + " 温度仕様 : " + HeaderData.TempSpec.ToString("") + " 周波数仕様 : " + HeaderData.FreqSpec.ToString(""));


                        // ケーブル孔更新 
                        cmbCableHole25.Text = CableHole[0].ToString("");
                        cmbCableHole50.Text = CableHole[1].ToString("");
                        cmbCableHole100.Text = CableHole[2].ToString("");
                        cmbCableHoleFlat.Text = CableHole[3].ToString("");

                        // データグリッド更新
                        dataGridViewMain.Refresh();
                        dataGridViewMain.DataSource = HeaderData.Option;

                        Console.WriteLine(dataGridViewMain.Columns.Count);
                        if (dataGridViewMain.Columns.Count != 0)
                        {
                            // カラム名を指定
                            dataGridViewMain.Columns[0].HeaderText = "Code";
                            dataGridViewMain.Columns[1].HeaderText = "名称";
                            dataGridViewMain.Columns[2].HeaderText = "種類";
                            dataGridViewMain.Columns[3].HeaderText = "Value";
                            dataGridViewMain.CurrentCell = null;
                            dataGridViewMain.ClearSelection();
                        }
                        if (ChkNowSpec.WindowUse) Console.WriteLine(" 観測窓仕様 : 有り");
                        if (!ChkNowSpec.WindowUse) Console.WriteLine(" 観測窓仕様 : 無し");
                        if (ChkNowSpec.IndoorUse) Console.WriteLine(" 内扉仕様 : 有り");
                        if (!ChkNowSpec.IndoorUse) Console.WriteLine(" 内扉仕様 : 無し");
                        if (ChkNowSpec.OpeWindowUse) { Console.WriteLine(" 側面操作孔仕様 : 有り"); cmbOperationHole.SelectedIndex = 1; }
                        if (!ChkNowSpec.OpeWindowUse) { Console.WriteLine(" 側面操作孔仕様 : 無し"); cmbOperationHole.SelectedIndex = 0; }


                        // 検査プログラム特定　　                       String ChkModel = HeaderData.PuroductModel;
                        String ChkTemp = HeaderData.TempSpec.ToString("");
                        String ChkSpec = "標準";
                        string ChkName = "";
                        if (ChkNowSpec.WindowUse) ChkSpec = "観窓";
                        if (ChkNowSpec.IndoorUse) ChkSpec = "内扉";
                        foreach (Json.MODELS Md in Mashine.Models)
                        {
                            if (ChkTemp != "150")
                            {
                                if (Md.ProgramName.Contains(ChkModel) && Md.ProgramName.Contains(ChkTemp) && Md.ProgramName.Contains(ChkSpec))
                                {
                                    Console.WriteLine(Md.ProgramName);
                                    ChkName = Md.ProgramName;
                                    break;
                                }
                            }
                            else
                            {
                                if (Md.ProgramName.Contains(ChkModel) && Md.ProgramName.Contains(ChkSpec))
                                {
                                    Console.WriteLine(Md.ProgramName);
                                    ChkName = Md.ProgramName;
                                    break;
                                }
                            }
                        }
                        // 製品情報更新
                        txtUserName.Text = HeaderData.UserName;
                        cmbTestVolt.Text = HeaderData.VoltSpec.ToString("");
                        cmbTestFreq.Text = HeaderData.FreqSpec.ToString("");

                        // 検査プログラム検索更新
                        for (int i = 0; i < cmbModel.Items.Count; i++)
                        {
                            if (ChkName == cmbModel.Items[i].ToString())
                            {
                                cmbModel.SelectedIndex = i;
                                ModelSelInitial(i);
                                break;
                            }
                        }
                    }
                    else 
                    {
                        if (SerialNo.Substring(0, 2) == "00")
                        {
                            SerialNo = SerialNo.Substring(2, SerialNo.Length - 2);
                            HeaderData.CreateNo = SerialNo;
                            ExcelValue.SerialNo = SerialNo;
                        }

                        // 基幹システムから情報を取り込めなかった（レンタル品）
                        DialogResult result = DialogResult.OK;
                        string cMsg = "基幹システムより製品情報を取得できませんでした。\r\n";
                        cMsg += "レンタル品と思われます。\r\n";
                        cMsg += "ユーザー名・検査プログラム名等を手動で設定してください。";
                        result = MessageBox.Show(cMsg, "登録確認警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }

            }

            // エクセル処理
            if (bExcelWriteBusy)
            {
                // UIスレッドでキューからアイテムを取り出して処理する
                while (Excel_Operation.DataQueue.TryDequeue(out string item))
                {
                    // ここでフォームのコントロールを安全に操作できる
                    TextMsgWrite(item);
                }
                if (bExcelWriteEnd)
                {
                    bExcelWriteBusy = false;
                    if (bExcelWriteFail) 
                    {
                        // エクセル処理異常
                        txtMessage.BackColor = Color.Red;
                    }
                }
            }
        }
        private void btnPatternCheck_Click(object sender, EventArgs e)
        {
            string cMsg = "";
            // 検査パターン表示
            if (ChkNowSpec.HumiUse) 
            {
                cMsg += "温湿度制御パターン\r\n\r\n";
                for (int i = 0;i< HUMISETSTEPS; ++i) 
                {
                    cMsg += "ステップ " + i.ToString() + 
                            "制御設定 = " + Humiset[i].SetTemp.ToString("F1") + " ℃ ／" + Humiset[i].SetHumi.ToString("") + " %RH " +
                            // "制御時間 = " + Humiset[i].SetTime.ToString("F1") + " Hr " + "\r\n";
                            "制御時間 = " + TimeFormatChange(Humiset[i].SetTime) + "\r\n";
                }
                cMsg += "\r\n";
            }
            cMsg += "温度制御パターン\r\n\r\n";
            for (int i = 0; i < TEMPSETSTEPS; ++i)
            {
                cMsg += "ステップ " + i.ToString() +
                        "制御設定 = " + Tempset[i].SetTemp.ToString("").PadLeft(3, ' ') + " ℃ " +
                        // "制御時間 = " + Tempset[i].SetTime.ToString("F1") + " Hr " + "\r\n";
                        "制御時間 = " + TimeFormatChange(Tempset[i].SetTime) + "\r\n";
            }

            sTestCheckMsg1 = cMsg;

            cMsg = "温度精度および温度変動幅\r\n\r\n";
            cMsg += "低温 精度 = " + ChkNowSpec.AccuracyTemp1Spec.ToString("F1").PadLeft(4, ' ') + " ℃ ";
            cMsg += "温度変動幅 = " + ChkNowSpec.VarTemp1Spec.ToString("F1") + " ℃ " + "\r\n";
            cMsg += "中温 精度 = " + ChkNowSpec.AccuracyTemp2Spec.ToString("F1").PadLeft(4, ' ') + " ℃ ";
            cMsg += "温度変動幅 = " + ChkNowSpec.VarTemp2Spec.ToString("F1") + " ℃ " + "\r\n";
            if (ChkNowSpec.Temp180Use) 
            { 
                cMsg += "高温 精度 = " + ChkNowSpec.AccuracyTemp3Spec.ToString("F1").PadLeft(4, ' ') + " ℃ ";
                cMsg += "温度変動幅 = " + ChkNowSpec.VarTemp3Spec.ToString("F1") + " ℃ " + "\r\n";
            }
            if (ChkNowSpec.HumiUse) 
            {
                cMsg += "\r\n湿度精度および湿度変動幅\r\n\r\n";
                for (int i = 0; i < HUMISETSTEPS; ++i)
                {
                    cMsg += Humiset[i].SetTemp.ToString("F0") + "℃/" + Humiset[i].SetHumi.ToString("F0") + "% " + "精度 = " 
                        + Humiset[i].HumiAccuracy.ToString("F1") + " %RH " + "\r\n";
                }
            }
            sTestCheckMsg2 = cMsg;

            var form2 = new SetCheck();
            form2.ShowDialog();
            form2.Dispose();
        }

        public string TimeFormatChange(float decimalHours) 
        {
            TimeSpan timeSpan = TimeSpan.FromHours(decimalHours);
            string formattedTime = timeSpan.ToString(@"hh\:mm");
            return formattedTime;
        }

        private void formsPlot1_MouseUp(object sender, MouseEventArgs e)
        {

            // マウスの座標を取得します
            (double mouseCoordX, double mouseCoordY) = formsPlot1.GetMouseCoordinates();

            // グラフのX軸/Y軸の単位から比率を計算します
            double xyRatio = formsPlot1.Plot.XAxis.Dims.PxPerUnit / formsPlot1.Plot.YAxis.Dims.PxPerUnit;

            // マウスの座標・グラフの縦横比から、カーソル近傍のポイント座標を取得します
            double[] pointX = new double[6];
            double[] pointY = new double[6];
            int[] pointIndex = new int[6];
            bool bNull = false;

            for (int i = 0; i < 6; ++i)
            {
                if (scat[i] != null)
                {
                    (pointX[i], pointY[i], pointIndex[i]) = scat[i].GetPointNearest(mouseCoordX, mouseCoordY, xyRatio);
                }
                else bNull = true;
            }

            string sMsg = "";
            //sMsg += "経過時間:" + pointX[0].ToString("") + "\r\n";
            sMsg += "制御温度:" + pointY[0].ToString("F1") + "℃\r\n";
            if(!bNull && ChkNowSpec.HumiUse) sMsg += "制御湿度:" + pointY[1].ToString("F1") + "\r\n";
            sMsg += "乾球温度:" + pointY[2].ToString("F1") + "℃\r\n";
            if (!bNull && ChkNowSpec.HumiUse) sMsg += "湿球湿度:" + pointY[3].ToString("F1") + "℃\r\n";
            if (!bNull && ChkNowSpec.HumiUse) sMsg += "相対湿度:" + pointY[5].ToString("F1") + "%RH\r\n";
            sMsg += "周囲温度:" + pointY[4].ToString("F1") + "℃\r\n";
            Console.WriteLine(sMsg);
            //formsPlot1.Render();
            lblToolTip.Text = sMsg;
            lblToolTip.Visible = true;
            lblToolTip.Top = e.Y;
            lblToolTip.Left = e.X;

            // 強調表示用の散布図のポイントを、上で取得したポイントの座標にセットします。
            // これにより、カーソル近傍のポイントに赤丸がついたように見せることができます
            //HighlightedPoint.Xs[0] = pointX;
            //HighlightedPoint.Ys[0] = pointY;
            //HighlightedPoint.IsVisible = true;

            // カーソル近傍のポイントが変わった場合のみ、描画を行います
            //if (LastHighlightedIndex != pointIndex)
            //{
            //    LastHighlightedIndex = pointIndex;
            //    formsPlot1.Render();
            //}
        }
        private void lblToolTip_Click(object sender, EventArgs e)
        {
            lblToolTip.Visible = false;
        }
        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            lblToolTip.Visible = false;
        }
        private string CheckBarCode() 
        {
            if (BarCodeData == "") return "";
            string ChekDigit = ""; string ErrorMsg = "";string ReturnMsg = "";
            string Value = BarCodeData.Substring(0, BarCodeData.Count() - 1);
            string rDigit = BarCodeData.Substring(BarCodeData.Count()-1 ,1);

            ChekDigit = GetModulus43(Value);
            Console.WriteLine("受信コード = " + BarCodeData + " データ = " + Value + " チェックデジット = " + ChekDigit);

            // チェックデジット確認
            if (rDigit == ChekDigit)
            {
                if (BarCodeData.Substring(0, 1) == "S")
                {
                    // 検査プログラム名取得
                    bool fExist = false;
                    for(int i = 0 ; i < Mashine.Models.Count ; ++i)
                    {
                        Json.MODELS model = Mashine.Models[i];
                        if(model.BarCode == Value) 
                        {
                            DialogResult result = DialogResult.No;
                            string cMsg = "検査プログラム名: " + Value + " を登録します。宜しいですか？";
                            result = MessageBox.Show(cMsg, "登録確認警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                            if (result == DialogResult.Yes)
                            {
                                ReturnMsg = Value;
                                fExist = true;
                                break;
                            }

                        }
                    }
                    if (!fExist) 
                    {
                        // 検査プログラム名不一致
                        ErrorMsg += "検査プログラム名不一致エラー";
                    }
                }
                else
                {
                    // 製造番号取得
                    // 10文字以上確認
                    if (Value.Count() < 10)
                    {
                        ErrorMsg += "製造番号文字数取得エラー";
                    }
                    else 
                    {
                        // 製造番号１０文字取出し登録(数値確認必要か？)
                        string SerialNo = Value.Substring(0,10);
                        DialogResult result = DialogResult.No;
                        string cMsg = "製造番号: " + SerialNo +  " を登録します。宜しいですか？";
                        result = MessageBox.Show(cMsg, "登録確認警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                        if (result == DialogResult.Yes)
                        {
                            ReturnMsg = SerialNo;
                        }
                    }
                }
            }
            else 
            {
                // チェックディジット不一致
                ErrorMsg += "チェックディジット不一致エラー";
            }
            if (ErrorMsg != "") 
            {
                // バーコード読取エラー
                ReturnMsg = "";
                string wMsg = "バーコード読み取りエラー\r\n" + ErrorMsg + "\r\n"  + "再度、読み取りしてください。\r\n";
                MessageBox.Show(wMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            BarCodeData = "";
            BarCodeStart = false;
            return ReturnMsg;
        }
        private void Form1_KeyPress(object sender, KeyPressEventArgs e)
        {

#if DEBUG
            string cMsg = e.KeyChar.ToString();
            if (!BarCodeStart) 
            {
                if (BarCodeData == "" && cMsg == "*")
                {
                    // スタートコード受信
                    BarCodeStart = true;
                }
            }
            else 
            {
                if(cMsg == "*" ) 
                {
                    // ストップコード受信
                    // チェックデジット確認
                    ReturnBarCodeData = CheckBarCode();
                }
                else 
                {
                    BarCodeData += cMsg;
                }
            }
#endif
        }
        private void Form1_Click(object sender, EventArgs e)
        {
            this.Activate();
        }
        private void BarCodeReader_DataReceived(object sender, System.IO.Ports.SerialDataReceivedEventArgs e)
        {
            string cMsg = "";

            int bytesToRead = BarCodeReader.BytesToRead;
            if (bytesToRead != 0)
            {
                byte[] receivedData = new byte[bytesToRead];
                BarCodeReader.Read(receivedData, 0, bytesToRead);
                //BarCodeReader.ReadLine();
                cMsg = Encoding.ASCII.GetString(receivedData);
                Console.WriteLine(cMsg);

                if (cMsg.Contains('\r'))
                {
                    BarCodeData = cMsg.Substring(0, cMsg.IndexOf("\r"));
                    // チェックデジット確認
                    ReturnBarCodeData = CheckBarCode();
                }
            }
        }
        private void btnSerialNoInput_Click(object sender, EventArgs e)
        {
            string ErrMsg = "";
            try
            {
                // 入力内容確認
                string StringSerialNo = txtSerial_number.Text;
                ExcelValue.SerialNo = StringSerialNo;

                if (StringSerialNo.Count() != 10)
                {
                    ErrMsg += "入力文字数が１０文字ではありません。\r\n";
                }
                else if (double.IsNaN(ValueCheck(StringSerialNo)))
                {
                    ErrMsg += "入力文字が数値ではありません。\r\n";
                }

                if (ErrMsg != "")
                {
                    TextMsgWrite(ErrMsg);
                    WriteLog(ErrMsg);
                    MessageBox.Show(ErrMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                else
                {
                    // 製造番号登録
                    lblSerial_number.Text = StringSerialNo;
                    ReturnBarCodeData = StringSerialNo;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
            }

        }
        private void ChrSizeTextBoxChg(System.Windows.Forms.TextBox control) 
        {
            // テキストボックスの幅
            int textBoxWidth = control.Width;

            // 現在のフォントサイズ
            float currentFontSize = control.Font.Size;

            // テキストの長さを測定
            System.Drawing.Font font = control.Font;
            System.Drawing.Size textSize = TextRenderer.MeasureText(control.Text, font);

            // テキストがテキストボックスの幅を超えた場合にフォントを縮小する
            if (textSize.Width > textBoxWidth && currentFontSize > 1)
            {
                // 適切なフォントサイズを計算
                float newFontSize = currentFontSize * (float)textBoxWidth / textSize.Width;

                // フォントサイズを更新
                control.Font = new System.Drawing.Font(control.Font.FontFamily, newFontSize);
            }
            // テキストが短くなり、フォントサイズを大きくできる場合に元のサイズに戻す
            //else if (textSize.Width < textBoxWidth && currentFontSize < currentFontSize)
            //{
                // 元のフォントサイズに戻すロジックを実装
            //}
        }
        private void ChrSizeLabelChg(System.Windows.Forms.Label control)
        {
            // テキストボックスの幅
            int textBoxWidth = control.Width;

            // 現在のフォントサイズ
            float currentFontSize = control.Font.Size;

            // テキストの長さを測定
            System.Drawing.Font font = control.Font;
            System.Drawing.Size textSize = TextRenderer.MeasureText(control.Text, font);

            // テキストがテキストボックスの幅を超えた場合にフォントを縮小する
            if (textSize.Width > textBoxWidth && currentFontSize > 1)
            {
                // 適切なフォントサイズを計算
                float newFontSize = currentFontSize * (float)textBoxWidth / textSize.Width;

                // フォントサイズを更新
                control.Font = new System.Drawing.Font(control.Font.FontFamily, newFontSize);
            }
            // テキストが短くなり、フォントサイズを大きくできる場合に元のサイズに戻す
            //else if (textSize.Width < textBoxWidth && currentFontSize < currentFontSize)
            //{
            // 元のフォントサイズに戻すロジックを実装
            //}
        }
        private void txtUserName_TextChanged(object sender, EventArgs e)
        {
            txtUserName.Font = new System.Drawing.Font(txtUserName.Font.FontFamily, 18);
            ChrSizeTextBoxChg(txtUserName);
            UserName = txtUserName.Text;
            NowExcelFileName = "検査データ_" + SerialNo + "_" + UserName + "_" + ModelName;
            lblCreateExcelFileName.Text = NowExcelFileName;
        }
        private void lblCreateExcelFileName_TextChanged(object sender, EventArgs e)
        {
            lblCreateExcelFileName.Font = new System.Drawing.Font(lblCreateExcelFileName.Font.FontFamily, 20);
            ChrSizeLabelChg(lblCreateExcelFileName);
        }

        private void cmbTestVolt_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selectedText = cmbTestVolt.Text;

            HeaderData.VoltSpec = int.Parse(selectedText);
        }
    }
}

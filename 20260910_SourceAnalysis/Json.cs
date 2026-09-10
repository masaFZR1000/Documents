using Newtonsoft.Json;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using ScottPlot;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static EMS_Small_Chb.Form1;
using static EMS_Small_Chb.Json;

namespace EMS_Small_Chb
{
    internal static class Json
    {
        #region 変数定義
        public class HUMISET 
        {
            [JsonProperty("ステップ")]
            public int SetNo { get; set; } = 0;
            [JsonProperty("設定温度")]
            public float SetTemp { get; set; } = float.MinValue;
            [JsonProperty("設定湿度")]
            public float SetHumi { get; set; } = float.MinValue;
            [JsonProperty("湿度精度")]
            public float HumiAccuracy { get; set; } = float.MinValue;
            [JsonProperty("制御時間")]
            public float SetTime { get; set; } = float.MinValue;
        }
        public class TEMPSET
        {
            [JsonProperty("ステップ")]
            public int SetNo { get; set; } = 0;
            [JsonProperty("設定温度")]
            public float SetTemp { get; set; } = float.MinValue;
            [JsonProperty("制御時間")]
            public float SetTime { get; set; } = float.MinValue;
        }
        public class POLELIMIT
        {
            [JsonProperty("開始温度")]
            public float StartTemp { get; set; } = float.MinValue;
            [JsonProperty("終了温度")]
            public float EndTemp { get; set; } = float.MinValue;
            [JsonProperty("判定基準")]
            public float Judgment { get; set; } = float.MinValue;
        }
        public class HUMIACCURACY
        {
            [JsonProperty("設定温度")]
            public float SetTemp { get; set; } = float.MinValue;
            [JsonProperty("設定湿度")]
            public float SetHumi { get; set; } = float.MinValue;
            [JsonProperty("判定基準")]
            public float Judgment { get; set; } = float.MinValue;
        }
        public class MODELS
        {
            [JsonProperty("バーコード")]
            public string BarCode { get; set; } = "";
            [JsonProperty("型式")]
            public string Model { get; set; } = "";
            [JsonProperty("検査プログラム名")]
            public string ProgramName { get; set; } = "";
            [JsonProperty("検査チェックシート番号")]
            public string CheckSheetName { get; set; } = "";
            [JsonProperty("180℃仕様")]
            public bool b180Temp { get; set; } = false;
            [JsonProperty("観測窓")]
            public bool bChkWindow { get; set; } = false;
            [JsonProperty("内扉")]
            public bool bInWindow { get; set; } = false;
            [JsonProperty("湿度有")]
            public bool bHumi { get; set; } = false;
            [JsonProperty("下限温度")]
            public float LowTemp { get; set; } = float.MinValue;
            [JsonProperty("上限温度")]
            public float HighTemp { get; set; } = float.MaxValue;
            [JsonProperty("下限湿度")]
            public float LowHumi { get; set; } = float.MinValue;
            [JsonProperty("上限湿度")]
            public float HighHumi { get; set; } = float.MaxValue;
            [JsonProperty("温度湿度設定")]
            public IList<HUMISET> HumiSteps { get; set; } = new List<HUMISET>();
            [JsonProperty("温度設定")]
            public IList<TEMPSET> TempSteps { get; set; } = new List<TEMPSET>();
            [JsonProperty("温度極値到達時間上昇")]
            public POLELIMIT PolelLmitUp { get; set; } = new POLELIMIT();
            [JsonProperty("温度極値到達時間降下")]
            public POLELIMIT PolelLmitDown { get; set; } = new POLELIMIT();
            [JsonProperty("温度変化速度上昇")]
            public POLELIMIT ChgLimitUp { get; set; } = new POLELIMIT();
            [JsonProperty("温度変化速度降下")]
            public POLELIMIT ChgLimitDown { get; set; } = new POLELIMIT();

            [JsonProperty("温度精度低温")]
            public POLELIMIT AccuracyLowTemp { get; set; } = new POLELIMIT();
            [JsonProperty("温度精度中温")]
            public POLELIMIT AccuracyMiddleTemp { get; set; } = new POLELIMIT();
            [JsonProperty("温度精度高温")]
            public POLELIMIT AccuracyHighTemp { get; set; } = new POLELIMIT();
            [JsonProperty("温度変動低温")]
            public POLELIMIT VarLimitLowTemp { get; set; } = new POLELIMIT();
            [JsonProperty("温度変動中温")]
            public POLELIMIT VarLimitMiddleTemp { get; set; } = new POLELIMIT();
            [JsonProperty("温度変動高温")]
            public POLELIMIT VarLimitHighTemp { get; set; } = new POLELIMIT();
            [JsonProperty("湿度変動")]
            public POLELIMIT VarLimitHumi { get; set; } = new POLELIMIT();

            [JsonProperty("電圧群")]
            public IList<VOLTCURRENTS> Volts { get; set; } = new List<VOLTCURRENTS>();

        }
        public class MASHINE
        {
            [JsonProperty("製品仕様")]
            public IList<MODELS> Models { get; set; } = new List<MODELS>();
        }
        public class VOLTCURRENTS
        {
            [JsonProperty("電圧")]
            public string Volt { get; set; } = "";
            [JsonProperty("電流群")]
            public CURRENTS VoltCurrents { get; set; } = new CURRENTS();
        }

        public class CURRENTS
        {
            [JsonProperty("電流群")]
            public IList<CURRENT> Currents { get; set; } = new List<CURRENT>();
        }
        public class CURRENT 
        {
            [JsonProperty("電圧名称")]
            public string VoltName { get; set; } = "";
            [JsonProperty("名称")]
            public string Name { get; set; } = "";
            [JsonProperty("下限")]
            public float LowLimit { get; set; } = 0f;
            [JsonProperty("上限")]
            public float HighLimit { get; set; } = 0f;
            [JsonProperty("電力")]
            public float Power { get; set; } = 0f;
        }

        public class MEMBERS
        {
            [JsonProperty("社員番号")]
            public string No { get; set; } = "";
            [JsonProperty("名前")]
            public string Name { get; set; } = "";
            [JsonProperty("英文")]
            public string EnName { get; set; } = "";
        }
        public class MEMBER
        {
            [JsonProperty("担当者")]
            public IList<MEMBERS> Member { get; set; }
        }
        public class RECORDERS
        {
            [JsonProperty("品番")]
            public string No { get; set; } = "";
            [JsonProperty("型式")]
            public string Model { get; set; } = "";
            [JsonProperty("製造番号")]
            public string SerialNo { get; set; } = "";
        }
        public class RECORDER
        {
            [JsonProperty("記録計")]
            public IList<RECORDERS> Recorder { get; set; }
        }
        public class SYSTEM
        {
            [JsonProperty("電圧01")]
            public string Volt01 { get; set; } = "";
            [JsonProperty("電圧02")]
            public string Volt02 { get; set; } = "";
            [JsonProperty("電圧03")]
            public string Volt03 { get; set; } = "";
            [JsonProperty("電圧04")]
            public string Volt04 { get; set; } = "";
            [JsonProperty("電圧05")]
            public string Volt05 { get; set; } = "";
            [JsonProperty("電圧06")]
            public string Volt06 { get; set; } = "";
            [JsonProperty("電圧07")]
            public string Volt07 { get; set; } = "";
            [JsonProperty("電圧08")]
            public string Volt08 { get; set; } = "";
            [JsonProperty("COM_Port")]
            public string Com_Port { get; set; } = "";
            [JsonProperty("ＣＦカードドライブ")]
            public string CfDrive { get; set; } = "";
            [JsonProperty("最大レコード長")]
            public int Rec_Length { get; set; } = 1200;
            [JsonProperty("エクセルフォルダ")]
            public string ExcelPath { get; set; } = "";
            [JsonProperty("基幹システム取り込み")]
            public bool OracleRead { get; set; } = false;
            [JsonProperty("試験検査報告書書出し")]
            public bool TestReportRequired { get; set; } = false;
            [JsonProperty("検査データ集計書出し")]
            public bool Test_data_collection_Required { get; set; } = false;
            [JsonProperty("GraphRequired")]
            public bool GraphRequired { get; set; } = false;
            [JsonProperty("ClosedXMLUsed")]
            public bool ClosedXMLUsed { get; set; } = false;
            [JsonProperty("SerialNoCheck")]
            public bool SerialNoCheck { get; set; } = false;
            [JsonProperty("マスタフォルダ")]
            public string MasterPath { get; set; } = "";
            [JsonProperty("ＣＳＶフォルダ")]
            public string CsvPath { get; set; } = "";
            [JsonProperty("自動判定結果フォルダ")]
            public string ResultPath { get; set; } = "";
            [JsonProperty("検査データ集計_小型環境フォルダ")]
            public string Test_data_collection_Path { get; set; } = "";
            [JsonProperty("検査データ集計_小型環境ファイル")]
            public string Test_data_collection_FileName { get; set; } = "";
            [JsonProperty("温湿度特性判定")]
            public bool TempSpecJdge { get; set; } = false;
            [JsonProperty("第一ステップ")]
            public bool First_Step { get; set; } = true;
            [JsonProperty("制御温度ライン色")]
            public int OpeTempColor { get; set; } = 0;
            [JsonProperty("制御湿度ライン色")]
            public int OpeHumiColor { get; set; } = 0;
            [JsonProperty("中央乾球温度ライン色")]
            public int CentTempColor { get; set; } = 0;
            [JsonProperty("中央湿球温度ライン色")]
            public int CentHumiColor { get; set; } = 0;
            [JsonProperty("室内温度ライン色")]
            public int RoomTempColor { get; set; } = 0;
            [JsonProperty("相対湿度ライン色")]
            public int HumiColor { get; set; } = 0;
        }
        public class MASHINEDB
        {

            [JsonProperty("SeverAdr")]
            public String SeverAdr { get; set; } = "localhost";
            [JsonProperty("PortNo")]
            public int PortNo { get; set; } = 5432;
            [JsonProperty("UserID")]
            public string UserID { get; set; } = "espec";
            [JsonProperty("DataBaseName")]
            public string DataBaseName { get; set; } = "chamberdatas";
            [JsonProperty("Password")]
            public string Password { get; set; } = "espec";
            [JsonProperty("TableName")]
            public string TableName { get; set; } = "header";
            [JsonProperty("OptionTableName")]
            public string OptionTableName { get; set; } = "option";
            [JsonProperty("SerialNo")]
            public string SerialNo_FLName { get; set; } = "SerialNo";
            [JsonProperty("Model")]
            public string Model_FLName { get; set; } = "Model";
            [JsonProperty("Order")]
            public string Order_FLName { get; set; } = "Order";
            [JsonProperty("Volt")]
            public string Volt_FLName { get; set; } = "Volt";
            [JsonProperty("Freq")]
            public string Freq_FLName { get; set; } = "Freq";
            [JsonProperty("UserName")]
            public string UserName_FLName { get; set; } = "UserName";
            [JsonProperty("OptionCode_FLName")]
            public string OptionCode_FLName { get; set; } = "";
            [JsonProperty("OptionName_FLName")]
            public string OptionName_FLName { get; set; } = "";
        }
        public class HEADER
        {
            [JsonProperty("TableName")]
            public String TableName { get; set; } = "t_header_record";
            [JsonProperty("SerialNo")]
            public String SerialNo_FLName { get; set; } = "seri_no";
            [JsonProperty("created_at")]
            public String Createted_DateTime_FLName { get; set; } = "created_at";
        }
        public class DATAINFO
        {
            [JsonProperty("TableName")]
            public String TableName { get; set; } = "t_body_record";
            [JsonProperty("measured_at")]
            public String Measured_DateTime_FLName { get; set; } = "measured_at";
            [JsonProperty("created_at")]
            public String Createted_DateTime_FLName { get; set; } = "created_at";
            [JsonProperty("temp_sv")]
            public String temp_sv_FLName { get; set; } = "temp_sv";
            [JsonProperty("temp_pv")]
            public String temp_pv_FLName { get; set; } = "temp_pv";
            [JsonProperty("humi_sv")]
            public String humi_sv_FLName { get; set; } = "humi_sv";
            [JsonProperty("humi_pv")]
            public String humi_pv_FLName { get; set; } = "humi_pv";
            [JsonProperty("Room_Temp")]
            public String Room_Temp_FLName { get; set; } = "ai_00";
            [JsonProperty("center_temp")]
            public String Center_Temp_FLName { get; set; } = "ai_01";
            [JsonProperty("center_wet")]
            public String Center_Wet_FLName { get; set; } = "ai_02";

            [JsonProperty("den_data")]
            public String den_data { get; set; } = "den_data";
            [JsonProperty("rec_data1")]
            public String rec_data1 { get; set; } = "rec_data1";
            [JsonProperty("rec_data2")]
            public String rec_data2 { get; set; } = "rec_data2";

        }
        public class MODELDB
        {
            [JsonProperty("SeverAdr")]
            public String SeverAdr { get; set; } = "localhost";
            [JsonProperty("PortNo")]
            public int PortNo { get; set; } = 5432;
            [JsonProperty("UserID")]
            public string UserID { get; set; } = "espec";
            [JsonProperty("DataBaseName")]
            public string DataBaseName { get; set; } = "pj";
            [JsonProperty("Password")]
            public string Password { get; set; } = "espec";

            [JsonProperty("user_id")]
            public string user_id { get; set; } = "00000";

            [JsonProperty("test_category")]
            public string test_category { get; set; } = "1";

            [JsonProperty("ヘッダー情報")]
            public HEADER Header { get; set; } = new HEADER();
            [JsonProperty("データ情報")]
            public DATAINFO Data { get; set; } = new DATAINFO();
        }
        public class DATEBASE
        {
            [JsonProperty("製品情報")]
            public MASHINEDB MashineDb { get; set; } = new MASHINEDB();
            [JsonProperty("装置情報")]
            public MODELDB ModelDb { get; set; } = new MODELDB();
        }
        public class JUDGEMENT
        {
            [JsonProperty("目標温度")]
            public float CheckTemp { get; set; } = 1.0f;
            [JsonProperty("変化判定温度")]
            public float ChgJudgeTemp { get; set; } = 2.0f;
            [JsonProperty("最大最小幅")]
            public float MaxMinWidth { get; set; } = 0.5f;
            [JsonProperty("安定確認")]
            public float StabCheckCount { get; set; } = 10f;
            [JsonProperty("変化許容回数")]
            public int ToleranceLimit { get; set; } = 5;
            [JsonProperty("変動判定時間")]
            public int FluctJdgTimes { get; set; } = 30;
            [JsonProperty("判定後安定時間")]
            public int FluctJdgAfterTimes { get; set; } = 10;
        }
        public class OPMODELS
        {
            [JsonProperty("Name")]
            public String Name { get; set; } = "";
        }
        public class OPTIONS
        {
            [JsonProperty("OpCode")]
            public String OpCode { get; set; } = "";
            [JsonProperty("OpName")]
            public String OpName { get; set; } = "";
            [JsonProperty("Kind")]
            public String Kind { get; set; } = "";
            [JsonProperty("Value")]
            public String Value { get; set; } = "";
        }
        public class OPTIONCDS
        {
            [JsonProperty("ModelCD")]
            public String ModelCD { get; set; }
            [JsonProperty("Option")]
            public IList<OPTIONS> Option { get; set; }
            [JsonProperty("Model")]
            public IList<OPMODELS> Model { get; set; }
        }
        public class OPTION
        {
            [JsonProperty("OptionCD")]
            public IList<OPTIONCDS> OptionCD { get; set; }
        }

        #endregion 変数定義
        // 製品仕様ファイルの読み書き
        public static MASHINE MashineSpecRead(String cPath) 
        {
            MASHINE Mashine = new MASHINE();
            // 製品仕様ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "製品仕様.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                Mashine = JsonConvert.DeserializeObject<Json.MASHINE>(jsonReadData);
            }
            return Mashine;
        }
        public static void MashineSpecWrite(String cPath,MASHINE Mashine)
        {
            // 製品仕様ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(Mashine, Formatting.Indented);
            // 製品仕様.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "製品仕様.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        // 担当者ファイルの読み書き
        public static MEMBER MemberRead(String cPath)
        {
            MEMBER Member = new MEMBER();
            // 製品仕様ファイル読込み(JSON)
            using(var sr = new StreamReader(cPath + "担当者.Json", System.Text.Encoding.UTF8)) {
                var jsonReadData = sr.ReadToEnd();
                Member = JsonConvert.DeserializeObject<Json.MEMBER>(jsonReadData); 
            }
            return Member;
        }
        public static void MemberWrite(String cPath,MEMBER Member)
        {
            // 製品仕様ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(Member, Formatting.Indented);
            // 製品仕様.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "担当者.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        // 記録計ファイルの読み書き
        public static RECORDER RecorderRead(String cPath)
        {
            RECORDER Recorder = new RECORDER();
            // 製品仕様ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "記録計.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                Recorder = JsonConvert.DeserializeObject<Json.RECORDER>(jsonReadData);
            }
            return Recorder;

        }
        public static void RecorderWrite(String cPath,RECORDER Recorder)
        {
            // 製品仕様ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(Recorder, Formatting.Indented);
            // 製品仕様.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "記録計.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        public static DATEBASE DateBaseRead(String cPath)
        {
            DATEBASE DataBaseSet = new DATEBASE();
            // データベース設定ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "データベース情報.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                DataBaseSet = JsonConvert.DeserializeObject<Json.DATEBASE>(jsonReadData);
            }
            return DataBaseSet;
        }
        public static void DataBaseWrite(String cPath, DATEBASE DataBaseSet)
        {
            // データベース情報ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(DataBaseSet, Formatting.Indented);
            // データベース情報.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "データベース情報.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        // システム設定ファイルの読み書き
        public static SYSTEM SystemRead(String cPath)
        {
            SYSTEM SystemSet = new SYSTEM();
            // 製品仕様ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "システム設定.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                SystemSet = JsonConvert.DeserializeObject<Json.SYSTEM>(jsonReadData);
            }
            return SystemSet;
        }
        public static void SystemWrite(String cPath,SYSTEM SystemSet)
        {
            // 製品仕様ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(SystemSet, Formatting.Indented);
            // 製品仕様.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "システム設定.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        // 判定条件ファイルの読み書き
        public static JUDGEMENT JudgmentRead(String cPath)
        {
            JUDGEMENT Judment = new JUDGEMENT();
            // 製品仕様ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "判定条件.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                Judment = JsonConvert.DeserializeObject<Json.JUDGEMENT>(jsonReadData);
            }
            return Judment;
        }
        public static void JudgmentWrite(String cPath, JUDGEMENT Judment)
        {
            // 製品仕様ファイル書込み(JSON)
            // 1. JSON データにシリアライズ
            var jsonWriteData = JsonConvert.SerializeObject(Judment, Formatting.Indented);
            // 製品仕様.json を UTF-8 で書き込み用でオープン
            using (var sw = new StreamWriter(cPath + "判定条件.Json", false, System.Text.Encoding.UTF8))
            {
                // JSON データをファイルに書き込み
                sw.Write(jsonWriteData);
            }
        }
        public static OPTION OptionRead(String cPath)
        {
            OPTION OptionSet = new OPTION();
            // 製品仕様ファイル読込み(JSON)
            using (var sr = new StreamReader(cPath + "Option.Json", System.Text.Encoding.UTF8))
            {
                var jsonReadData = sr.ReadToEnd();
                OptionSet = JsonConvert.DeserializeObject<Json.OPTION>(jsonReadData);
            }
            return OptionSet;
        }
    }
}

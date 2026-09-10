
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;
using System.Xml;
using System.Xml.Linq;
using MathNet.Numerics;
using MathNet.Numerics.RootFinding;
using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.FormulaParsing.Excel.Functions.DateTime;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Math;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Numeric;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
using PdfSharp.Drawing;
using PdfSharp.Pdf;
using ScottPlot;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel;
using static EMS_Small_Chb.Form1;
using static EMS_Small_Chb.Json;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using OfficeOpenXml.Drawing;

namespace EMS_Small_Chb
{
    internal static class Excel_Operation
    {
        // スレッドセーフなキュー
        public static ConcurrentQueue<string> DataQueue = new ConcurrentQueue<string>();

        public static string sSendMsg = "";

        public static MASHINE MashineSpecRead(String cPath)
        {
            int PRGNSPOS = 0; int BRCSPOS = 1; int HUMISPOS = 2; int TEMPSPOS = 14; int POLESPOS = 29;
            int LOWTEMP = 22; int HIGHTEMP = 14; int LOWHUMI = 9; int HIGHHUMI = 12;
            int CHGSPOS = 35; int TEMPACUSPOS = 43; int TEMPFRUCSPOS = 46; int HUMIACUSPOS = 49;
            int HUMIFRUCSPOS = 53; int CHKSHEETSPOS = 54; int MODELSPOS = 55; int CURRENTSPOS = 56;
            string[] V_Name = new String[]{ "100V","115V","200V","220V","230V"};
            bool bHeader = false;

            MASHINE Mashine = new MASHINE();
            // 製品仕様ファイル読込み(Excel)
            string file = cPath + "製品仕様Jsonファイル設定マスタ.xlsx";
            string sheet = "検査条件_判定条件表";
            var allRows = new List<Dictionary<string, string>>();
            var headers = new List<string>();
            Console.WriteLine($"読み込み中: {Path.GetFileName(file)}");

            // ファイル情報を読み込む
            var InputExcelFile = new FileInfo(file);
            using (ExcelPackage package = new ExcelPackage(InputExcelFile))
            {
                ExcelWorkbook wb = package.Workbook;
                // 取得 (シート名指定)
                ExcelWorksheet ws = wb.Worksheets[sheet];

                // B1から右下の範囲を取得
                // シートの使用範囲（右下）を取得
                var dimension = ws.Dimension;

                // B1から右下の範囲を取得
                int startRow = 1;
                int startCol = 2; // "B"列
                int endRow = dimension.End.Row;
                int endCol = dimension.End.Column;

                var range = ws.Cells[startRow, startCol, endRow, endCol];

                // 2次元リストとしてデータを取得
                var table = new List<List<string>>();

                string sCell = "";
                for (int row = 1; row < endRow; row++)
                {
                    MODELS ms = new MODELS();
                    var rowData = new List<string>();
                    for (int col = 1; col < endCol; col++)
                    {
                        if (range[row + 1, col + 1].Value != null) sCell = range[row + 1, col + 1].Value.ToString();
                        else sCell = "";
                        Console.WriteLine(sCell);
                        if (sCell == "END" || sCell == "SH関連検査プログラム名" || sCell == "SU関連検査プログラム名") { bHeader = true; break; }
                        rowData.Add(sCell);
                        bHeader = false;
                    }
                    if (!bHeader) 
                    {
                        // tableに1行（製品データ）追加
                        table.Add(rowData);

                        // 製品情報データ設定
                        ms.ProgramName = rowData[PRGNSPOS].ToString();
                        ms.BarCode = rowData[BRCSPOS].ToString();
                        ms.CheckSheetName = rowData[CHKSHEETSPOS].ToString();
                        ms.Model = rowData[MODELSPOS].ToString();

                        if (ms.ProgramName.Contains("180")) ms.b180Temp = true; else ms.b180Temp = false;
                        if (ms.BarCode.Contains("SH")) ms.bHumi = true; else ms.bHumi = false;
                        if (ms.ProgramName.Contains("観窓")) ms.bChkWindow = true; else ms.bChkWindow = false;
                        if (ms.ProgramName.Contains("内扉")) ms.bInWindow = true; else ms.bInWindow = false;

                        ms.LowTemp = CheckFloatValue(rowData[LOWTEMP].ToString());
                        ms.HighTemp = CheckFloatValue(rowData[HIGHTEMP].ToString());
                        ms.LowHumi = CheckFloatValue(rowData[LOWHUMI].ToString());
                        ms.HighHumi = CheckFloatValue(rowData[HIGHHUMI].ToString());

                        // 湿度制御設定登録
                        for (int i = 0; i < 4; ++i)
                            {
                                if (rowData[HUMISPOS + i] == "") break;
                                HUMISET hs = new HUMISET();
                                hs.SetTemp = CheckFloatValue(rowData[HUMISPOS + i * 3].ToString());
                                hs.SetHumi = CheckFloatValue(rowData[HUMISPOS + 1 + i * 3].ToString());
                                hs.HumiAccuracy = CheckFloatValue(rowData[HUMIACUSPOS + i].ToString());
                                hs.SetTime = CheckFloatValue(rowData[HUMISPOS + 2 + i * 3].ToString());
                                ms.HumiSteps.Add(hs);
                            }
                        // 温度制御設定登録
                        for (int i = 0; i < 7; ++i)
                        {
                            if (rowData[TEMPSPOS + i * 2] == "") break;
                            TEMPSET hs = new TEMPSET();
                            hs.SetTemp = CheckFloatValue(rowData[TEMPSPOS + i * 2].ToString());
                            hs.SetTime = CheckFloatValue(rowData[TEMPSPOS + 1 + i * 2].ToString());
                            ms.TempSteps.Add(hs);
                        }
                        // 極値設定登録
                        ms.PolelLmitUp.StartTemp = CheckFloatValue(rowData[POLESPOS].ToString());
                        ms.PolelLmitUp.EndTemp = CheckFloatValue(rowData[POLESPOS + 1].ToString());
                        ms.PolelLmitUp.Judgment = CheckFloatValue(rowData[POLESPOS + 2].ToString());
                        ms.PolelLmitDown.StartTemp = CheckFloatValue(rowData[POLESPOS + 3].ToString());
                        ms.PolelLmitDown.EndTemp = CheckFloatValue(rowData[POLESPOS + 4].ToString());
                        ms.PolelLmitDown.Judgment = CheckFloatValue(rowData[POLESPOS + 5].ToString());

                        // 変加速度上昇設定登録
                        ms.ChgLimitUp.StartTemp = CheckFloatValue(rowData[CHGSPOS].ToString());
                        ms.ChgLimitUp.EndTemp = CheckFloatValue(rowData[CHGSPOS + 1].ToString());
                        // ??? 時間か？変化速度か？
                        ms.ChgLimitUp.Judgment = CheckFloatValue(rowData[CHGSPOS + 3].ToString());

                        // 変加速度降下設定登録
                        ms.ChgLimitDown.StartTemp = CheckFloatValue(rowData[CHGSPOS + 4].ToString());
                        ms.ChgLimitDown.EndTemp = CheckFloatValue(rowData[CHGSPOS + 5].ToString());
                        // ??? 時間か？変化速度か？
                        ms.ChgLimitDown.Judgment = CheckFloatValue(rowData[CHGSPOS + 7].ToString());

                        // 温度精度
                        ms.AccuracyLowTemp.Judgment = CheckFloatValue(rowData[TEMPACUSPOS + 0].ToString());
                        ms.AccuracyMiddleTemp.Judgment = CheckFloatValue(rowData[TEMPACUSPOS + 1].ToString());
                        ms.AccuracyHighTemp.Judgment = CheckFloatValue(rowData[TEMPACUSPOS + 2].ToString());

                        ms.AccuracyLowTemp.StartTemp = ms.LowTemp;
                        ms.AccuracyLowTemp.EndTemp = 100f;
                        ms.AccuracyMiddleTemp.StartTemp = 100.1f;
                        ms.AccuracyMiddleTemp.EndTemp = ms.HighTemp;
                        if (ms.HighTemp == 180) 
                        {
                            ms.AccuracyMiddleTemp.EndTemp = 150f;
                            ms.AccuracyHighTemp.StartTemp = 150.1f;
                            ms.AccuracyHighTemp.EndTemp = ms.HighTemp;
                        }

                        // 温度変動
                        ms.VarLimitLowTemp.StartTemp = ms.LowTemp;
                        ms.VarLimitLowTemp.EndTemp = 100f;
                        ms.VarLimitMiddleTemp.StartTemp = 100.1f;
                        ms.VarLimitMiddleTemp.EndTemp = ms.HighTemp;

                        ms.VarLimitLowTemp.Judgment = CheckFloatValue(rowData[TEMPFRUCSPOS + 0].ToString());
                        ms.VarLimitMiddleTemp.Judgment = CheckFloatValue(rowData[TEMPFRUCSPOS + 1].ToString());
                        ms.VarLimitHighTemp.Judgment = CheckFloatValue(rowData[TEMPFRUCSPOS + 2].ToString());

                        // 湿度変動
                        ms.VarLimitHumi.Judgment = CheckFloatValue(rowData[HUMIFRUCSPOS + 0].ToString());

                        // 電流判定情報
                        int iPos = CURRENTSPOS;
                        for (int j = 0; j < 5; ++j)
                        {
                            VOLTCURRENTS Vc = new VOLTCURRENTS();
                            Vc.Volt = V_Name[j];
                            CURRENTS Cus = new CURRENTS();
                            for (int i = 0; i < 5; ++i)
                            {
                                CURRENT Cu = new CURRENT();
                                Cu.VoltName = Vc.Volt + "_電流_" + (i+1).ToString();
                                Cu.Name = rowData[iPos].ToString();
                                Cu.LowLimit = CheckFloatValue(rowData[iPos + 1].ToString());
                                Cu.HighLimit = CheckFloatValue(rowData[iPos + 2].ToString());
                                Cu.Power = CheckFloatValue(rowData[iPos + 3].ToString());
                                Cus.Currents.Add(Cu);
                                iPos += 4;
                            }
                            Vc.VoltCurrents = Cus;
                            ms.Volts.Add(Vc);

                        }
                        Mashine.Models.Add(ms); 
                    }
                }

                // 動作確認：中身を表示
                foreach (var row in table)
                {
                    Console.WriteLine(string.Join(", ", row));
                }
                // 製品仕様に設定
            }
            return Mashine;
        }
        public static float CheckFloatValue(string str) 
        {
            float fValue = 0;
            if (str.Contains(":")) 
            {
                string[] hm = new string[2];
                hm =str.Split(':');
                fValue = float.Parse(hm[0]) + float.Parse(hm[1]) / 60;

            }
            else if(str != "") float.TryParse(str, out fValue);
            return fValue;
        }
        public static async Task<bool> ExcelFileWrite(Json.SYSTEM SystemSet, ExcelData Ev, RecData RecValue, List<CURRENT> CurrentInfo) 
        {
            bool bResult = false;
            DialogResult result = DialogResult.Yes;

            string cMsg = "";
            // 解析結果データ準備
            cMsg = "エクセル処理起動\r\n";
            WriteLog(cMsg);
            DataQueue.Enqueue(cMsg);

            String DocPath = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal);

            String DistFileName = @DocPath + "\\" + "一時ファイル.xlsx";

            if (System.IO.File.Exists(DistFileName)) System.IO.File.Delete(DistFileName);

            String cPath = SystemSet.MasterPath;

            String SourceFileName = cPath;

            bool bGrapfCopyErr = false;

            if (ChkNowSpec.HumiUse) SourceFileName += "検査データ_湿度用.xlsx";
            else SourceFileName += "検査データ_温度用.xlsx";
            string sheet = "測定データ";

            try
            {
                #region 検査結果書込み
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
                // Excel Book
                // ファイル情報を読み込む
                // 解析結果データ準備
                cMsg = "検査データ書き込みファイル準備\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);
                var InputExcelFile = new FileInfo(DistFileName);
                using (ExcelPackage package = new ExcelPackage(InputExcelFile))
                {
                    ExcelWorkbook wb = package.Workbook;
                    // 取得 (シート名指定)
                    ExcelWorksheet ws = wb.Worksheets[sheet];

                    // 解析結果データ準備
                    cMsg = "解析結果データ準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);

                    // ファイル情報
                    ws.Cells["B27"].Value = Ev.HumiEndRecord;
                    ws.Cells["B28"].Value = Ev.TempEndRecord;
                    ws.Cells["B29"].Value = Ev.NowCSVFilePath;
                    ws.Cells["J28"].Value = SystemSet.ExcelPath;
                    ws.Cells["J29"].Value = Ev.NowExcelFileName;

                    // 装置情報
                    ws.Cells["B3"].Value = Ev.ModelName;
                    ws.Cells["B4"].Value = Ev.SerialNo;
                    ws.Cells["B5"].Value = Ev.UserName;

                    ws.Cells["G27"].Value = ChkNowSpec.SpecMinTemp;
                    ws.Cells["G28"].Value = ChkNowSpec.SpecMaxTemp;

                    // オプション情報
                    ws.Cells["B6"].Value = "なし";
                    if (Ev.CheckWindow) ws.Cells["B6"].Value = "あり";
                    ws.Cells["B7"].Value = "なし";
                    if (Ev.InnerDoor) ws.Cells["B7"].Value = "あり";
                    ws.Cells["B8"].Value = "なし";
                    if (Ev.OperationHole) ws.Cells["B8"].Value = "あり";
                    ws.Cells["B21"].Value = "なし";
                    if (ChkNowSpec.Temp180Use) ws.Cells["B21"].Value = "あり";

                    // 検査プログラム名
                    ws.Cells["B22"].Value = ChkNowSpec.InspectPrgName;

                    // 検査チェックシート名
                    ws.Cells["B24"].Value = ChkNowSpec.CheckSheetName;

                    // 2024/11/06　改善の為、解除
                    // 2024/08/22　改善の為、一時保留
                    // 平均室温書出し
                    ws.Cells["B23"].Value = sValue(Ev.RoomTempAverage,1);
                    // 2024/08/22　改善の為、一時保留
                    // 2024/11/06　改善の為、解除

                    // 付加仕様情報
                    ws.Cells["B9"].Value = Ev.AdditionalSpecifications;
                    // ケーブル孔仕様情報
                    ws.Cells["B10"].Value = Ev.CableHole[0];
                    ws.Cells["B11"].Value = Ev.CableHole[1];
                    ws.Cells["B12"].Value = Ev.CableHole[2];
                    ws.Cells["B13"].Value = Ev.CableHole[3];
                    // 検査状況情報
                    ws.Cells["B14"].Value = Ev.TestingPerson;
                    ws.Cells["B25"].Value = Ev.TestingPersonEn;

                    ws.Cells["B15"].Value = DateTime.Parse(Ev.TestingDate);
                    ws.Cells["B16"].Value = Ev.TestingVolt;
                    ws.Cells["B17"].Value = Ev.TestingFreq;
                    // 記録計情報
                    ws.Cells["B18"].Value = Ev.TestingRecorderNo;
                    ws.Cells["B19"].Value = Ev.TestingRecorderModel;
                    ws.Cells["B20"].Value = Ev.TestingRecorderSerialNo;

                    // 解析結果データ準備
                    cMsg = "自動判定条件準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);

                    // 自動判定条件

                    ws.Cells["W2"].Value = JdgCnd.ChgJudgeTemp.ToString("F1");

                    ws.Cells["W3"].Value = JdgCnd.MaxMinWidth.ToString("F1");
                    ws.Cells["W4"].Value = JdgCnd.StableCount.ToString("");
                    ws.Cells["W5"].Value = JdgCnd.ChgToleranceLimit.ToString("");
                    ws.Cells["W6"].Value = JdgCnd.JdgTimes.ToString("");
                    ws.Cells["W7"].Value = JdgCnd.JdgAfterTime.ToString("");

                    // 製品仕様（極値到達時間上昇）
                    ws.Cells["E3"].Value = float.Parse(ChkNowSpec.PoleUpSpec.ToString(""));
                    ws.Cells["F3"].Value = float.Parse(ChkNowSpec.PoleUpStartTemp.ToString("F1")).Round(2);
                    ws.Cells["G3"].Value = float.Parse(ChkNowSpec.PoleUpEndTemp.ToString("F1")).Round(2);
                    // 製品仕様（極値到達時間降下）
                    ws.Cells["E4"].Value = float.Parse(ChkNowSpec.PoleDownSpec.ToString(""));
                    ws.Cells["F4"].Value = float.Parse(ChkNowSpec.PoleDownStartTemp.ToString("F1")).Round(2);
                    ws.Cells["G4"].Value = float.Parse(ChkNowSpec.PoleDownEndTemp.ToString("F1")).Round(2);
                    // 製品仕様（変化速度上昇）
                    ws.Cells["E5"].Value = float.Parse(ChkNowSpec.ChgUpSpec.ToString("F1")).Round(2);
                    ws.Cells["F5"].Value = float.Parse(ChkNowSpec.ChgUpStartTemp.ToString("F1")).Round(2);
                    ws.Cells["G5"].Value = float.Parse(ChkNowSpec.ChgUpEndTemp.ToString("F1")).Round(2);
                    // 製品仕様（変化速度降下）
                    ws.Cells["E6"].Value = float.Parse(ChkNowSpec.ChgDownSpec.ToString("F1")).Round(2);
                    ws.Cells["F6"].Value = float.Parse(ChkNowSpec.ChgDownStartTemp.ToString("F1")).Round(2);
                    ws.Cells["G6"].Value = float.Parse(ChkNowSpec.ChgDownEndTemp.ToString("F1")).Round(2);
                    // 温度変動低温製品仕様
                    ws.Cells["E7"].Value = float.Parse(ChkNowSpec.VarTemp1Spec.ToString("F1")).Round(2);
                    ws.Cells["F7"].Value = float.Parse(ChkNowSpec.VarTemp1StartTemp.ToString("F1")).Round(2);
                    ws.Cells["G7"].Value = float.Parse(ChkNowSpec.VarTemp1EndTemp.ToString("F1")).Round(2);
                    // 温度変動高温製品仕様
                    ws.Cells["E8"].Value = float.Parse(ChkNowSpec.VarTemp2Spec.ToString("F1")).Round(2);
                    ws.Cells["F8"].Value = float.Parse(ChkNowSpec.VarTemp2StartTemp.ToString("F1")).Round(2);
                    ws.Cells["G8"].Value = float.Parse(ChkNowSpec.VarTemp2EndTemp.ToString("F1")).Round(2);
                    // 湿度変動製品仕様
                    ws.Cells["E9"].Value = float.Parse(ChkNowSpec.VarHumiSpec.ToString("F1")).Round(2);

                    // 検査結果（極値到達時間上昇）
                    ws.Cells["E12"].Value = float.Parse(ExtUpValue.ResultValue.ToString(""));
                    ws.Cells["G12"].Value = "NG";
                    if (ChkNowSpec.PoleUpSpec >= ExtUpValue.ResultValue) ws.Cells["G12"].Value = "GOOD";
                    else { if (Math.Abs(ExtUpValue.ResultValue) > 1000) ws.Cells["E12"].Value = "--"; }

                    // 検査結果（極値到達時間降下）
                    ws.Cells["E13"].Value = float.Parse(ExtDownValue.ResultValue.ToString(""));
                    ws.Cells["G13"].Value = "NG";
                    if (ChkNowSpec.PoleDownSpec >= ExtDownValue.ResultValue) ws.Cells["G13"].Value = "GOOD";
                    else { if (Math.Abs(ExtDownValue.ResultValue) > 1000) ws.Cells["E13"].Value = "--"; }

                    // 検査結果（変化速度上昇）
                    ws.Cells["E14"].Value = float.Parse((ChgUpValue.EndXPos - ChgUpValue.StartXPos).ToString(""));
                    ws.Cells["F14"].Value = float.Parse(ChgUpValue.ResultValue.ToString("F1")).Round(2);
                    ws.Cells["G14"].Value = "NG";
                    if (ChkNowSpec.ChgUpSpec <= ChgUpValue.ResultValue) ws.Cells["G14"].Value = "GOOD";
                    else { if (Math.Abs(ChgUpValue.EndXPos - ChgUpValue.StartXPos) > 1000) ws.Cells["E14"].Value = "---"; }

                    // 検査結果（変化速度降下）
                    ws.Cells["E15"].Value = float.Parse((ChgDownValue.EndXPos - ChgDownValue.StartXPos).ToString(""));
                    ws.Cells["F15"].Value = float.Parse(ChgDownValue.ResultValue.ToString("F1")).Round(2);
                    ws.Cells["G15"].Value = "NG";
                    if (ChkNowSpec.ChgDownSpec <= ChgDownValue.ResultValue) ws.Cells["G15"].Value = "GOOD";
                    else { if (Math.Abs(ChgDownValue.EndXPos - ChgDownValue.StartXPos) > 1000) ws.Cells["E15"].Value = "---"; }

                    // 解析結果データ準備
                    cMsg = "検査パターン準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);

                    // 検査パターン
                    for (int i = 0; i < TEMPSETSTEPS; ++i)
                    {
                        ws.Cells["J" + (i + 4).ToString()].Value = sValue(Tempset[i].SetTemp,0);
                        ws.Cells["K" + (i + 4).ToString()].Value = Tempset[i].SetTime.ToString("F1");
                    }
                    if (ChkNowSpec.HumiUse)
                    {
                        for (int i = 0; i < HUMISETSTEPS; ++i)
                        {
                            ws.Cells["J" + (i + 12).ToString()].Value = sValue(Humiset[i].SetTemp,0);
                            ws.Cells["k" + (i + 12).ToString()].Value = Humiset[i].SetHumi.ToString("");
                            ws.Cells["L" + (i + 12).ToString()].Value = Humiset[i].SetTime.ToString("F1");
                        }
                    }

                    // 温度精度・変動幅
                    ws.Cells["N4"].Value = ChkNowSpec.AccuracyTemp1StartTemp.ToString("F1");
                    ws.Cells["O4"].Value = ChkNowSpec.AccuracyTemp1EndTemp.ToString("F1");
                    ws.Cells["P4"].Value = "±" + ChkNowSpec.AccuracyTemp1Spec.ToString("F1");

                    ws.Cells["R4"].Value = ChkNowSpec.VarTemp1StartTemp.ToString("F1");
                    ws.Cells["S4"].Value = ChkNowSpec.VarTemp1EndTemp.ToString("F1");
                    ws.Cells["T4"].Value = ChkNowSpec.VarTemp1Spec.ToString("F1");

                    ws.Cells["N5"].Value = ChkNowSpec.AccuracyTemp2StartTemp.ToString("F1");
                    ws.Cells["O5"].Value = ChkNowSpec.AccuracyTemp2EndTemp.ToString("F1");
                    ws.Cells["P5"].Value = "±" + ChkNowSpec.AccuracyTemp2Spec.ToString("F1");

                    ws.Cells["R5"].Value = ChkNowSpec.VarTemp2StartTemp.ToString("F1");
                    ws.Cells["S5"].Value = ChkNowSpec.VarTemp2EndTemp.ToString("F1");
                    ws.Cells["T5"].Value = ChkNowSpec.VarTemp2Spec.ToString("F1");

                    if (ChkNowSpec.Temp180Use)
                    {
                        ws.Cells["N6"].Value = ChkNowSpec.AccuracyTemp3StartTemp.ToString("F1");
                        ws.Cells["O6"].Value = ChkNowSpec.AccuracyTemp3EndTemp.ToString("F1");
                        ws.Cells["P6"].Value = "±" + ChkNowSpec.AccuracyTemp3Spec.ToString("F1");

                        ws.Cells["R6"].Value = ChkNowSpec.VarTemp3StartTemp.ToString("F1");
                        ws.Cells["S6"].Value = ChkNowSpec.VarTemp3EndTemp.ToString("F1");
                        ws.Cells["T6"].Value = ChkNowSpec.VarTemp3Spec.ToString("F1");
                    }

                    // 解析結果データ準備
                    cMsg = "検査結果準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);

                    // 湿度精度・変動幅
                    if (ChkNowSpec.HumiUse)
                    {
                        for (int i = 0; i < HumiFluct.Count; ++i)
                        {
                            ws.Cells["P" + (i + 12).ToString()].Value = "±" + Humiset[i].HumiAccuracy.ToString("F1");
                        }
                        ws.Cells["Q12"].Value = "3.0 %RH";
                    }
                    ws.Cells["I20"].Value = "+" + ChkNowSpec.SpecMaxTemp.ToString();
                    ws.Cells["I22"].Value = ChkNowSpec.SpecMinTemp.ToString();
                    ws.Cells["I25"].Value = (ChkNowSpec.SpecMinTemp - 3).ToString() + " ℃に到達している事";
                    ws.Cells["I26"].Value = "+157 ℃に到達している事";
                    if (ChkNowSpec.Temp180Use) ws.Cells["I26"].Value = "+185 ℃に到達している事";
                    ws.Cells["N25"].Value = ChkNowSpec.SpecMinTemp.ToString() + " ℃に到達している事";

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

                            ws.Cells["L20"].Value = "NG";
                            ws.Cells["L21"].Value = "NG";
                            ws.Cells["L22"].Value = "NG";
                            ws.Cells["M25"].Value = "NG";
                            ws.Cells["M26"].Value = "NG";
                            ws.Cells["R25"].Value = "NG";

                            foreach (WIDTH_FLUCT fluct in TempFluct)
                            {
                                float TempWidth = ChkNowSpec.AccuracyTemp1Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp1EndTemp) TempWidth = ChkNowSpec.AccuracyTemp2Spec;
                                if (fluct.TempSvAve > ChkNowSpec.AccuracyTemp2EndTemp) TempWidth = ChkNowSpec.AccuracyTemp3Spec;

                                if (fluct.TempSvAve >= ChkNowSpec.SpecMaxTemp - TempWidth / 2 && fluct.TempSvAve <= ChkNowSpec.SpecMaxTemp + TempWidth / 2)
                                {
                                    // 最大温度の安定があった。
                                    MaxSVValue = ChkNowSpec.SpecMaxTemp;
                                    ws.Cells["J20"].Value = fluct.TempSvAve.ToString("F1");
                                    ws.Cells["K20"].Value = (fluct.TempSvAve - MaxSVValue).ToString("F1");
                                    if (Math.Abs(fluct.TempSvAve - MaxSVValue) < 1.2) ws.Cells["L20"].Value = "GOOD";
                                }

                                else if (fluct.TempSvAve >= ChkNowSpec.SpecMinTemp - TempWidth && fluct.TempSvAve <= ChkNowSpec.SpecMinTemp + TempWidth)
                                {
                                    MinSVValue = ChkNowSpec.SpecMinTemp;
                                    // 最小温度の安定があった。
                                    ws.Cells["J22"].Value = fluct.TempSvAve.ToString("F1");
                                    ws.Cells["K22"].Value = (fluct.TempSvAve - MinSVValue).ToString("F1") + " ℃";
                                    if (Math.Abs(fluct.TempSvAve - MinSVValue) < 1.2) ws.Cells["L22"].Value = "GOOD";

                                    // 槽中央温度で判定
                                    ws.Cells["Q25"].Value = fluct.TempPvAve.ToString("F1");
                                    if (Math.Abs(fluct.TempPvAve - ChkNowSpec.SpecMinTemp) < TempWidth) ws.Cells["R25"].Value = "GOOD";
                                }
                                else if (fluct.TempSvAve > 20 - TempWidth && fluct.TempSvAve < 20 + TempWidth)
                                {
                                    // 20℃の安定があった。
                                    ws.Cells["J21"].Value = fluct.TempSvAve.ToString("F1");
                                    ws.Cells["K21"].Value = (fluct.TempSvAve - 20).ToString("F1");
                                    if (Math.Abs(fluct.TempSvAve - 20) < 1.2) ws.Cells["L21"].Value = "GOOD";
                                }

                                if (fluct.TempPvStd > MaxValue && fluct.Exist) MaxValue = fluct.TempPvStd;
                                if (!fluct.Judge && fluct.Exist) fNG = true;

                                // 取得データ書出し
                                if (fluct.Exist)
                                {
                                    int chkTemp = FluctSetTemp(fluct.EndValue);
                                    if (chkTemp == int.MinValue) chkTemp = (int)fluct.EndValue;
                                    ws.Cells["D" + (19 + i).ToString()].Value = chkTemp.ToString("");
                                    ws.Cells["E" + (19 + i).ToString()].Value = fluct.TempPvStd.ToString("F1");
                                }
                                ++i;
                            }
                            ws.Cells["E16"].Value = MaxValue.ToString("F1");
                            ws.Cells["F16"].Value = "NG";
                            if (!fNG) ws.Cells["F16"].Value = "GOOD";

                            // 最高最高温度確認
                            ws.Cells["M26"].Value = "NG";
                            if (MaxTemp.Exist)
                            {
                                ws.Cells["L26"].Value = MaxTemp.TempSvAve.ToString("F1");
                                ws.Cells["M26"].Value = "GOOD";
                            }

                            // 最低最低温度確認
                            ws.Cells["M25"].Value = "NG";
                            if (MinTemp.Exist)
                            {
                                ws.Cells["L25"].Value = MinTemp.TempSvAve.ToString("F1");
                                ws.Cells["M25"].Value = "GOOD";
                            }
                        }

                        // 検査結果（湿度変動）
                        if (HumiFluct.Count != 0)
                        {
                            float MaxValue = float.MinValue;
                            bool fNG = false; int i = 0;
                            ws.Cells["V20"].Value = "NG";
                            ws.Cells["V22"].Value = "NG";
                            ws.Cells["V24"].Value = "NG";
                            ws.Cells["V26"].Value = "NG";
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
                                        ws.Cells["U" + (j * 2 + 20).ToString()].Value = fluct.TempSvAve.ToString("F1");
                                        ws.Cells["U" + (j * 2 + 21).ToString()].Value = fluct.HumiSvAve.ToString("F1");

                                        ws.Cells["W" + (j * 2 + 20).ToString()].Value = fluct.TempPvAve.ToString("F1");
                                        ws.Cells["W" + (j * 2 + 21).ToString()].Value = fluct.WetPvAve.ToString("F1");
                                        ws.Cells["X" + (j * 2 + 21).ToString()].Value = fluct.HumiPvAve.ToString("F1");
                                        if (fluct.HumiSvAve > Humiset[j].SetHumi - 3 && fluct.HumiSvAve < Humiset[j].SetHumi + 3) ws.Cells["V" + (j * 2 + 20).ToString()].Value = "GOOD";
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
                                    ws.Cells["F" + (19 + i).ToString()].Value = chkHumi.ToString("");
                                    ws.Cells["G" + (19 + i).ToString()].Value = fluct.HumiPvStd.ToString("F1");
                                }
                                ++i;
                            }
                            ws.Cells["E17"].Value = MaxValue.ToString("F1");
                            ws.Cells["F17"].Value = "NG";
                            if (!fNG) ws.Cells["F17"].Value = "GOOD";
                        }
                    }
#if False
                    // 解析結果データ準備
                    cMsg = "確認グラフ画像複写準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);

                    #region グラフデータ
                    int iErrCnt = 0;
                    while (true)
                    {
                        try
                        {
                            // 確認グラフデータ複写
                            Bitmap bmpImage = Ev.formsPlot1.Plot.GetBitmap();
                            Clipboard.SetImage(bmpImage);
                            // セル位置（例：B2）に画像を挿入
                            var picture = ws.Drawings.AddPicture("ClipboardImage", bmpImage);
                            // セル位置を指定
                            picture.SetPosition(30, 0, 8, 0); // 行=30, 列=8 → "I31"
                            // サイズを変更（比率指定）
                            //picture.SetSize(67); // 67%
                            break;
                        }
                        catch (Exception ex)
                        {
                            ++iErrCnt;
                            string sMsg = "確認グラフデータ複写異常\r\n" + ex.Message;
                            WriteLog(cMsg);
                            DataQueue.Enqueue(cMsg);
                            Console.WriteLine(iErrCnt.ToString() + ":" + sMsg);
                            if (iErrCnt > 3)
                            {
                                bGrapfCopyErr = true;

                                WriteLog(sMsg);
                                // 異常時の背景色を赤色
                                // txtMessage.BackColor = Color.Red;
                                DataQueue.Enqueue(cMsg);
                                Console.WriteLine(sMsg);
                                break;
                            }
                        }
                    }
                    string SheetName = "";
                    // グラフデータ準備
                    cMsg = "グラフデータ準備\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);
                    #endregion グラフデータ
#endif

                    // xlApp.Run("CSVファイル読込");
                    // CSVのデータを読み込んでセルに書き込む必要が有る。
                    // 測定データシートにCSVファイルのデータを転記
                    ws = wb.Worksheets[sheet]; int iLineNo = 31;
                    ExcelWorksheet wstemp = wb.Worksheets["温度グラフデータ"]; int iTempNo = 4;
                    ExcelWorksheet wshumi = wb.Worksheets["湿度グラフデータ"]; int iHumiNo = 4;
                    ExcelWorksheet wsData = wb.Worksheets["検査結果"];
                    ExcelWorksheet wsDataEn = wb.Worksheets["検査結果（英文）"];

                    int iRecNo = 0;
                    while (true) 
                    {
                        if (iRecNo >= Ev.EndRecord) break;
                        { // 測定データ記録
                            ws.Cells["A" + iLineNo.ToString()].Value = RecValue.Date[iRecNo];
                            ws.Cells["B" + iLineNo.ToString()].Value = RecValue.Time[iRecNo];
                            ws.Cells["C" + iLineNo.ToString()].Value = RecValue.TempSV[iRecNo];
                            ws.Cells["D" + iLineNo.ToString()].Value = RecValue.HumiSV[iRecNo];
                            ws.Cells["E" + iLineNo.ToString()].Value = RecValue.TempPV[iRecNo];
                            ws.Cells["F" + iLineNo.ToString()].Value = RecValue.WetTempPV[iRecNo];
                            ws.Cells["G" + iLineNo.ToString()].Value = RecValue.RoomTemp[iRecNo];
                            iLineNo++;
                        }// 湿度データ記録
                        if ( ChkNowSpec.HumiUse && iRecNo < Ev.HumiEndRecord) 
                        {
                            wshumi.Cells["C" + iHumiNo.ToString()].Value = RecValue.Date[iRecNo];
                            wshumi.Cells["D" + iHumiNo.ToString()].Value = RecValue.Time[iRecNo];
                            wshumi.Cells["E" + iHumiNo.ToString()].Value = RecValue.TempSV[iRecNo];
                            wshumi.Cells["F" + iHumiNo.ToString()].Value = RecValue.HumiSV[iRecNo];
                            wshumi.Cells["G" + iHumiNo.ToString()].Value = RecValue.TempPV[iRecNo];
                            wshumi.Cells["H" + iHumiNo.ToString()].Value = RecValue.WetTempPV[iRecNo];
                            wshumi.Cells["I" + iHumiNo.ToString()].Value = RecValue.RoomTemp[iRecNo];
                            wshumi.Cells["L" + iHumiNo.ToString()].Value = RecValue.HumiPV[iRecNo];
                            iHumiNo++;
                        }
                        // 温度データ記録
                        if (iRecNo >= Ev.HumiEndRecord) 
                        {
                            wstemp.Cells["C" + iTempNo.ToString()].Value = RecValue.Date[iRecNo];
                            wstemp.Cells["D" + iTempNo.ToString()].Value = RecValue.Time[iRecNo];
                            wstemp.Cells["E" + iTempNo.ToString()].Value = RecValue.TempSV[iRecNo];
                            wstemp.Cells["F" + iTempNo.ToString()].Value = RecValue.HumiSV[iRecNo];
                            wstemp.Cells["G" + iTempNo.ToString()].Value = RecValue.TempPV[iRecNo];
                            wstemp.Cells["H" + iTempNo.ToString()].Value = RecValue.WetTempPV[iRecNo];
                            wstemp.Cells["I" + iTempNo.ToString()].Value = RecValue.RoomTemp[iRecNo];
                            iTempNo++;
                        }
                        ++iRecNo;
                    }
                    // グラフにデータ範囲設定
                    // 🔹 データを書き込んだあとに、グラフ設定
                    // 温度グラフ準備
                    int iLength = Ev.EndRecord - Ev.HumiEndRecord + 4;
                    var chart1 = wsData.Drawings["グラフ 5"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                    var chart3 = wsDataEn.Drawings["グラフ 5"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                    double vMax = ChkNowSpec.SpecMaxTemp + 20;
                    double vMin = ChkNowSpec.SpecMinTemp - 20;
                    if (ChkNowSpec.SpecMaxTemp == 150) vMax = 180;

                    chart1.YAxis.MaxValue = vMax;
                    chart1.YAxis.MinValue = vMin;
                    chart1.Axis[0].MinValue = 0; chart1.Axis[0].MaxValue = iLength;

                    chart3.YAxis.MaxValue = vMax;
                    chart3.YAxis.MinValue = vMin;
                    chart3.Axis[0].MinValue = 0; chart3.Axis[0].MaxValue = iLength;

                    if (ChkNowSpec.HumiUse)
                    {
                        // 湿度グラフ準備
                        iLength = Ev.HumiEndRecord + 4;
                        var chart2 = wsData.Drawings["グラフ 1"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                        chart2.Axis[0].MinValue = 0; chart2.Axis[0].MaxValue = iLength;
                        var chart4 = wsDataEn.Drawings["グラフ 1"] as OfficeOpenXml.Drawing.Chart.ExcelChart;
                        chart4.Axis[0].MinValue = 0; chart4.Axis[0].MaxValue = iLength;
                    }


                    cMsg = "検査データファイル書込み開始\r\n";
                    WriteLog(cMsg);
                    DataQueue.Enqueue(cMsg);
                    package.Save();
                }
                #endregion 検査結果書込み

                #region グラフファイル書込み
                // グラフファイル準備
                string GraphDistFileName = "";
                string GraphDistPDFFileName = "";
                string GraphSourceFileName = "";
                string GraphSourcePDFFileName = "";
                if (SystemSet.GraphRequired)
                {
                    // 確認グラフデータ複写
                    Bitmap bmpImage = Ev.formsPlot1.Plot.GetBitmap();
                    GraphDistPDFFileName = @DocPath + "\\" + "Graph一時ファイル.pdf";
                    if (System.IO.File.Exists(GraphDistPDFFileName)) System.IO.File.Delete(GraphDistPDFFileName);
#if False

                    GraphDistFileName = @DocPath + "\\" + "Graph一時ファイル.xlsx";

                    if (System.IO.File.Exists(GraphDistFileName)) System.IO.File.Delete(GraphDistFileName);
                    cPath = SystemSet.MasterPath + "グラフファイル.xlsx";

                    GraphSourceFileName = cPath;
                    sheet = "グラフファイル";

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
                    InputExcelFile = new FileInfo(GraphDistFileName);
                    using (ExcelPackage package = new ExcelPackage(InputExcelFile))
                    {
                        ExcelWorkbook wb = package.Workbook;
                        // 取得 (シート名指定)
                        ExcelWorksheet ws = wb.Worksheets[sheet];
                        ws.Cells["W29"].Value = ChkNowSpec.CheckSheetName;

                        // グラフ複写
                        int iErrCnt = 0;
                        while (true)
                        {
                            try
                            {
                                Clipboard.SetImage(bmpImage);
                                // セル位置（例：B2）に画像を挿入
                                var picture = ws.Drawings.AddPicture("ClipboardImage", bmpImage);
                                // セル位置を指定
                                picture.SetPosition(1, 0, 1, 0); // 行=1, 列=1 → "B2"
                                picture.SetSize(57); // サイズを変更（比率指定）67%

                                break;
                            }
                            catch (Exception ex)
                            {
                                ++iErrCnt;
                                string sMsg = "確認グラフデータ複写異常：" + ex.Message + "\r\n";
                                WriteLog(sMsg);
                                DataQueue.Enqueue(cMsg);
                                Console.WriteLine(iErrCnt.ToString() + ":" + sMsg);
                                if (iErrCnt > 3)
                                {
                                    bGrapfCopyErr = true;

                                    WriteLog(sMsg);
                                    // 異常時の背景色を赤色
                                    //txtMessage.BackColor = Color.Red;
                                    DataQueue.Enqueue(cMsg);
                                    Console.WriteLine(sMsg);
                                    break;
                                }
                            }
                        }
                        // Top表示
                        ws.View.ActiveCell = "A1";
                        ws.Select("A1");
                        cMsg = "グラフファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        DataQueue.Enqueue(cMsg);
                        package.Save();
                    }
#endif
                    // PDFファイル出力
                    // 確認グラフデータ複写
                    bmpImage = Ev.formsPlot1.Plot.GetBitmap();
                    ConvertBitmapToPdf(bmpImage, GraphDistPDFFileName);
                    // Bitmapオブジェクトを解放
                    if (bmpImage != null)
                    {
                        bmpImage.Dispose();
                    }

                    // PDFファイル出力
                    //var excelApp = new Excel.Application();
                    //var wb1 = excelApp.Workbooks.Open(GraphDistFileName);
                    //wb1.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, GraphDistPDFFileName);
                    //wb1.Close(false);
                    //excelApp.Quit();
                }
#endregion グラフファイル書込み

                #region 試験成績書処理
                string TestDistFileName = "";
                    string TestSourceFileName = "";
                if (SystemSet.TestReportRequired)
                {
                    TestDistFileName = @DocPath + "\\" + "Test一時ファイル.xlsx";

                    if (System.IO.File.Exists(TestDistFileName)) System.IO.File.Delete(TestDistFileName);
                    cPath = SystemSet.MasterPath + "試験検査報告書.xlsx";

                    TestSourceFileName = cPath;
                    sheet = "試験検査報告書";

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
                    InputExcelFile = new FileInfo(TestDistFileName);
                    using (ExcelPackage package = new ExcelPackage(InputExcelFile))
                    {
                        ExcelWorkbook wb = package.Workbook;
                        // 取得 (シート名指定)
                        ExcelWorksheet ws = wb.Worksheets[sheet];

                        // ヘッダー情報
                        ws.Cells["AA3"].Value = HeaderData.PuroductModel;
                        ws.Cells["AA4"].Value = HeaderData.CreateNo;

                        string strTime = Ev.TestingDate + " 00:00:00";

                        DateTime dTime = DateTime.Parse(strTime);
                        ws.Cells["AA5"].Value = dTime;

                        ws.Cells["AA6"].Value = Ev.RoomTempAverage;
                        ws.Cells["AA7"].Value = HeaderData.VoltSpec.ToString();
                        ws.Cells["AA8"].Value = HeaderData.FreqSpec.ToString();
                        ws.Cells["AA9"].Value = Ev.TestingPersonEn;
                        ws.Cells["AA10"].Value = Ev.TestingPerson;

                        // 湿度仕様有無
                        ws.Cells["AE19"].Value = "FALSE";
                        if (ChkNowSpec.HumiUse) ws.Cells["AE19"].Value = "TRUE";

                        // 観測窓仕様有無
                        ws.Cells["AE7"].Value = "FALSE";
                        if (ChkNowSpec.WindowUse) ws.Cells["AE7"].Value = "TRUE";

                        // 内扉仕様有無
                        ws.Cells["AG7"].Value = "FALSE";
                        if (ChkNowSpec.IndoorUse) ws.Cells["AG7"].Value = "TRUE";

                        // 側面操作孔仕様有無
                        ws.Cells["AI7"].Value = "FALSE";
                        if (ChkNowSpec.OpeWindowUse) ws.Cells["AI7"].Value = "TRUE";

                        // 冷凍機リークテスト仕様有無
                        ws.Cells["AD89"].Value = false;
                        
                        if (ChkNowSpec.RefLeakTestUse)
                        {
                            ws.Cells["AD89"].Value = true;
                        }
                        else 
                        {
                            // Row 88, Col 1 の図形を非表示にする
                            //HideShapeByCoordinate(ws, 88, 1);
                            HideShapeByCoordinateV45(ws, 88, 1);
                        }


                        // 温度範囲
                        ws.Cells["AE21"].Value = ChkNowSpec.SpecMinTemp.ToString();
                        ws.Cells["AF21"].Value = ChkNowSpec.SpecMaxTemp.ToString();
                        double dValue = 0;
                        bool bRet = double.TryParse(Ev.lblTempOpeMemL, out dValue);
                        if (bRet) ws.Cells["AG21"].Value = dValue.ToString("F1");

                        bRet = double.TryParse(Ev.lblTempOpeMemH, out dValue);
                        if (bRet) ws.Cells["AH21"].Value = dValue.ToString("F1");

                        ws.Cells["AI21"].Value = "NG";
                        if (Ev.lblTempOpeJdgeL == "GOOD" && Ev.lblTempOpeJdgeH == "GOOD") ws.Cells["AI21"].Value = "GOOD";

                        if (ChkNowSpec.HumiUse)
                        {
                            // 湿度範囲  ？？？？
                            //ws.Cells["AE22"].Value = ChkNowSpec.SpecMinHumi.ToString();
                            //ws.Cells["AF22"].Value = ChkNowSpec.SpecMaxHumi.ToString();
                            //ws.Cells["AG22"].Value = lblHumiLowValue.Text;
                            //ws.Cells["AH22"].Value = lblHumiHighValue.Text;
                            // 上限湿度：95％　下限湿度：30%　の測定値で判断か？
                            ws.Cells["AE22"].Value = "30";
                            ws.Cells["AF22"].Value = "95";

                            bRet = double.TryParse(Ev.lblHumiOpeHumi3, out dValue);
                            if (bRet) ws.Cells["AG22"].Value = dValue.ToString("F1");

                            bRet = double.TryParse(Ev.lblHumiOpeHumi4, out dValue);
                            if (bRet) ws.Cells["AH22"].Value = dValue.ToString("F1");

                            ws.Cells["AI22"].Value = "GOOD";

                            float Humi30 = float.Parse(Ev.lblHumiOpeHumi3);
                            float Humi95 = float.Parse(Ev.lblHumiOpeHumi4);
                            if (Humi30 > 33 || Humi30 < 27 || Humi95 > 98 || Humi95 < 92) ws.Cells["AI22"].Value = "NG";
                        }
                        // 温度変化上昇
                        //if (ChkNowSpec.ChgUpUse)
                        //{
                        ws.Cells["AE29"].Value = Ev.lblChgUpStart;
                        ws.Cells["AF29"].Value = Ev.lblChgUpEnd;
                        ws.Cells["AG29"].Value = Ev.lblChgUpSpec;
                        bRet = double.TryParse(Ev.lblChgUpValue, out dValue);
                        if (bRet) ws.Cells["AH29"].Value = dValue.ToString("F1");
                        ws.Cells["AI29"].Value = Ev.lblChgUpStr;
                        //}
                        //else ws.Cells["AJ29"].Value = "FALSE";

                        // 温度到達上昇
                        //if (ChkNowSpec.PoleUpUse)
                        //{
                        ws.Cells["AE30"].Value = Ev.lblPoleUpStart;
                        ws.Cells["AF30"].Value = Ev.lblPoleUpEnd;
                        ws.Cells["AG30"].Value = Ev.lblPoleUpSpec;
                        bRet = double.TryParse(Ev.lblPoleUpValue, out dValue);
                        if (bRet) ws.Cells["AH30"].Value = dValue.ToString("F0");
                        ws.Cells["AI30"].Value = Ev.lblPoleUpStr;
                        //}
                        //else ws.Cells["AJ30"].Value = "FALSE";

                        // 温度変化降下
                        //if (ChkNowSpec.ChgDownUse)
                        //{
                        ws.Cells["AE31"].Value = Ev.lblChgDownStart;
                        ws.Cells["AF31"].Value = Ev.lblChgDownEnd;
                        ws.Cells["AG31"].Value = Ev.lblChgDownSpec;
                        bRet = double.TryParse(Ev.lblChgDownValue, out dValue);
                        if (bRet) ws.Cells["AH31"].Value = dValue.ToString("F1");
                        ws.Cells["AI31"].Value = Ev.lblChgDownStr;
                        //}
                        //else ws.Cells["AJ31"].Value = "FALSE";

                        // 温度到達降下
                        //if (ChkNowSpec.PoleDownUse)
                        //{
                        ws.Cells["AE32"].Value = Ev.lblPoleDownStart;
                        ws.Cells["AF32"].Value = Ev.lblPoleDownEnd;
                        ws.Cells["AG32"].Value = Ev.lblPoleDownSpec;
                        bRet = double.TryParse(Ev.lblPoleDownValue, out dValue);
                        if (bRet) ws.Cells["AH32"].Value = dValue.ToString("F0");
                        ws.Cells["AI32"].Value = Ev.lblPoleDownStr;
                        //}
                        //else ws.Cells["AJ32"].Value = "FALSE";

                        // 温度変動１
                        ws.Cells["AE46"].Value = Ev.lblVarTemp1Start;
                        ws.Cells["AF46"].Value = Ev.lblVarTemp1End;

                        bRet = double.TryParse(Ev.lblVarTemp1Spec, out dValue);
                        if (bRet) ws.Cells["AG46"].Value = dRoundValue(dValue,1);

                        bRet = double.TryParse(Ev.lblTempFluctMemM, out dValue);
                        if (bRet) bRet = double.TryParse(Ev.lblTempFluctMemL, out dValue);
                        if (bRet)
                        {
                            float MTemp = fRoundValue(float.Parse(Ev.lblTempFluctMemM),1);
                            float LTemp = fRoundValue(float.Parse(Ev.lblTempFluctMemL),1);

                            ws.Cells["AH46"].Value = fRoundValue(LTemp,1);
                            if (MTemp > LTemp) ws.Cells["AH46"].Value = fRoundValue(MTemp,1);
                            ws.Cells["AI46"].Value = "NG";
                            if (Ev.lblTempFluctJdgeL == "GOOD" && Ev.lblTempFluctJdgeM == "GOOD") ws.Cells["AI46"].Value = "GOOD";
                        }
                        else ws.Cells["AI46"].Value = "NG";

                        // 温度変動２
                        ws.Cells["AJ47"].Value = "FALSE";
                        bRet = double.TryParse(Ev.lblVarTemp2Start, out dValue);
                        if (bRet) ws.Cells["AE47"].Value = dRoundValue(dValue,1);

                        bRet = double.TryParse(Ev.lblVarTemp2End, out dValue);
                        if (bRet) ws.Cells["AF47"].Value = dRoundValue(dValue,0);

                        bRet = double.TryParse(Ev.lblVarTemp2Spec, out dValue);
                        if (bRet) ws.Cells["AG47"].Value = dRoundValue(dValue,1);

                        bRet = double.TryParse(Ev.lblTempFluctMemH, out dValue);
                        if (bRet) ws.Cells["AH47"].Value = dRoundValue(dValue,1);

                        ws.Cells["AI47"].Value = Ev.lblTempFluctJdgeH;
                        if (HeaderData.TempSpec > 100) ws.Cells["AJ47"].Value = "TRUE";

                        if (ChkNowSpec.HumiUse)
                        {
                            // 湿度変動
                            float fValue1 = float.Parse(Ev.lblHumiFluctStd1);
                            float fValue2 = float.Parse(Ev.lblHumiFluctStd2);
                            float fValue3 = float.Parse(Ev.lblHumiFluctStd3);
                            float fValue4 = float.Parse(Ev.lblHumiFluctStd4);
                            float max_value = Math.Max(Math.Max(fValue1, fValue2), Math.Max(fValue3, fValue4));

                            bRet = double.TryParse(Ev.lblVarHumiSpec, out dValue);
                            if (bRet) ws.Cells["AG48"].Value = dRoundValue(dValue,1);
                            ws.Cells["AH48"].Value = max_value.ToString("F1");

                            ws.Cells["AI48"].Value = "NG";
                            if (Ev.lblHumiFluctJdge1 == "GOOD" && Ev.lblHumiFluctJdge2 == "GOOD" &&
                                Ev.lblHumiFluctJdge3 == "GOOD" && Ev.lblHumiFluctJdge4 == "GOOD") ws.Cells["AI48"].Value = "GOOD";
                        }
                        // 電流判定情報の登録
                        int iLineNo = 70;int iCurName = 13;
                        foreach(CURRENT cur in CurrentInfo) 
                        {
                            // 名称が空白か”額縁”が含まれていたら、未使用とする
                            if (cur.Name == "" || cur.Name.Contains("額縁")) ws.Cells["AE" + iLineNo.ToString()].Value = "FALSE";
                            else 
                            {
                                ws.Cells["AD" + iLineNo.ToString()].Value = cur.Name;
                                ws.Cells["Z" + iCurName.ToString()].Value = cur.Name;
                                ws.Cells["AE" + iLineNo.ToString()].Value = "TRUE"; 
                            }
                            ws.Cells["AF" + iLineNo.ToString()].Value = fRoundValue(cur.Power,1);
                            ws.Cells["AG" + iLineNo.ToString()].Value = fRoundValue(cur.LowLimit,1);
                            ws.Cells["AH" + iLineNo.ToString()].Value = fRoundValue(cur.HighLimit, 1);
                            ++iLineNo;++iCurName;
                            if (iLineNo == 73) iLineNo = 79;
                        }
#if Flase

                    // 温度設定差
                    ws.Cells["AI52"].Value = "GOOD";
                    ws.Cells["AH52"].Value = TempSettingDiffValue.ToString("F1");
                    if (TempSettingDiffValue > ChkNowSpec.AccuracyTemp1Spec)
                    {
                        ws.Cells["AI52"].Value = "NG";
                    }
                    ws.Cells["AG52"].Value = "1.3";
                    if (!ChkNowSpec.HumiUse)
                    {
                        ws.Cells["AE52"].Value = ChkNowSpec.SpecMaxTemp.ToString("F0");
                        ws.Cells["AE53"].Value = "---";
                        ws.Cells["AG53"].Value = "---";
                        ws.Cells["AH53"].Value = "---";
                        ws.Cells["AI53"].Value = "---";
                    }
                    else
                    {
                        ws.Cells["AE52"].Value = lblSettingTempSetValue.Text;
                        ws.Cells["AE53"].Value = lblSettingHumiSetValue.Text;
                        ws.Cells["AG53"].Value = "3.0";
                        ws.Cells["AI53"].Value = "GOOD";
                        ws.Cells["AH53"].Value = HumiSettingDiffValue.ToString("F1");
                        if (HumiSettingDiffValue > ChkNowSpec.AccuracyHumiSpec)
                        {
                            ws.Cells["AI53"].Value = "NG";
                        }

                    }
                    if (SystemSet.CurrentRequired)
                    {
                        // 積算電力量表示
                        ws.Cells["AD89"].Value = "---";
                        if (PowerInfo.EndExist)
                        {
                            ws.Cells["AD89"].Value = PowerInfo.HourPower.ToString("F1");
                        }
                        // 電流判定結果
                        int j = 0;
                        for (int i = 0; i < 6; ++i)
                        {
                            if (i > 2) j = 6;
                            if (CurInfos[i].Use)
                            {
                                ws.Cells["AD" + (70 + i + j).ToString()].Value = CurInfos[i].name;
                                ws.Cells["AE" + (70 + i + j).ToString()].Value = "TRUE";
                                ws.Cells["AF" + (70 + i + j).ToString()].Value = CurInfos[i].Power.ToString("F1");
                                ws.Cells["AG" + (70 + i + j).ToString()].Value = CurInfos[i].Low.ToString("F1");
                                ws.Cells["AH" + (70 + i + j).ToString()].Value = CurInfos[i].High.ToString("F1");
                                ws.Cells["AI" + (70 + i + j).ToString()].Value = CurInfos[i].Value.ToString("F1");
                                ws.Cells["AJ" + (70 + i + j).ToString()].Value = "NG";
                                if (CurInfos[i].Judge) ws.Cells["AJ" + (70 + i + j).ToString()].Value = "GOOD";
                            }
                            else ws.Cells["AE" + (70 + i + j).ToString()].Value = "FALSE";
                        }
                    }
#endif
                        // シート保護
                        // 2. シートの保護を有効にする
                        ws.Protection.AllowSelectLockedCells = false;
                        ws.Protection.IsProtected = true;
                        // 3. パスワードを設定する（省略可能）
                        ws.Protection.SetPassword("password");
                        // Top表示
                        ws.View.ActiveCell = "AA12";
                        ws.Select("AA12");

                        cMsg = "試験検査報告書ファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        DataQueue.Enqueue(cMsg);

                        //警告メッセージを非表示にする
                        package.Save();
                    } 
                }
                #endregion 試験成績書処理

                #region アース導通保存
                string EarthDistFileName = "";
                    string EarthSourceFileName = "";
                    sheet = "";
                if (SystemSet.TestReportRequired)
                {
                    EarthDistFileName = @DocPath + "\\" + "Earth一時ファイル.xlsx";

                    if (System.IO.File.Exists(EarthDistFileName)) System.IO.File.Delete(EarthDistFileName);
                    cPath = SystemSet.MasterPath + "アース導通試験報告書.xlsx";

                    EarthSourceFileName = cPath;
                    sheet = "アース導通";

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
                    InputExcelFile = new FileInfo(EarthDistFileName);
                    using (ExcelPackage package = new ExcelPackage(InputExcelFile))
                    {
                        ExcelWorkbook wb = package.Workbook;
                        // 取得 (シート名指定)
                        ExcelWorksheet ws = wb.Worksheets[sheet];
                        // ヘッダー情報
                        ws.Cells["Z5"].Value = HeaderData.PuroductModel;
                        ws.Cells["Z6"].Value = HeaderData.CreateNo;

                        string strTime = Ev.TestingDate;

                        DateTime dTime = DateTime.Parse(strTime);
                        ws.Cells["Z7"].Value = dTime;
                        ws.Cells["Z8"].Value = Ev.TestingPersonEn;
                        ws.Cells["Z9"].Value = Ev.TestingPerson;

                        // 不要な行の削除
                        // 型式検索
                        int iRow = 29;int iCnt = 14;
                        while (true) 
                        {
                            string sValue= ws.Cells[4, iRow].Value.ToString();
                            if (sValue == "") break;
                            if (Ev.ModelName == sValue) { iCnt = int.Parse(ws.Cells[3,iRow].Text); break; }
                            ++iRow;
                        }
                        if(iCnt != 18) ws.DeleteRow(24 + iCnt, 18 - iCnt);
                        var drawing = ws.Drawings["図 6"];
                        drawing.SetPosition(24 + iCnt, 0, 14, 0);

                        // シート保護
                        // 2. シートの保護を有効にする
                        ws.Protection.AllowSelectLockedCells = false;
                        ws.Protection.IsProtected = true;
                        // 3. パスワードを設定する（省略可能）
                        ws.Protection.SetPassword("password");
                        // Top表示
                        ws.View.ActiveCell = "A1";
                        ws.Select("A1");

                        // マクロを実行できないのでコードで処理する
                        //xlApp.Run("HideBlankRows");

                        cMsg = "アース導通試験報告書ファイル書込み開始\r\n";
                        WriteLog(cMsg);
                        DataQueue.Enqueue(cMsg);
                        //警告メッセージを非表示にする
                        package.Save();
                    } 
                }
                #endregion アース導通保存


                // Excel を終了する
                cMsg = "エクセル処理を終了\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);

                #region 自動判定ログ保存

                string path = System.Environment.GetFolderPath(Environment.SpecialFolder.Personal) + "\\";
                path += "CSV";
                String CheckDataSourceFileName = @path + "\\" + NowExcelFileName + "_自動判定ログ.CSV";

                // 検査データ保存フォルダ作成
                DateTime NowDt = DateTime.Now;
                string m = NowDt.ToString("MM");
                int Year =  int.Parse(NowDt.ToString("yyyy"));
                int Month = int.Parse(m);
                if (Month < 4) --Year;

                string UserPath = Ev.SerialNo + " " + Ev.ModelName + " " + Ev.UserName;

                string DistPath = @SystemSet.ResultPath + Year.ToString() + "年度" + "\\" + UserPath;

                if (!Directory.Exists(DistPath))
                {
                    // ディレクトリ作成
                    Directory.CreateDirectory(DistPath);
                }

                // 検査データ移動

                TestSourceFileName = TestDistFileName;

                string sFileName = NowExcelFileName.Replace("検査データ_", "");
                String CheckDataDistFileName = @DistPath + "\\自動判定ログ_" + sFileName + ".csv";

                TestDistFileName = @DistPath + "\\" + "試験検査報告書_" + sFileName + ".xlsx";
                EarthSourceFileName = EarthDistFileName;
                EarthDistFileName = @DistPath + "\\" + "アース試験報告書_" + sFileName + ".xlsx";

                GraphSourceFileName = GraphDistFileName;
                GraphDistFileName = @DistPath + "\\グラフ_" + sFileName + ".xlsx";

                GraphSourcePDFFileName = GraphDistPDFFileName;
                GraphDistPDFFileName = @DistPath + "\\グラフ_" + sFileName + ".pdf";

                SourceFileName = DistFileName;
                DistFileName = @DistPath + "\\" + NowExcelFileName + ".xlsx";

                result = DialogResult.Yes;
                if (System.IO.File.Exists(DistFileName))
                {
                    // 保存先にファイルが存在していた場合
                    cMsg = NowExcelFileName + ".xlsx" + "ファイルが存在します。上書きしますか？";
                    result = MessageBox.Show(cMsg, "上書き警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                }
                cMsg = "検査データファイル書込み中\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);
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
                        //System.IO.File.Copy(GraphSourceFileName, GraphDistFileName, true);
                        System.IO.File.Copy(GraphSourcePDFFileName, GraphDistPDFFileName, true);
                    }
                    else
                    {
                        //if (System.IO.File.Exists(GraphDistFileName)) System.IO.File.Delete(GraphSourceFileName);
                        if (System.IO.File.Exists(GraphDistPDFFileName)) System.IO.File.Delete(GraphSourcePDFFileName);
                    }

                    System.IO.File.Copy(CheckDataSourceFileName, CheckDataDistFileName, true);

                    // CSV ファイル複写
                    string cFileName = DistPath + "\\" + Ev.lblCsvFileName + ".CSV";
                    System.IO.File.Copy(NowCSVFilePath, @cFileName, true);
                }
                #endregion 自動判定ログ保存

                // 一時ファイル削除
                System.IO.File.Delete(SourceFileName);
                //if (SystemSet.TestReportRequired) File.Delete(TestSourceFileName);
                if (SystemSet.GraphRequired)
                {
                    //System.IO.File.Delete(GraphSourceFileName);
                    System.IO.File.Delete(GraphSourcePDFFileName);
                }
                System.IO.File.Delete(CheckDataSourceFileName);
                cMsg = "検査データファイル書込み完了\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);

                #region 検査データ集計_小型環境保存
                cMsg = "検査データ集計_小型環境ファイル書込み開始\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);
                System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("ja-JP");

                try
                {
                    string Collection_Source_FileName = "";
                    sheet = "";
                    if (SystemSet.Test_data_collection_Required)
                    {
                        Collection_Source_FileName = @SystemSet.Test_data_collection_Path + SystemSet.Test_data_collection_FileName;

                        sheet = "一覧";

                        // xlApplication から WorkBooks を取得する
                        // 既存の Excel ブックを開く
                        InputExcelFile = new FileInfo(Collection_Source_FileName);
                        using (ExcelPackage package = new ExcelPackage(InputExcelFile))
                        {
                            ExcelWorkbook wb = package.Workbook;
                            // 取得 (シート名指定)
                            ExcelWorksheet ws = wb.Worksheets[sheet];

                            ws.Cells["A7:I" + ws.Dimension.End.Row.ToString()].Style.Numberformat.Format = "General";
                            ws.Cells["M7:Y" + ws.Dimension.End.Row.ToString()].Style.Numberformat.Format = "General";
                            ws.Cells["N4:N4"].Style.Numberformat.Format = "General";

                            // 製造番号検索
                            // 指定された列範囲のセルに対してLINQクエリを実行
                            // 列範囲は (開始行, 開始列, 終了行, 終了列) で指定
                            int column = 1; // A列
                            int row = 7; // 7行目から
                            string searchTerm = Ev.SerialNo;
                            var cellsInColumn = ws.Cells[row, column, ws.Dimension.End.Row, column];

                            var foundCell = (from cell in ws.Cells[ws.Dimension.Start.Row, column, ws.Dimension.End.Row, column]
                                             where cell.Text.Equals(searchTerm)
                                             select cell)
                                            .FirstOrDefault(); // 最初の一致セルを取得

                            // 結果の処理
                            if (foundCell != null)
                            {
                                int findrow = foundCell.Start.Row;
                                int findcolumn = foundCell.Start.Column;

                                // 製造番号が見つかった
                                Console.WriteLine($"文字列 \"{searchTerm}\" が見つかりました:");
                                string model = "";
                                if (ws.Cells[findrow, 2].Value != null) 
                                {
                                    model = ws.Cells[findrow, 2].Value.ToString();
                                }
                                // 上書き確認
                                result = DialogResult.Yes;
                                if (model != "")
                                {
                                    // 形式が存在していた場合
                                    cMsg = "検査データ集計_小型環境ファイルに既に型式が登録されております。上書きしますか？";
                                    result = MessageBox.Show(cMsg, "上書き警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                }
                                if (result == DialogResult.Yes)
                                {
                                    ws.Cells[findrow, 2].Value = Ev.ModelName;
                                    ws.Cells[findrow, 3].Value = int.Parse(Ev.TestingVolt);
                                    ws.Cells[findrow, 5].Value = int.Parse(Ev.TestingFreq);

                                    ws.Cells[findrow, 6].Value = int.Parse(Ev.lblPoleUpValue);
                                    ws.Cells[findrow, 7].Value = int.Parse(Ev.lblPoleDownValue);

                                    ws.Cells[findrow, 8].Value = int.Parse(Ev.lblChgUpTime.Replace("分",""));
                                    ws.Cells[findrow, 9].Value = int.Parse(Ev.lblChgDownTime.Replace("分", ""));

                                    ws.Cells[findrow, 10].Value = Ev.PoleDownRoomTempAverage;

                                    ws.Cells[findrow, 11].Value = float.Parse(Ev.lblChgUpValue);
                                    ws.Cells[findrow, 12].Value = float.Parse(Ev.lblChgDownValue);

                                    // ケーブル孔
                                    for (int i = 0; i < 4; ++i) 
                                    {
                                        ws.Cells[findrow, 14 + i].Value = "";
                                        if (Ev.CableHole[i] != 0) ws.Cells[findrow, 14 + i].Value = Ev.CableHole[i].ToString() + "個";
                                    }

                                    // 操作孔　
                                    ws.Cells[findrow, 18].Value = "無";
                                    if (Ev.OperationHole) ws.Cells[findrow, 18].Value = "有";

                                    // 内扉　
                                    ws.Cells[findrow, 19].Value = "無";
                                    if (Ev.InnerDoor) ws.Cells[findrow, 19].Value = "有";

                                    // 観測窓　
                                    ws.Cells[findrow, 20].Value = "無";
                                    if (Ev.CheckWindow) ws.Cells[findrow, 20].Value = "有";

                                    // 上限温度　
                                    ws.Cells[findrow, 21].Value = ChkNowSpec.SpecMaxTemp;

                                    // 天井観測窓　
                                    ws.Cells[findrow, 22].Value = "無";

                                    // 検査日　
                                    ws.Cells[findrow, 24].Value = DateTime.Parse(Ev.TestingDate).ToString("yy.M.d");

                                    // その他オプション
                                    ws.Cells[findrow, 25].Value = Ev.AdditionalSpecifications;

                                    // 書き込んだセルTop表示　未対応　EPPlus4.5では対応方法が無い
                                    // ws.Select("A" + findrow.ToString());

                                    package.Save();

                                    cMsg = "検査データ集計_小型環境ファイル書込み完了\r\n";
                                    WriteLog(cMsg);
                                    DataQueue.Enqueue(cMsg);
                                }
                                cMsg = "検査データ集計_小型環境ファイル処理終了\r\n";
                                WriteLog(cMsg);
                                DataQueue.Enqueue(cMsg);
                            }
                            else
                            {
                                // 製造番号が見つからなかった
                                result = DialogResult.Yes;
                                cMsg = $"検査データ集計_小型環境ファイルに製造番号:\"{Ev.SerialNo}\"が見つかりませんでした。\r\n" +
                                        "検査データ集計_小型環境ファイルの内容を確認して、再度、判定を実施してください。";
                                result = MessageBox.Show(cMsg, "製造番号無し警告", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                                Console.WriteLine($"文字列 \"{searchTerm}\" は見つかりませんでした。");
                                throw new Exception(cMsg);
                            }
                        }
                    }
                }
                catch (Exception ex) 
                {
                    // 形式が存在していた場合
                    cMsg = "検査データ集計_小型環境ファイルが保存出来ませんでした!!\r\n" + 
                           "ファイルを閉じて、再度判定から実施してください。";
                    MessageBox.Show(cMsg, "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                    throw new Exception(cMsg);
                }
                #endregion 検査データ集計_小型環境保存


                // 検査保存フォルダをエクスプローラで開く
                cMsg = "検査データ保存フォルダをエクスプローラで開きます。\r\n";
                WriteLog(cMsg);
                DataQueue.Enqueue(cMsg);
                System.Diagnostics.Process.Start(@DistPath);
                bExcelWriteFail = false;
                bResult = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex);
                // データ読込エラー
                cMsg = "エクセルファイル書出しに失敗しました。\r\n";
                cMsg += "設定されている内容を再度確認して下さい。\r\n";
                DataQueue.Enqueue(cMsg);
                MessageBox.Show(cMsg, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, MessageBoxOptions.DefaultDesktopOnly);
                bExcelWriteFail = true;
            }
            bExcelWriteEnd = true;

            return bResult;
        }

        private static string sValue(float fValue,int dot) 
        {
            string sRet = fValue.ToString("");
            if(dot!=0) sRet = fValue.ToString("F" + dot.ToString());
            if (fValue > 0) sRet = "+" + sRet;
            return sRet;
        }
        private static float fRoundValue(float fValue, int dot)
        {
            float sRet = (float)Math.Round((double)fValue, dot+1);
            return sRet;
        }
        private static double dRoundValue(double fValue, int dot)
        {
            double sRet = (double)Math.Round((double)fValue, dot + 1);
            return sRet;
        }

        private static void ConvertBitmapToPdf(Bitmap bitmap, string pdfPath)
        {
            // 新しいPDFドキュメントを作成
            PdfDocument document = new PdfDocument();
            // ページを追加
            PdfPage page = document.AddPage();
            // サイズをA4に設定
            page.Size = PdfSharp.PageSize.A4;
            // 向きを横（Landscape）に設定
            // これにより、WidthとHeightが自動的に入れ替わります
            page.Orientation = PdfSharp.PageOrientation.Landscape;
            // GDI+ Image (Bitmap) から XImage を作成
            XImage image = XImage.FromGdiPlusImage(bitmap); // ここでエラーが解消するはずです
            // 描画オブジェクトを取得
            XGraphics gfx = XGraphics.FromPdfPage(page);
            // 画像を描画
            // (0, 0)から始まり、画像の元の幅と高さで描画
            gfx.DrawImage(image, 20, 100, image.PixelWidth * 0.425, image.PixelHeight * 0.425);
            // ドキュメントを保存
            document.Save(pdfPath);
        }

        public static void SetShapeVisibilityReflection(ExcelDrawing shape, bool isVisible)
        {
            if (shape == null) return;

            XmlDocument xmlDoc = null;

            // 型の中にある「XmlDocument」型のフィールドをすべて取得して、最初に見つかったものを使う
            // これにより、内部変数が "_drawingXml" でも "xml" でも自動対応できます
            var fields = shape.GetType().GetFields(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            foreach (var field in fields)
            {
                if (field.FieldType == typeof(XmlDocument))
                {
                    xmlDoc = (XmlDocument)field.GetValue(shape);
                    if (xmlDoc != null) break;
                }
            }

            // プロパティ経由でも探す
            if (xmlDoc == null)
            {
                var props = shape.GetType().GetProperties(BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
                foreach (var prop in props)
                {
                    if (prop.PropertyType == typeof(XmlDocument))
                    {
                        xmlDoc = (XmlDocument)prop.GetValue(shape);
                        if (xmlDoc != null) break;
                    }
                }
            }

            if (xmlDoc == null)
            {
                // デバッグ用：どうしても見つからない場合
                Console.WriteLine("XmlDocumentが見つかりませんでした。");
                return;
            }

            // 名前空間の設定
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

            // cNvPr ノードを特定
            XmlNode nvPrNode = xmlDoc.SelectSingleNode("//xdr:cNvPr", nsmgr);

            if (nvPrNode != null)
            {
                XmlElement element = (XmlElement)nvPrNode;
                if (!isVisible)
                {
                    element.SetAttribute("hidden", "1");
                }
                else
                {
                    element.RemoveAttribute("hidden");
                }
            }
        }

        public static void HideShapeByCoordinate(ExcelWorksheet sheet, int row, int col)
        {
            // 1. シートの全図形を管理している Drawings コレクションから、リフレクションで XML を取得
            var drawingsField = typeof(ExcelWorksheet).GetField("_drawings", BindingFlags.Instance | BindingFlags.NonPublic);
            if (drawingsField == null) return;

            var drawingsManager = drawingsField.GetValue(sheet);
            if (drawingsManager == null) return;

            // ExcelDrawings クラス内の DrawingXml プロパティを取得
            var xmlProp = drawingsManager.GetType().GetProperty("DrawingXml", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);
            if (xmlProp == null) return;

            XmlDocument xmlDoc = (XmlDocument)xmlProp.GetValue(drawingsManager);
            if (xmlDoc == null) return;

            // 2. 名前空間の設定
            XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
            nsmgr.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
            nsmgr.AddNamespace("a", "http://schemas.openxmlformats.org/drawingml/2006/main");

            // 3. XML内から、指定した座標 (from row/col) を持つ図形を探す
            // Excelは0始まりなので注意
            string xpath = string.Format("//xdr:from[xdr:row={0} and xdr:col={1}]/..", row - 1, col - 1);
            XmlNode shapeNode = xmlDoc.SelectSingleNode(xpath, nsmgr);

            if (shapeNode != null)
            {
                // 4. その図形の非表示属性ノード (cNvPr) を探す
                XmlNode nvPrNode = shapeNode.SelectSingleNode(".//xdr:cNvPr", nsmgr);
                if (nvPrNode != null)
                {
                    XmlElement element = (XmlElement)nvPrNode;
                    element.SetAttribute("hidden", "1"); // 非表示設定
                    Console.WriteLine("座標 {0},{1} の図形を非表示に設定しました。", row, col);
                }
            }
            else
            {
                Console.WriteLine("指定された座標の図形がXML内に見つかりませんでした。");
            }
        }

        public static void HideShapeByCoordinateV45(ExcelWorksheet sheet, int row, int col)
        {
            try
            {
                // 1. sheet.Drawings は公開されているので直接取得
                var drawings = sheet.Drawings;
                if (drawings == null)
                {
                    Console.WriteLine("Drawingsコレクションが取得できません。");
                    return;
                }

                // 2. ExcelDrawingsクラスから内部XML (DrawingXml) をリフレクションで取得
                // 4.5.2.1ではプロパティとして存在することが多いですが、フィールドも念のためチェック
                XmlDocument xmlDoc = null;
                var prop = drawings.GetType().GetProperty("DrawingXml", BindingFlags.Instance | BindingFlags.NonPublic | BindingFlags.Public);

                if (prop != null)
                {
                    xmlDoc = (XmlDocument)prop.GetValue(drawings);
                }
                else
                {
                    var field = drawings.GetType().GetField("_drawingXml", BindingFlags.Instance | BindingFlags.NonPublic);
                    if (field != null) xmlDoc = (XmlDocument)field.GetValue(drawings);
                }

                if (xmlDoc == null)
                {
                    Console.WriteLine("Drawings内のXmlDocumentが見つかりませんでした。");
                    return;
                }

                // 3. 名前空間の設定
                XmlNamespaceManager nsmgr = new XmlNamespaceManager(xmlDoc.NameTable);
                nsmgr.AddNamespace("xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");

                // 4. 指定座標 (Row 88, Col 1 -> XML上は 87, 0) を持つノードを特定
                // 座標指定が「From」の情報を元にしていることを想定
                string xpath = string.Format("//xdr:from[xdr:row={0} and xdr:col={1}]/..", row - 1, col - 1);
                XmlNode shapeNode = xmlDoc.SelectSingleNode(xpath, nsmgr);

                if (shapeNode != null)
                {
                    // 5. 非表示属性を付与
                    XmlNode nvPrNode = shapeNode.SelectSingleNode(".//xdr:cNvPr", nsmgr);
                    if (nvPrNode != null)
                    {
                        XmlElement element = (XmlElement)nvPrNode;
                        element.SetAttribute("hidden", "1");
                        Console.WriteLine("座標 {0},{1} の図形を非表示にしました。", row, col);
                    }
                }
                else
                {
                    Console.WriteLine("指定座標に該当するXMLノードが見つかりませんでした。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("エラー発生: " + ex.Message);
            }
        }

    }
}

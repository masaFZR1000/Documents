using ScottPlot;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EMS_Small_Chb
{
    public partial class SetCheck : Form
    {
        public SetCheck()
        {
            InitializeComponent();
            this.Text = "検査プログラム名：" + Form1.ChkNowSpec.InspectPrgName;
            lblTestCheckMsg1.Text = "検査プログラム\r\n\r\n" + Form1.sTestCheckMsg1;
            lblTestCheckMsg2.Text = "温度特性判定基準　＋　温湿度変動判定基準\r\n\r\n" + Form1.sTestCheckMsg2;
            ViewGrapf();
        }

        private void ViewGrapf() 
        {
            double AllTime = 0;

            formsPlot1.Plot.Clear();

            formsPlot1.Plot.Title("小型環境試験器　検査プログラム");
            // X軸, Y軸に名前を付けます
            formsPlot1.Plot.XLabel("経過時間(Hr)");
            formsPlot1.Plot.YLabel("温度(℃)");

            // Ｙ軸の軸範囲設定
            formsPlot1.Plot.YAxis.Dims.SetAxis(-100, 200);

            // 温度パターン制御時間
            for (int i = 0; i < Form1.TEMPSETSTEPS; ++i) 
            {
                AllTime += Form1.Tempset[i].SetTime;
            }
            if (Form1.ChkNowSpec.HumiUse)
            {
                formsPlot1.Plot.YAxis2.Label("湿度");
                formsPlot1.Plot.YAxis2.Dims.SetAxis(0, 100);
                // 目盛り線の表示を設定します(trueにしないと表示されません)
                formsPlot1.Plot.YAxis2.Ticks(true);
                // 湿度パターン制御時間
                for (int i = 0; i < Form1.HUMISETSTEPS; ++i)
                {
                    AllTime += Form1.Humiset[i].SetTime;
                }
            }
            int HumiStep = 0;
            int TempStep = 13;
            if (Form1.ChkNowSpec.HumiUse)
            {
                HumiStep = 9;
                TempStep = 12;
            }

            // 温度パターン
            double[] TempTime = new double[HumiStep+ TempStep]; double[] TempSV = new double[HumiStep+ TempStep];
            // 湿度パターン
            double[] HumiTime = new double[HumiStep]; double[] HumiSV = new double[HumiStep];int iHumiCnt = 0;
            if (Form1.ChkNowSpec.HumiUse)
            {
                TempTime[0] = 0; TempSV[0] = 0; HumiTime[0] = 0; HumiSV[0] = 0;

                for (int i = 0; i < Form1.HUMISETSTEPS; ++i)
                {
                    TempTime[2*i+1] = TempTime[2*i]; TempSV[2*i+1] = Form1.Humiset[i].SetTemp;
                    HumiTime[2*i+1] = HumiTime[2*i]; HumiSV[2*i+1] = Form1.Humiset[i].SetHumi;

                    TempTime[2*i+2] = TempTime[2*i+1] + Form1.Humiset[i].SetTime; TempSV[2*i+2] = Form1.Humiset[i].SetTemp;
                    HumiTime[2*i+2] = HumiTime[2*i+1] + Form1.Humiset[i].SetTime; HumiSV[2*i+2] = Form1.Humiset[i].SetHumi;
                }
                iHumiCnt = 4;
            }
            // 温度パターン
            for (int i = 0; i < Form1.TEMPSETSTEPS; ++i)
            {
                int j = i + iHumiCnt;
                TempTime[2 * j + 1] = TempTime[2 * j]; TempSV[2 * j + 1] = Form1.Tempset[i].SetTemp;

                TempTime[2 * j + 2] = TempTime[2 * j + 1] + Form1.Tempset[i].SetTime; TempSV[2 * j + 2] = Form1.Tempset[i].SetTemp;
            }


            // 分布図プロット
            if (Form1.ChkNowSpec.HumiUse) 
            { 
                var scat1 = formsPlot1.Plot.AddScatterLines(HumiTime, HumiSV,Color.Blue, 5, label: "湿度設定");
                scat1.OnNaN = ScottPlot.Plottable.ScatterPlot.NanBehavior.Gap;
                scat1.YAxisIndex = 1;
            }
            var scat2 = formsPlot1.Plot.AddScatterLines(TempTime, TempSV, Color.Red, 5, label: "温度設定");
            scat2.OnNaN = ScottPlot.Plottable.ScatterPlot.NanBehavior.Gap;

            // Ｘ軸の軸範囲設定
            formsPlot1.Plot.XAxis.Dims.SetAxis(0, AllTime);

            // Y軸０にLineを引く
            formsPlot1.Plot.AddHorizontalLine(0);
            // Legend(凡例)の表示を指定します(指定しないと表示されません)
            formsPlot1.Plot.Legend();

            // ScottPlotのコントロールに描画(表示)します。
            formsPlot1.Render();
            formsPlot1.Visible = true;

        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}

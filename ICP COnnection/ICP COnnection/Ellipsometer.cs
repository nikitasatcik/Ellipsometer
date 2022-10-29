using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Collections;
using System.Threading;
using System.IO.Ports;
using ZedGraph;
using Telerik.WinControls.UI;

namespace ICP_COnnection
{
    public partial class Ellipsometer : Telerik.WinControls.UI.RadForm
    {
        public Ellipsometer()
        {
            InitializeComponent();
            Telerik.WinControls.ThemeResolutionService.ApplicationThemeName = "Windows8";
        }

        private static double phi = 1.27749;
        private string temp;
        private IntPtr hPort;
        private int iSlot = PACNET.PAC_IO.PAC_REMOTE_IO(1);
        private int iAI_TotalCh = 8;
        private float fValue = 0;
        private volatile bool isWorking;
        private double n;
        private double k;
        private float I0;
        private float I45;
        private float I90;
        private float I;
        private int currentWave;
        private int voltage;
        private const int delay_ms = 3000;
        private ArrayList arrayLamda;
        private ArrayList arrayN;
        private ArrayList arrayK;
        private ArrayList Intensity;

        private ArrayList arrayI0;
        private ArrayList arrayI45;
        private ArrayList arrayI90;

        float readADC(int iChannel)
        {
            bool iRet = PACNET.PAC_IO.ReadAI(hPort, iSlot, iChannel, iAI_TotalCh, ref fValue);
            return Math.Abs(fValue);
        }

        float readSequenceADC(int iChannel, int iterations)
        {
            float result = 0;
            for (int i = 0; i < iterations; i++)
            {
                result += readADC(iChannel);
                Thread.Sleep(150);
            }
            result /= iterations;
            return result;
        }

        /************************************ START ************************************/
		
		// Start program in  ellipsometric mode
        private void radButton2_Click(object sender, EventArgs e)
        {
            arrayLamda = new ArrayList();
            arrayN = new ArrayList();
            arrayK = new ArrayList();

            arrayI0 = new ArrayList();
            arrayI45 = new ArrayList();
            arrayI90 = new ArrayList();

            progressBar1.Value = 0;

            this.MeasurementThread = new System.Threading.Thread(Worker);
            isWorking = true;
            MeasurementThread.Start();
        }

        // Start program in  Spectrophotometric mode
        private void radButton3_Click(object sender, EventArgs e)
        {
            arrayLamda = new ArrayList();
            Intensity = new ArrayList();
            progressBar1.Value = 0;
            this.MeasurementThread = new System.Threading.Thread(Worker2);
            isWorking = true;
            MeasurementThread.Start();
        }

        //  Drift check start
        private void radButton5_Click(object sender, EventArgs e)
        {
            Intensity = new ArrayList();
            this.MeasurementThread = new System.Threading.Thread(Worker3);
            isWorking = true;
            MeasurementThread.Start();
        }

        private void Worker()
        {
            while (isWorking == true)
            {
                try
                {
                    currentWave = Convert.ToInt32(textBox1.Text);
                    int counter = Convert.ToInt32(textBox2.Text) - Convert.ToInt32(textBox1.Text); ;
                    this.Invoke((MethodInvoker)(() => progressBar1.Maximum = counter));

                    for (int i = 0; i < counter; i++)
                    {
                        minControl();
						
                        Thread.Sleep(500);
                        this.Invoke((MethodInvoker)(() => Log.Items.Add(temp + "Analyser at 90°")));
                        Thread.Sleep(delay_ms);
                        I90 = readSequenceADC(3, 10);
                        arrayI90.Add(I90);
                        radRadialGauge2.Value = I90;
                        this.Invoke((MethodInvoker)(() => Log.Items.Add("I90 = " + I90)));

                        serialPort1.Write("backward2");
                        temp = serialPort1.ReadLine();
                        this.Invoke((MethodInvoker)(() => Log.Items.Add(temp + "Analyser at 45°")));
                        Thread.Sleep(delay_ms);
                        I45 = readSequenceADC(3, 10);
                        arrayI45.Add(I45);
                        radRadialGauge2.Value = I45;
                        this.Invoke((MethodInvoker)(() => Log.Items.Add("I45 = " + I45)));

                        serialPort1.Write("backward2");
                        temp = serialPort1.ReadLine();
                        this.Invoke((MethodInvoker)(() => Log.Items.Add(temp + "Analyser at 0°")));
                        Thread.Sleep(delay_ms);
                        I0 = readSequenceADC(3, 10);
                        arrayI0.Add(I0);
                        radRadialGauge2.Value = I0;
                        this.Invoke((MethodInvoker)(() => Log.Items.Add("I0 = " + I0)));

                        serialPort1.Write("forward2");
                        temp = serialPort1.ReadLine();
                        count(I0, I45, I90);
                        String s = String.Format("n = {0:F} k = {1:F} \t\t", n, k);
                        this.Invoke((MethodInvoker)(() => Log.Items.Add(s)));

                        arrayN.Add(n);
                        arrayK.Add(k);
                        arrayLamda.Add(currentWave);

                        DrawGraph(currentWave, n, k);

                        this.Invoke((MethodInvoker)(() => progressBar1.Value++));
                        currentWave++;

                        serialPort1.Write("forward1");
                        temp = serialPort1.ReadLine();

                        maxControl();
                    }
                    PrintExel.ExportToExcelIntensity(arrayLamda, arrayI0, arrayI45, arrayI90);
                    PrintExel.ExportToExcel(arrayLamda, arrayN, arrayK);
                    serialPort1.Write("setVoltage"); // Reset High voltage
                    radRadialGauge1.Value = 0;
                    radRadialGauge2.Value = 0;
                    isWorking = false;
                }
                catch (Exception e) { MessageBox.Show("Error:\n" + e.Message); }
            }

        }

        private void Worker2()
        {
            while (isWorking == true)
            {
                try
                {
                    currentWave = Convert.ToInt32(textBox1.Text);
                    int counter = Convert.ToInt32(textBox2.Text) - Convert.ToInt32(textBox1.Text); ;
                    this.Invoke((MethodInvoker)(() => progressBar1.Maximum = counter));

                    for (int i = 0; i < counter; i++)
                    {
                        //this.Invoke((MethodInvoker)(() => Log.Items.Add(temp + "Analyser at 90°")));
                        Thread.Sleep(3000); // signal integrating for 3 seconds
                        I90 = readSequenceADC(3, 10);
                        Intensity.Add(I90);
                        this.Invoke((MethodInvoker)(() => Log.Items.Add("I = " + I90)));

                        arrayLamda.Add(currentWave);
                        // DrawGraph(currentWave, n, k);
                        this.Invoke((MethodInvoker)(() => progressBar1.Value++));
                        currentWave++;
                        serialPort1.Write("forward1");
                        temp = serialPort1.ReadLine();
                    }
                    isWorking = false;
                }
                catch (Exception e) { MessageBox.Show("Error:\n" + e.Message); }
            }

        }

        // Intensity Drifting Check
        private void Worker3()
        {
            while (isWorking == true)
            {
                try
                {
                    for (int i = 0; i < 1000; i++)
                    {
                        Thread.Sleep(200);
                        I = readADC(3);
                        Intensity.Add(I);
                        this.Invoke((MethodInvoker)(() => Log.Items.Add("I = " + I)));
                        DrawGraph(i, I, 0);
                    }
                    PrintExel.ExportText(Intensity);
                    ClearGraph();
                    isWorking = false;
                }
                catch (Exception e) { MessageBox.Show("Error:\n" + e.Message); }
            }

        }



        private void maxControl()
        {
            double val = 1;
            while (val > 0.85)
            {
                this.radRadialGauge1.Value -= 7;
                serialPort1.Write("decrease");
                val = Convert.ToInt32(Math.Abs(readADC(3)));
                Thread.Sleep(1000);
            }
        }

        private void minControl()
        {
            double val = 0;
            while (val < 0.4) // val < 300
            {
                this.radRadialGauge1.Value += 7;
                serialPort1.Write("increase");
                val = Convert.ToInt32(Math.Abs(readADC(3)));
                Thread.Sleep(1000);
            }
        }

        private void DrawGraph(int x, double n, double k)
        {
            // Получим панель для рисования
            GraphPane pane = zedGraphControl1.GraphPane;

            // Создадим список точек
            PointPairList listN = new PointPairList();
            PointPairList listK = new PointPairList();
            LineItem myCurve;

            listN.Add(x, n);
            listK.Add(x, k);

            myCurve = pane.AddCurve("", listN, Color.Blue, SymbolType.Default);
            myCurve = pane.AddCurve("", listK, Color.Red, SymbolType.Default);

            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        // Clear Graph
        private void radButton1_Click(object sender, EventArgs e)
        {
            ClearGraph();
        }

        private void ClearGraph()
        {
            GraphPane pane = zedGraphControl1.GraphPane;
            pane.CurveList.Clear();
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private void DrawGraphOnLoad()
        {
            GraphPane pane = zedGraphControl1.GraphPane;
            SetColors(pane);
            pane.XAxis.Title.Text = "WaveLength";
            pane.YAxis.Title.Text = "n, k";
            pane.Title.Text = "Optical Constants";
            pane.CurveList.Clear();
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();
        }

        private static void SetColors(GraphPane pane)
        {
            // Установим цвет рамки для всего компонента
            pane.Border.Color = Color.DarkGray;
            // Установим цвет рамки вокруг графика
            pane.Chart.Border.Color = Color.Black;
            // Закрасим фон всего компонента ZedGraph
            // Заливка будет сплошная
            pane.Fill.Type = FillType.Solid;
            pane.Fill.Color = Color.WhiteSmoke;
            // Закрасим область графика (его фон) в черный цвет
            pane.Chart.Fill.Type = FillType.Solid;
            pane.Chart.Fill.Color = Color.White;
            // Установим цвет осей
            pane.XAxis.Color = Color.Black;
            pane.YAxis.Color = Color.Black;
            // Установим цвет для подписей рядом с осями
            pane.XAxis.Title.FontSpec.FontColor = Color.Black;
            pane.YAxis.Title.FontSpec.FontColor = Color.Black;
            // Установим цвет подписей под метками
            pane.XAxis.Scale.FontSpec.FontColor = Color.Black;
            pane.YAxis.Scale.FontSpec.FontColor = Color.Black;
            // Установим цвет заголовка над графиком
            pane.Title.FontSpec.FontColor = Color.Black;
            int titleXFontSize = 10;
            int titleYFontSize = 10;
            int legendFontSize = 12;
            int mainTitleFontSize = 12;
            // Установим размеры шрифтов для подписей по осям
            pane.XAxis.Title.FontSpec.Size = titleXFontSize;
            pane.YAxis.Title.FontSpec.Size = titleYFontSize;
            // Установим размеры шрифта для легенды
            pane.Legend.FontSpec.Size = legendFontSize;
            // Установим размеры шрифта для общего заголовка
            pane.Title.FontSpec.Size = mainTitleFontSize;
        }

        /*************** Count Optical Constants *******************************
         * ********************************************************************/
        public void count(float a, float b, float c)
        {
            // I0 = x, I45 = y, I90 = z
            double x = a;
            double y = b;
            double z = c;
            //count delta and ksi
            double ksi = Math.Atan(Math.Sqrt(x / z));
            double delta = Math.Acos((x + z - 2 * y) / (2 * Math.Sqrt(x * z)));

            //count nk
            double denominator = Math.Pow((1 - Math.Sin(2 * ksi) * Math.Cos(delta)), 2);
            double numerator = Math.Pow(Math.Sin(phi), 2) * Math.Pow(Math.Tan(phi), 2) *
                Math.Sin(2 * ksi) * Math.Cos(2 * ksi) * Math.Sin(delta);
            double nk = numerator / denominator;

            //count n^2-k^2
            double numerator1 = Math.Pow(Math.Cos(2 * ksi), 2) - Math.Pow(Math.Sin(2 * ksi), 2) *
                Math.Pow(Math.Sin(delta), 2);
            double nk2 = Math.Pow(Math.Sin(phi), 2) + (Math.Pow(Math.Sin(phi), 2) * Math.Pow(Math.Tan(phi), 2) *
                (numerator1 / denominator));

            //find n, k
            double x1 = 0;
            double x2 = 0;
            double D = nk2 * nk2 - 4 * ((-1) * nk * nk);
            if (D < 0)
            {
                // D < 0
            }
            else if (D == 0)
            {
                x1 = -nk2 / 2;
                x2 = -nk2 / 2;
            }
            else if (D > 0)
            {
                x1 = (-nk2 + Math.Sqrt(D)) / 2;
                x2 = (-nk2 - Math.Sqrt(D)) / 2;
            }
            if (x1 < 0 && x2 < 0)
            {
                // "X1, x2 < 0"
            }
            else if (x1 > 0)
            {
                k = Math.Sqrt(x1);
            }
            else if (x2 > 0)
            {
                k = Math.Sqrt(x2);
            }
            n = Math.Sqrt(nk2 + k * k);
        }



        /********************** Setup Com Ports ********************************
        ***********************************************************************/
        private void Form1_Load_1(object sender, EventArgs e)
        {
            DrawGraphOnLoad();
            radToggleSwitch1.Value = false;
            string[] ports = SerialPort.GetPortNames();
            foreach (string port in ports)
            {
                comboBox1.Items.Add(port);
                ArduinoCom.Items.Add(port);
            }
        }
        private void Form1_FormClosing_1(object sender, FormClosingEventArgs e)
        {
            if (serialPort1.IsOpen) serialPort1.Close();
            else if (PACNET.UART.Close(hPort)) PACNET.UART.Close(hPort);
        }

        // Open ADC Com
        private void button2_Click(object sender, EventArgs e)
        {
            String s = "";
            if (comboBox1.Text != "")
            {
                s += comboBox1.Text;
            }

            try
            {
                hPort = PACNET.UART.Open(s + ",9600,N,8,1");
                Log.Items.Add("Порт АЦП открыт");
                button2.Enabled = false;
                comboBox1.Enabled = false;
            }
            catch
            {
                MessageBox.Show("Порт " + serialPort1.PortName +
                    " неможливо відкрити!", "Помилка!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Log.Items.Add("Порт неможливо відкрити!");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            if (serialPort1.IsOpen)
            {
                ArduinoBtn.Enabled = false;
                ArduinoCom.Hide();
                serialPort1.Close();
            }

            if (ArduinoCom.Text != "")
            {
                serialPort1.PortName = ArduinoCom.Text;
                serialPort1.BaudRate = 9600;
            }
            try
            {
                serialPort1.Open();
                Log.Items.Add("Порт Arduino открыт");
                ArduinoBtn.Enabled = false;
                ArduinoCom.Enabled = false;
            }
            catch
            {
                MessageBox.Show("Port " + serialPort1.PortName + " can't be opened",
                    "Error!", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                Log.Items.Add("Arduino Port Can't Be Opened!");
            }
        }

        /********************** Save To Excel File ****************************
        ***********************************************************************/
        class PrintExel
        {
            public static void ExportToExcel(ArrayList arrLamda, ArrayList arrN, ArrayList arrK)
            {
                try
                {
                    // Загрузить Excel, затем создать новую пустую рабочую книгу
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;
                    excelApp.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                    // Установить заголовки столбцов в ячейках
                    workSheet.Cells[1, "A"] = "WaveLength";
                    workSheet.Cells[1, "B"] = "n";
                    workSheet.Cells[1, "C"] = "k";
                    int row = 1;

                    for (int i = 0; i < arrN.Count; i++)
                    {
                        row++;
                        workSheet.Cells[row, "A"] = arrLamda[i];
                        workSheet.Cells[row, "B"] = arrN[i];
                        workSheet.Cells[row, "C"] = arrK[i];
                    }

                    workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    excelApp.DisplayAlerts = false;
                     workSheet.SaveAs(string.Format(@"{0}\OpticalConstants.xlsx", Environment.CurrentDirectory));
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:\n" + ex.Message);
                }
            }

            public static void ExportText(ArrayList arr)
            {
                FileStream file = new FileStream("d:\\test.txt", FileMode.Create);
                StreamWriter writer = new StreamWriter(file, Encoding.UTF8);
                for (int i = 0; i < arr.Count; i++)
                {
                    writer.WriteLine(arr[i]);
                }
                writer.Close();
            }

            public static void ExportToExcel2(ArrayList arrLamda, ArrayList Intensity)
            {
                try
                {
                    // Загрузить Excel, затем создать новую пустую рабочую книгу
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;
                    excelApp.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                    // Установить заголовки столбцов в ячейках
                    workSheet.Cells[1, "A"] = "WaveLength";
                    workSheet.Cells[1, "B"] = "I";
                    int row = 1;

                    for (int i = 0; i < Intensity.Count; i++)
                    {
                        row++;
                        workSheet.Cells[row, "A"] = arrLamda[i];
                        workSheet.Cells[row, "B"] = Intensity[i];
                    }

                    workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    excelApp.DisplayAlerts = false;
                    //workSheet.SaveAs(string.Format(@"{0}\OpticalIntensity.xlsx", Environment.CurrentDirectory));
                    workSheet.SaveAs(string.Format(@"{0}\OpticalIntensity_" + DateTime.Now.ToString("hh:mm:ss tt", System.Globalization.DateTimeFormatInfo.InvariantInfo) + ".xlsx", Environment.CurrentDirectory));
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:\n" + ex.Message);
                }
            }

            public static void ExportToExcelIntensity(ArrayList arrLamda, ArrayList arrI0, ArrayList arrI45, ArrayList arrI90)
            {
                try
                {
                    // Загрузить Excel, затем создать новую пустую рабочую книгу
                    Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                    excelApp.Visible = true;
                    excelApp.Workbooks.Add();
                    Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                    // Установить заголовки столбцов в ячейках
                    workSheet.Cells[1, "A"] = "WaveLength";
                    workSheet.Cells[1, "B"] = "I0";
                    workSheet.Cells[1, "C"] = "I45";
                    workSheet.Cells[1, "D"] = "I90";
                    int row = 1;

                    for (int i = 0; i < arrLamda.Count; i++)
                    {
                        row++;
                        workSheet.Cells[row, "A"] = arrLamda[i];
                        workSheet.Cells[row, "B"] = arrI0[i];
                        workSheet.Cells[row, "C"] = arrI45[i];
                        workSheet.Cells[row, "D"] = arrI90[i];
                    }

                    workSheet.Range["A1"].AutoFormat(Microsoft.Office.Interop.Excel.XlRangeAutoFormat.xlRangeAutoFormatClassic2);
                    excelApp.DisplayAlerts = false;
                    workSheet.SaveAs(string.Format(@"{0}\OpticalIntensities.xlsx", Environment.CurrentDirectory));
                    excelApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Error:\n" + ex.Message);
                }
            }
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            PrintExel.ExportToExcel(arrayLamda, arrayN, arrayK);
        }

        // Input only digits
        private void textBox1_KeyPress_1(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        // Input only digits
        private void textBox2_KeyPress(object sender, KeyPressEventArgs e)
        {
            if ((e.KeyChar <= 47 || e.KeyChar >= 58) && e.KeyChar != 8)
                e.Handled = true;
        }

        // Info About
        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            MessageBox.Show("Ellipsometer Version 1.0 \n © 2015 Kolchiba Nikita",
               "About Ellipsometer", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }


        private void radButton4_Click(object sender, EventArgs e)
        {
            PrintExel.ExportToExcel2(arrayLamda, Intensity);
        }

        private void saveIntensitiesToolStripMenuItem_Click(object sender, EventArgs e)
        {
            PrintExel.ExportToExcelIntensity(arrayLamda, arrayI0, arrayI45, arrayI90);
        }
    }
}



using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Windows.Devices.Bluetooth;
using Windows.Devices.Bluetooth.Advertisement;
using Windows.Storage.Streams;
using Windows.Devices.Bluetooth.GenericAttributeProfile;
using Windows.Security.Cryptography;
using System.Runtime.InteropServices;
using System.Timers;
using System.Windows.Forms.DataVisualization.Charting;
using System.Drawing;

namespace WindowsFormsApp1
{

    public partial class Form1 : Form
    {
        System.Timers.Timer timer = new System.Timers.Timer();

        public UInt16 deltaT = 0;
        public double timeVal = 0;


        public Form1()
        {
            InitializeComponent();
            timer.AutoReset = true;
            timer.Elapsed += new ElapsedEventHandler(Reconnect);
        }


        public void startBLEwatcher()
        {

            // Create Bluetooth Listener
            var watcher = new BluetoothLEAdvertisementWatcher();

            watcher.ScanningMode = BluetoothLEScanningMode.Active;

            // Only activate the watcher when we're recieving values >= -80
            watcher.SignalStrengthFilter.InRangeThresholdInDBm = -100;

            // Stop watching if the value drops below -90 (user walked away)
            watcher.SignalStrengthFilter.OutOfRangeThresholdInDBm = -100;

            // Register callback for when we see an advertisements
            watcher.Received += OnAdvertisementReceived;

            // Wait 5 seconds to make sure the device is really out of range
            watcher.SignalStrengthFilter.OutOfRangeTimeout = TimeSpan.FromMilliseconds(5000);
            watcher.SignalStrengthFilter.SamplingInterval = TimeSpan.FromMilliseconds(500);

            // Starting watching for advertisements
            watcher.Start();
            SetText("Scanning...");
            this.Invoke((MethodInvoker)delegate { chart1.Series[0].Points.Clear(); });
        }

        public UInt16 ecFlex_idx;
        public UInt16 ecFlex_timer;
        public UInt16 ecFlex_temp;
        public UInt16 ecFlex_adc;
        public BluetoothLEDevice device = null;
        public string fileName = $"{Path.GetDirectoryName(Application.ExecutablePath)}\\Autosave.csv";
        public string newFileName = $"{Path.GetDirectoryName(Application.ExecutablePath)}\\Autosave.csv";
        public bool BLEdisconnectFlag = true;
        public Object[] sensorParam = new object[27]; // your initial array
        public bool senParamsFilled = false;


        public static class Globals
        {
            public static bool DBLCLICK = false;
            public static string UUID_SERVICE = "00002d8d00001000800000805f9b34fb";
            public static string UUID_CHAR = "00002da700001000800000805f9b34fb";

        }

        private void OnAdvertisementReceived(BluetoothLEAdvertisementWatcher watcher, BluetoothLEAdvertisementReceivedEventArgs eventArgs)
        {
            // Thread error workaround
            // Check if there are any manufacturer-specific sections.
            // If there is, print the raw data of the first manufacturer section (if there are multiple).
            string manufacturerDataString = "";
            var manufacturerSections = eventArgs.Advertisement.ManufacturerData;
            // eventArgs.Advertisement.ManufacturerData.Compan
            if (manufacturerSections.Count > 0)
            {
                var manufacturerData = manufacturerSections[0];
                var data = new byte[manufacturerData.Data.Length];
                using (var reader = DataReader.FromBuffer(manufacturerData.Data))
                {
                    reader.ReadBytes(data);
                }
                // Print the company ID + the raw data in hex format.
                manufacturerDataString = string.Format("0x{0}: {1}",
                    manufacturerData.CompanyId.ToString("X"),
                    BitConverter.ToString(data));
            }
            string str = string.Format("\n[{0}] [{1}]: Rssi={2}dBm, localName={3}, manufacturerData=[{4}]",
                eventArgs.Timestamp.ToString("hh\\:mm\\:ss\\.fff"),
                eventArgs.AdvertisementType.ToString(),
                eventArgs.RawSignalStrengthInDBm.ToString(),
                eventArgs.Advertisement.LocalName,
                manufacturerDataString);


            if (eventArgs.Advertisement.LocalName.Contains("ecFlex") || eventArgs.Advertisement.LocalName.Contains("SensorBLEPeripheral"))
            {
                listView1.Invoke((MethodInvoker)delegate ()
                {
                    //var selIndex = listView1.SelectedIndices[1];

                    // Update list with new items, refresh old
                    ListViewItem oldItem = listView1.FindItemWithText(Regex.Replace(eventArgs.BluetoothAddress.ToString("X"), "(.{2})(.{2})(.{2})(.{2})(.{2})(.{2})", "$1:$2:$3:$4:$5:$6"));
                    ListViewItem newItem = (new ListViewItem(new string[] { string.Format("{0}", eventArgs.Advertisement.LocalName), string.Format("{0}", eventArgs.AdvertisementType.ToString()), Regex.Replace(eventArgs.BluetoothAddress.ToString("X"), "(.{2})(.{2})(.{2})(.{2})(.{2})(.{2})", "$1:$2:$3:$4:$5:$6"), Convert.ToString(eventArgs.RawSignalStrengthInDBm), Convert.ToString(eventArgs.BluetoothAddress) }));
                    if (oldItem == null)
                    {
                        listView1.Items.Add(newItem);
                    }
                    else
                    {
                        // Remove from list if power too weak / offline
                        if (eventArgs.RawSignalStrengthInDBm < -130)
                        {
                            listView1.Items[oldItem.Index].Remove();
                        }
                        else
                        {
                            listView1.Items[oldItem.Index].SubItems[0].Text = string.Format("{0}", eventArgs.Advertisement.LocalName);
                            listView1.Items[oldItem.Index].SubItems[3].Text = Convert.ToString(eventArgs.RawSignalStrengthInDBm);
                        }
                    }

                });
            }

            //if(Globals.DBLCLICK)
            // watcher.Stop();




        }


        private void button1_Click(object sender, EventArgs e)
        {
            startBLEwatcher();
        }

        delegate void SetTextCallback(string text);

        private void SetText(string text)
        {
            // InvokeRequired required compares the thread ID of the
            // calling thread to the thread ID of the creating thread.
            // If these threads are different, it returns true.
            if (this.textBox1.InvokeRequired)
            {
                SetTextCallback d = new SetTextCallback(SetText);
                this.Invoke(d, new object[] { text });
            }
            else
            {
                this.textBox1.Text = text;
            }
        }

        public async Task subscribeToNotify()
        {
            // Service
            Guid serviceGuid = Guid.Parse(Globals.UUID_SERVICE);
            GattDeviceServicesResult serviceResult = await device.GetGattServicesForUuidAsync(serviceGuid);

            device.ConnectionStatusChanged += OnConnectionChange;

            // Characteristic (Handle 17 - Sensor Command)
            Guid charGiud = Guid.Parse(Globals.UUID_CHAR);
            var characs = await serviceResult.Services.Single(s => s.Uuid == serviceGuid).GetCharacteristicsAsync();
            var charac = characs.Characteristics.Single(c => c.Uuid == charGiud);

            GattCharacteristicProperties properties = charac.CharacteristicProperties;


            //Write the CCCD in order for server to send notifications.               
            var notifyResult = await charac.WriteClientCharacteristicConfigurationDescriptorAsync(
                                                      GattClientCharacteristicConfigurationDescriptorValue.Notify);

            charac.ValueChanged += Charac_ValueChangedAsync;

        }

        public async Task Connect(ulong addr)
        {
            SetText($"Connecting to {listView1.Items[listView1.FocusedItem.Index].Text}...");

            int retryCounter = 1;
            for (int i = 0; i < retryCounter; i++)
            {
                // Devices
                device = await BluetoothLEDevice.FromBluetoothAddressAsync(addr);


                //device.ConnectionStatus

                if (device.ConnectionStatus != BluetoothConnectionStatus.Connected)
                {
                    await Task.Delay(250);

                }
                else
                {
                    i = retryCounter;
                    SetText($"BLEWATCHER Found: {device.Name}");
                }


            }

            //// Service
            Guid serviceGuid = Guid.Parse(Globals.UUID_SERVICE);
            GattDeviceServicesResult serviceResult = await device.GetGattServicesForUuidAsync(serviceGuid);

            // Subscribe to connection change
            if (serviceResult.Status == GattCommunicationStatus.Success)
            {
                BLEdisconnectFlag = false;
                SetText($"Communicating: Success!");
            }
            device.ConnectionStatusChanged += OnConnectionChange;

            Guid charGiud = Guid.Parse(Globals.UUID_CHAR);
            var characs = await serviceResult.Services.Single(s => s.Uuid == serviceGuid).GetCharacteristicsAsync();

            int readCounter = 0;
            // Read all readable characteristics in this service
            foreach (var character in characs.Characteristics)
            {
                GattCharacteristicProperties properties = character.CharacteristicProperties;
                if ((properties.HasFlag(GattCharacteristicProperties.Read)) && (senParamsFilled == false)) // Chech if not already read
                {
                    var result = await character.ReadValueAsync();

                    CryptographicBuffer.CopyToByteArray(result.Value, out byte[] data);

                    switch (data.Length)
                    {
                        case 1:
                            sensorParam[readCounter] = data[0];
                            break;
                        case 2:
                            sensorParam[readCounter] = BitConverter.ToUInt16(data, 0);
                            break;
                        case 4:
                            sensorParam[readCounter] = BitConverter.ToInt32(data, 0);
                            break;
                        default:
                            sensorParam[readCounter] = Encoding.UTF8.GetString(data);
                            break;
                    }



                    SetText($"Read handle {character.AttributeHandle}: {sensorParam[readCounter]}");
                    if (readCounter > 25)
                    {
                        senParamsFilled = true;
                        this.Invoke((MethodInvoker)delegate { chart1.ChartAreas[0].Axes[1].IsLogarithmic = Convert.ToBoolean(sensorParam[18]); });
                        //this.Invoke((MethodInvoker)delegate { chart1.ChartAreas[0].AxisX.Minimum = calcSensorVal((UInt16)sensorParam[9]);
                        //                                      chart1.ChartAreas[0].AxisX.Maximum = calcSensorVal((UInt16)sensorParam[10]);
                        this.Invoke((MethodInvoker)delegate
                        {
                            chart1.ChartAreas[0].Name = Convert.ToString(sensorParam[4]);
                            chart1.Titles.Add(Convert.ToString(sensorParam[4]));
                            //chart1.ChartAreas[0].AxisX.Title = Convert.ToString(sensorParam[6]);
                            chart1.ChartAreas[0].AxisX.Title = "Time";
                            chart1.ChartAreas[0].AxisY.Title = Convert.ToString(sensorParam[7]);
                        });

                        File.WriteAllText(fileName, $"Start time: {DateTime.Now.ToString("yyyy-MM-ddTHH:mm:sszzz")}{Environment.NewLine}");
                        File.WriteAllText(fileName, $"Counter,Relative time / s,Temperature / degC,{Convert.ToString(sensorParam[7])}{Environment.NewLine}");

                        //});


                    }

                    readCounter++;

                }

                // these are other sorting flags that can be used so sort characterisics.
                if (properties.HasFlag(GattCharacteristicProperties.Write))
                {
                    //SetText("This characteristic supports writing.");
                }
                if (properties.HasFlag(GattCharacteristicProperties.Notify))
                {
                    // SetText("This characteristic supports subscribing to notifications.");
                    if (character.Uuid == charGiud) // sensorCommand
                    {
                        //Write the CCCD in order for server to send notifications.               
                        var notifyResult = await character.WriteClientCharacteristicConfigurationDescriptorAsync(
                                                                  GattClientCharacteristicConfigurationDescriptorValue.Notify);
                        character.ValueChanged += Charac_ValueChangedAsync;

                    }
                }

            }
        }

        public async Task Disconnect()
        {
            BLEdisconnectFlag = true;
            //senParamsFilled = false;
            timer.Stop();
            // Service
            Guid serviceGuid = Guid.Parse(Globals.UUID_SERVICE);
            GattDeviceServicesResult serviceResult = await device.GetGattServicesForUuidAsync(serviceGuid);

            device.ConnectionStatusChanged -= OnConnectionChange;

            // Characteristic (Handle 17 - Sensor Command)
            Guid charGiud = Guid.Parse(Globals.UUID_CHAR);
            var characs = await serviceResult.Services.Single(s => s.Uuid == serviceGuid).GetCharacteristicsAsync();
            var charac = characs.Characteristics.Single(c => c.Uuid == charGiud);


            charac.ValueChanged -= Charac_ValueChangedAsync;
            //  device.Dispose();
            // device = null;
            //GC.Collect();
        }





        public struct Data
        {
            public short counter;
            public short time;
            public short temp;
            public short sens;

        }
        public void Charac_ValueChangedAsync(GattCharacteristic sender, GattValueChangedEventArgs args)
        {

            // Dont run if we do not wish to be connected anymore OR we are still waiting for all parameters
            if (BLEdisconnectFlag || !senParamsFilled) return;


            if (deltaT > 0)
            {
                timer.Interval = 3 * deltaT;
                timer.Start();
            }

            CryptographicBuffer.CopyToByteArray(args.CharacteristicValue, out byte[] data);



            //Asuming Encoding is in ASCII, can be UTF8 or other!
            string dataFromNotify = Encoding.ASCII.GetString(data);
            byte[] bytes = Encoding.UTF8.GetBytes(dataFromNotify);
            GCHandle handle = GCHandle.Alloc(bytes, GCHandleType.Pinned);
            Data data1 = (Data)Marshal.PtrToStructure(handle.AddrOfPinnedObject(), typeof(Data));
            handle.Free();
            // SetText($"Handle 17: {data1.counter} - {(double)da ta1.time/1000} - {(double)data1.temp/10} - {data1.sens}");

            UInt16 old_ecFlex_timer = ecFlex_timer;
            UInt16 oldDeltaT = deltaT;
            deltaT = (UInt16)((UInt16)((data[3] << 8) + data[2]) - ecFlex_timer);
            if (deltaT == 0)
                deltaT = oldDeltaT;

            double oldTimeVal = (double)(deltaT * ecFlex_idx / 1000);
            ecFlex_idx = (UInt16)((data[1] << 8) + data[0]);
            ecFlex_timer = (UInt16)((data[3] << 8) + data[2]);
            ecFlex_temp = (UInt16)((data[5] << 8) + data[4]);
            ecFlex_adc = (UInt16)((data[7] << 8) + data[6]);
            double sensorVal = calcSensorVal(ecFlex_adc);




            if (oldTimeVal > 0) // No point of logging data if not ready
            {
                // Workaround for nonsensical time values
                if (((double)deltaT * (double)ecFlex_idx / 1000) > (3 * deltaT / 1000 + timeVal))
                {
                    timeVal = timeVal + (double)deltaT / 1000;
                }
                else
                {
                    timeVal = ((double)deltaT * (double)ecFlex_idx) / 1000;
                }
                SetText($"{ecFlex_idx} - {(deltaT * (float)ecFlex_idx / 1000)} - {(Single)ecFlex_temp / 10} - {ecFlex_adc}");

                this.Invoke((MethodInvoker)delegate { chart1.Series[0].Points.AddY(sensorVal); chart1.Series[0].Points.AddXY((Single)timeVal, sensorVal); });



                try
                {
                    Invoke((MethodInvoker)delegate { File.AppendAllText(fileName, $"{ecFlex_idx},{timeVal},{(Single)ecFlex_temp / 10},{sensorVal}{Environment.NewLine}"); });
                }
                catch (IOException er)
                {
                    //  if (er.Message.Contains("being used"))
                    { // Just append date if unable to access file.


                        string oldFileName = newFileName;
                        newFileName = ($"{fileName.Substring(0, fileName.Length - 4)}_{DateTime.Now.ToString("yyyyMMddHHmmss")}.csv").Replace("/", "_");
                        File.Copy(oldFileName, newFileName, true);

                        // Dont forget to log the new sample in the new file
                        this.Invoke((MethodInvoker)delegate { File.AppendAllText(newFileName, $"{ecFlex_idx},{timeVal},{(Single)ecFlex_temp / 10},{sensorVal}{Environment.NewLine}"); });

                    }
                }


            }



        }

        public void Reconnect(object sender, ElapsedEventArgs e)
        {
            timer.Stop();
            subscribeToNotify();
            SetText("Reconnected");

        }

        public void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView1.FocusedItem.Index >= 0)
            {//watcher
                Globals.DBLCLICK = true;
                ulong addr = Convert.ToUInt64(listView1.Items[listView1.FocusedItem.Index].SubItems[4].Text);

                // We do not need to set it up again if the user just clicks to reconnect
                if (senParamsFilled == false)
                {
                    this.Invoke((MethodInvoker)delegate { chart1.Series[0].Points.Clear(); });
                    this.Invoke((MethodInvoker)delegate
                    {

                        //chart1.Series[0].XValueType = System.Windows.Forms.DataVisualization.Charting.ChartValueType.DateTime;
                        //chart1.ChartAreas[0].AxisX.IntervalType = DateTimeIntervalType.Milliseconds;
                    });


                    Connect(addr);
                }
                else
                {
                    BLEdisconnectFlag = false;
                    subscribeToNotify();
                }
                chart1.MouseWheel += chart1_MouseWheel;
                //textBox1.Text = "Connected!";
                //subscribeToNotify();
            }
        }


        public async void OnConnectionChange(BluetoothLEDevice bluetoothLEDevice, object args)
        {
            // Dont run if we do not wish to be connected anymore or we are still waiting for all parameters
            if (BLEdisconnectFlag || !senParamsFilled) return;

            SetText($"The device is now: {bluetoothLEDevice.ConnectionStatus}");
            subscribeToNotify();

        }


        private void button2_Click(object sender, EventArgs e)
        {
            Disconnect();
        }

        private double calcSensorVal(UInt32 ADC)
        {
            double D0 = Convert.ToDouble(sensorParam[0]); // ADC resolution * 100
            double N0 = Convert.ToDouble(sensorParam[1]); // ADC reference voltage in V * 100
            double X0 = Convert.ToDouble(sensorParam[2]); // Virtual ground level in V * 100
            double D1 = Convert.ToDouble(sensorParam[3]); // R_TIA * 100
            double N1 = Convert.ToDouble(sensorParam[25]); // Scale factor
            double D2 = Convert.ToDouble(sensorParam[26]); // Scale factor
            //  Int32 D2 = sensorParam[23]; // 
            D1 = D1; // Account for resistor tolerance

            double Vout = (ADC / D0) * N0 - X0 / 100; // Volts



            double val = -100 * (Vout * N1 / D1) / D2;
            //double val = Vout * N1 / D1 / D2;
            this.Invoke((MethodInvoker)delegate {
                listView2.Items[0].SubItems[1].Text = deltaT.ToString("F1");
                listView2.Items[1].SubItems[1].Text = timeVal.ToString("F2");
                listView2.Items[2].SubItems[1].Text = (0.1 * ecFlex_temp).ToString("F1");
                listView2.Items[3].SubItems[1].Text = (1000 * Vout).ToString("F2"); // mV

                if (X0 != 0) // if there is no internal zero it is likely OCP measuremet
                {
                    listView2.Items[4].SubItems[1].Text = val.ToString("F2"); // µA
                }
            });
            return val;
        }

        private void button3_Click(object sender, EventArgs e)
        {
            // Disconnect(); // Make sure we are not still logging data

            // Displays a SaveFileDialog so the user can save the file
            SaveFileDialog saveFileDialog1 = new SaveFileDialog();
            saveFileDialog1.Filter = "Comma separated values|*.csv";
            saveFileDialog1.Title = "Save data";
            saveFileDialog1.ShowDialog();

            // If the file name is not an empty string open it for saving.  
            if (saveFileDialog1.FileName != "")
            {
                try
                {
                    File.Copy(fileName, saveFileDialog1.FileName, true);
                }
                catch (IOException er)
                {
                    if (er.Message.Contains("being used"))
                    {
                        SetText($"Error: Target file in use. File retreiveable at {fileName}");
                    }
                }
            }
        }

        Point? prevPosition = null;
        ToolTip tooltip = new ToolTip();



        private void Form1_Load(object sender, EventArgs e)
        {
            //chart1.ChartAreas["ChartArea1"].AxisX.Interval = 10.0;
            chart1.ChartAreas[0].AxisX.Minimum = 0;

            chart1.ChartAreas[0].AxisX.ScaleView.Zoomable = true;
            chart1.ChartAreas[0].AxisY.ScaleView.Zoomable = true;
            chart1.ChartAreas[0].CursorX.LineColor = Color.Black;
            chart1.ChartAreas[0].CursorX.LineWidth = 1;
            chart1.ChartAreas[0].CursorX.LineDashStyle = ChartDashStyle.Dot;
            chart1.ChartAreas[0].CursorX.Interval = 1;
            chart1.ChartAreas[0].CursorY.LineColor = Color.Black;
            chart1.ChartAreas[0].CursorY.LineWidth = 1;
            chart1.ChartAreas[0].CursorY.LineDashStyle = ChartDashStyle.Dot;
            chart1.ChartAreas[0].CursorY.Interval = 1;
            chart1.ChartAreas[0].AxisX.ScrollBar.Enabled = false;
            chart1.ChartAreas[0].AxisY.ScrollBar.Enabled = false;

            //chart1.ChartAreas[0].IsSameFontSizeForAllAxes = true;
            chart1.ChartAreas[0].AxisX.TitleFont = new Font("Arial", 10, FontStyle.Bold);
            chart1.ChartAreas[0].AxisY.TitleFont = new Font("Arial", 10, FontStyle.Bold);
            listView2.Items.Add($"\u2206Time/ms"); listView2.Items[0].SubItems.Add("");
            listView2.Items.Add("Time/s"); listView2.Items[1].SubItems.Add("");
            listView2.Items.Add("Temp/\u2103"); listView2.Items[2].SubItems.Add("");
            listView2.Items.Add("Vout/mV"); listView2.Items[3].SubItems.Add("");
            listView2.Items.Add("I/µA"); listView2.Items[4].SubItems.Add("");
            // listView2.Items.Add("Distance/m"); listView2.Items[3].SubItems.Add("");

        }

        private void chart1_MouseClick(object sender, MouseEventArgs e)
        {
            var pos = e.Location;
            if (prevPosition.HasValue && pos == prevPosition.Value)
                return;
            tooltip.RemoveAll();
            prevPosition = pos;
            var results = chart1.HitTest(pos.X, pos.Y, false,
                                            ChartElementType.DataPoint);
            foreach (var result in results)
            {
                if (result.ChartElementType == ChartElementType.DataPoint)
                {
                    var prop = result.Object as DataPoint;
                    if (prop != null)
                    {
                        var pointXPixel = result.ChartArea.AxisX.ValueToPixelPosition(prop.XValue);
                        var pointYPixel = result.ChartArea.AxisY.ValueToPixelPosition(prop.YValues[0]);

                        // check if the cursor is really close to the point (2 pixels around the point)
                        if (Math.Abs(pos.X - pointXPixel) < 2 &&
                            Math.Abs(pos.Y - pointYPixel) < 2)
                        {
                            tooltip.Show("X=" + prop.XValue.ToString("F2") + ", Y=" + prop.YValues[0].ToString("F2"), this.chart1,
                                            pos.X, pos.Y - 15);
                        }
                    }
                }
            }
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {

        }

        private void chart1_MouseMove(object sender, MouseEventArgs e)
        {
            Point mousePoint = new Point(e.X, e.Y);

            chart1.ChartAreas[0].CursorX.SetCursorPixelPosition(mousePoint, true);
            chart1.ChartAreas[0].CursorY.SetCursorPixelPosition(mousePoint, true);

            // ...
        }

        private void chart1_MouseWheel(object sender, MouseEventArgs e)
        {
            var chart = (Chart)sender;
            var xAxis = chart.ChartAreas[0].AxisX;
            var yAxis = chart.ChartAreas[0].AxisY;

            try
            {
                if (e.Delta < 0) // Scrolled down.
                {
                    xAxis.ScaleView.ZoomReset();
                    yAxis.ScaleView.ZoomReset();
                }
                else if (e.Delta > 0) // Scrolled up.
                {
                    var xMin = xAxis.ScaleView.ViewMinimum;
                    var xMax = xAxis.ScaleView.ViewMaximum;
                    var yMin = yAxis.ScaleView.ViewMinimum;
                    var yMax = yAxis.ScaleView.ViewMaximum;

                    int zoomSens = 2;

                    var posXStart = xAxis.PixelPositionToValue(e.Location.X) - (xMax - xMin) / zoomSens;
                    var posXFinish = xAxis.PixelPositionToValue(e.Location.X) + (xMax - xMin) / zoomSens;
                    var posYStart = yAxis.PixelPositionToValue(e.Location.Y) - (yMax - yMin) / zoomSens;
                    var posYFinish = yAxis.PixelPositionToValue(e.Location.Y) + (yMax - yMin) / zoomSens;

                    xAxis.ScaleView.Zoom(posXStart, posXFinish);
                    yAxis.ScaleView.Zoom(posYStart, posYFinish);
                }
            }
            catch { }
        }
    }
}

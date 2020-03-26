using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Diagnostics;
using OpenCvSharp;
using OpenCvSharp.UserInterface;
using OpenCvSharp.Extensions;
using Basler.Pylon;
using EasyModbus;
using MCProtocol;

namespace img_save
{
    
    public partial class Form1 : Form
    {
        
        public Form1()
        {
            InitializeComponent();
           
        }
        OpenCvSharp.Point[][] srtpnts;
        bool reinspect = false;     
        string cont_copy;
        //int Model = 0;
        //string time1;
        string foldername = DateTime.Now.ToString("yyyy-MM-dd");
        Mat original = new Mat();
        bool img_flag;
        Camera usbcam;   // camera object create
        bool y1k;
        ModbusClient iomodule;  // easy modbus object creation 
        PixelDataConverter converter = new PixelDataConverter();
        bool btnrun = false;
        bool sftTrig = true;
        bool extTrig = false;
        int t_count, p_count, f_count = 0;
        int i;
        bool trig1;
        bool trig2;
        bool trig3;
        bool trig4;  
        bool re_inspect_enable = false;
        Mat imgInput = new Mat();
        //Stopwatch sw = new Stopwatch();
        Mat iimg = new Mat();
        Mat imgT = new Mat();
        Mat inimg = new Mat();
        Mat disp_img = new Mat();
        FileInfo[] files = null;
        //int totalcount;
        string filenamea;
        string filenamec;
        string filenameyl7a;
        string filenameyl7c;
        string filenameb;
        string filenameyl7b;
        string filecountb;
        string filecounta;
        string filecountc;
        string filename_reinspect;
        string filename_reinspect1;
        int re_count,re_count1;
        int n;
        
        //bool lineactivate = false;
        string delimiter = ",";
        int y1kcount, yl7count;
        
        string appendtext;
        string control;
        
        string time;
        string part_name;
        Mat disp = new Mat();
        int T_y1k=0, P_y1k=0, F_y1k = 0;
        int T_yl7=0, P_yl7=0, F_yl7 = 0;
        Mitsubishi.McProtocolTcp mcProtocolTcp = new Mitsubishi.McProtocolTcp();    // MC protocol object declaration
        int recount = 0;
        int recountf = 0;
        bool recount_flag;
        

        private void Form1_Load(object sender, EventArgs e)
        {
           
            try
            {


                mcProtocolTcp.HostName = "192.168.14.55"; // mitsubhishi PLC ip address
                mcProtocolTcp.PortNumber = 24577;  // plc port number
                mcProtocolTcp.Open(); // open the communication
                if (mcProtocolTcp.Connected)
                {
                    plc_status.Text = "connected";
                }
                else
                {
                    plc_status.Text = "Not Connected";
                }

                // load_settings();
                // Ask the camera finder for a list of camera devices.
                List<ICameraInfo> allCameras = CameraFinder.Enumerate();
                if (allCameras.Count == 0)
                {
                    throw new Exception("No devices found.");
                }
                usbcam = new Camera();
                iomodule = new ModbusClient("COM4");             //create object with com4 modbusclient 
                iomodule.Parity = System.IO.Ports.Parity.Even;      // set parity bit
                iomodule.Connect();                              //connection establish

                // Before accessing camera device parameters, the camera must be opened.
                usbcam.Open();
                //tlStrpPrgsBar1.Value = 20;
                // Set an enum parameter.
                string oldPixelFormat = usbcam.Parameters[PLCamera.PixelFormat].GetValue(); // Remember the current pixel format.
                Console.WriteLine("Old PixelFormat  : {0} ({1})", usbcam.Parameters[PLCamera.PixelFormat].GetValue(), oldPixelFormat);
                //isAvail = Pylon.DeviceFeatureIsAvailable(hDev, "EnumEntry_PixelFormat_Mono8");
                Console.WriteLine("count  : {0}", usbcam.Parameters[PLCamera.AcquisitionMode].GetValue());
                if (!usbcam.Parameters[PLCamera.PixelFormat].TrySetValue(PLCamera.PixelFormat.Mono8))
                {
                    /* Feature is not available. */
                    throw new Exception("Device doesn't support the RGB8 pixel format.");
                }

                // tlStrpPrgsBar1.Value = 50;
                usbcam.Parameters[PLCamera.TriggerMode].TrySetValue(PLCamera.TriggerMode.Off);
                if (extTrig)
                {
                    usbcam.Parameters[PLCamera.TriggerSource].TrySetValue(PLCamera.TriggerSource.Line1);
                    usbcam.Parameters[PLCamera.TriggerMode].TrySetValue(PLCamera.TriggerMode.On);
                }
                
                usbcam.StreamGrabber.ImageGrabbed += OnImageGrabbed;
                usbcam.StreamGrabber.GrabStopped += OnGrabStopped;
                // tlStrpPrgsBar1.Value = 100;
                toolStripTextBox1.Text = "Camera connected..";
                //tlStrpStsLbl1.Image = global::Group_Pharma_OCR.Properties.Resources.ok;
                if (sftTrig | extTrig)
                {

                    //grab_once.Enabled = true;
                    grab_countinue.Enabled = true;
                }

                /* Grab some images in a loop. */


            }
            catch (Exception ex)
            {

                /* Retrieve the error message. */
                toolStripTextBox1.Text = "Exception caught:";
                toolStripTextBox1.Text += ex.Message;
                if (ex.Message != " ")
                {
                    //tlStrpStsLbl1.Text += "Last error message:" + ex.Message;
                    //tlStrpStsLbl1.Text += msg;
                }
                try
                {
                    // Close the camera.
                    if (usbcam != null)
                    {
                        usbcam.Close();
                        usbcam.Dispose();
                    }
                }
                catch (Exception)
                {
                    /*No further handling here.*/
                }

                // Pylon.Terminate();  /* Releases all pylon resources. */
                grab_countinue.Enabled = false;


            }


        }
        
        private void OnImageGrabbed(Object sender, ImageGrabbedEventArgs e)
        {
            //draw = false;
            //draw_roi = false;
            img_flag = true;
           
            if (InvokeRequired)
            {
                // If called from a different thread, we must use the Invoke method to marshal the call to the proper GUI thread.
                // The grab result will be disposed after the event call. Clone the event arguments for marshaling to the GUI thread.
                BeginInvoke(new EventHandler<ImageGrabbedEventArgs>(OnImageGrabbed), sender, e.Clone());
                return;
            }
            // Thread.Sleep(200);
            // iomodule.WriteSingleCoil(1283, false);
            //tlStrpPrgsBar1.Value = 20;
            Mat img, grayimg = new Mat();
            
            if (btnrun == true)
            {
                Byte min, max;
                //PylonGrabResult_t grabResult;
                //SetText("Grab Image\r\n");
                /* Grab one single frame from stream channel 0. The
                camera is set to "single frame" acquisition mode.
                Wait up to 500 ms for the image to be grabbed.
                If imgBuf is null a buffer is automatically created with the right size.*/
                //IGrabResult res = usbcam.StreamGrabber.GrabOne(10000000, TimeoutHandling.Return);
                IGrabResult res = e.GrabResult;
                if (!res.IsValid)
                {
                    /* Timeout occurred. */
                    //SetText(String.Format("Frame {0}: timeout.\r\n", i + 1));
                }

                /* Check to see if the image was grabbed successfully. */
                if (res.GrabSucceeded)
                {
                    img_flag = false;
                    /* Success. Perform image processing. */
                    Bitmap bitmap = new Bitmap(res.Width, res.Height, PixelFormat.Format32bppRgb);
                    // Lock the bits of the bitmap.
                    BitmapData bmpData = bitmap.LockBits(new Rectangle(0, 0, bitmap.Width, bitmap.Height), ImageLockMode.ReadWrite, bitmap.PixelFormat);
                    // Place the pointer to the buffer of the bitmap.
                    converter.OutputPixelFormat = PixelType.BGRA8packed;
                    IntPtr ptrBmp = bmpData.Scan0;
                    converter.Convert(ptrBmp, bmpData.Stride * bitmap.Height, res); //Exception handling TODO
                    bitmap.UnlockBits(bmpData);
                    img = new Mat();
                    img = BitmapConverter.ToMat(bitmap);
                    img.CopyTo(imgInput);
                    toolStripTextBox1.Text = "Image " + n.ToString() + " Captured";
                    control_nor.Text = "";
                    
                  /*
                   * read string here
                   * 
                   */
                    string time1= DateTime.Now.ToString();

                    int len = time1.Length;           
                    for (int j = 0; j < len; j++)                   //creat time as name of the image file to save
                    {
                        if (time1[j].ToString() == "/" || time1[j].ToString() == ":")
                        {
                            textBox2.AppendText("_");

                        }
                        else if (time1[j].ToString() == " ")
                        {
                            textBox2.AppendText("_");
                        }
                        else
                        {
                            textBox2.AppendText(time1[j].ToString());
                        }


                    }
                    time = textBox2.Text.ToString();
                    int[] data = new int[1];
                    mcProtocolTcp.ReadDeviceBlock("D14001", 1, data); //control number will be prasent in 2 register left most bit will be in D14001
                    DectoHec(data[0]); //received data will be in HEX , this will convert HEX to DEC
                    control = control_nor.Text.ToString();
                    if (control == "")
                    {
                        control_nor.AppendText("0");
                    }
                    int[] data1 = new int[1];
                    mcProtocolTcp.ReadDeviceBlock("D14000", 1, data1); // right most 4 bit will be in D14000, it 16 bit data transfer
                    DectoHec(data1[0]); // received data will be in HEX , this will convert HEX to DEC
                   
                    
                   // control_nor.Text = data[0].ToString();
                   
                    control = control_nor.Text.ToString();  // read control nor from text control nor text box
                    
                    if (control.Length ==4)     // to check control is having "0" at second position,  
                    {
                       

                            control_nor.Text = "";
                        for (int k = 0; k < 4; k++)
                        {

                            if (k == 1)
                            {
                                control_nor.AppendText("0");
                                control_nor.AppendText(control[k].ToString());// adds "0" at second position 
                            }
                            else
                            {
                                control_nor.AppendText(control[k].ToString());
                            }


                        }
                       

                    }
                    else if (control.Length == 3)  // to check control is having "00" at second position,  
                    {
                        control_nor.Text = "";
                        for (int k = 0; k < 3; k++)
                        {

                            if (k == 1)
                            {
                                control_nor.AppendText("00");
                                control_nor.AppendText(control[k].ToString());// adds "0" at second position 
                            }
                            else
                            {
                                control_nor.AppendText(control[k].ToString());
                            }


                        }
                    }
                    else if (control.Length == 2)  // to check control is having "000" at second position,  
                    {
                        control_nor.Text = "";
                        for (int k = 0; k < 2; k++)
                        {

                            if (k == 1)
                            {
                                control_nor.AppendText("000");
                                control_nor.AppendText(control[k].ToString());// adds "0" at second position 
                            }
                            else
                            {
                                control_nor.AppendText(control[k].ToString());
                            }


                        }
                    }
                    else if (control.Length == 1)  // to check control is having "0000" at second position,  
                    {
                        control_nor.Text = "";
                      
                        control_nor.AppendText(control[0].ToString());
                        control_nor.AppendText("0000");
                        
                    }
                   
                    control = control_nor.Text.ToString(); // control number data will be in theform of string data type

                    //
                    //

                    //if (re_inspect_enable)
                    //{
                    t_count++;
                    if (cont_copy == control) // this condition will check the retrigger
                    { 
                        t_count--;
                        reinspect = true;
                        if (recount == 0)
                        {
                            if (y1k)
                            {
                                T_y1k--;
                            }
                            else { T_yl7--; }
                            
                        }
                       
                      

                    }
                    if (cont_copy != control)
                    {
                        recountf = 0;
                        recount_flag = false;
                    }
                    cont_copy = control; // this will uses to next cycle to check retrigger
                    
                    


                    part_name = time + "_" + control_nor.Text.ToString();
                   // MessageBox.Show("yes");
                    process_image(img, true); // original images pass to process_images function
                    
                    reinspect = false;



                    bitmap = BitmapConverter.ToBitmap(img);
                    pictureBox1.Image = bitmap;
                    pictureBox1.Image.RotateFlip(RotateFlipType.Rotate270FlipNone); // rotate the picture box with 270 degree
                    textBox2.Text = "";


                    //try
                    //{
                    //    iomodule.WriteSingleCoil(1282, false);
                    //}
                    //catch (Exception ex) { }

                    ////



                }
                else if (!res.GrabSucceeded)
                {
                    //SetText(String.Format("Frame {0} wasn't grabbed successfully.  Error code = {1}\r\n", i + 1, res.ErrorCode));
                }
                ++n;

            }
            //tlStrpStsLbl1.Text = String.Format("Frame {0} grabbed.", i + 1);
            // Dispose the grab result if needed for returning it to the grab loop.
            e.DisposeGrabResultIfClone();
        }
        private void OnGrabStopped(Object sender, GrabStopEventArgs e)
        {
            if (InvokeRequired)
            {
                // If called from a different thread, we must use the Invoke method to marshal the call to the proper thread.
                BeginInvoke(new EventHandler<GrabStopEventArgs>(OnGrabStopped), sender, e);
                return;
            }
            btnrun = false;
            //ckBox_SH.Enabled = true;

            stop.Enabled = false;
            grab_countinue.Enabled = true;
           // grab_once.Enabled = true;
        }

        private void grab_once_Click(object sender, EventArgs e)  // this is not using, this was extra backup feature
        {
            if (re_inspect_enable)
            {
                re_inspect_enable = false;
                grab_once.BackColor = Color.WhiteSmoke;
            }
            else
            {
                grab_once.BackColor = Color.Green;
                re_inspect_enable = true;
            }
            
            

        }

        private void grab_countinue_Click(object sender, EventArgs e)  // continue button click event
        {
            timer1.Enabled = true; // timer will enabled to read trigger for Y1K port "1024"
            timer2.Enabled = true; // timer will enabled to read trigger for Yl7 port "1025"
                                  
            grab();
        }
        private void grab()
        {
           
            if (foldercheck.Checked)
            {
                usbcam.Parameters[PLCamera.TriggerMode].TrySetValue(PLCamera.TriggerSelector.FrameStart);
                usbcam.Parameters[PLCamera.TriggerMode].TrySetValue(PLCamera.TriggerMode.On);
            }
           
            try
            {
                usbcam.Parameters[PLCamera.TriggerSource].TrySetValue(PLCamera.TriggerSource.Software);
                usbcam.Parameters[PLCamera.TriggerMode].TrySetValue(PLCamera.TriggerMode.On);

                // Starts the grabbing of one image.
                usbcam.Parameters[PLCamera.AcquisitionMode].SetValue(PLCamera.AcquisitionMode.Continuous);
                btnrun = true;

                stop.Enabled = true;
                grab_countinue.Enabled = false;
                //grab_once.Enabled = false;
                usbcam.StreamGrabber.Start(GrabStrategy.OneByOne, GrabLoop.ProvidedByStreamGrabber);


            }
            catch (Exception exception)
            {
                MessageBox.Show("ERROR", exception.ToString(), MessageBoxButtons.OK);
            }

        }
        private void stop_Click(object sender, EventArgs e)
        {
            timer1.Enabled = false; // all timer will stop
            timer2.Enabled = false;
            timer4.Enabled = false;
            stop.Enabled = false;
            grab_countinue.Enabled = true;
            //grab_once.Enabled = true;
            try
            {
                usbcam.StreamGrabber.Stop();
            }
            catch (Exception exception)
            {
                //throw (exception);
            }
        }

        private void log_image_Click(object sender, EventArgs e)
        {

            folderBrowserDialog1.ShowDialog();
            if (folderBrowserDialog1.SelectedPath == "")
            {

                MessageBox.Show("error", "unable to log image", MessageBoxButtons.OK);
                return;
            }
           

        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            try
            {
                
                
                bool[] data = iomodule.ReadDiscreteInputs(1024, 2);

                for (int i = 0; i < data.Length; i++)
                {
                    Console.WriteLine("output is: " + data[i].ToString());
                }

                trig1 = trig2;
                trig2 = data[0];

                if (trig1 && !trig2) 
                {
                    //iomodule.WriteSingleCoil(1280, false);
                    //iomodule.WriteSingleCoil(1281, false);
                    T_y1k++;
                    recount = 0;
                    y1k = true;
                    D_Models.SelectedIndex = 0;
                    if (sftTrig)
                    {

                        //sw.Reset();
                        //sw.Start();
                        //usbcam.Parameters[PLCamera.ExposureTimeRaw].SetValue(8995);
                        img_flg.Enabled = true;
                        usbcam.ExecuteSoftwareTrigger();
                        
                        //sw.Stop();
                        //time_lapsed.Text = sw.ElapsedMilliseconds.ToString();

                    }

                }

            }
            catch (Exception ex)
            {
                //tlStrpStsLbl1.Image = global::Group_Pharma_OCR.Properties.Resources.error;
                //tlStrpStsLbl1.Text = "IO Module Error:" + ex.Message;

            }
        }

        private void brows_Click(object sender, EventArgs e)
        {
            openFileDialog1.ShowDialog();
            if (openFileDialog1.FileName == "")
            {
                MessageBox.Show("Error", "unable to load", MessageBoxButtons.OK);
                return;
            }
            imgInput = new Mat(openFileDialog1.FileName, ImreadModes.Color);
            //path.Text = openFileDialog1.FileName.ToString();
            part_name = openFileDialog1.FileName;
            Bitmap bitmap = BitmapConverter.ToBitmap(imgInput);
            pictureBox1.Image = bitmap;

        }

        private void Reset_count_Click(object sender, EventArgs e)
        {
            f_count = 0;
            t_count = 0;
            p_count = 0;
        }


        private void test_Click(object sender, EventArgs e)
        {
            Mat fck = new Mat();
            if (foldercheck.Checked)
            {

                folderBrowserDialog2.ShowDialog();
                if (folderBrowserDialog2.SelectedPath == "")
                {
                    MessageBox.Show("error", "folder not selected cant inspect", MessageBoxButtons.OK);
                    return;
                }
                //t_count = 0;
                //p_count = 0;
                //f_count = 0;
                DirectoryInfo dir = new DirectoryInfo(folderBrowserDialog2.SelectedPath);
                files = dir.GetFiles();
                foreach (FileInfo file in files)
                {
                    fck = new Mat(file.FullName, ImreadModes.Color);
                    part_name = file.FullName;
                    process_image(fck,true);
                }
                return;
            }
            //sw.Reset();
            //sw.Start();
            process_image(imgInput,true);
            //Bitmap bitmapOld = pictureBox1.Image as Bitmap;
            //if (bitmapOld != null)
            //{
            //    // Dispose the bitmap.
            //    bitmapOld.Dispose();
            //}

            Bitmap bitmap = BitmapConverter.ToBitmap(imgInput);
            //Cv2.NamedWindow("a", WindowMode.FreeRatio);
            //Cv2.ImShow("a", disp);
            pictureBox1.Image = bitmap;
            pictureBox1.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
            ////sw.Stop();
            //time_lapsed.Text = sw.ElapsedMilliseconds.ToString();
        }

        
        private void button1_Click(object sender, EventArgs e)
        {

            textBox2.Text = "";

          
            try
            {
                iomodule.WriteSingleCoil(1282, true);
                Thread.Sleep(2000);

                iomodule.WriteSingleCoil(1282, false);
            }
            catch
            {

            }
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            //int[] data = new int[2];
            //// mcProtocolTcp.ReadDeviceBlock("D1400", 2, data);
            //control_nor.Text = data[0].ToString();
        }

        private void timer2_Tick(object sender, EventArgs e)
        {
            try
            {
                
               
                bool[] data1 = iomodule.ReadDiscreteInputs(1025, 2);

                for (int i = 0; i < data1.Length; i++)
                {
                    Console.WriteLine("output is: " + data1[i].ToString());
                }

                trig3 = trig4;
                trig4 = data1[0];

                if (trig3 && !trig4)
                {
                   // iomodule.WriteSingleCoil(1280, false);
                    //iomodule.WriteSingleCoil(1281, false);
                    recount = 0;
                    T_yl7++;
                    iomodule.WriteSingleCoil(1282, true);
                    y1k = false;   // here it will decides model YL7 or Y1K
                    D_Models.SelectedIndex = 1;
                    timer4.Enabled = true; // this is only for YL7 to turn on the light befor grabbing image
                    Thread.Sleep(600); // wait for 600 ms to turn on light 
                    if (sftTrig)
                    {
                        //sw.Reset();
                        //sw.Start();
                        //  usbcam.Parameters[PLCamera.ExposureTimeRaw].SetValue(12250);
                        img_flg.Enabled = true;
                        usbcam.ExecuteSoftwareTrigger();
                        //sw.Stop();
                        //time_lapsed.Text = sw.ElapsedMilliseconds.ToString();

                    }

                }

            }
            catch (Exception ex)
            {
                //tlStrpStsLbl1.Image = global::Group_Pharma_OCR.Properties.Resources.error;
                //tlStrpStsLbl1.Text = "IO Module Error:" + ex.Message;

            }
        }

        private void D_Models_SelectedIndexChanged(object sender, EventArgs e)
        {

            int cmbxitem = D_Models.SelectedIndex;
            if (cmbxitem == 0)
            {
                y1k = true;
            }
            else
            {
                y1k = false;
            }

            
        }
        int today,next = 0;
        
        private void timer3_Tick(object sender, EventArgs e)  // used to check data and time for changeing shifts
        {
            DateTime tim = DateTime.Now;
            int t = tim.Hour;
            int m = tim.Minute;
            toolStripLabel1.Text = tim.ToString();
          // 6:30AM - 3:15 PM shift A
          // 3:15PM - 12 AM shift B
            if (t < 6) 
            {
                shift_box.SelectedIndex = 2;
            }
            else if (t == 6)
            {
                if (m <= 30)
                {
                    shift_box.SelectedIndex = 2;
                }
                else
                {
                    shift_box.SelectedIndex = 0;
                }
            }
            else if (t > 6 && t < 15)
            {
                shift_box.SelectedIndex = 0;
            }
            else if (t == 15)
            {
                if (m <= 15)
                {
                    shift_box.SelectedIndex = 0;
                }
                else
                {
                    shift_box.SelectedIndex = 1;
                }
            }
            else if (t > 15)
            {
                shift_box.SelectedIndex = 1;
            }

            //today = next;
            next = today;
            
            today = tim.Date.Day;
            if (today != next) // decide the date changed or not //if changed it will create all necessory files in Perticular folders
            {
                n = 0;
                foldername = DateTime.Now.ToString("yyyy-MM-dd");
                if (!Directory.Exists(@"D:\Y1K\Fail_images\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\Y1K\Fail_images\" + foldername);
                }
                if (!Directory.Exists(@"D:\Y1K\Pass_images\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\Y1K\Pass_images\" + foldername);
                }
                if (!Directory.Exists(@"D:\YL7\Pass_images"))
                {
                    Directory.CreateDirectory(@"D:\YL7\Pass_images\");
                }
                if (!Directory.Exists(@"D:\YL7\Fail_images\"))
                {
                    Directory.CreateDirectory(@"D:\YL7\Fail_images\");

                }
                if (!Directory.Exists(@"D:\YL7\Fail_images\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\YL7\Fail_images\" + foldername);
                }
                if (!Directory.Exists(@"D:\YL7\Pass_images\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\YL7\Pass_images\" + foldername);
                }
                if (!Directory.Exists(@"D:\reinspect\Y1K\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\reinspect\Y1K\" + foldername);
                }
                if (!Directory.Exists(@"D:\reinspect\YL7\" + foldername))
                {
                    Directory.CreateDirectory(@"D:\reinspect\YL7\" + foldername);
                }
                filenamea = string.Format(@"D:\Y1K\{0:yyyy-MM-dd}_shift_A.csv", DateTime.Now);//create shift-A Y1K CSV file variable data 
                filenameb = string.Format(@"D:\Y1K\{0:yyyy-MM-dd}_shift_B.csv", DateTime.Now);//create shift-B Y1K CSV file variable data 
                filenamec = string.Format(@"D:\Y1K\{0:yyyy-MM-dd}_shift_C.csv", DateTime.Now);//create shift-C y1k csv file
                filenameyl7a = string.Format(@"D:\YL7\{0:yyyy-MM-dd}_shift_A.csv", DateTime.Now);//create shift-A YL7 CSV file variable data 
                filenameyl7b = string.Format(@"D:\YL7\{0:yyyy-MM-dd}_shift_B.csv", DateTime.Now);//create shift-B YL7 CSV file variable data 
                filenameyl7c = string.Format(@"D:\YL7\{0:yyyy-MM-dd}_shift_C.csv", DateTime.Now);//create shift-B YL7 CSV file variable data 
                filecounta = string.Format(@"D:\Count\{0:yyyy-MM-dd}_shift_A.csv", DateTime.Now);// create shift_A count varible 
                filecountb = string.Format(@"D:\Count\{0:yyyy-MM-dd}_shift_B.csv", DateTime.Now);// create shift_B count varible 
                filecountc = string.Format(@"D:\Count\{0:yyyy-MM-dd}_shift_C.csv", DateTime.Now);// create shift_B count varible 
                filename_reinspect = string.Format(@"D:\reinspect\Y1K\{0:yyyy-MM-dd}_reinspected.csv",DateTime.Now);// create reinspect file
                filename_reinspect1 = string.Format(@"D:\reinspect\YL7\{0:yyyy-MM-dd}_reinspected.csv", DateTime.Now);// create reinspect file

                if (!File.Exists(filename_reinspect))    // check file exist or not
                {
                    string d = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;

                    File.WriteAllText(filename_reinspect, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filename_reinspect).Last();
                    var value = lastlin.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filename_reinspect1))    // check file exist or not
                {
                    string d = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;

                    File.WriteAllText(filename_reinspect1, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filename_reinspect1).Last();
                    var value = lastlin.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }

                if (!File.Exists(filenameyl7a))    // check file exist or not
                {
                    string d = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;

                    File.WriteAllText(filenameyl7a, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filenameyl7a).Last();
                    var value = lastlin.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }

                if (!File.Exists(filenameyl7b))    // check file exist or not
                {
                    string d = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;

                    File.WriteAllText(filenameyl7b, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filenameyl7b).Last();
                    var value = lastlin.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filenameyl7c))    // check file exist or not
                {
                    string d = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;

                    File.WriteAllText(filenameyl7c, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filenameyl7c).Last();
                    var value = lastlin.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filenamea))      // check file exist or not
                {
                    string newline = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;
                    File.WriteAllText(filenamea, newline);
                }
                else
                {
                    var lastline = File.ReadAllLines(filenamea).Last();
                    var values = lastline.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filenameb))    // check file exist or not
                {
                    string newline = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;
                    File.WriteAllText(filenameb, newline);
                }
                else
                {
                    var lastline = File.ReadAllLines(filenameb).Last();
                    var values = lastline.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filenamec))      // check file exist or not
                {
                    string newline = "S.No" + delimiter + "filename" + delimiter + "Time" + delimiter + "Result" + Environment.NewLine;
                    File.WriteAllText(filenamec, newline);
                }
                else
                {
                    var lastline = File.ReadAllLines(filenamec).Last();
                    var values = lastline.Split(',');
                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filecounta))                // check file exist or not
                {
                    string d = "0" + delimiter + "0" + delimiter + "0" + Environment.NewLine;

                    File.WriteAllText(filecounta, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filecounta).Last();
                    var valueb = lastlin.Split(',');
                    if (shift_box.SelectedIndex == 0)
                    {
                        T_c.Text = valueb[0];
                        P_c.Text = valueb[1];
                        F_c.Text = valueb[2];
                        t_count = Convert.ToInt16(valueb[0]);
                        p_count = Convert.ToInt16(valueb[1]);
                        f_count = Convert.ToInt16(valueb[2]);

                    }
                    T_c.Text = valueb[0];
                    P_c.Text = valueb[1];
                    F_c.Text = valueb[2];


                    //counts = Convert.ToInt32(values[0]);
                }
                if (!File.Exists(filecountb))              // check file exist or not
                {
                    string d = "0" + delimiter + "0" + delimiter + "0" + Environment.NewLine;

                    File.WriteAllText(filecountb, d);

                }
                else
                {
                    var lastlin = File.ReadAllLines(filecountb).Last();
                    var valueb = lastlin.Split(',');
                    if (shift_box.SelectedIndex == 1)
                    {
                        T_b_c.Text = valueb[0];
                        P_b_c.Text = valueb[1];
                        F_b_c.Text = valueb[2];
                        t_count = Convert.ToInt16(valueb[0]);
                        p_count = Convert.ToInt16(valueb[1]);
                        f_count = Convert.ToInt16(valueb[2]);
                    }
                    T_b_c.Text = valueb[0];
                    P_b_c.Text = valueb[1];
                    F_b_c.Text = valueb[2];
                }
                if (!File.Exists(filecountc))                // check file exist or not
                {
                    string d = "0" + delimiter + "0" + delimiter + "0" + Environment.NewLine;

                    File.WriteAllText(filecountc, d);
                }
                else
                {
                    var lastlin = File.ReadAllLines(filecountc).Last();
                    var valueb = lastlin.Split(',');
                    if (shift_box.SelectedIndex == 2)
                    {
                        T_c_c.Text = valueb[0];
                        P_c_c.Text = valueb[1];
                        F_c_c.Text = valueb[2];
                        t_count = Convert.ToInt16(valueb[0]);
                        p_count = Convert.ToInt16(valueb[1]);
                        f_count = Convert.ToInt16(valueb[2]);

                    }
                    T_c_c.Text = valueb[0];
                    P_c_c.Text = valueb[1];
                    F_c_c.Text = valueb[2];


                    //counts = Convert.ToInt32(values[0]);
                }


            }
           
           
        }

        private void path_TextChanged(object sender, EventArgs e)
        {

        }

        private void process_image(Mat imgIn, bool offline = false)
        {

            
            Mat gray = new Mat();
            Mat thresh = new Mat();
            //imgIn.CopyTo(disp);

            Cv2.CvtColor(imgIn, gray, ColorConversionCodes.BGR2GRAY); // converts gray formate
            gray.CopyTo(disp);
            OpenCvSharp.Point[][] pointss;

            HierarchyIndex[] heyy;

            part_name = time + "_" + control_nor.Text.ToString();
            if (y1k)
            {
                
                y1kcount++;

                Cv2.AdaptiveThreshold(gray, imgT, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 7, 1);
                
                Cv2.FindContours(imgT, out pointss, out heyy, RetrievalModes.Tree, ContourApproximationModes.ApproxNone); // contour will apply to check get connected area
               
                Y1k(pointss, true); // contour results will sends to Y1k function 
                pictureBox1.Image = BitmapConverter.ToBitmap(disp);
               
                pictureBox1.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
            }
            else
            {
                Mat aa = new Mat();
                Mat img_dsp = new Mat();
                Mat save = new Mat();
                imgIn.CopyTo(aa);
                aa.CopyTo(save);
                Bitmap b = BitmapConverter.ToBitmap(aa);
                b.RotateFlip(RotateFlipType.Rotate90FlipXY); // images need to rotate 90
                aa = BitmapConverter.ToMat(b);
                yl7count++;
                
                Cv2.CvtColor(aa, aa, ColorConversionCodes.BGR2GRAY);
                Cv2.AdaptiveThreshold(aa, aa, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 11, 1); // adaptive threshold will help to get the contour because no uniform light condtion
                aa.CopyTo(img_dsp);
                Cv2.CvtColor(img_dsp, img_dsp, ColorConversionCodes.GRAY2BGR); // convert to back to BGR to draw ractangle

                Rect ts = new Rect(424,122,868,226);
                Rect bs = new Rect(220,864,1288,248);
                Rect tr = new Rect(1395,322,220,353);
                Rect tl = new Rect(110,328,250,450);
                Rect bl = new Rect(92,875,68,200);  //875:1075,92:160
                Rect br = new Rect(1560,675,76,315);  // 675:990,1560:1636 
                Mat its = new Mat(aa, ts); // crop the images top side
                Mat ibs = new Mat(aa, bs); // bottom side
                Mat itr = new Mat(aa, tr); // top right
                Mat itl = new Mat(aa, tl); //top left
                Mat ibl = new Mat(aa, bl); // bottom left
                Mat ibr = new Mat(aa, br);  // bottom right
                OpenCvSharp.Point sp = new OpenCvSharp.Point(0, 0);
                OpenCvSharp.Point ep = new OpenCvSharp.Point();
                OpenCvSharp.Size ele = new OpenCvSharp.Size(3, 3);
                //Mat element = Cv2.GetStructuringElement(MorphShapes.Ellipse, ele);
                ep.X = its.Width; ep.Y = its.Height;
                Cv2.Rectangle(its, sp, ep, Scalar.White, 4); // drawing rectangle to separate bead from edge
                ep.X = ibs.Width;ep.Y = ibs.Height;
                Cv2.Rectangle(ibs, sp, ep, Scalar.Black, 4); // drawing rectangle to separate bead from edge
                ep.X = itr.Width;ep.Y = itr.Height;
                

                ep.X = ibl.Width; ep.Y = ibl.Height;

                Cv2.Rectangle(ibl, sp, ep, Scalar.White, 4); // drawing rectangle to separate bead from edge

                int psd_cnt = 0;
                bool flag = false;
               
               
                OpenCvSharp.Point[][] points;
                HierarchyIndex[] hey;
                //
                //top side inspection
                //
                Cv2.FindContours(its, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int c = 0; c < points.Length; c++)
                {
                    Rect r = Cv2.BoundingRect(points[c]);
                    double area = r.Width * r.Height;
                    if (area > 60000 && area < 120000)
                    {
                        if (Cv2.ContourArea(points[c]) > 3500 && Cv2.ContourArea(points[c]) < 6500)
                        {
                            psd_cnt++;
                            flag = true;
                            break;
                            //Cv2.DrawContours(its,points,c,Scalar.Black,3);
                        }
                        //Cv2.DrawContours(its, points, c, Scalar.Black, 3);
                    }
                    
                    //double ca = Cv2.ContourArea(points[c]);
                }
                if (!flag)
                {

                    Cv2.Rectangle(img_dsp, ts, Scalar.Red, 3);
                }
               
                //
                //bottom side inspection
                //
                flag = false;
                Cv2.FindContours(ibs, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int c = 0; c < points.Length; c++)
                {
                    Rect r = Cv2.BoundingRect(points[c]);
                    double area = r.Width * r.Height;
                    if (area > 200000 && area < 300000)
                    {
                        
                            psd_cnt++;
                            flag = true;
                        break;
                        //double ca = Cv2.ContourArea(points[c]);
                    }
                }
                if (!flag)
                {

                    Cv2.Rectangle(img_dsp, bs, Scalar.Red, 3);
                }
                //
                // top right side
                //
                flag = false;
                OpenCvSharp.Point tr_st = new OpenCvSharp.Point(0,100);
                OpenCvSharp.Point tr_ep = new OpenCvSharp.Point(35,170);
                Cv2.Rectangle(itr, tr_st, tr_ep, Scalar.White, -1); // drawing rectangle to separate bead from shadow
                tr_st.X = 0;tr_st.Y = 170;
                tr_ep.X = 88;tr_ep.Y = 220;
                Cv2.Rectangle(itr, tr_st, tr_ep, Scalar.White, -1); // drawing rectangle to separate bead from shadow
                tr_st.X = 0; tr_st.Y = 270;
                tr_ep.X = 139; tr_ep.Y = 352;
                Cv2.Rectangle(itr, tr_st, tr_ep, Scalar.White, -1);  // drawing rectangle to separate bead from shadow
                tr_st.X = 0; tr_st.Y = 0;
                tr_ep.X = 0; tr_ep.Y = 352;
                Cv2.Line(itr, tr_st, tr_ep, Scalar.Black, 4); // drawing rectangle to separate bead from edge
                tr_st.X = 0; tr_st.Y = 353;
                tr_ep.X = 220; tr_ep.Y = 353;
                Cv2.Line(itr, tr_st, tr_ep, Scalar.Black, 4);   // drawing rectangle to separate bead from edge
                tr_st.X = 111; tr_st.Y = 170; 
                tr_ep.X = 220; tr_ep.Y = 170;
                Cv2.Line(itr, tr_st, tr_ep, Scalar.Black, 4);    // drawing rectangle to separate bead from edge
                //Cv2.ImShow("itr", itr);
                Cv2.FindContours(itr, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int c = 0; c < points.Length; c++)
                {
                    Rect r = Cv2.BoundingRect(points[c]);
                    double area = r.Width * r.Height;
                    if (area > 50000 && area < 60000)
                    {
                        if (Cv2.ContourArea(points[c]) > 25000 && Cv2.ContourArea(points[c]) < 35000)
                        {
                            psd_cnt++;
                            flag = true;
                            break;
                        }
                    }
                    double ca = Cv2.ContourArea(points[c]);
                }
                if (!flag)
                {

                    Cv2.Rectangle(img_dsp, tr, Scalar.Red, 3);
                }
                
                //
                //top left side
                //
                flag = false;
                OpenCvSharp.Point tl_st = new OpenCvSharp.Point(155, 128);
                OpenCvSharp.Point tl_ep = new OpenCvSharp.Point(176, 280);
                OpenCvSharp.Point tl_st1 = new OpenCvSharp.Point(208, 46);
                OpenCvSharp.Point tl_ep1 = new OpenCvSharp.Point(247,272);
                Cv2.Rectangle(itl, tl_st, tl_ep, Scalar.White,-1);
                Cv2.Rectangle(itl, tl_st1, tl_ep1, Scalar.White, -1);
                OpenCvSharp.Point tl_st2 = new OpenCvSharp.Point(34, 356);
                OpenCvSharp.Point tl_ep2 = new OpenCvSharp.Point(234, 440);
                Cv2.Rectangle(itl, tl_st2, tl_ep2, Scalar.White, -1);
                ep.X = itl.Width; ep.Y = itl.Height;
                Cv2.Rectangle(itl, sp, ep, Scalar.Black, 4);   // drawing rectangle to separate bead from edge
                //Cv2.ImShow("itl", itl);
                //return;
                Cv2.FindContours(itl, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int c = 0; c < points.Length; c++)
                {
                    Rect r = Cv2.BoundingRect(points[c]);
                    double area = r.Width * r.Height;
                    if (area > 101000 && area < 105500)
                    {
                        if (Cv2.ContourArea(points[c]) > 50000 && Cv2.ContourArea(points[c]) < 70000)
                        {
                            psd_cnt++;
                            flag = true;

                            break;
                        }
                        Cv2.DrawContours(itl, points, c, Scalar.Black, 3);
                    }
                    double ca = Cv2.ContourArea(points[c]);
                }
                if (!flag)
                {

                    Cv2.Rectangle(img_dsp, tl, Scalar.Red, 3);
                }
              
                //
                //bottom right
                //
                flag = false;
                ep.X = ibr.Width; ep.Y = ibr.Height;
                Cv2.Rectangle(ibr, sp, ep, Scalar.White, 4);  // drawing rectangle to separate bead from edge
                Cv2.FindContours(ibr, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                
                for (int c = 0; c < points.Length; c++)
                {
                    Rect r = Cv2.BoundingRect(points[c]);
                    double area = r.Width * r.Height;
                    if (area > 5000 && area < 15000 && r.Height > 310 && r.Height < 340)   // changed
                    {
                        if (Cv2.ContourArea(points[c]) > 500 && Cv2.ContourArea(points[c]) < 3500) // changed
                        {
                            psd_cnt++;
                            flag = true;
                            break;
                        }
                        double ca = Cv2.ContourArea(points[c]);
                    }
                }
                if (!flag)
                {

                    Cv2.Rectangle(img_dsp, br, Scalar.Red, 3);
                }
                //
                //bottom left
                //

                //flag = false;
                //Cv2.FindContours(ibl, out points, out hey, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                //for (int c = 0; c < points.Length; c++)
                //{
                //    Rect r = Cv2.BoundingRect(points[c]);
                //    double area = r.Width * r.Height;
                //    if (area > 5000 && area < 10000 && r.Height > 150 && r.Height < 180)   //10000 > a > 5000 and 180 > h > 150 
                //    {
                //        if (Cv2.ContourArea(points[c]) > 500 && Cv2.ContourArea(points[c]) < 1500)
                //        {
                //            psd_cnt++;
                //            flag = true;
                //            break;
                //        }
                //    }
                //    double ca = Cv2.ContourArea(points[c]);
                //}
                //if (!flag)
                //{

                //    Cv2.Rectangle(img_dsp, bl, Scalar.Red, 3);
                //}

                pictureBox1.Image = BitmapConverter.ToBitmap(imgIn);
                pictureBox1.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);  // rotate the picture box with 270 degree
               

                if(shift_box.SelectedIndex == 0)
                {
                    T_c.Text = t_count.ToString();
                    t_yl7_a.Text = T_yl7.ToString();
                }
                else if(shift_box.SelectedIndex == 1)
                {
                    T_b_c.Text = t_count.ToString();
                    t_yl7_b.Text = T_yl7.ToString();
                }
                else
                {
                    T_c_c.Text = t_count.ToString();
                    t_yl7_c.Text = T_yl7.ToString();
                }
                if (psd_cnt == 5)   //  count the nor of area good
                {
                    result.Text = "PASS";

                    if (reinspect)    // retriggered condition
                    {
                        recount_flag = true;
                       // f_count = f_count - recountf;
                        //F_yl7 = F_yl7 - recountf;
                        P_yl7++;
                        p_count++;

                        iomodule.WriteSingleCoil(1280, true);
                        //iomodule.WriteSingleCoil(1280, false);
                        timer5.Enabled = true;
                        if (shift_box.SelectedIndex == 0)
                        {
                            P_c.Text = p_count.ToString();
                            p_yl7_a.Text = P_yl7.ToString();

                        }
                        else if (shift_box.SelectedIndex == 1)
                        {
                            P_b_c.Text = p_count.ToString();
                            p_yl7_b.Text = P_yl7.ToString();
                        }
                        else
                        {
                            P_c_c.Text = p_count.ToString();
                            p_yl7_c.Text = P_yl7.ToString();
                        }

                        reinspect = false;
                        result.Text = "PASS";
                        result.BackColor = Color.LawnGreen;

                        
                        Cv2.ImWrite(@"D:\reinspect\\YL7\\" + foldername + "\\" + part_name + "origin.bmp", save);
                        //Cv2.CvtColor(save, save, ColorConversionCodes.BGR2GRAY);
                        //Cv2.AdaptiveThreshold(save, save, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 7, 1);
                        Cv2.ImWrite(@"D:\reinspect\\YL7\\" + foldername + "\\" + part_name + ".bmp", img_dsp);
                        pictureBox3.Image = BitmapConverter.ToBitmap(img_dsp);
                        //pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        appendtext = re_count.ToString() + delimiter + part_name + delimiter + time + delimiter + delimiter + "OK" + Environment.NewLine;
                        File.AppendAllText(filename_reinspect1, appendtext);

                    }
                    else
                    {   // normal first time pass
                        recount_flag = true;
                       // f_count = f_count - recountf;
                        //F_yl7 = F_yl7 - recountf;
                        P_yl7++;
                        p_count++;
                        iomodule.WriteSingleCoil(1280, true);
                        timer5.Enabled = true;
                       
                        result.BackColor = Color.LawnGreen;
                       
                        if (shift_box.SelectedIndex == 0)
                        {
                            P_c.Text = p_count.ToString();
                            p_yl7_a.Text = P_yl7.ToString();

                        }
                        else if (shift_box.SelectedIndex == 1)
                        {
                            P_b_c.Text = p_count.ToString();
                            p_yl7_b.Text = P_yl7.ToString();
                        }
                        else
                        {
                            P_c_c.Text = p_count.ToString();
                            p_yl7_c.Text = P_yl7.ToString();
                        }


                        Cv2.ImWrite(@"D:\YL7\\Pass_images\\" + foldername + "\\" + part_name + "origin.bmp", save);
                        //Cv2.CvtColor(save, save, ColorConversionCodes.BGR2GRAY);
                       // Cv2.AdaptiveThreshold(save, save, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 7, 1);
                        Cv2.ImWrite(@"D:\YL7\\Pass_images\\" + foldername + "\\" + part_name + ".bmp", img_dsp);
                        pictureBox3.Image = BitmapConverter.ToBitmap(img_dsp);
                       // pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                        appendtext = yl7count.ToString() + delimiter + part_name + delimiter + time + delimiter + delimiter + "OK" + Environment.NewLine;
                        if (shift_box.SelectedIndex == 0)
                        {
                            File.AppendAllText(filenameyl7a, appendtext);
                        }
                        else if(shift_box.SelectedIndex ==1)
                        {
                            File.AppendAllText(filenameyl7b, appendtext);
                        }
                        else
                        {
                            File.AppendAllText(filenameyl7c, appendtext);
                        }
                    }
                    

                }
                else
                {
                    if (reinspect)  // retrigger condtion fail
                    {
                        recountf++;  // count the nor of fails to avoid the extra count
                        iomodule.WriteSingleCoil(1281, true);  
                        timer5.Enabled = true;
                        result.Text = "FAIL";
                        result.BackColor = Color.Red;
                        
                        pictureBox3.Image = BitmapConverter.ToBitmap(img_dsp);
                        Cv2.ImWrite(@"D:\YL7\\Fail_images\\" + foldername + "\\" + part_name + "origin.bmp", save);
                        Cv2.ImWrite(@"D:\YL7\\Fail_images\\" + foldername + "\\" + part_name + ".bmp", img_dsp);
                        pictureBox3.Image = BitmapConverter.ToBitmap(img_dsp);
                        appendtext = yl7count.ToString() + delimiter + part_name + delimiter + time +  delimiter + "NG" + Environment.NewLine;
                        if (shift_box.SelectedIndex == 0)
                        {
                            File.AppendAllText(filenameyl7a, appendtext);
                        }
                        else if (shift_box.SelectedIndex == 1)
                        {
                            File.AppendAllText(filenameyl7b, appendtext);
                        }
                        else
                        {
                            File.AppendAllText(filenameyl7c, appendtext);
                        }
                        //pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    }
                    else
                    {
                        
                            recountf++;   // count the nor of fails to avoid the extra count
                            
                            iomodule.WriteSingleCoil(1281, true);
                            timer5.Enabled = true;
                            result.Text = "FAIL";
                            result.BackColor = Color.Red;
                          


                            Cv2.ImWrite(@"D:\YL7\\Fail_images\\" + foldername + "\\" + part_name + "origin.bmp", save);
                            //Cv2.CvtColor(save, save, ColorConversionCodes.BGR2GRAY);
                            //Cv2.AdaptiveThreshold(save, save, 255, AdaptiveThresholdTypes.MeanC, ThresholdTypes.Binary, 7, 1);
                            Cv2.ImWrite(@"D:\YL7\\Fail_images\\" + foldername + "\\" + part_name + ".bmp", img_dsp);
                            pictureBox3.Image = BitmapConverter.ToBitmap(img_dsp);
                            // pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                            appendtext = yl7count.ToString() + delimiter + part_name + delimiter + time +  delimiter + "NG" + Environment.NewLine;
                            if (shift_box.SelectedIndex == 0)
                            {
                                File.AppendAllText(filenameyl7a, appendtext);
                            }
                            else if (shift_box.SelectedIndex == 1)
                            {
                                File.AppendAllText(filenameyl7b, appendtext);
                            }
                            else
                            {
                                File.AppendAllText(filenameyl7c, appendtext);
                            }

                        
                       
                    }
                   
                }

            }
                f_count = t_count-p_count; 
                if (y1k)
                {
                    F_y1k = T_y1k - P_y1k;
                }
                else
                {
                    F_yl7 = T_yl7 - P_yl7; 
                }
               
               
            
            
            if (shift_box.SelectedIndex == 0)
            {
                F_c.Text = f_count.ToString();
                f_y1k_a.Text = F_y1k.ToString();
                f_yl7_a.Text = F_yl7.ToString();
                P_c.Text = p_count.ToString();
                p_y1k_a.Text = P_y1k.ToString();
                p_yl7_a.Text = P_yl7.ToString();
            }
            else if (shift_box.SelectedIndex == 1)
            {
                F_b_c.Text = f_count.ToString();
                f_y1k_b.Text = F_y1k.ToString();
                f_yl7_b.Text = F_yl7.ToString();
                P_b_c.Text = p_count.ToString();
                p_y1k_b.Text = P_y1k.ToString();
                p_yl7_b.Text = P_yl7.ToString();
            }
            else
            {
                F_c_c.Text = f_count.ToString();
                f_y1k_c.Text = F_y1k.ToString();
                f_yl7_c.Text = F_yl7.ToString();
                P_c_c.Text = p_count.ToString();
                p_y1k_c.Text = P_y1k.ToString();
                p_yl7_c.Text = P_yl7.ToString();
            }
            //Thread.Sleep(3000);
            if (shift_box.SelectedIndex == 0)
            {
                string f = "TOTAL" + delimiter + "PASS" + delimiter + "FAIL" + Environment.NewLine;
                string d = t_count + delimiter + p_count + delimiter + f_count + Environment.NewLine;
                File.WriteAllText(filecounta, f);
                File.AppendAllText(filecounta, d);
            }
            else if(shift_box.SelectedIndex == 1)
            {
                string f = "TOTAL" + delimiter + "PASS" + delimiter + "FAIL" + Environment.NewLine;
                string d = t_count + delimiter + p_count + delimiter + f_count + Environment.NewLine;
                File.WriteAllText(filecountb, f);
                File.AppendAllText(filecountb, d);
            }
            else
            {
                string f = "TOTAL" + delimiter + "PASS" + delimiter + "FAIL" + Environment.NewLine;
                string d = t_count + delimiter + p_count + delimiter + f_count + Environment.NewLine;
                File.WriteAllText(filecountc, f);
                File.AppendAllText(filecountc, d);
            }
        }

        private void shift_box_SelectedIndexChanged(object sender, EventArgs e)
        {  
            
            // shift changes all counts will be reset

            if (shift_box.SelectedIndex == 0)
            {
                
                t_count = 0;
                p_count = 0;
                f_count = 0;
                F_y1k = T_y1k = F_y1k = T_yl7 = F_yl7 = P_yl7 =P_y1k= 0;

               
            }
            else if (shift_box.SelectedIndex == 1)
            {
                t_count = 0;
                p_count = 0;
                f_count = 0;
                F_y1k = T_y1k = F_y1k = T_yl7 = F_yl7 = P_yl7 =P_y1k= 0;
            }
            else
            {
                t_count = 0;
                p_count = 0;
                f_count = 0;
                F_y1k = T_y1k = F_y1k = T_yl7 = F_yl7 = P_yl7 =P_y1k= 0;
            }
        }

        private void timer4_Tick(object sender, EventArgs e)
        { 
            Thread.Sleep(2500); // YL7 side light will On For 2500ms
            iomodule.WriteSingleCoil(1282, false);
            timer4.Enabled = false;
        }

        private void img_flg_Tick(object sender, EventArgs e)
        { 
            // this function will check the image grabed or not if not grabbed it will stream grabber will restart
            Thread.Sleep(200);
            if (img_flag)
            {
                try
                {
                    usbcam.StreamGrabber.Stop();
                    grab();
                    usbcam.ExecuteSoftwareTrigger();
                }
                catch(Exception ex) { }
                
            }
            img_flg.Enabled = false;
        }

        private void timer5_Tick(object sender, EventArgs e)
        {
            
            Thread.Sleep(1000); // out put signal will turn on for 1000ms
            iomodule.WriteSingleCoil(1280, false);
            iomodule.WriteSingleCoil(1281, false);
            timer5.Enabled = false;

        }

        private void pictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void timer6_Tick(object sender, EventArgs e)
        {

        }

        private void Form1_Closed(object sender, FormClosingEventArgs e)
        {

            try
            {
                iomodule.Disconnect(); // easyModbus communication will disconnect
                usbcam.Close(); // usbcam will close
                usbcam.Dispose(); // usbcam will dispose
                //iomodule.Disconnect();
            }
            catch(Exception ex)
            {
            }

        }

        
        private void Y1k(OpenCvSharp.Point[][] points, bool ascending)
        {
           
            double area,area1;
            string time = DateTime.Now.ToString("HH:mm:ss tt");
           
            if (shift_box.SelectedIndex == 0)
            {
                T_c.Text = t_count.ToString();
                t_y1k_a.Text = T_y1k.ToString();

            }
            else if(shift_box.SelectedIndex ==1)
            {
                T_b_c.Text = t_count.ToString();
                t_y1k_b.Text = T_y1k.ToString();
            }
            else
            {
                T_c_c.Text = t_count.ToString();
                t_y1k_c.Text = T_y1k.ToString();
            }
            OpenCvSharp.Rect rect;
            Mat imgt = new Mat();
            imgT.CopyTo(imgt);
            Cv2.CvtColor(imgT, imgT, ColorConversionCodes.GRAY2RGB);
            Mat label = imgT.EmptyClone();
            Scalar col = new Scalar(255, 0, 0);

            for (i = 0; i < points.Length; i++)
            {
                rect = Cv2.BoundingRect(points[i]);
                
                area1 = rect.Width * rect.Height;  // consider the bounding rect area to check sealer
                if  (area1 > 1320000 && area1 < 1390000)
                {
                    area = Cv2.ContourArea(points[i]);  // consider the contour area to check sealer
                    Cv2.DrawContours(imgT, points, i, col, 3);
                    Cv2.DrawContours(label , points, i, col, 3);
                    //Cv2.NamedWindow("A", WindowMode.FreeRatio);
                    //Cv2.ImShow("A", label);
                    if (1050000 <= area && area <= 1075000)
                    {
                        recount_flag = true;
                        if (reinspect)  // retrigger condition
                        {
                            p_count++;
                            P_y1k++;
                            if (shift_box.SelectedIndex == 0)
                            {
                                P_c.Text = p_count.ToString();
                                p_y1k_a.Text = P_y1k.ToString();
                            }
                            else if (shift_box.SelectedIndex == 1)
                            {
                                P_b_c.Text = p_count.ToString();
                                p_y1k_b.Text = P_y1k.ToString();
                            }
                            else
                            {
                                P_c_c.Text = p_count.ToString();
                                p_y1k_c.Text = P_y1k.ToString();
                            }
                            result.Text = "PASS";
                            result.BackColor = Color.LawnGreen;
                            
                            iomodule.WriteSingleCoil(1280, true);
                            timer5.Enabled = true;
                            //re_count1++;
                            reinspect = false;
                            Cv2.ImWrite(@"D:\reinspect\\Y1K\\" + foldername + "\\" + part_name + "origin.bmp", imgInput);
                            Cv2.ImWrite(@"D:\reinspect\\Y1K\\" + foldername + "\\" + part_name + ".bmp", imgT);
                            //Cv2.ImWrite(@"D:\YL7\\Pass_images\\" + foldername + "\\" + part_name + ".bmp", disp);
                            pictureBox3.Image = BitmapConverter.ToBitmap(imgT);
                            pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                            appendtext = re_count1.ToString() + delimiter + part_name + delimiter + time + delimiter + delimiter + "OK" + Environment.NewLine;
                            File.AppendAllText(filename_reinspect, appendtext);
                            return;
                        }
                        else  // normal condition
                        {
                            p_count++;
                            P_y1k++;

                            iomodule.WriteSingleCoil(1280, true);
                            timer5.Enabled = true;
                            //Cv2.NamedWindow("ad", WindowMode.FreeRatio);
                            //Cv2.ImShow("ad", imgT);
                            //p_count++;
                            //P_y1k++;
                            if (shift_box.SelectedIndex == 0)
                            {
                                P_c.Text = p_count.ToString();
                                p_y1k_a.Text = P_y1k.ToString();
                            }
                            else if (shift_box.SelectedIndex == 1)
                            {
                                P_b_c.Text = p_count.ToString();
                                p_y1k_b.Text = P_y1k.ToString();
                            }
                            else
                            {
                                P_c_c.Text = p_count.ToString();
                                p_y1k_c.Text = P_y1k.ToString();
                            }
                            result.Text = "PASS";
                            result.BackColor = Color.LawnGreen;

                            Cv2.ImWrite(@"D:\Y1K\\Pass_images\\" + foldername + "\\" + part_name + "origin.bmp", imgInput);
                            Mat na = new Mat();
                            Cv2.Threshold(imgInput, na, 26, 255, ThresholdTypes.BinaryInv);
                            Cv2.Add(imgT, na, imgT);
                            Cv2.ImWrite(@"D:\Y1K\\Pass_images\\" + foldername + "\\" + part_name + ".bmp", imgT);
                            pictureBox3.Image = BitmapConverter.ToBitmap(imgT);
                            pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                            appendtext = y1kcount.ToString() + delimiter + part_name + delimiter + time +  delimiter + "OK" + Environment.NewLine;
                            if (shift_box.SelectedIndex == 0)
                            {
                                File.AppendAllText(filenamea, appendtext);
                            }
                            else if (shift_box.SelectedIndex == 1)
                            {
                                File.AppendAllText(filenameb, appendtext);
                            }
                            else
                            {
                                File.AppendAllText(filenamec, appendtext);
                            }

                            //iomodule.WriteSingleCoil(1280, false);
                            return;  // if PASS it will directly return
                        }
                       
                    }
                  

                }
                
               
            }
            
            // if got fail it will divide in 4 parts and check the missed area

                Rect rh = new Rect(810, 85, 365, 1600);
                Rect bs = new Rect(294, 1576, 532, 104); //1576:1680,294:826
                Rect ts = new Rect(331, 42, 465, 81); //42:123,331:795
                Rect lh = new Rect(208, 132, 336, 1456);
               
                
                Mat irhi = new Mat(imgt, rh);
                Mat ibsi = new Mat(imgt, bs);
                Mat itsi = new Mat(imgt, ts);
                Mat ilhi = new Mat(imgt, lh);
                Mat irh = new Mat();
                irhi.CopyTo(irh);
                Mat ibs = new Mat();
                ibsi.CopyTo(ibs);
                Mat its = new Mat();
                itsi.CopyTo(its);
                Mat ilh = new Mat();
                ilhi.CopyTo(ilh);

                bool rrh = false, rbs = false, rts = false;
                bool rect_done = false;
                Scalar rcol = new Scalar(0, 0, 255);
                OpenCvSharp.Point[][] pts; HierarchyIndex[] hierarchies;
                Mat n = new Mat();
                Cv2.Threshold(imgInput, n, 24, 255, ThresholdTypes.BinaryInv);
                Cv2.Add(imgT, n, imgT);
                OpenCvSharp.Point sp = new OpenCvSharp.Point(0, 0);
                OpenCvSharp.Point ep;
                ep.X = irh.Width; ep.Y = irh.Height;
                Cv2.Rectangle(irh, sp, ep, Scalar.White, 2);
                ep.X = ibs.Width; ep.Y = ibs.Height;
                Cv2.Rectangle(ibs, sp, ep, Scalar.White, 2);
                ep.X = its.Width; ep.Y = its.Height;
                Cv2.Rectangle(its, sp, ep, Scalar.White, 2);
                ep.X = ilh.Width; ep.Y = ilh.Height;
                Cv2.Rectangle(ilh, sp, ep, Scalar.White, 2);
                //
                // Right side
                //
                Cv2.FindContours(irh, out pts, out hierarchies, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int a = 0; a < pts.Length; a++)
                {
                    Rect ba = Cv2.BoundingRect(pts[a]);   //550000 > a > 40000 and 1409 < h < 1550 
                    double baa = ba.Width * ba.Height;
                    if (550000 > baa && baa < 400000 && 1400 > ba.Height && ba.Height < 1550)
                    {
                        double ca = Cv2.ContourArea(pts[a]);
                        if (ca < 20000 && ca > 5000)
                        {

                            rrh = true;
                            break;
                        }
                    }
                }
                if (!rrh)
                {
                    rect_done = true;
                    Cv2.Rectangle(imgT, rh, rcol, 3);
                    //Cv2.NamedWindow("a", WindowMode.FreeRatio);
                    //Cv2.ImShow("a", irh);
                }
                // 
                //bottom side
                //
                Cv2.FindContours(ibs, out pts, out hierarchies, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int a = 0; a < pts.Length; a++)
                {
                    Rect ba = Cv2.BoundingRect(pts[a]);
                    double baa = ba.Width * ba.Height;
                    if (30000 > baa && baa > 20000 && ba.Width > 525 && ba.Width < 540)
                    {
                        double ca = Cv2.ContourArea(pts[a]);
                        if (ca < 6000 && ca > 500)
                        {

                            rbs = true;
                            break;
                        }
                    }
                }
                if (!rbs)
                {
                    rect_done = true;
                    Cv2.Rectangle(imgT, bs, rcol, 3);
                    //Cv2.NamedWindow("b", WindowMode.FreeRatio);
                    //Cv2.ImShow("b", ibs);
                }
                //
                //top side
                //
                Cv2.FindContours(its, out pts, out hierarchies, RetrievalModes.Tree, ContourApproximationModes.ApproxNone);
                for (int a = 0; a < pts.Length; a++)
                {
                    Rect ba = Cv2.BoundingRect(pts[a]);
                    double baa = ba.Width * ba.Height;
                    if (20000 > baa && baa > 8000 && ba.Width >450 && ba.Width < 470)
                    {
                        double ca = Cv2.ContourArea(pts[a]);
                        if (ca < 5000 && ca > 1000)
                        {

                            rts = true;
                            break;
                        }
                    }
                }
                if (!rts)
                {
                    rect_done = true;
                    Cv2.Rectangle(imgT, ts, rcol, 3);
                    //Cv2.NamedWindow("c", WindowMode.FreeRatio);
                    //Cv2.ImShow("c", its);
                }
                // 
                //left side
                //
                // it wont check for left side if due to improper sealer
                if (!rect_done)
                {
                    
                    Cv2.Rectangle(imgT, lh, rcol, 3);
                   
                }
                if (reinspect)
                {
                  
                    recountf++; // counts fail times for the same part
                    result.Text = "FAIL";
                    result.BackColor = Color.Red;
                    iomodule.WriteSingleCoil(1281, true); // send the signal
                    timer5.Enabled = true;


                    Cv2.ImWrite(@"D:\Y1k\Fail_images\" + foldername + "\\" + part_name + "origin.bmp", imgInput);
                    Cv2.ImWrite(@"D:\Y1K\Fail_images\" + foldername + "\\" + part_name + ".bmp", imgT);
                    pictureBox3.Image = BitmapConverter.ToBitmap(imgT);
                    pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    appendtext = y1kcount.ToString() + delimiter + part_name + delimiter + time +  delimiter + "NG" + Environment.NewLine;
                    if (shift_box.SelectedIndex == 0)
                    {
                        File.AppendAllText(filenamea, appendtext);
                    }
                    else if (shift_box.SelectedIndex == 1)
                    {
                        File.AppendAllText(filenameb, appendtext);
                    }
                    else
                    {
                        File.AppendAllText(filenamec, appendtext);
                    }
                }
                else
                {
                   
                    recountf++; // // counts fail times for the same part

                result.Text = "FAIL";
                    result.BackColor = Color.Red;
                    iomodule.WriteSingleCoil(1281, true);
                    timer5.Enabled = true;

                    Cv2.ImWrite(@"D:\Y1k\Fail_images\" + foldername + "\\" + part_name + "origin.bmp", imgInput);
                    Cv2.ImWrite(@"D:\Y1K\Fail_images\" + foldername + "\\" + part_name + ".bmp", imgT);
                    pictureBox3.Image = BitmapConverter.ToBitmap(imgT);
                    pictureBox3.Image.RotateFlip(RotateFlipType.Rotate270FlipNone);
                    appendtext = y1kcount.ToString() + delimiter + part_name + delimiter + time + delimiter + "NG" + Environment.NewLine;
                    if (shift_box.SelectedIndex == 0)
                    {
                        File.AppendAllText(filenamea, appendtext);
                    }
                    else if (shift_box.SelectedIndex == 1)
                    {
                        File.AppendAllText(filenameb, appendtext);
                    }
                    else
                    {
                        File.AppendAllText(filenamec, appendtext);
                    }

                }

            




        }
        

        // this will convert dec to hex received data will be in decimal but the designed data will be in HEX
        private void DectoHec(int a)
        {
            string output = "";
            if (a < 0)
            {

                string binary = Convert.ToString(a, 2);


                for (int i = 0; i <= binary.Length - 4; i += 4)
                {
                    output += string.Format("{0:X}", Convert.ToByte(binary.Substring(i, 4), 2));
                }
                control_nor.Text = output;
            }
            else
            {
                int n;
                n = a;
                char[] hexaDeciNum = new char[100];

                // counter for hexadecimal number array 
                int i = 0;
                while (n != 0)
                {
                    // temporary variable to  
                    // store remainder 
                    int temp = 0;

                    // storing remainder in temp 
                    // variable. 
                    temp = n % 16;

                    // check if temp < 10 
                    if (temp < 10)
                    {
                        hexaDeciNum[i] = (char)(temp + 48);
                        i++;
                    }
                    else
                    {
                        hexaDeciNum[i] = (char)(temp + 55);
                        i++;
                    }

                    n = n / 16;
                }

                // printing hexadecimal number  
                // array in reverse order 
                for (int j = i - 1; j >= 0; j--)
                    control_nor.AppendText(hexaDeciNum[j].ToString());
                    

            }
        }
       
        
    }
}

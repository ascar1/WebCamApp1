using OpenCvSharp;
using OpenCvSharp.Extensions;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Drawing.Imaging;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WebCamApp1
{
    public partial class Form1 : Form
    {
        VideoCapture capture;
        Mat frame;
        Bitmap image;
        private Thread camera;
        bool isCameraRunning = false;

        private void CaptureCamera()
        {
            camera = new Thread(new ThreadStart(CaptureCameraCallback));
            camera.Start();
        }

        private void CaptureCameraCallback()
        {

            frame = new Mat();
            capture = new VideoCapture(0);
            //capture.Set(CaptureProperty.FrameHeight, 2000);
            //capture.Set(CaptureProperty.FrameHeight, 1000);
            
            
            capture.Open(0);

            //MessageBox.Show(capture.FrameHeight.ToString());
            if (capture.IsOpened())
            {
                while (isCameraRunning)
                {
                    capture.FrameWidth = 1280;
                    capture.FrameHeight = 1044;


                    capture.Read(frame);
                    image = BitmapConverter.ToBitmap(frame);
                    if (pictureBox1.Image != null)
                    {
                        pictureBox1.Image.Dispose();
                    }
                    pictureBox1.Image = image;
                    
                }
            }
        }

        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Text.Equals("Start"))
            {
                CaptureCamera();
                button1.Text = "Stop";
                isCameraRunning = true;
            }
            else
            {
                capture.Release();
                button1.Text = "Start";
                isCameraRunning = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (isCameraRunning)
            {
                // Take snapshot of the current image generate by OpenCV in the Picture Box
                Bitmap snapshot = new Bitmap(pictureBox1.Image);

                // Save in some directory
                // in this example, we'll generate a random filename e.g 47059681-95ed-4e95-9b50-320092a3d652.png
                // snapshot.Save(@"C:\Users\sdkca\Desktop\mysnapshot.png", ImageFormat.Png);
                snapshot.Save(string.Format(@"C:\2\{0}.png", Guid.NewGuid()), ImageFormat.Png);
                image.Save(string.Format(@"C:\2\{0}.png", Guid.NewGuid()), ImageFormat.Png);

            }
            else
            {
                Console.WriteLine("Cannot take picture if the camera isn't capturing image!");
            }
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            camera.Abort();
            
           
        }
    }
}

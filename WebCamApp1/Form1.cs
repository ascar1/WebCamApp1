using MigraDoc.DocumentObjectModel;
using MigraDoc.Rendering;
using OpenCvSharp;
using OpenCvSharp.Extensions;

using PdfSharp;
using PdfSharp.Pdf;
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
using System.Timers;
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

        BindingSource source = new BindingSource();

        private void CaptureCamera()
        {


            camera = new Thread(new ThreadStart(CaptureCameraCallback));
            camera.Start();
        }

        private void CaptureCameraCallback()
        {

            frame = new Mat();
            //MessageBox.Show(capture.FrameHeight.ToString());
            if (capture.IsOpened())
            {
                capture.FrameWidth = 1280;
                capture.FrameHeight = 720;
                capture.AutoFocus = true;
                
               while (isCameraRunning)
               {

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
            source.DataSource = pages;
            dataGridView1.DataSource = source;



        }

        private void button1_Click(object sender, EventArgs e)
        {
            capture = new VideoCapture(0);
            capture.Open(0);
            if (button1.Text.Equals("Start"))
            {
                CaptureCamera();
                button1.Text = "Stop";
                isCameraRunning = true;
            }
            else
            {
                camera.Abort();
                capture.Release();
                button1.Text = "Start";
                isCameraRunning = false;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            TakePhoto();
        }

        private void Form1_FormClosed(object sender, FormClosedEventArgs e)
        {
            if (camera!= null)
            {
                camera.Abort();
            }
           
            
        }

        static System.Windows.Forms.Timer myTimer = new System.Windows.Forms.Timer();
        private void button3_Click(object sender, EventArgs e)
        {
            isCameraRunning = true;
            myTimer.Tick += new EventHandler(OnTimedEvent);
            myTimer.Interval = 3000;
            
            if( myTimer.Enabled)
            {
                myTimer.Stop();
                myTimer.Dispose();
                button3.Text = "Start";
            }
            else
            {
                //camera.Abort ();
                myTimer.Start();
                button3.Text = "Stop";
                
            }
        }

        private void OnTimedEvent(Object myObject, EventArgs myEventArgs)
        {            
            this.TakePhoto();
            //MessageBox.Show("!");
        }

        List<Page> pages = new List<Page>();
        int i = 1;
        private void TakePhoto ()
        {
            if (isCameraRunning)
            {
                frame = new Mat();
                capture.Read(frame);

                
                // Take snapshot of the current image generate by OpenCV in the Picture Box
                Bitmap snapshot = BitmapConverter.ToBitmap(frame);
                snapshot.RotateFlip(RotateFlipType.Rotate90FlipXY);
                if (pictureBox2.Image != null)
                {
                    pictureBox2.Image.Dispose();
                }
                pictureBox2.Image = snapshot;

                string name = string.Format(@"C:\2\{0}{1}.jpeg", textBox1.Text, i);
                snapshot.Save(name, ImageFormat.Jpeg);
                
                source.Add(new Page() { Num = i, FName = name });
                i++;
                /*name = string.Format(@"C:\2\{0}.jpeg", Guid.NewGuid());
                snapshot.Save(name, ImageFormat.Jpeg);*/
                
            }
            else
            {
                Console.WriteLine("Cannot take picture if the camera isn't capturing image!");
            }
        }

        private void tmp123 (Object tmp)
        {
            MigraDoc.DocumentObjectModel.Shapes.Image image ;
            image = (MigraDoc.DocumentObjectModel.Shapes.Image) tmp;
            image.Height = "297mm";
            image.Width = "210mm";
        }

        private void GetPDF ()
        {
            string name = textBox1.Text + ".pdf";
            Document document = new Document();
            Section section = document.AddSection();
            section.PageSetup.PageFormat = PageFormat.A4;//стандартный размер страницы
            section.PageSetup.Orientation = MigraDoc.DocumentObjectModel.Orientation.Portrait;//ориентация
            section.PageSetup.BottomMargin = 0;//нижний отступ
            section.PageSetup.TopMargin = 0;//верхний отступ
            section.PageSetup.RightMargin = 0; //
            section.PageSetup.LeftMargin = 0;


            foreach (Page page in pages)
            {
                //MigraDoc.DocumentObjectModel.Shapes.Image image = new MigraDoc.DocumentObjectModel.Shapes.Image(page.FName)
                Paragraph paragraph = new Paragraph();               
                
                paragraph.AddImage (page.FName);
                var tmp = paragraph.Elements.LastObject;
                if (tmp is MigraDoc.DocumentObjectModel.Shapes.Image)
                {
                    tmp123(tmp); 
                    
                    
                }
                
                
                section.Add(paragraph);
                
            }
            

            PdfDocumentRenderer pdfRenderer = new PdfDocumentRenderer(true);
            pdfRenderer.Document = document;
            pdfRenderer.RenderDocument();
            string name1 = string.Format(@"C:\2\{0}.pdf", Guid.NewGuid());
            pdfRenderer.PdfDocument.Save(name1);
            MessageBox.Show(name1);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            GetPDF();
        }

        private void button2_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void button2_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                TakePhoto();
            }
        }
    }



    class Page
    {
        public int Num { get; set; }
        public string FName { get; set; }

    }
}

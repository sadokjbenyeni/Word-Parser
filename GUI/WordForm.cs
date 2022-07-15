using Services.Services;
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Threading;
using System.Windows.Forms;
using Timer = System.Timers;

namespace GUI
{
    public partial class WordParser : Form
    {
        public WordParser()
        {
            InitializeComponent();
        }




        private void UploadButtonAction(object sender, EventArgs e)
        {



            // Dislpay the loading spinner
            spinner.Show();

            //Create and instance of Service
            Service service = new Service();
            //Sho< the Dialog
            DialogResult result = openFileDialog1.ShowDialog();
            {
                //Test the result
                if (result == DialogResult.OK )
                {
                    //Get the file name from the choosen one by the user
                    string file = openFileDialog1.FileName;

                    try
                    {
                        //Create an dinstance of the global object 
                        FunctionalTestDocument obj = new FunctionalTestDocument();
                        obj = service.ExtractGlobalFile(file);

                        //Call the create test method to Post the word file 
                        service.CreateTest(obj);

                        MessageBox MessageBox = new MessageBox();
                        MessageBox.Show();
                        Thread.Sleep(3000);

                    }

                    catch (IOException exceptionIO)

                    {
                        Console.WriteLine("Cannot open the document" + exceptionIO.Message);
                    }
                }
            
            }

            //When the test is successful call the Message box of success for 3 seconds
          

            //Exit the application by the end of upload
            Application.Exit();


        }

        private void WordParser_Load(object sender, EventArgs e)
        {

            const int margin = 5;
            int x = Screen.PrimaryScreen.WorkingArea.Right -
                this.Width - margin;
            int y = Screen.PrimaryScreen.WorkingArea.Bottom -
                this.Height + margin;
            this.Location = new Point(x, y);


        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {

        }

        private void CancelButtonAction(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void CloseButtonAction(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void HideButtonAction(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Normal)
            {
                SystemTray.Visible = true;
                SystemTray.BalloonTipText = "WordParser";
                SystemTray.ShowBalloonTip(500);
                SystemTray.BalloonTipTitle = "NEW ACCESS";
                Hide();

            }

        }

        private void SystemTray_MouseClick(object sender, MouseEventArgs e)
        {
            Show();
            this.WindowState = FormWindowState.Normal;
            SystemTray.Visible = false;
        }



    }




}



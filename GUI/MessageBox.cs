using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace GUI
{
    public partial class MessageBox : Form
    {


        public MessageBox()
        {
            InitializeComponent();
        }




        static MessageBox newMessageBox;



        public static void ShowBox()
        {
            newMessageBox = new MessageBox();
            //newMessageBox.ShowDialog();

        }



        private void label1_Click(object sender, EventArgs e)
        {

        }

        public void btnClose_Click()
        {


        }

    }
}

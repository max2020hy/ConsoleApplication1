using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WindowsFormsApplication1
{
    public partial class Form2 : Form
    {
        public static Form2 frm2;
        public Form2()
        {
            InitializeComponent();
        }
        Model mymodel;
        private void Form2_Load(object sender, EventArgs e)
        {
            mymodel = new Model(print);
        }

        public static void print(string str)
        {

            //Text = str;

        }

        private void button1_Click(object sender, EventArgs e)
        {

        }
    }


}

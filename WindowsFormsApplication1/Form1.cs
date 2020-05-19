
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


    public delegate void Model(string str);
    
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }
        public Model mymodel;
        private void Form1_Load(object sender, EventArgs e)
        {
           
            mymodel = new Model(Form2.print);
            mymodel("F1-F2");


        }

        public void print(string str)
        {

            this.Text = str;


        }

        private void button1_Click(object sender, EventArgs e)
        {
            Form2 frm2 = new Form2();

            frm2.Show();
        }
    }
}

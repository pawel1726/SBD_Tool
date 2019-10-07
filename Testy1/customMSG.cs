using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Testy1
{
    public partial class customMSG : Form
    {
        public int BaczNumCustom { get; set; }

        public customMSG()
        {
            InitializeComponent();
        }

        static customMSG CstmMSG;
        static DialogResult result = DialogResult.No;
        
        public static DialogResult Show(object h)
        {
            CstmMSG = new customMSG();
            CstmMSG.ComboCustom.DataSource = h;
            
            CstmMSG.ShowDialog();
            return result;
        }

        

        public void button1_Click(object sender, EventArgs e)
        {
            //MessageBox.Show("wybrano batch (z selected)" + ComboCustom.SelectedItem);
            //BaczNumCustom = Convert.ToInt32(ComboCustom.SelectedItem);

            //BaczNumCustom = ComboCustom.SelectedItem;


            //result = DialogResult.OK;
            //CstmMSG.Close();
            
        }
        
        

    }
}

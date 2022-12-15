using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace InvAddIn
{
    public partial class InputForm : Form
    {
        private Inventor.Application m_inventorApplication;

        //_3DModels oModels = new _3DModels();
        public InputForm(Inventor.Application oInvApp)
        {
            InitializeComponent();

            
            //txtDesigner.Text = UserPrincipal.Current.DisplayName;
            m_inventorApplication = oInvApp;
        }




        private void Form1_Load(object sender, EventArgs e)
        {

        }
    }
}

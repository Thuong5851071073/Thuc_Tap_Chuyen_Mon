using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Design_Dashboard_Modern.DAO;

namespace Design_Dashboard_Modern
{
    public partial class XemHoadon : UserControl
    {
        public XemHoadon()
        {
            InitializeComponent();
           
        }

        private void XemHoadon_Load(object sender, EventArgs e)
        {
            datahoadon.DataSource = Hoadondao.Instance.layhoadon();
            //datahoadon.DataSource = txttimkiem.Text;

        }

     

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string sql = "SELECT * FROM dbo.HoaDon WHERE Tinhtrang LIKE N'%" + textBox1.Text + "%'";
            datahoadon.DataSource = DataProvider.Instance.ExecuteQuery(sql);
        }
    }
}

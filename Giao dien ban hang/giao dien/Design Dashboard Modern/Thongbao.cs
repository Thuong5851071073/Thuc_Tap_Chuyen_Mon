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
    public partial class Thongbao : UserControl
    {
        public Thongbao()
        {
            InitializeComponent();
        }

        private void bunifuThinButton23_Click(object sender, EventArgs e)
        {

        }

        void loadsanphamhethan()
        {
            Dtsanphamhethan.DataSource = sanphamDAO.Instance.ThongbaoSanPhamHetHan();
        }
        private void Thongbao_Load(object sender, EventArgs e)
        {
            dtsp.DataSource = sanphamDAO.Instance.thongbaosplay();
            loadsanphamhethan();
        }

        private void dtsp_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            idhang.Text=dtsp.CurrentRow.Cells[0].Value.ToString();
            txttenhang.Text = dtsp.CurrentRow.Cells[1].Value.ToString();
            txtsl.Text = dtsp.CurrentRow.Cells[4].Value.ToString();
            textBox1.Text= dtsp.CurrentRow.Cells[5].Value.ToString();
            MessageBox.Show("Bạn cần lấy hàng này về !", "Thông báo", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void dtsp_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
    }
}

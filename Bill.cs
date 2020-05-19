using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace chatupon
{
    public partial class Bill : Form
    {
        public Bill()
        {
            InitializeComponent();
        }

        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();

        private void Bill_Load(object sender, EventArgs e)
        {
            DataTable dt = ElesD.Select();
            DGVbill.DataSource = dt;
        }
        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                double x = double.Parse(TBbill1.Text);
                double y = double.Parse(TBbill3.Text);
                double z = z = x * y;
                TBbill4.Text = z.ToString();
            }
            catch
            {
                MessageBox.Show("กรอกข้อมูลผิดพลาด \n กรุณาลองใหม่อีกครั้ง");
            }
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Eles.Number = int.Parse(TBbill0.Text);

            bool success = ElesD.Delete(Eles);

            if (success == true)
            {
                MessageBox.Show("ลบรายการสำเร็จแล้ว");
                Clear();
                DataTable dt = ElesD.Select();
                DGVbill.DataSource = dt;
            }
            else
            {
                MessageBox.Show("กรุณาลองใหม่อีกครั้ง");
            }

        }
        private void Clear()
        {
            TBbill0.Text = "- ไม่สามารถระบุได้ -";
            TBbill1.Text = "- กรุณาระบุ -";
            TBbill2.Text = "- กรุณาระบุ -";
            TBbill3.Text = "- กรุณาระบุ -";
            TBbill4.Text = "- กรุณาคำนวณจำนวนเงิน -";
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Eles.Number = int.Parse(TBbill0.Text);
            Eles.Quantity = int.Parse(TBbill1.Text);
            Eles.Descriiption = TBbill2.Text;
            Eles.UnitPrice = int.Parse(TBbill3.Text);
            Eles.Amount = int.Parse(TBbill4.Text);

            bool success = ElesD.Update(Eles);

            if (success == true)
            {
                MessageBox.Show("อัพเดทรายการสำเร็จแล้ว");
                Clear();
                DataTable dt = ElesD.Select();
                DGVbill.DataSource = dt;
            }
            else
            {
                MessageBox.Show("กรุณาลองใหม่อีกครั้ง");
            }
        }
        private void DGVbill_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //คลิกตัวไหน
            int rowIndex = e.RowIndex;
            //แสดงค่าบน textBox
            TBbill0.Text = DGVbill.Rows[rowIndex].Cells[0].Value.ToString();
            TBbill1.Text = DGVbill.Rows[rowIndex].Cells[1].Value.ToString();
            TBbill2.Text = DGVbill.Rows[rowIndex].Cells[2].Value.ToString();
            TBbill3.Text = DGVbill.Rows[rowIndex].Cells[3].Value.ToString();
            TBbill4.Text = DGVbill.Rows[rowIndex].Cells[4].Value.ToString();      
        }
    }
}

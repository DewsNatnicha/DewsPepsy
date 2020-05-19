using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace chatupon
{
    public partial class Else : Form
    {
        public Else()
        {
            InitializeComponent();
        }

        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();

        private void Else_Load(object sender, EventArgs e)
        {
            DataTable dt = ElesD.Select();
            DGVbill.DataSource = dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Eles.Quantity = int.Parse(textBox1.Text);
            Eles.Descriiption = textBox2.Text;
            Eles.UnitPrice = int.Parse(textBox3.Text);
            Eles.Amount = int.Parse(textBox4.Text);

            //สร้างตัวแปรบูลีน เพื่อตรวจสอบว่ามีการเพิ่มรายการสำเร็จหรือไม่
            bool success = ElesD.Insert(Eles);

            if (success == true)
            {
                //เพิ่มรายการสำเร็จแล้ว
                MessageBox.Show("เพิ่มรายการสำเร็จแล้ว");

                //ลบMethodทิ้ง
                Clear();
                //รีเฟรชตาราง
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
            textBox0.Text = "- ไม่สามารถระบุได้ -";
            textBox1.Text = "- กรุณาระบุ -";
            textBox2.Text = "- กรุณาระบุ -";
            textBox3.Text = "- กรุณาระบุ -";
            textBox4.Text = "- กรุณาคำนวณจำนวนเงิน -";
        }
        private void dataGridView1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            //คลิกตัวไหน
            int rowIndex = e.RowIndex;
            //แสดงค่าบน textBox
            textBox0.Text = DGVbill.Rows[rowIndex].Cells[0].Value.ToString();
            textBox1.Text = DGVbill.Rows[rowIndex].Cells[1].Value.ToString();
            textBox2.Text = DGVbill.Rows[rowIndex].Cells[2].Value.ToString();
            textBox3.Text = DGVbill.Rows[rowIndex].Cells[3].Value.ToString();
            textBox4.Text = DGVbill.Rows[rowIndex].Cells[4].Value.ToString();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Eles.Number = int.Parse(textBox0.Text);
            Eles.Quantity = int.Parse(textBox1.Text);
            Eles.Descriiption = textBox2.Text;
            Eles.UnitPrice = int.Parse(textBox3.Text);
            Eles.Amount = int.Parse(textBox4.Text);

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
        private void button3_Click(object sender, EventArgs e)
        {
            Eles.Number = int.Parse(textBox0.Text);

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
        private void button4_Click(object sender, EventArgs e)
        {
            double x = double.Parse(textBox1.Text);
            double y = double.Parse(textBox3.Text);
            double z = x * y;
            textBox4.Text = z.ToString();
        }
    }
}

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
    public partial class Card : Form
    {
        public Card()
        {
            InitializeComponent();
        }

        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();

        private void button2_Click(object sender, EventArgs e)
        {

            if (button2.Visible)
            {
                #region ราคาต่อหน่วย
                int unit = 0;
                //5บาท
                if (RBp1.Checked)
                { unit = 5; }

                //6บาท
                if (RBp2.Checked)
                { unit = 6; }

                //8บาท
                if (RBp3.Checked)
                { unit = 8; }

                //12บาท
                if (RBp4.Checked)
                { unit = 12; }
                #endregion

                //คำนวณราคา
                int sheet = int.Parse(TBsheet.Text);
                int sum = unit * sheet;
                TBsum.Text = sum.ToString();

                //ขนาดกระดาษ
                string paper = "";
                string page = "";

                if (RBc1.Checked)
                { paper = "พิมพ์การ์ดงานบวช"; }
                if (RBc2.Checked)
                { paper = "พิมพ์การ์ดงานแต่งงาน"; }
                if (RBc3.Checked)
                { paper = "พิมพ์การ์ดงานศพ"; }
                if (RBc4.Checked)
                { paper = "พิมพ์การ์ดงานขึ้นบ้านใหม่"; }

                //รายการ
                string list = paper + page;
                TBcard2.Text = list.ToString();

                //จำนวนแผ่น
                int sheetpaper = int.Parse(TBsheet.Text);
                TBcard1.Text = sheetpaper.ToString();

                //จำนวนต่อหน่วย
                TBcard3.Text = unit.ToString();

                //รวมเงิน
                int sumall = int.Parse(TBsum.Text);
                TBcard4.Text = sumall.ToString();
            }
            else
            { MessageBox.Show("คุณทำรายการผิดพลาด กรุณาลองใหม่อีกครั้ง"); }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            {
                Eles.Quantity = int.Parse(TBcard1.Text);
                Eles.Descriiption = TBcard2.Text;
                Eles.UnitPrice = int.Parse(TBcard3.Text);
                Eles.Amount = int.Parse(TBcard4.Text);

                //สร้างตัวแปรบูลีน เพื่อตรวจสอบว่ามีการเพิ่มรายการสำเร็จหรือไม่
                bool success = ElesD.Insert(Eles);

                if (success == true)
                {
                    //เพิ่มรายการสำเร็จแล้ว
                    MessageBox.Show("เพิ่มรายการสำเร็จแล้ว");

                    //ลบMethodทิ้ง
                    Clear();

                }
                else
                {
                    //เพิ่มรายการผิด
                    MessageBox.Show("คุณทำรายการผิดพลาด กรุณาลองใหม่อีกครั้ง");
                }
            }
        }
        private void Clear()
        {
            TBcard1.Text = "- กรุณาระบุ -";
            TBcard2.Text = "- กรุณาระบุ -";
            TBcard3.Text = "- กรุณาระบุ -";
            TBcard4.Text = "- กรุณาระบุ -";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Bill bill = new Bill();
            bill.ShowDialog();
        }
    }
}
    

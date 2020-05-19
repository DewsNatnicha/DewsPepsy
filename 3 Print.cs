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
    public partial class print : Form
    {
        public print()
        {
            InitializeComponent();
        }

        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();

        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Visible)
            {
                try
                {
                    //ขาว-ดำ
                    if (RBbw.Checked)
                    {
                        //หน้าเดียว
                        if (RBonepage.Checked)
                        { TBunit.Text = "2"; }

                        //หน้า-หลัง
                        if (RBtwopage.Checked)
                        { TBunit.Text = "3"; }
                    }

                    //สี
                    if (RBcolor.Checked)
                    {
                        //หน้าเดียว
                        if (RBonepage.Checked)
                        { TBunit.Text = "3"; }

                        //หน้า-หลัง
                        if (RBtwopage.Checked)
                        { TBunit.Text = "5"; }
                    }
                }
                catch (Exception)
                { 
                    TBunit.Text = "0"; 
                }

                //คำนวณราคา
                int unit = int.Parse(TBunit.Text);
                int sheet = int.Parse(TBsheet.Text);
                int sum = unit * sheet;
                TBsum.Text = sum.ToString();

                //ขนาดกระดาษ
                string paper = "";
                string page = "";

                if (RBbw.Checked)
                { paper = "ปริ้นเอกสาร กระดาษ A4 ขาว-ดำ"; }
                if (RBcolor.Checked)
                { paper = "ปริ้นเอกสาร กระดาษ A4 สี"; }

                //รูปแบบ
                if (RBonepage.Checked)
                { page = " (หน้าเดียว)"; }
                if (RBtwopage.Checked)
                { page = " (หน้า-หลัง)"; }

                //รายการ
                string list = paper + page;
                TBprint2.Text = list.ToString();

                //จำนวนแผ่น
                int sheetpaper = int.Parse(TBsheet.Text);
                TBprint1.Text = sheetpaper.ToString();

                //จำนวนต่อหน่วย
                int unitone = int.Parse(TBunit.Text);
                TBprint3.Text = unitone.ToString();

                //รวมเงิน
                int sumall = int.Parse(TBsum.Text);
                TBprint4.Text = sumall.ToString();
            }
            else
            { MessageBox.Show("คุณทำรายการผิดพลาด กรุณาลองใหม่อีกครั้ง"); }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Eles.Quantity = int.Parse(TBprint1.Text);
            Eles.Descriiption = TBprint2.Text;
            Eles.UnitPrice = int.Parse(TBprint3.Text);
            Eles.Amount = int.Parse(TBprint4.Text);

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
        private void Clear()
        {
            TBprint1.Text = "- กรุณาระบุ -";
            TBprint2.Text = "- กรุณาระบุ -";
            TBprint3.Text = "- กรุณาระบุ -";
            TBprint4.Text = "- กรุณาระบุ -";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Bill bill = new Bill();
            bill.ShowDialog();
        }
    }
}


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
    public partial class Papha : Form
    {
        public Papha()
        {
            InitializeComponent();
        }
        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();
        private void button2_Click(object sender, EventArgs e)
        {
            if (button2.Visible)
            {

                if (RBa4.Checked)
                {
                    //หน้าเดียว
                    if (CBonepage.Checked)
                    { TBonepage.Text = "0.5"; }
                    else { TBonepage.Text = "0"; }

                    //หน้า-หลัง
                    if (CBtwopage.Checked)
                    { TBtwopage.Text = "1"; }
                    else { TBtwopage.Text = "0"; }
                }
                if (RBf14.Checked)
                { 
                    //หน้าเดียว
                    if (CBonepage.Checked)
                    { TBonepage.Text = "1"; }
                    else { TBonepage.Text = "0"; }

                    //หน้า-หลัง
                    if (CBtwopage.Checked)
                    { TBtwopage.Text = "1.5"; }
                    else { TBtwopage.Text = "0"; }
                }

                //ราคาต่อหน่วย
                double onepage = double.Parse(TBonepage.Text);
                double twopage = double.Parse(TBtwopage.Text);
                double sheetonepage = double.Parse(TBopp.Text);
                double sheettwopage = double.Parse(TBtpp.Text);
                double unitsum = Convert.ToDouble(onepage * sheetonepage) + Convert.ToDouble(twopage * sheettwopage);
                string str = String.Format("{0:F2}", unitsum);
                TBunit.Text = str.ToString();
                
                //คำนวณราคา
                double unit = double.Parse(TBunit.Text);
                int sheet = int.Parse(TBsheet.Text);
                double sum = unit * sheet;
                string sumstr = String.Format("{0:F2}", sum);
                TBsum.Text = sumstr.ToString();

                //ขนาดกระดาษ
                string paper = "";
                string page = "";

                if (RBk1.Checked)
                { paper = "พิมพ์ใบฎีกา"; }
                if (RBk2.Checked)
                { paper = "พิมพ์ผ้าป่า"; }
                if (RBk3.Checked)
                { paper = "พิมพ์กฐิน"; }

                if (RBa4.Checked)
                { page = " ขนาดกระดาษ A4"; }
                if (RBf14.Checked)
                { page = " ขนาดกระดาษ F14"; }

                //รายการ
                string list = paper + page;
                TBpapha2.Text = list.ToString();

                //จำนวนแผ่น
                int sheetpaper = int.Parse(TBsheet.Text);
                TBpapha1.Text = sheetpaper.ToString();

                //จำนวนต่อหน่วย
                TBpapha3.Text = unit.ToString();

                //รวมเงิน
                double sumall = double.Parse(TBsum.Text);
                TBpapha4.Text = sumall.ToString();
            }
            else
            { MessageBox.Show("กรุณาลองใหม่อีกครั้ง"); }
        }
        private void CBonepage_CheckedChanged(object sender, EventArgs e)
        {
            if (CBonepage.Checked)
            { TBopp.Enabled = true; }
            else { TBopp.Enabled = false; }
        }
        private void CBtwopage_CheckedChanged(object sender, EventArgs e)
        {
            if (CBtwopage.Checked)
            { TBtpp.Enabled = true; }
            else { TBtpp.Enabled = false; }
        }
        private void button1_Click(object sender, EventArgs e)
        {
            {
                Eles.Quantity = int.Parse(TBpapha1.Text);
                Eles.Descriiption = TBpapha2.Text;

                double UnitPrice = double.Parse(TBpapha3.Text);
                string Unitsum = String.Format("{0:F2}", UnitPrice);
                Eles.UnitPrice = double.Parse(Unitsum);

                double Amount = double.Parse(TBpapha4.Text);
                string Amountsum = String.Format("{0:F2}", Amount);
                Eles.Amount = double.Parse(Amountsum);

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
            TBpapha1.Text = "- กรุณาระบุ -";
            TBpapha2.Text = "- กรุณาระบุ -";
            TBpapha3.Text = "- กรุณาระบุ -";
            TBpapha4.Text = "- กรุณาระบุ -";
        }
        private void button3_Click(object sender, EventArgs e)
        {
            Bill bill = new Bill();
            bill.ShowDialog();
        }
    }
}

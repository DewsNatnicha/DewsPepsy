using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DataTable = Microsoft.Office.Interop.Excel.DataTable;

namespace chatupon
{
    public partial class Copy : Form
    {
        public Copy()
        {
            InitializeComponent();
        }

        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();

        private void button1_Click(object sender, EventArgs e)
        {
            if (button1.Visible)
            {
                try
                {
                    //A4
                    if (RBa4.Checked)
                    {
                        //หน้าเดียว
                        if (RBonepage.Checked)
                        { TBunit.Text = "1"; }

                        //หน้า-หลัง
                        if (RBtwopage.Checked)
                        { TBunit.Text = "2"; }
                    }
                    //A3
                    if (RBa3.Checked)
                    {
                        //หน้าเดียว
                        if (RBonepage.Checked)
                        { TBunit.Text = "3"; }

                        //หน้า-หลัง
                        if (RBtwopage.Checked)
                        { TBunit.Text = "5"; }
                    }
                    //F14
                    if (RBf14.Checked)
                    {
                        //หน้าเดียว
                        if (RBonepage.Checked)
                        { TBunit.Text = "2"; }

                        //หน้า-หลัง
                        if (RBtwopage.Checked)
                        { TBunit.Text = "4"; }
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

                if (RBa4.Checked)
                { paper = "ถ่ายเอกสาร กระดาษ A4"; }
                if (RBa3.Checked)
                { paper = "ถ่ายเอกสาร กระดาษ A3"; }
                if (RBf14.Checked)
                { paper = "ถ่ายเอกสาร กระดาษ F14"; }

                //รูปแบบ
                if (RBonepage.Checked)
                { page = " (หน้าเดียว)"; }
                if (RBtwopage.Checked)
                { page = " (หน้า-หลัง)"; }

                //รายการ
                string list = paper + page;
                TBcopy2.Text = list.ToString();

                //จำนวนแผ่น
                int sheetpaper = int.Parse(TBsheet.Text);
                TBcopy1.Text = sheetpaper.ToString();

                //จำนวนต่อหน่วย
                int unitone = int.Parse(TBunit.Text);
                TBcopy3.Text = unitone.ToString();

                //รวมเงิน
                int sumall = int.Parse(TBsum.Text);
                TBcopy4.Text = sumall.ToString();
            }
            else
            { MessageBox.Show("คุณทำรายการผิดพลาด กรุณาลองใหม่อีกครั้ง"); }
        }
        private void button2_Click(object sender, EventArgs e)
        {
            Eles.Quantity = int.Parse(TBcopy1.Text);
            Eles.Descriiption = TBcopy2.Text;
            Eles.UnitPrice = int.Parse(TBcopy3.Text);
            Eles.Amount = int.Parse(TBcopy4.Text);

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
            TBcopy1.Text = "- กรุณาระบุ -";
            TBcopy2.Text = "- กรุณาระบุ -";
            TBcopy3.Text = "- กรุณาระบุ -";
            TBcopy4.Text = "- กรุณาระบุ -";
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Bill bill = new Bill();
            bill.ShowDialog();
        }

        private void Copy_Load(object sender, EventArgs e)
        {

        }
    }
    
}

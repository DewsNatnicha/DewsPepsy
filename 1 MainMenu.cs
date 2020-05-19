using DGVPrinterHelper;
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
    public partial class Form1 : Form
    {
        #region เครื่องคิดเลข
        Double resultValue = 0;
        String operationPerformed = "";
        bool isOperationPerformed = false;
        #endregion
        public Form1()
        {
            InitializeComponent();
        }
        ElseBill Eles = new ElseBill();
        ElseDel ElesD = new ElseDel();
        DGVPrinter printer = new DGVPrinter();
        #region เครื่องคิดเลข
        private void button_click(object sender, EventArgs e)
        {
            if ((TBresult.Text == "0") || (isOperationPerformed))
                TBresult.Clear();

            isOperationPerformed = false;
            Button button = (Button)sender;
            if (button.Text == ".")
            {
                if (!TBresult.Text.Contains("."))
                    TBresult.Text = TBresult.Text + button.Text;

            }
            else
                TBresult.Text = TBresult.Text + button.Text;
        }
        private void operator_click(object sender, EventArgs e)
        {
            Button button = (Button)sender;

            if (resultValue != 0)
            {
                button15.PerformClick();
                operationPerformed = button.Text;
                label1.Text = resultValue + " " + operationPerformed;
                isOperationPerformed = true;
            }
            else
            {

                operationPerformed = button.Text;
                resultValue = Double.Parse(TBresult.Text);
                label1.Text = resultValue + " " + operationPerformed;
                isOperationPerformed = true;
            }
        }
        private void button6_Click(object sender, EventArgs e)
        {
            TBresult.Text = "0";
        }
        private void button11_Click(object sender, EventArgs e)
        {
            TBresult.Text = "0";
            resultValue = 0;
        }
        private void button21_Click(object sender, EventArgs e)
        {
            switch(operationPerformed)
            {
                case "+":
                    TBresult.Text = (resultValue + Double.Parse(TBresult.Text)).ToString();
                break;
                case "-":
                    TBresult.Text = (resultValue - Double.Parse(TBresult.Text)).ToString();
                break;
                case "*":
                    TBresult.Text = (resultValue * Double.Parse(TBresult.Text)).ToString();
                break;
                case "/":
                    TBresult.Text = (resultValue / Double.Parse(TBresult.Text)).ToString();
                break;
                default:
                    break;
            }
            resultValue = Double.Parse(TBresult.Text);
            label1.Text = "";
        }
        #endregion
        private void Form1_Load(object sender, EventArgs e)
        {
            DataTable dt = ElesD.Select();
            DGVbill.DataSource = dt;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            Copy two = new Copy();
            two.ShowDialog();
        }
        private void button2_Click(object sender, EventArgs e)
        {
            print three = new print();
            three.ShowDialog();
        }
        private void button16_Click(object sender, EventArgs e)
        {
            Card five = new Card();
            five.ShowDialog();
        }
        private void button5_Click(object sender, EventArgs e)
        {
            Papha six = new Papha();
            six.ShowDialog();
        }
        private void button4_Click(object sender, EventArgs e)
        {
            Else seven = new Else();
            seven.ShowDialog();
        }
        private void button24_Click(object sender, EventArgs e)
        {
            Bill bill = new Bill();
            bill.ShowDialog();
        }

        #region ใบเสร็จรับเงิน & ใบส่งสินค้า
        private void button27_Click(object sender, EventArgs e)
        {
            printer.Title = "\r\n\r\n\r\n ใบเสร็จรับเงิน \r\n\r\n ร้านเชตุพน \r\n";
            printer.SubTitle = "ที่อยู่ 91 หมู่ที่ 3 ต.ตาเป๊ก อ.เฉลิมพระเกียรติ จ.บุรีรัมย์ \r\n เลขประจำตัวเสียภาษี 3 3104 00395 641 \r\n โทรศัพท์/โทรสาร 0 4462 8120 และ 081-760-7451 \r\n\r\n";
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = false;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            printer.Footer = "ผู้รับเงิน .................................................................. \r\n วันที่ ......................................................................... \r\n\r\n หมายเหตุ : กรณีชำระด้วยเช็คสั่งจ่าย โปรดสั่งจ่ายในนาม ร้านเชตุพน และประทับตรา A/C PAYEE ONLY";
            printer.FooterSpacing = 15;
            printer.PrintDataGridView(DGVbill);
        }
        private void button28_Click(object sender, EventArgs e)
        {
            printer.Title = "\r\n\r\n\r\n ใบส่งสินค้า \r\n\r\n ร้านเชตุพน \r\n";
            printer.SubTitle = "ที่อยู่ 91 หมู่ที่ 3 ต.ตาเป๊ก อ.เฉลิมพระเกียรติ จ.บุรีรัมย์ \r\n เลขประจำตัวเสียภาษี 3 3104 00395 641 \r\n\r\n โทรศัพท์/โทรสาร 0 4462 8120 และ 081-760-7451 \r\n\r\n";
            printer.SubTitleFormatFlags = StringFormatFlags.LineLimit | StringFormatFlags.NoClip;
            printer.PageNumbers = false;
            printer.PageNumberInHeader = false;
            printer.PorportionalColumns = true;
            printer.HeaderCellAlignment = StringAlignment.Near;
            printer.Footer = "ผู้รับเงิน .................................................................. \r\n วันที่ ......................................................................... \r\n\r\n หมายเหตุ : กรณีชำระด้วยเช็คสั่งจ่าย โปรดสั่งจ่ายในนาม ร้านเชตุพน และประทับตรา A/C PAYEE ONLY";
            printer.FooterSpacing = 15;
            printer.PrintDataGridView(DGVbill);
        }
        #endregion
        private void button26_Click(object sender, EventArgs e)
        {
            Close();
            //Application.Exit();
            //Environment.Exit(0);
        }
    }
}

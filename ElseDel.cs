using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace chatupon
{
    class ElseDel
    {
        static string myconnstrng = ConfigurationManager.ConnectionStrings["connstrng"].ConnectionString;
        
        #region เลือก method รายการ 
        public DataTable Select()
        {
            //เชื่อมต่อ SQL กับ Databaes
            SqlConnection conn = new SqlConnection(myconnstrng);

            //DataTable เก็บข้อมูลจาก database
            DataTable dt = new DataTable();

            try
            {
                //เลือกรายการทั้งหมดจาก database
                String sql = "SELECT * FROM ElesBill";

                //สร้าง SQL Command ค้นหา
                SqlCommand cmd = new SqlCommand(sql, conn);

                //ตัวแปลงข้อมูล SQL เพื่อเก็บค่าจาก database ชั่วคราว
                SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                //เปิดการเชื่อมต่อ Database
                conn.Open();

                adapter.Fill(dt);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }

            return dt;
        }
        #endregion
        #region เพิ่มรายการใน database
        public bool Insert(ElseBill seven)
        {
            //สร้างตัวแปร Boolean และตั้งค่าเริ่มต้นเป็น false (ยังไม่เสร็จ = ยังไม่ทำ)
            bool isSuccess = false;

            //เชื่อมต่อ SQL กับ Database
            SqlConnection conn = new SqlConnection(myconnstrng);
            
            try
            {
                //SQL เพิ่มลง database
                String sql = "INSERT INTO ElesBill(จำนวน, รายการ, หน่วยละ, จำนวนเงิน) VALUES (@Quantity, @Descriiption, @UnitPrice, @Amount)";

                //สร้าง SQL Command เพื่อส่งผ่านค่า
                SqlCommand cmd = new SqlCommand(sql, conn);

                //การส่งผ่านค่าผ่านพารามิเตอร์
                cmd.Parameters.AddWithValue("@Quantity", seven.Quantity);
                cmd.Parameters.AddWithValue("@Descriiption", seven.Descriiption);
                cmd.Parameters.AddWithValue("@UnitPrice", seven.UnitPrice);
                cmd.Parameters.AddWithValue("@Amount", seven.Amount);

                //เปิดการเชื่อมต่อฐานข้อมูล
                conn.Open();

                int rows = cmd.ExecuteNonQuery();

                //ค้นหาสำเร็จ -> ค่าของแถวจะมากกว่า 0 และจะน้อยกว่า 0
                if (rows > 0)
                {
                    //ดำเนินการค้นหาสำเร็จแล้ว
                    isSuccess = true;
                }
                else
                {
                    //ไม่สามารถดำเนินการค้นหาได้
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }

            return isSuccess;
        }
        #endregion
        #region อัพเดทรายการ Database
        public bool Update(ElseBill seven)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(myconnstrng);

            try
            {
                //SQL อัพเดท database
                String sql = "UPDATE ElesBill SET จำนวน=@Quantity, รายการ = @Descriiption, หน่วยละ = @UnitPrice, จำนวนเงิน = @Amount WHERE ที่=@Number";

                //สร้าง SQL Command เพื่อส่งผ่านค่า
                SqlCommand cmd = new SqlCommand(sql, conn);

                //การส่งผ่านค่าผ่านพารามิเตอร์
                cmd.Parameters.AddWithValue("@Quantity", seven.Quantity);
                cmd.Parameters.AddWithValue("@Descriiption", seven.Descriiption);
                cmd.Parameters.AddWithValue("@UnitPrice", seven.UnitPrice);
                cmd.Parameters.AddWithValue("@Amount", seven.Amount);
                cmd.Parameters.AddWithValue("@Number", seven.Number);

                conn.Open();

                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }

            return isSuccess;
        }

        #endregion
        #region ลบรายการ Database
        public bool Delete(ElseBill seven)
        {
            bool isSuccess = false;
            SqlConnection conn = new SqlConnection(myconnstrng);

            try
            {
                String sql = "DELETE FROM ElesBill WHERE ที่=@Number";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@Number", seven.Number);
                conn.Open();

                int rows = cmd.ExecuteNonQuery();
                if (rows > 0)
                {
                    isSuccess = true;
                }
                else
                {
                    isSuccess = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            finally
            {
                conn.Close();
            }

            return isSuccess;
        }
        #endregion
    }
}

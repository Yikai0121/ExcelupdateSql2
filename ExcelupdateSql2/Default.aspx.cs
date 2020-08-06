using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using WebGrease.Activities;

namespace ExcelupdateSql2
{
    public static class MD5Extensions
    {
        public static string ToMD5(this string str)
        {
            using (var cryptoMD5 = System.Security.Cryptography.MD5.Create())
            {
                //將字串編碼成 UTF8 位元組陣列
                var bytes = Encoding.UTF8.GetBytes(str);

                //取得雜湊值位元組陣列
                var hash = cryptoMD5.ComputeHash(bytes);

                //取得 MD5
                var md5 = BitConverter.ToString(hash)
                  .Replace("-", String.Empty)
                  .ToUpper();

                return md5;
            }
        }
    }
    public partial class _Default : Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {

        }

        protected void Button1_Click(object sender, EventArgs e)
        {
            if (IsValid)
            {
                if (FileUpload1.HasFile)
                {
                    String filetype = Path.GetExtension(FileUpload1.FileName).ToLower();
                    if (filetype != ".xlsx")
                    {
                        String error = @"<script>alert('檔案格式錯誤，請上傳.xlsx檔');</script>";
                        ClientScript.RegisterStartupScript(GetType(), filetype, error);
                    }
                    else
                    {
                        String CustomerID, CustomerName;
                        String CompanyName;
                        String GetDate;

                        string path = Path.GetFileName(FileUpload1.FileName);
                        path = path.Replace(" ", "");
                        FileUpload1.SaveAs(Server.MapPath("~/ExcelFile/") + path);
                        String ExcelPath = Server.MapPath("~/ExcelFile/") + path;
                        Label1.Text = "檔案 : " + Path.GetFileName(FileUpload1.FileName);
                        using (OleDbConnection oleDb = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties='Excel 12.0; HDR=Yes;IMEX=1;'"))
                        {
                            oleDb.Open();
                            OleDbCommand cmd = new OleDbCommand("select * from [工作表1$]", oleDb);
                            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("select * from [工作表1$]", oleDb);
                            OleDbDataReader dr = cmd.ExecuteReader();
                            int count = 0;
                            while (dr.Read())
                            {
                                count++;
                                CustomerID = dr[0].ToString().Trim();
                                CustomerName = dr[1].ToString().Trim().ToMD5();
                                CompanyName = dr[2].ToString().Trim().ToMD5();
                                GetDate = dr[3].ToString().Trim();

                                if (string.IsNullOrEmpty(CustomerID))
                                {
                                    
                                    int SuceessCount = count - 1;
                                    Page.ClientScript.RegisterStartupScript(GetType(), "", "<script>alert('查無第" + count + " 筆資料。" + SuceessCount + "筆資料匯入成功')</script>");
                                    this.GridView1.Visible = false;
                                    return;
                                }
                                Savedata(CustomerID, CustomerName, CompanyName, GetDate);

                            }
                            dr.Close();
                            oleDb.Close();
                        }

                        this.GridView1.Visible = false;
                        ClientScript.RegisterStartupScript(GetType(), "", "<script> alert('上傳成功!') </script> ");
                        

                    }
                }
                else
                {

                    ClientScript.RegisterStartupScript(GetType(), "", "<script> alert('請上傳檔案!') </script> ");


                }
            }


        }
        private void Savedata(string CustomerID, string CustomerName, String CompanyName, string GetDate)
        {
            String query = "insert into Customer(CustomerID,CustomerName,CompanyName,GetDate) values(" + CustomerID + ",'" + CustomerName + "','" + CompanyName + "','" + GetDate + "')";
            String mycon = "Server=PC-MISI5\\SQLEXPRESS;Database=Customer;Trusted_Connection=True;MultipleActiveResultSets=True";
            using (SqlConnection con = new SqlConnection(mycon))
            {
                con.Open();
                SqlCommand cmd = new SqlCommand(query, con);
                cmd.ExecuteNonQuery();
                con.Close();
            }


        }

        protected void Button2_Click(object sender, EventArgs e)
        {
            if (IsValid)
            {

                if (FileUpload1.HasFile)
                {
                    String filetype = Path.GetExtension(FileUpload1.FileName).ToLower();
                    if (filetype != ".xlsx")
                    {
                        String error = @"<script>alert('檔案格式錯誤，請上傳.xlsx檔');</script>";
                        ClientScript.RegisterStartupScript(GetType(), filetype, error);
                    }
                    else
                    {
                        string path = Path.GetFileName(FileUpload1.FileName);
                        path = path.Replace(" ", "");
                        FileUpload1.SaveAs(Server.MapPath("~/ExcelFile/") + path);
                        String ExcelPath = Server.MapPath("~/ExcelFile/") + path;
                        Label1.Text = "檔案 : " + Path.GetFileName(FileUpload1.FileName);
                        using (OleDbConnection oleDb = new OleDbConnection("Provider = Microsoft.ACE.OLEDB.12.0; Data Source = " + ExcelPath + "; Extended Properties='Excel 12.0; HDR=Yes;IMEX=1;'"))
                        {
                            oleDb.Open();
                            OleDbCommand cmd = new OleDbCommand("select * from [工作表1$]", oleDb);
                            OleDbDataAdapter oleDbDataAdapter = new OleDbDataAdapter("select * from [工作表1$]", oleDb);
                            DataSet ds = new DataSet();
                            oleDbDataAdapter.Fill(ds, "Excel");
                            this.GridView1.DataSource = ds.Tables["Excel"].DefaultView;
                            this.GridView1.DataBind();

                            oleDb.Close();
                        }


                    }
                }
                else
                {

                    ClientScript.RegisterStartupScript(GetType(), "", "<script> alert('請上傳檔案!') </script> ");


                }

            }
        }

    }
}
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Forms;
using ClosedXML.Excel;
using MessageBox = System.Windows.Forms.MessageBox;

namespace ShippingControl
{
   public class cs_Connect
    {
        private static SqlConnection con = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlSystem"].ConnectionString);

        public static string[,] Built_Variables(string listvar, string listval)
        {
            string[] var = listvar.Split('|');
            string[] val = listval.Split('|');
            string[,] variable = new string[var.Length, 2];
            for (int i = 0; i < var.Length; i++)
            {
                variable[i, 0] = var[i]; variable[i, 1] = val[i];
            }
            return variable;
        }
        //lấy data từ sql
        public static DataTable ExecuteQuery(string query, object[] parameter = null)
        {
            DataTable data = new DataTable();

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlSystem"].ConnectionString))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                SqlDataAdapter adapter = new SqlDataAdapter(command);

                adapter.Fill(data);

                connection.Close();
            }

            return data;
        }

        public static object ExecuteScalar(string query, object[] parameter = null)
        {
            object data = 0;

            using (SqlConnection connection = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlSystem"].ConnectionString))
            {
                connection.Open();

                SqlCommand command = new SqlCommand(query, connection);

                if (parameter != null)
                {
                    string[] listPara = query.Split(' ');
                    int i = 0;
                    foreach (string item in listPara)
                    {
                        if (item.Contains('@'))
                        {
                            command.Parameters.AddWithValue(item, parameter[i]);
                            i++;
                        }
                    }
                }

                data = command.ExecuteScalar();

                connection.Close();
            }

            return data;
        }
        public static string Get_DataTable(ref DataTable table, string proc, string[,] variable)
        {

            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };
            if (variable != null)
                for (int i = 0; i < variable.GetLength(0); i++)
                    command.Parameters.AddWithValue(variable[i, 0], variable[i, 1]);

            try
            {
                con.Open();
                // create data adapter
                SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                // this will query your database and return the result to your datatable
                dataAdapter.Fill(table);
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString();
            }
        }
        //Hàm thực thi câu lệnh Insert, Update vào DB
        public static string Insert_Data(string proc, string[,] variable)
        {

            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };
            for (int i = 0; i < variable.GetLength(0); i++)
                command.Parameters.AddWithValue(variable[i, 0], variable[i, 1]);

            try
            {
                con.Open();
                IAsyncResult result = command.BeginExecuteNonQuery();
                while (!result.IsCompleted)
                {
                    System.Threading.Thread.Sleep(100);
                }
                int count = command.EndExecuteNonQuery(result);

                //command.ExecuteNonQuery();
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString().Split('.')[0];
            }
        }
        //Hàm Update cả DataTable vào DB
        public static string Insert_Table(string proc, DataTable table, string user)
        {
            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };
          
            // command.Parameters.AddWithValue("@table", table);
            var dcm = command.Parameters.AddWithValue("@table", table);
            dcm.SqlDbType = SqlDbType.Structured;

            command.Parameters.AddWithValue("@user", user);

            try
            {
                con.Close();
                con.Open();                
                command.ExecuteNonQuery();
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString();
            }
        }

        //Hàm Update cả DataTable vào DB
        public static string Insert_Table_RFF_Data(string proc, DataTable table, string user,string RFFNO, string RFFLine)
        {
            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };

            // command.Parameters.AddWithValue("@table", table);
            var dcm = command.Parameters.AddWithValue("@table", table);
            dcm.SqlDbType = SqlDbType.Structured;

            command.Parameters.AddWithValue("@user", user);
            command.Parameters.AddWithValue("@RFFNo", RFFNO);
            command.Parameters.AddWithValue("@RFFLine", RFFLine);

            try
            {
                con.Close();
                con.Open();
                command.ExecuteNonQuery();
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString();
            }
        }

        public static string Insert_And_Get_Table(ref DataTable tb, string proc, DataTable table, string user)
        {
            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };

            // command.Parameters.AddWithValue("@table", table);
            var dcm = command.Parameters.AddWithValue("@table", table);
            dcm.SqlDbType = SqlDbType.Structured;

            command.Parameters.AddWithValue("@user", user);

            try
            {
                con.Close();
                con.Open();             
                SqlDataAdapter dataAdapter = new SqlDataAdapter(command);                
                dataAdapter.Fill(tb);
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString();
            }

        }
        public static string Insert_And_Get_Table1(ref DataTable tb, string proc, DataTable table, string user, string section_id, string department_id, string allocation_id)
        {
            SqlCommand command = new SqlCommand(proc, con)
            {
                CommandType = CommandType.StoredProcedure
            };

            // command.Parameters.AddWithValue("@table", table);
            var dcm = command.Parameters.AddWithValue("@table", table);
            dcm.SqlDbType = SqlDbType.Structured;

            command.Parameters.AddWithValue("@user", user);
            command.Parameters.AddWithValue("@SECTION_ID", section_id);
            command.Parameters.AddWithValue("@DEPARTMENT_ID", department_id);
            command.Parameters.AddWithValue("@ALLOCATION_ID", allocation_id);

            try
            {
                con.Close();
                con.Open();

                SqlDataAdapter dataAdapter = new SqlDataAdapter(command);
                // this will query your database and return the result to your datatable
                dataAdapter.Fill(tb);
                con.Close();
                return "";
            }
            catch (Exception ex)
            {
                con.Close();
                return ex.Message.ToString();
            }

        }
        public static void Export_Excel_To_Tem(string templateLocalPath, string ouputname, int startRow, int StartCols, DataTable table)
        {

            //Get Template extensions
            string[] name = templateLocalPath.Split('.');
            string exten = name[name.Count() - 1];
            //Open and Copy Template
            var workbook = new XLWorkbook(templateLocalPath);
            var ws = workbook.Worksheet(1);
            //Save Data to Template
            ws.Cell(startRow, StartCols).InsertData(table.AsEnumerable());
            workbook.SaveAs(ouputname);
        }

        public static void ExportExcels(string templateLocalPath, DataTable table, string exportName, int startRow, int startColum, string nameWorksheet)
        {
            //Get Template extensions
            string[] name = templateLocalPath.Split('.');
            string exten = name[name.Count() - 1];
            //Open and Copy Template
            var workbook = new XLWorkbook(templateLocalPath);
            var ws = workbook.Worksheet(nameWorksheet);
            ws.Cell(startRow, startColum).InsertData(table.AsEnumerable());
            ws.ShowGridLines = true;
            ExportDiaglog(workbook, exportName + DateTime.Now.ToString("_yyyyMMdd_HHmmssfff") + ".xlsx");

        }

        public static void exportExcel(string path, string filename, string pathTemp,int StartRow,int StartColunm, DataTable dt)
        {
            try
            {

                using (var fldrDlg = new FolderBrowserDialog())
                {
                    if (fldrDlg.ShowDialog() == DialogResult.OK)
                    {
                        path = fldrDlg.SelectedPath.ToString();

                        Export_Excel_To_Tem(pathTemp, path + filename, StartRow, StartColunm, dt);

                    }
                }
                DialogResult dialogResult = MessageBox.Show("Download successful!\r\n Do you want open file?", "Thông báo", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);

                if (dialogResult == DialogResult.OK)
                {
                    Process.Start(path + filename);
                }
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show(ex.Message);
            }
        }
        public void editQtyDetail(string requisition, string id, double qty, double amount)
        {
            SqlCommand cmd = new SqlCommand("[CMS].[sp_Edit_Qty_Detail]", con);
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.AddWithValue("@requisition_id", requisition);
            cmd.Parameters.AddWithValue("@id", id);
            cmd.Parameters.AddWithValue("@qty", qty);
            cmd.Parameters.AddWithValue("@amount", amount);
            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();
        }

        public static void ExportExcel(string templateLocalPath, DataTable table, string exportName)
        {
            //Get Template extensions
            string[] name = templateLocalPath.Split('.');
            string exten = name[name.Count() - 1];
            //Open and Copy Template
            var workbook = new XLWorkbook(templateLocalPath);
            var ws = workbook.Worksheet(1);
            ws.Cell(6, 1).InsertData(table.AsEnumerable());
            ws.ShowGridLines = true;
            ExportDiaglog(workbook, exportName + DateTime.Now.ToString("_yyyyMMdd_HHmmssfff") + ".xlsx");

        }

        public static void ExportExcel(string templateLocalPath, DataTable table, string exportName,int startRow, int startColumn)
        {
            //Get Template extensions
            string[] name = templateLocalPath.Split('.');
            string exten = name[name.Count() - 1];
            //Open and Copy Template
            var workbook = new XLWorkbook(templateLocalPath);
            var ws = workbook.Worksheet(1);
            ws.Cell(startRow, startColumn).InsertData(table.AsEnumerable());
            ws.ShowGridLines = true;
            ExportDiaglog(workbook, exportName + DateTime.Now.ToString("_yyyyMMdd_HHmmssfff") + ".xlsx");

        }

        protected static void ExportDiaglog(XLWorkbook wb, string filename)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "Excel files|*.xlsx",
                Title = "Save an Excel File"
            };
            saveFileDialog.FileName = filename;
            if (saveFileDialog.ShowDialog() == DialogResult.Cancel) return;
            if (!String.IsNullOrWhiteSpace(saveFileDialog.FileName))
                wb.SaveAs(saveFileDialog.FileName);
            MessageBox.Show("Xuất file excel thành công!");
            wb.Dispose();
        }
    }
}

using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;
namespace ConsoleApp1
{
    class insert_into_db_service
    {
        /// <summary>
        /// this function will read the excel file and insert the data into data base
        /// </summary>
        /// <param name="filePath"></param>
        public void readFileAndInsertDb(string filePath)
        {
            //check if the pass is not null
            if (filePath != null)
            {
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                using (var package = new ExcelPackage(new FileInfo(filePath)))
                {

                    var firstSheet = package.Workbook.Worksheets[0];
                    int column = firstSheet.Dimension.Columns;
                    int row = firstSheet.Dimension.Rows;
                    for (int r = 2; r <= row; r++)
                    {
                        string name = "";
                        string mobile = "";
                        string title = "";
                        string email = "";
                        string address = "";
                        string net_salary = "";
                        string gross_salary = "";
                        string gender = "";


                        //Console.WriteLine(firstSheet.Cells[r, c].Value.ToString()?? string.Empty);
                        name = firstSheet.Cells[r, 1].Value.ToString() ?? string.Empty;
                        mobile = firstSheet.Cells[r, 2].Value.ToString() ?? string.Empty;
                        title = firstSheet.Cells[r, 3].Value.ToString() ?? string.Empty;
                        email = firstSheet.Cells[r, 4].Value.ToString() ?? string.Empty;
                        address = firstSheet.Cells[r, 5].Value.ToString() ?? string.Empty;
                        net_salary = firstSheet.Cells[r, 6].Value.ToString() ?? string.Empty;
                        gross_salary = firstSheet.Cells[r, 7].Value.ToString() ?? string.Empty;
                        gender = firstSheet.Cells[r, 8].Value.ToString() ?? string.Empty;


                        try
                        {
                            Console.WriteLine(name + mobile + title + email + address);
                            string conn = "Data Source=.;Database=task;Integrated Security=True;";
                            //create instanace of database connection to insert the changes
                            using (SqlConnection con = new SqlConnection(conn))
                            {


                                con.Open();
                                var sql = "INSERT INTO [dbo].[employees]([name],[mobile_number],[job_title],[email],[address],[net_salary],[gross_salary],[gender]) VALUES " +
                                    " (@name, @mobile, @title, @email, @address, @net_salary, @gross_salary, @gender)";
                                using (var cmd = new SqlCommand(sql, con))
                                {
                                    cmd.Parameters.AddWithValue("@name", name);
                                    cmd.Parameters.AddWithValue("@mobile", mobile);
                                    cmd.Parameters.AddWithValue("@title", title);
                                    cmd.Parameters.AddWithValue("@email", email);
                                    cmd.Parameters.AddWithValue("@address", address);
                                    cmd.Parameters.AddWithValue("@net_salary", DBNull.Value);
                                    cmd.Parameters.AddWithValue("@gross_salary", DBNull.Value);
                                    cmd.Parameters.AddWithValue("@gender", gender);

                                    cmd.ExecuteNonQuery();
                                }


                                //  Console.WriteLine("Commands executed! Total rows affected are " + insertCommand.ExecuteNonQuery());

                            }
                        }
                        catch (Exception x)
                        {

                            throw;
                        }

                    }
                    Console.WriteLine("Sheet 1 Data");
                    Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A2"].Text}");

                }
                Console.WriteLine("this is were we read a file" + filePath);
            }
        }
        public bool checkIfRowExists()
        {
            return false;

        }
    }

    
}
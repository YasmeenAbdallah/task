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
                    for (int r=2; r<=row;r++)
                    {
                        string name = "";
                        string mobile = "";
                        string title = "";
                        string email = "";
                        string address = "";
                        string net_salary = "";
                        string gross_salary = "";
                        string gender = "";
                        
                        for (int c = 1; c <= column; c++)
                        {
                            //Console.WriteLine(firstSheet.Cells[r, c].Value.ToString()?? string.Empty);
                            if (c == 1) name = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 2) mobile = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 3) title = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 4) email = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 5) address = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 6) net_salary = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 7) gross_salary = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;
                            if (c == 8) gender = firstSheet.Cells[r, c].Value.ToString() ?? string.Empty;


                        }
                        Console.WriteLine(name+ mobile+title+email+address);
                        string conn = "Data Source=.;Database=task;Integrated Security=True;";
                        //create instanace of database connection to insert the changes
                        using (SqlConnection con = new SqlConnection(conn))
                        {
                            

                            con.Open();
                            var sql = "insert into[dbo].[employees] ([id],[name],[mobile_number],[job_title],[email],[address],[net_salary],[gross_salary],[gender]) VALUES(@name, @mobile, @title, @email, @address, @net_salary, @gross_salary, @gender)";
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
                    Console.WriteLine("Sheet 1 Data");
                    Console.WriteLine($"Cell A2 Value   : {firstSheet.Cells["A2"].Text}");
                  
                }
                Console.WriteLine("this is were we read a file" + filePath);
            }
        }
        public void InsertToDb(string filePath, string conn)
        {
          
        }
        }


    }



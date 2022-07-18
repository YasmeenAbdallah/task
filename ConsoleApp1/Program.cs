using System;
using System.Data.SqlClient;
using System.IO;

namespace ConsoleApp1
{
    class Program
    {
       
      static  ConnectionString connObj = new ConnectionString();
        static void Main()
        {
            using var watcher = new FileSystemWatcher(@"D:\yasmeen's files\ConsoleApp1\ConsoleApp1");

            watcher.NotifyFilter = NotifyFilters.Attributes
                                 | NotifyFilters.CreationTime
                                 | NotifyFilters.DirectoryName
                                 | NotifyFilters.FileName
                                 | NotifyFilters.LastAccess
                                 | NotifyFilters.LastWrite
                                 | NotifyFilters.Security
                                 | NotifyFilters.Size;

            watcher.Changed += OnChanged;
            watcher.Created += OnCreated;
            watcher.Deleted += OnDeleted;
            watcher.Renamed += OnRenamed;
          

            watcher.Filter = "*.xlsx";
            watcher.IncludeSubdirectories = true;
            watcher.EnableRaisingEvents = true;

            Console.WriteLine("Press enter to exit.");
            Console.ReadLine();
        }
        /// <summary>
        /// when there's a change happaened on file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnChanged(object sender, FileSystemEventArgs e)
        {
            if (e.ChangeType != WatcherChangeTypes.Changed)
            {
                return;
            }
            Console.WriteLine(" there's a change happened");
           
           
        }

        /// <summary>
        /// run when we open a file or create file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnCreated(object sender, FileSystemEventArgs e)
        {
            insert_into_db_service obj = new insert_into_db_service();

            string conn = "Data Source=.;Database=task;Integrated Security=True;";
            //create instanace of database connection to insert the changes
            using (SqlConnection con = new SqlConnection(conn))
            {
                try
                {
                    con.Open();
                    obj.readFileAndInsertDb(e.FullPath);

                }
                catch (Exception x)
                {
                    Console.WriteLine($"Failed to insert: {e.Name} into the database");
                }
            }

            string value = $"Created: {e.FullPath}";
             Console.WriteLine(value);

                //insert query for database

            

        }

        /// <summary>
        /// when close or delete file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnDeleted(object sender, FileSystemEventArgs e) =>
            Console.WriteLine($"Deleted: {e.FullPath}");
        /// <summary>
        /// when renamed the file
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private static void OnRenamed(object sender, RenamedEventArgs e)
        {
            Console.WriteLine($"Renamed:");
            Console.WriteLine($"    Old: {e.OldFullPath}");
            Console.WriteLine($"    New: {e.FullPath}");
        }

       
    }
}
   


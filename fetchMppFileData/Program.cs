using System;
using System.Data;


namespace fetchMppFileData
{
    class Program
    {
        static void Main(string[] args)
        {
            //MS Project File Path
            string file_name = @"C:\Users\LJI006\workspace\dotnet\MSProjectFiles\b4ubuild_sample_07.mpp";
            //DataTable table = new DataTable();
            //MPPFileProcess fileProcess = new MPPFileProcess(file_name, table);
            //Console.WriteLine(fileProcess.Load_MS_Project());
            //table = fileProcess.TasksTable;

            MPPFileProcess fileProcess = new MPPFileProcess(file_name);
            //fileProcess.Export_DB_Excel(table, "myExcel");
            Console.WriteLine("Try to edit MS Project...");
            //Edit Task Name, Start date, Finis date and which row of the task is.
            Console.WriteLine(fileProcess.Edit_MS_Project("test", DateTime.Now, DateTime.Now, 2));
        }
    }
}
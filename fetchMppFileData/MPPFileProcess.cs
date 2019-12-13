using Microsoft.Office.Interop.MSProject;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Diagnostics;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;




namespace fetchMppFileData
{
    class MPPFileProcess
    {
        string newLine = System.Environment.NewLine;
        object readOnly = true;
        object readAndWrite = false;
        object missing = Type.Missing;
        Microsoft.Office.Interop.MSProject.PjPoolOpen pool = Microsoft.Office.Interop.MSProject.PjPoolOpen.pjPoolReadOnly;
        Microsoft.Office.Interop.MSProject.ApplicationClass appclass = new Microsoft.Office.Interop.MSProject.ApplicationClass();

        private string my_fileName;
        public string FileName
        {
            get { return my_fileName; }
            set { my_fileName = value; }
        }

        private DataTable my_table;

        public DataTable TasksTable
        {
            get { return my_table; }
            set { my_table = value; }
        }


        public MPPFileProcess(string fileName, DataTable table)
        {
            this.my_fileName = fileName;
            this.my_table = table;
        }

        public MPPFileProcess(string fileName)
        {
            this.my_fileName = fileName;
        }

        /*.mpp file to Datatable
         */
        public string Load_MS_Project()
        {
            ApplicationClass app = null;
            //return value
            string result = "";
            ArrayList tasks = new ArrayList();

            try
            {
                //excute the Microsoft Project Application
                app = new ApplicationClass();

                //Do not display Microsoft Project
                app.Visible = false;

                //Open the project file
                if (app.FileOpen(my_fileName, readOnly, missing,
                    missing, missing, missing, missing,
                    missing, missing, missing, missing,
                    pool, missing, missing, missing, missing))
                {
                    Tasks taskTable = new Tasks(my_table);
                    my_table = taskTable.createTable();
                    //Get active project
                    Project proj = app.ActiveProject;
                    //Go through the tasks in the project
                    foreach (Task task in proj.Tasks)
                    {
                        string date_start = task.Start.ToString();
                        string date_finish = task.Finish.ToString();
                        bool isSummary = false;
                        //whether the task is a summary.
                        if (task.OutlineChildren.Count != 0)
                        {
                            isSummary = true;
                        }
                        else
                        {
                            isSummary = false;
                        }

                        //whether the start date is 'NA'
                        if (date_start == "NA")
                        {
                            date_start = "0000-00-00";
                        }

                        //whether the finish date is 'NA'
                        if (date_finish == "NA")
                        {
                            date_finish = "0000-00-00";
                        }

                        Tasks my_task = new Tasks(task.ID, task.Name,
                            Int32.Parse(task.Duration.ToString()) / 480,
                            DateTime.Parse(date_start),
                            DateTime.Parse(date_finish),
                            task.ResourceNames, isSummary);
                        //add Microsoft Project Task to arraylist
                        tasks.Add(my_task);

                        //add row to datatable
                        my_table = my_task.addRow(my_table);
                    }
                    result = "Read data from MS Project file" + my_fileName + "successfully!";
                }

                else
                {
                    result = "The MS Project file " + my_fileName + " Could not be opened";
                }
            }
            catch (System.Exception ex)
            {
                result = "Could not process the MS Project file " +
                    my_fileName + "." + newLine +
                    ex.Message + newLine +
                    ex.StackTrace;
            }

            //Close the application if was opened.
            if (app != null)
            {
                app.Quit(PjSaveType.pjDoNotSave);
            }
            return result;
        }

        /*Edit MS Project
         */
        public string Edit_MS_Project(string task_name, DateTime start_date, DateTime finish_date, int task_ID)
        {
            ApplicationClass app = null;
            Project project = null;
            //return value
            string result = "";

            try
            {
                //excute the Microsoft Project Application
                app = new ApplicationClass();
                if (app.FileOpen(my_fileName, readAndWrite, missing,
                    missing, missing, missing, missing,
                    missing, missing, missing, missing,
                    pool, missing, missing, missing, missing))
                {
                    project = app.ActiveProject;
                    //Do not display Microsoft Project
                    app.Visible = false;

                    //Go through the tasks in the project
                    foreach (Task task in project.Tasks)
                    {
                        if(task.ID == task_ID)
                        {
                            if (task.Name != task_name)
                            {
                                task.Name = task_name;
                            }
                           if(task.Start.ToString() != start_date.ToString())
                            {
                                task.Start = start_date.ToString();
                            }
                            if (task.Finish.ToString() != finish_date.ToString())
                            {
                                task.Finish = finish_date.ToString();
                            }
                            break;
                        }
                    }

                    project.SaveAs(my_fileName);
                    //app.FileClose(PjSaveType.pjSave, missing);
                    result = "Edit data to MS Project file" + my_fileName + "successfully!";
                }
                else
                {
                    result = "The MS Project file " + my_fileName + " Could not be opened";
                }
            }
            catch (System.Exception ex)
            {
                result = "Could not process the MS Project file " +
                    my_fileName + "." + newLine +
                    ex.Message + newLine +
                    ex.StackTrace;
            }

            //Close the application if was opened.
            if (app != null)
            {
                app.Quit(PjSaveType.pjDoNotSave);
            }
            return result;
        }

        /*Datatable to Excel
         */
        public void Export_DB_Excel(DataTable table, string excelFilename)
        {
            // Here is main process
            Excel.Application objexcelapp = new Excel.Application();
            objexcelapp.Application.Workbooks.Add(Type.Missing);
            objexcelapp.Columns.AutoFit();
            for (int i = 1; i < table.Columns.Count + 1; i++)
            {
                //Do not write "is Summary" to Colunm.
                if (table.Columns[i - 1].ColumnName != "isSummary")
                {
                    Excel.Range xlRange = (Excel.Range)objexcelapp.Cells[1, i];
                    xlRange.Font.Bold = -1;
                    xlRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                    xlRange.Borders.Weight = 1d;
                    xlRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
                    objexcelapp.Cells[1, i] = table.Columns[i - 1].ColumnName;
                }
            }
            /*For storing Each row and column value to excel sheet*/
            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++)
                {
                    //Do not write "is Summary" to Rows.
                    if (table.Rows[i][j] != null && table.Columns[j].ColumnName != "isSummary")
                    {
                        Excel.Range xlRange = (Excel.Range)objexcelapp.Cells[i + 2, j + 1];
                        xlRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;
                        xlRange.Borders.Weight = 1d;
                        if ((bool)table.Rows[i]["isSummary"] == true)
                        {
                            xlRange.Font.Bold = -1;
                        }
                        objexcelapp.Cells[i + 2, j + 1] = table.Rows[i][j].ToString();
                    }
                }
            }
            objexcelapp.Columns.AutoFit(); // Auto fix the columns size
            if (Directory.Exists(@"C:\Users\LJI006\workspace\dotnet\MSProjectFiles\")) // Folder dic
            {
                objexcelapp.ActiveWorkbook.SaveCopyAs(@"C:\Users\LJI006\workspace\dotnet\MSProjectFiles\" + excelFilename + ".xlsx");
            }
            else
            {
                Directory.CreateDirectory(@"C:\Users\LJI006\workspace\dotnet\MSProjectFiles\");
                objexcelapp.ActiveWorkbook.SaveCopyAs(@"C:\Users\LJI006\workspace\dotnet\MSProjectFiles\" + excelFilename + ".xlsx");
            }
            objexcelapp.ActiveWorkbook.Saved = true;
            foreach (Process proc in Process.GetProcessesByName("EXCEL"))
            {
                proc.Kill();
            }

        }
    }
}

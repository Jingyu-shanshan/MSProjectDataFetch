using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace fetchMppFileData
{
    class Tasks
    {
        public int task_id { get; set; }
        public string project_name { get; set; }
        public int duration_in_days { get; set; }
        public DateTime start { get; set; }
        public DateTime finish { get; set; }
        public double percent_complete { get; set; }
        public DateTime actual_finish { get; set; }
        public string resource_name { get; set; }
        public DataTable table { get; set; }

        public bool isSummary { get; set; }

        private DataSet dataSet;

        public Tasks(int task_id, string project_name, int duration_in_days, DateTime start, DateTime finish, string resource_name, bool isSummary)
        {
            this.task_id = task_id;
            this.project_name = project_name;
            this.duration_in_days = duration_in_days;
            this.start = start;
            this.finish = finish;
            this.resource_name = resource_name;
            this.isSummary = isSummary;
        }

        public Tasks(DataTable table)
        {
            this.table = table;
        }

        public DataTable createTable()
        {
            table = new DataTable("TasksTable");
            DataColumn column;

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "ID";
            column.ReadOnly = true;
            column.Unique = true;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Task Name";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Int32");
            column.ColumnName = "Duration";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.DateTime");
            column.ColumnName = "Start";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.DateTime");
            column.ColumnName = "Finish";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.String");
            column.ColumnName = "Resource Names";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            column = new DataColumn();
            column.DataType = System.Type.GetType("System.Boolean");
            column.ColumnName = "isSummary";
            column.ReadOnly = false;
            column.Unique = false;
            table.Columns.Add(column);

            dataSet = new DataSet();
            dataSet.Tables.Add(table);

            return table;
        }

        public DataTable addRow(DataTable tableAddRows)
        {
            DataRow row;

            row = tableAddRows.NewRow();
            row["ID"] = task_id;
            row["Task Name"] = project_name;
            row["Duration"] = duration_in_days;
            row["Start"] = start;
            row["Finish"] = finish;
            row["Resource Names"] = resource_name;
            row["isSummary"] = isSummary;
            tableAddRows.Rows.Add(row);

            return tableAddRows;
        }

        public override string ToString()
        {
            return "Project Name: " + project_name + " -- "
                 + "Duration: " + duration_in_days + " days -- "
                 + "Start Date: " + start + " -- "
                 + "Finish Date: " + finish + " -- "
                 + "Percent Complete: " + percent_complete + " -- "
                 + "Resource Name: " + resource_name + " -- "
                 + "Actual Finish Date: " + actual_finish + " --\n\n";
        }
    }

}


using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

using System;
using System.Data.OleDb;
using System.Windows;

namespace ConnectionDB
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    { 
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // 2019-09-29
            // 23-00
            // 23-03
            //23-05
            // 23-12
            //string dbFilePath = "D:\\_DB_project.mdb";
            string dbFilePath = "D:\\_DB_project_poligon.accdb";


            string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + dbFilePath;


            OleDbConnection dbConnection = new OleDbConnection();
            dbConnection.ConnectionString = connectionString;


            OleDbCommand command = new OleDbCommand();
            command.Connection = dbConnection;

            try
            {
                dbConnection.Open();

                MessageBox.Show("Подключились к БД");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка получения данных в БД: " +
                                 Environment.NewLine +
                                 ex.Message,
                                 "Ошибка");
            }
            finally
            {
                dbConnection.Close();
                dbConnection.Dispose();
            }


            return Result.Succeeded;
        }



    }




}

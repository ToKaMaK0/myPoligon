
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Autodesk.Revit.UI.Selection;
using System;
using System.Data.OleDb;
using System.Windows;

namespace POLIGON
{
    [Transaction(TransactionMode.Manual)]
    public class ACommand : IExternalCommand
    {
        public static string RP_name1;
        public static UIApplication uiApp;
        public static UIDocument uiDoc;
        public static Autodesk.Revit.ApplicationServices.Application app;
        public static Document doc;
        public static Selection sel;

        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {

            uiApp = commandData.Application;
            uiDoc = uiApp.ActiveUIDocument;
            app = uiApp.Application;
            doc = uiDoc.Document;
            sel = uiDoc.Selection;



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

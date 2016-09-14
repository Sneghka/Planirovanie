using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Data.OleDb;

namespace Planirovanie
{
    public class WorkWithExcelFile
    {

        public static string sMessage { get; private set; }

        ///
        /// Преобразует Excel-файл в DataTable
        ///
        ///Таблица для загрузки данных
        ///Полный путь к Excel-файлу
        ///SQL-запрос. Используйте $SHEETS$ для выбоки по всем листам
        public static void ExcelFileToDataTable(out DataTable dtData, string sFile, string sRequest)
        {
            DataSet dsData = new DataSet();

            string sConnStr =
                String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"{1};HDR=YES\";",
                    sFile, sFile.EndsWith(".xlsx") ? "Excel 12.0 Xml" : "Excel 8.0");

            using (OleDbConnection odcConnection = new OleDbConnection(sConnStr))
            {
                odcConnection.Open();
                if (sRequest.IndexOf("$SHEETS$") != -1)
                {
                    using (
                        DataTable dtMetadata = odcConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                            new object[4] { null, null, null, "TABLE" }))
                    {
                        for (int i = 0; i < dtMetadata.Rows.Count; i++)
                            if (dtMetadata.Rows[i]["TABLE_NAME"].ToString().IndexOf("$") == -1)
                                dtMetadata.Rows.Remove(dtMetadata.Rows[i]);

                        foreach (DataRow drRow in dtMetadata.Rows)
                        {
                            string sLocalRequest = sRequest.Replace("$SHEETS$",
                                String.Format("[{0}]", drRow["TABLE_NAME"]));
                            OleDbCommand odcCommand = new OleDbCommand(sLocalRequest, odcConnection);
                            using (OleDbDataAdapter oddaAdapter = new OleDbDataAdapter(((OleDbCommand)odcCommand)))
                                oddaAdapter.Fill(dsData);
                        }
                    }
                }
                else
                {
                    OleDbCommand odcCommand = new OleDbCommand(sRequest, odcConnection);
                    using (OleDbDataAdapter oddaAdapter = new OleDbDataAdapter(odcCommand))
                        oddaAdapter.Fill(dsData);
                }
                odcConnection.Close();
            }

            dtData = dsData.Tables[0];
        }
    }
}

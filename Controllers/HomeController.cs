using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Web.Mvc;
using System.Linq;
using System.Web;
using Newtonsoft.Json;
using System.IO;
using System.Text;
using System.IO.Compression;

namespace DynamicExcelGenerate.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public JsonResult GetAttendanceData(string startDate, string endDate)
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["xlContext"].ConnectionString;

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("GetAttendanceData", connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add(new SqlParameter("@startDate", SqlDbType.Date) { Value = DateTime.Parse(startDate) });
                    cmd.Parameters.Add(new SqlParameter("@endDate", SqlDbType.Date) { Value = DateTime.Parse(endDate) });
                    cmd.CommandTimeout = 300;

                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    List<Dictionary<string, object>> dataRows = new List<Dictionary<string, object>>();

                    foreach (DataRow row in dt.Rows)
                    {
                        Dictionary<string, object> rowData = new Dictionary<string, object>();

                        foreach (DataColumn col in dt.Columns)
                        {
                            rowData.Add(col.ColumnName, row[col]);
                        }

                        dataRows.Add(rowData);
                    }
                    return Json(new { DataRows = dataRows }, JsonRequestBehavior.AllowGet);
                }
            }
        }

        [HttpGet]
        public JsonResult GetAttendanceStatusReport(string startDate, string endDate)
        {
            string connectionString = System.Configuration.ConfigurationManager.ConnectionStrings["xlContext"].ConnectionString;
            List<Dictionary<string, object>> dataRows = new List<Dictionary<string, object>>();

            using (SqlConnection connection = new SqlConnection(connectionString))
            {
                using (SqlCommand cmd = new SqlCommand("GetAttendanceReport", connection))
                {
                    cmd.CommandType = CommandType.StoredProcedure;

                    cmd.Parameters.Add(new SqlParameter("@startDate", SqlDbType.Date) { Value = DateTime.Parse(startDate) });
                    cmd.Parameters.Add(new SqlParameter("@endDate", SqlDbType.Date) { Value = DateTime.Parse(endDate) });
                    cmd.CommandTimeout = 900;

                    SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    foreach (DataRow row in dt.Rows)
                    {
                        Dictionary<string, object> rowData = new Dictionary<string, object>();
                        foreach (DataColumn col in dt.Columns)
                        {
                            if (row[col] != DBNull.Value)
                            {
                                var columnName = col.ColumnName;
                                if (columnName.Contains("Time"))
                                {
                                    if (DateTime.TryParse(row[col].ToString(), out DateTime dateTimeValue))
                                    {
                                        var formattedTime = dateTimeValue.ToString("h:mm tt");
                                        rowData.Add(columnName, formattedTime);
                                    }
                                    else
                                    {
                                        rowData.Add(columnName, row[col].ToString());
                                    }
                                }
                                else if (columnName.Contains("Early") || columnName.Contains("Late"))
                                {
                                    if (TimeSpan.TryParse(row[col].ToString(), out TimeSpan timeSpanValue))
                                    {
                                        var formattedTime = timeSpanValue.ToString("hh\\:mm");
                                        rowData.Add(columnName, formattedTime);
                                    }
                                    else
                                    {
                                        rowData.Add(columnName, row[col].ToString());
                                    }
                                }
                                else
                                {
                                    rowData.Add(col.ColumnName, row[col]);
                                }
                            }
                            else
                            {
                                rowData.Add(col.ColumnName, row[col]);
                            }
                        }
                        dataRows.Add(rowData);
                    }
                }
            }
            var jsonResult = Json(new { DataRows = dataRows }, JsonRequestBehavior.AllowGet);
            jsonResult.MaxJsonLength = int.MaxValue;
            return jsonResult;
        }
    }
}
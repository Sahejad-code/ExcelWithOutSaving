/*using ExcelWithOutSaving.Models;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace ExcelWithOutSaving.Controllers
{


    public class HomeController : Controller
    {

        public ActionResult Index()
        {
            ExcelUploadRequestModel model = new ExcelUploadRequestModel();
            return View(model); 
        }


        [HttpPost]
        public ActionResult Index(ExcelUploadRequestModel model)
        {
            string UploadStatusLabel = string.Empty;
            try
            {
                if (model.File != null && Path.GetExtension(model.File.FileName) == ".xlsx")
                {
                    ExcelPackage.LicenseContext = LicenseContext.Commercial;

                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage excel = new ExcelPackage(model.File.InputStream))
                    {
                        var tbl = new DataTable();
                        var ws = excel.Workbook.Worksheets.First();
                        var hasHeader = true;


                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                        }
                        int startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                            DataRow row = tbl.NewRow();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                            tbl.Rows.Add(row);
                        }
                        if (tbl.Columns.Count == model.MaxAllowedColumns)
                        {
                            var msg = String.Format("DataTable successfully created from excel-file",
                                                    tbl.Columns.Count, tbl.Rows.Count);

                            model.DataTable = tbl;
                            model.UploadStatusLabel = msg;
                            
                        }
                        else
                        {
                            model.UploadStatusLabel = $"Please Upload in Proper Format. Maximum allowed columns is {model.MaxAllowedColumns}.";
                        }
                        return View(model);
                    }

                }
            
                else
                {
                    model.UploadStatusLabel = "You did not specify a file to upload.";
                }
            }
            catch (Exception)
            {
                model.UploadStatusLabel = "";
            }

            return View(model);
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}*/

using ExcelWithOutSaving.Models;
using OfficeOpenXml;
using System;
using System.Data;
using System.IO;
using System.Linq;
using System.Web.Mvc;

namespace ExcelWithOutSaving.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            ExcelUploadRequestModel model = new ExcelUploadRequestModel();
            return View(model);
        }

        [HttpPost]
        public ActionResult Index(ExcelUploadRequestModel model)
        {
            string UploadStatusLabel = string.Empty;
            try
            {
                if (model.File != null && Path.GetExtension(model.File.FileName) == ".xlsx")
                {
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
                    using (ExcelPackage excel = new ExcelPackage(model.File.InputStream))
                    {
                        var tbl = new DataTable();
                        var ws = excel.Workbook.Worksheets.First();
                        var hasHeader = true;

                        foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                        {
                            tbl.Columns.Add(hasHeader ? firstRowCell.Text : $"Column {firstRowCell.Start.Column}");
                        }

                        int startRow = hasHeader ? 2 : 1;
                        for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                        {
                            var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                            DataRow row = tbl.NewRow();
                            foreach (var cell in wsRow)
                            {
                                row[cell.Start.Column - 1] = cell.Text;
                            }
                            tbl.Rows.Add(row);
                        }

                        if (tbl.Columns.Count == model.MaxAllowedColumns)
                        {
                            var msg = String.Format("DataTable successfully created from excel-file with {0} columns and {1} rows.", tbl.Columns.Count, tbl.Rows.Count);

                            model.DataTable = tbl;
                            model.UploadStatusLabel = msg;
                        }
                        else
                        {
                            model.UploadStatusLabel = $"Please Upload in Proper Format. Maximum allowed columns is {model.MaxAllowedColumns}.";
                        }

                        return View(model);
                    }
                }
                else
                {
                    model.UploadStatusLabel = "You did not specify a file to upload or the file is not in the .xlsx format.";
                }
            }
            catch (Exception ex)
            {
                model.UploadStatusLabel = $"An error occurred: {ex.Message}";
            }

            return View(model);
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            return View();
        }
    }
}

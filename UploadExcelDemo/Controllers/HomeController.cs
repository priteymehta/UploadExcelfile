using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace UploadExcelDemo.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
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

        public JsonResult UploadData()
        {
            int result = 0;
            try
            {
                if (Request.Files["FileUpload1"].ContentLength > 0)
                {
                    string extension = System.IO.Path.GetExtension(Request.Files["FileUpload1"].FileName).ToLower();
                    string query = null;
                    string connString = "";
                    //string[] validFileTypes = { ".xls", ".xlsx", ".csv" };
                    string[] validFileTypes = { ".csv" };
                    string path1 = string.Format("{0}/{1}", Server.MapPath("~/Content/Uploads"), Request.Files["FileUpload1"].FileName);
                    if (!Directory.Exists(path1))
                    {
                        Directory.CreateDirectory(Server.MapPath("~/Content/Uploads"));
                    }
                    if (validFileTypes.Contains(extension))
                    {
                        if (System.IO.File.Exists(path1))
                        { System.IO.File.Delete(path1); }
                        Request.Files["FileUpload1"].SaveAs(path1);
                        if (extension == ".csv")
                        {
                            DataTable dt = Utility.ConvertCSVtoDataTable(path1);
                            ViewBag.Data = dt;
                        }
                        //Connection String to Excel Workbook  
                        //else if (extension.Trim() == ".xls")
                        //{
                        //    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                        //    DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        //    ViewBag.Data = dt;
                        //}
                        //else if (extension.Trim() == ".xlsx")
                        //{
                        //    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                        //    DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                        //    ViewBag.Data = dt;
                        //}

                    }
                    else
                    {
                        ViewBag.Error = "Please Upload Files in .xls, .xlsx or .csv format";

                    }

                }

                result = 1;
                return Json(result, JsonRequestBehavior.AllowGet);
            }
            catch (Exception)
            {
                return Json(result, JsonRequestBehavior.AllowGet);
            }
        }
    }
}
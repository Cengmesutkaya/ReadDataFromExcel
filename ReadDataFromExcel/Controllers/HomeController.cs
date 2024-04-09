using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ReadDataFromExcel.Controllers
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
        public ActionResult CreateExcel()
        {
            return View();
        }

        [HttpPost]
        public ActionResult UploadExcel(HttpPostedFileBase uploadfile)
        {
            if (uploadfile != null)
            {
                Stream stream = uploadfile.InputStream;
                IExcelDataReader reader = null;
                if (uploadfile.FileName.EndsWith(".xls"))
                {
                    reader = ExcelReaderFactory.CreateBinaryReader(stream);
                }
                else if (uploadfile.FileName.EndsWith(".xlsx"))
                {
                    reader = ExcelReaderFactory.CreateOpenXmlReader(stream);
                }
                DataSet result = reader.AsDataSet();
                reader.Close();
                foreach (DataTable dt in result.Tables)
                {
                    List<ProductDTO> productList = new List<ProductDTO>();
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row.ItemArray[0].ToString() != "Code")
                        {
                            ProductDTO productDTO = new ProductDTO();
                            productDTO.Code = Convert.ToInt32(row.ItemArray[0]);
                            productDTO.Price = Convert.ToDecimal(row.ItemArray[1]);
                            productList.Add(productDTO);
                        }
                    }
                }

            }
            return View();

        }

      public class ProductDTO
        {
            public int Code { get; set; }
            public decimal Price { get; set; }
        }

    }
}
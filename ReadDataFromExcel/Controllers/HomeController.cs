using ClosedXML.Excel;
using ExcelDataReader;
using Microsoft.AspNetCore.Http;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;
using System.Web.UI.WebControls;

namespace ReadDataFromExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public ActionResult ExcelDowload()
        {

            List<ProductDTO> list = new List<ProductDTO>();
            ProductDTO productDTO = new ProductDTO();
            productDTO.Code = 101;
            productDTO.Price = 5;
            list.Add(productDTO);

            ProductDTO productDTO2 = new ProductDTO();
            productDTO2.Code = 105;
            productDTO2.Price = 50;
            list.Add(productDTO2);

            var file = ExcelHelper.CreateFile(list);
            return File(file, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "products.xlsx");
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
                IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);    

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

        public static class ExcelHelper
        {
            public static byte[] CreateFile<T>(List<T> source)
            {
                var workbook = new XSSFWorkbook();
                var sheet = workbook.CreateSheet("Sheet1");
                var rowHeader = sheet.CreateRow(0);

                var properties = typeof(T).GetProperties();

                //header
                var font = workbook.CreateFont();
                font.IsBold = true;
                var style = workbook.CreateCellStyle();
                style.SetFont(font);

                var colIndex = 0;
                foreach (var property in properties)
                {
                    var cell = rowHeader.CreateCell(colIndex);
                    cell.SetCellValue(property.Name);
                    cell.CellStyle = style;
                    colIndex++;
                }
                //end header


                //content
                var rowNum = 1;
                foreach (var item in source)
                {
                    var rowContent = sheet.CreateRow(rowNum);

                    var colContentIndex = 0;
                    foreach (var property in properties)
                    {
                        var cellContent = rowContent.CreateCell(colContentIndex);
                        var value = property.GetValue(item, null);

                        if (value == null)
                        {
                            cellContent.SetCellValue("");
                        }
                        else if (property.PropertyType == typeof(string))
                        {
                            cellContent.SetCellValue(value.ToString());
                        }
                        else if (property.PropertyType == typeof(int) || property.PropertyType == typeof(int?))
                        {
                            cellContent.SetCellValue(Convert.ToInt32(value));
                        }
                        else if (property.PropertyType == typeof(decimal) || property.PropertyType == typeof(decimal?))
                        {
                            cellContent.SetCellValue(Convert.ToDouble(value));
                        }
                        else if (property.PropertyType == typeof(DateTime) || property.PropertyType == typeof(DateTime?))
                        {
                            var dateValue = (DateTime)value;
                            cellContent.SetCellValue(dateValue.ToString("yyyy-MM-dd"));
                        }
                        else cellContent.SetCellValue(value.ToString());

                        colContentIndex++;
                    }

                    rowNum++;
                }

                //end content


                var stream = new MemoryStream();
                workbook.Write(stream);
                var content = stream.ToArray();

                return content;
            }

        }
    }
}
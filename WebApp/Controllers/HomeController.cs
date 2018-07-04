using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebApp.Models;
using System.Data.Entity;
using Excel = Microsoft.Office.Interop.Excel;
using System.Text;
using System.Runtime.InteropServices;
using System.Net.Mail;
using System.Net;
using System.Configuration;
using System.Threading;
using System.IO;

namespace WebApp.Controllers
{
    public class HomeController : Controller
    {
        NorthwindEntities db = new NorthwindEntities();
        public ActionResult Index()
        {
            var data = db.OrderDetail.Include(d => d.Product);
            return View(data.ToList());
        }
        [HttpPost]
        public ActionResult SendData(string fromDate, string toDate, string email)
        {
            try
            {
                List<OrderDetail> report = ReportOrderDetail(fromDate, toDate);

                Excel.Application application = new Excel.Application();
                Excel.Workbook workbook = 
                    application.Workbooks.Add(System.Reflection.Missing.Value);
                Excel.Worksheet worksheet = workbook.ActiveSheet;

                worksheet.Cells[1, 1] = "Номер заказа";
                worksheet.Cells[1, 2] = "Дата заказа";
                worksheet.Cells[1, 3] = "Название товара";
                worksheet.Cells[1, 4] = "Кол-во реализованных ед.";
                worksheet.Cells[1, 5] = "Цена реализации за ед.";

                int row = 2;
                foreach (var item in report)
                {
                    worksheet.Cells[row, 1] = item.Order.ID;
                    worksheet.Cells[row, 2] = item.Order.OrderDate;
                    worksheet.Cells[row, 3] = item.Product.Name;
                    worksheet.Cells[row, 4] = item.Quantity;
                    worksheet.Cells[row, 5] = item.UnitPrice;
                    row++;
                }

                worksheet.Cells[row, 1] = "Итого:";
                worksheet.Cells[row, 4].Formula = 
                    $"=Sum({worksheet.Cells[2,4].Address}:{worksheet.Cells[row-1,4].Address})";
                worksheet.Cells[row, 5].Formula =
                    $"=Sum({worksheet.Cells[2, 5].Address}:{worksheet.Cells[row - 1, 5].Address})";

                string filename = Server.MapPath($"~/Reports/Report{DateTime.Now.ToString("Hms")}.xls");
                workbook.SaveAs(filename);
                workbook.Close();

                Marshal.ReleaseComObject(workbook);
                application.Quit();
                Marshal.FinalReleaseComObject(application);

                SendMail(filename, email);

                if (Convert.ToDateTime(fromDate).Date > Convert.ToDateTime(toDate).Date)
                    ViewBag.Result = "Отчет был создан с ошибками.. Возможно, вы указали неверный временной период.";
                else
                    ViewBag.Result = "Отчет упешно создан!";
            }
            catch (Exception ex)
            {
                ViewBag.Result = ex.Message;
            }

            return View();
        }

        public List<OrderDetail> ReportOrderDetail(string fromDate, string toDate)
        {
            List<OrderDetail> list = new List<OrderDetail>();

            if (!string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate))
            {
                DateTime newFromDate = Convert.ToDateTime(fromDate).Date;
                DateTime newToDate = Convert.ToDateTime(toDate).Date;
                list = (from obj in db.OrderDetail
                            where obj.Order.OrderDate.Value >= newFromDate && obj.Order.OrderDate.Value <= newToDate
                                select obj).ToList();
            }

            if (!string.IsNullOrEmpty(fromDate) && string.IsNullOrEmpty(toDate))
            {
                DateTime newFromDate = Convert.ToDateTime(fromDate).Date;
                list = (from obj in db.OrderDetail
                            where obj.Order.OrderDate.Value >= newFromDate
                        select obj).ToList();
            }

            if(string.IsNullOrEmpty(fromDate) && !string.IsNullOrEmpty(toDate))
            {
                DateTime newToDate = Convert.ToDateTime(toDate).Date;
                list = (from obj in db.OrderDetail
                            where obj.Order.OrderDate.Value <= newToDate
                        select obj).ToList();
            }

            if (string.IsNullOrEmpty(fromDate) && string.IsNullOrEmpty(toDate))
            {
                list = (from obj in db.OrderDetail select obj).ToList();
            }

            return list;
        }

        public static void SendMail(string report, string email)
        {
            string senderEmail = ConfigurationManager.AppSettings["Email"].ToString();
            string senderPassword = ConfigurationManager.AppSettings["Password"].ToString();

            SmtpClient client = new SmtpClient("smtp.gmail.com", 587);
            client.EnableSsl = true;
            client.Timeout = 50000;
            client.DeliveryMethod = SmtpDeliveryMethod.Network;
            client.UseDefaultCredentials = false;
            client.Credentials = new NetworkCredential(senderEmail, senderPassword);

            MailMessage mailMessage = new MailMessage(senderEmail, email, "Отчет", "");
            mailMessage.IsBodyHtml = true;
            mailMessage.BodyEncoding = Encoding.UTF8;
            mailMessage.Attachments.Add(new Attachment(report));

            client.Send(mailMessage);
        }
        protected override void Dispose(bool disposing)
        {
            db.Dispose();
            base.Dispose(disposing);
        }
    }
}

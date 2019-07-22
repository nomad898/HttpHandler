using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Web;
using System.Xml.Linq;

namespace Http.Application
{
    public class ReportHandler : IHttpHandler
    {
        public bool IsReusable => true;

        private const string CUSTOMER = "customer";
        private const string DATE_FROM = "dateFrom";
        private const string DATE_TO = "dateTo";
        private const string TAKE = "take";
        private const string SKIP = "skip";
        private const string TYPE = "type";

        public void ProcessRequest(HttpContext context)
        {
            NorthwindEntities db = new NorthwindEntities();
            var queryString = context.Request.QueryString;
            ValidateQueryStringDate(queryString.AllKeys);

            var orders = db.Orders;
            string customer = context.Request[CUSTOMER];
            string dateFrom = context.Request[DATE_FROM];
            string dateTo = context.Request[DATE_TO]; 
            string take = context.Request[TAKE];
            string skip = context.Request[SKIP];
            string type = context.Request[TYPE];

            var ordersList = db.Orders.ToList();
            if (!string.IsNullOrEmpty(customer))
            {
                ordersList = ordersList.Where(x => x.CustomerID == customer).ToList();
            }
            if (!string.IsNullOrEmpty(dateFrom))
            {
                ordersList = ordersList.Where(x => x.OrderDate >= Convert.ToDateTime(dateFrom)).ToList();
            }
            if (!string.IsNullOrEmpty(dateTo))
            {
                ordersList = ordersList.Where(x => x.OrderDate <= Convert.ToDateTime(dateTo)).ToList();
            }
            if (!string.IsNullOrEmpty(take))
            {
                ordersList = ordersList.Take(Convert.ToInt32(take)).ToList();
            }
            if (!string.IsNullOrEmpty(skip))
            {
                ordersList = ordersList.Skip(Convert.ToInt32(skip)).ToList();
            }


            if (type == "xml")
            {
                context.Response.ContentType = "text/xml";
                XDocument xdoc = new XDocument();
                XElement xOrders = new XElement("Orders");
                xdoc.Add(xOrders);
                foreach (var item in ordersList)
                {
                    XElement xOrder = new XElement("Order");
                    XElement xCustomerId = new XElement("CustomerId", item.CustomerID);
                    XElement xOrderDate = new XElement("OrderDate", $"{item.OrderDate.Value.Day}-{item.OrderDate.Value.Month}-{item.OrderDate.Value.Year}");
                    xOrder.Add(xCustomerId);
                    xOrder.Add(xOrderDate);
                    xOrders.Add(xOrder);
                }
                using (MemoryStream stream = new MemoryStream())
                {
                    xdoc.Save(stream);
                    context.Response.OutputStream.Write(stream.ToArray(), 0, (int)stream.Length);
                }
            }
            else if (type == "excel")
            {
                context.Response.ContentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

                using (var workbook = new XLWorkbook())
                {
                    var worksheet = workbook.Worksheets.Add("Orders");
                    for (int i = 0; i < ordersList.Count; i++)
                    {
                        worksheet.Cell($"A{i + 1}").Value = ordersList[i].CustomerID;
                        if (ordersList[i].OrderDate.HasValue)
                        {
                            worksheet.Cell($"B{i + 1}").Value = $"{ordersList[i].OrderDate.Value.Day} - {ordersList[i].OrderDate.Value.Month} - {ordersList[i].OrderDate.Value.Year}";
                        }
                    }
                    worksheet.Columns().AdjustToContents();

                    using (MemoryStream stream = new MemoryStream())
                    {
                        workbook.SaveAs(stream);
                        context.Response.OutputStream.Write(stream.ToArray(), 0, (int)stream.Length);
                    }
                }
            }
            else
            {
                context.Response.Write($"Orders - Count: {ordersList.Count} <br />");
                foreach (var item in ordersList)
                {
                    if (item.OrderDate.HasValue)
                    {
                        context.Response.Write($"{item.CustomerID} - {item.OrderDate.Value.Day}-{item.OrderDate.Value.Month}-{item.OrderDate.Value.Year} <br />");
                    }
                }
            }
        }

        private bool ValidateQueryStringDate(string[] queryString)
        {
            var isDateExist = false;
            foreach (var item in queryString)
            {
                if (item == DATE_FROM || item == DATE_TO)
                {
                    return isDateExist = true;
                }
                if (isDateExist)
                {
                    throw new ArgumentException("Only one type of date supported");
                }
            }
            return false;
        }
    }
}
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.UI;
using System.Web.UI.WebControls;
using System.Data;
using System.Reflection;

namespace eLearning.Common.Utils
{
    public class SimpleExportToExcel
    {

        Logger logger = Logger.getInstance();
        string Module_NAME = "SimpleExportToExcel";

        // Added by AVANZA\Hamza on 04/09/2018
        public void GenerateExcel(string filename, DataTable table)
        {
            try
            {
                HttpContext.Current.Response.Clear();
                HttpContext.Current.Response.ClearContent();
                HttpContext.Current.Response.ClearHeaders();
                HttpContext.Current.Response.Buffer = true;

                // corrected by awais, now excel supports every language
                HttpContext.Current.Response.AddHeader("Content-Disposition", "attachment;filename=" + filename + "-" + DateTime.Now.ToString() + ".xls");
                HttpContext.Current.Response.ContentType = "application/ms-excel";
                HttpContext.Current.Response.ContentEncoding = System.Text.Encoding.Unicode;
                HttpContext.Current.Response.BinaryWrite(System.Text.Encoding.Unicode.GetPreamble());//HttpContext.Current.Response.ContentType = "application/ms-word";
                HttpContext.Current.Response.Write(@"<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">");
                
                HttpContext.Current.Response.Write("<font style='font-size:10.0pt; font-family:Calibri;'>");

				HttpContext.Current.Response.Write("<div style=\"margin:150px\"><BR/><h1>" + filename + "</h1></div>");

				////-----------------For Date-----------------------
                //HttpContext.Current.Response.Write("<BR>");
                HttpContext.Current.Response.Write("<Table border='0' bgColor='#ffffff' borderColor='#000000' cellSpacing='0' cellPadding='0' style='font-size:10.0pt; font-family:Calibri; background:white;'> <TR>");
                int dateCol = table.Columns.Count + 3;

                for (int x = 0; x < dateCol; x++)
                {
                    HttpContext.Current.Response.Write("<Td>");
                    HttpContext.Current.Response.Write("<B>");
                    HttpContext.Current.Response.Write(x == dateCol - 1? "Date: "+ DateTime.Now.ToString():" ");
                    HttpContext.Current.Response.Write("</B>");
                    HttpContext.Current.Response.Write("</Td>");
                }

                HttpContext.Current.Response.Write("</Table>");

                ////-----------------report table-----------------------
                HttpContext.Current.Response.Write("<BR>");
                HttpContext.Current.Response.Write("<BR>");
                HttpContext.Current.Response.Write("<Table border='1' bgColor='#ffffff' borderColor='#000000' cellSpacing='0' cellPadding='0' style='font-size:10.0pt; font-family:Calibri; background:white;'> <TR>");
                int columnscount = table.Columns.Count;
               

                //fetching header
                for (int j = 0; j < columnscount; j++)
                {
                    HttpContext.Current.Response.Write("<Td>");
                    HttpContext.Current.Response.Write("<B>");
                    HttpContext.Current.Response.Write(table.Columns[j].ToString());
                    HttpContext.Current.Response.Write("</B>");
                    HttpContext.Current.Response.Write("</Td>");
                }
                HttpContext.Current.Response.Write("</TR>");
                //fetching rows
                foreach (DataRow row in table.Rows)
                {
                    HttpContext.Current.Response.Write("<TR>");
                    for (int i = 0; i < table.Columns.Count; i++)
                    {
                        HttpContext.Current.Response.Write("<Td>");
                        HttpContext.Current.Response.Write(row[i].ToString());
                        HttpContext.Current.Response.Write("</Td>");
                    }

                    HttpContext.Current.Response.Write("</TR>");
                }
                HttpContext.Current.Response.Write("</Table>");
                HttpContext.Current.Response.Write("</font>");
              
                HttpContext.Current.Response.Flush();
                HttpContext.Current.Response.End();
            }

            catch (Exception ex)
            {
                logger.Error(Module_NAME, "GenerateExcel", ex);
                throw ex;
            }
        }
    }
}

using System;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.IO;
using System.Threading;
using System.Net.Mail;
using System.Net.Mime;
using System.Collections;

namespace Lead_Performance_Email
{
    class Program
    {
        enum TIMEPERIOD:int { currentMonth = 0, previousMonth = 1, currentFYTD = 2, previousFYTD = 3 };
        enum RESULTSET:int { first = 0, second = 1 };
        enum LEADTYPE:int { organic = 0, TPL = 1 };
        enum METRIC:int { valids = 0, uniques = 1, sales = 2, closeRate = 3 };

        static void Main(string[] args)
        {
            string subjectLine = args[1] + " Lead Performance - thru " + (DateTime.Today.AddDays(-1)).ToString("yyyy-MM-dd");

            string query = File.ReadAllText(QueryFetcher(args[0])); //Put your entire written query into a variable

            DataSet[] queryResults = new DataSet[4]; //The containers in which the query's results will be stored
            for (int i = 0; i < queryResults.Length; i++)
            {
                queryResults[i] = new DataSet(); //Initializing each dataset
            }

            
            Thread[] threads = new Thread[4]; //You need one thread for each report you are running
            for (int i = 0; i < threads.Length; i++)
            {
                ParameterizedThreadStart start = new ParameterizedThreadStart(QueryRunner);
                threads[i] = new Thread(start);
                threads[i].Start(new Tuple<string, DataSet, int>(query, queryResults[i], (i + 1)));
                Console.WriteLine("Thread" + i.ToString() + " started!");
                Thread.Sleep(500);
            }
            for (int i = 0; i < threads.Length; i++)
            {
                threads[i].Join(); //This will let all the threads finish before continuing with the rest of the program
            }


            string[] emailDestinations = args[2].Split(';');

            Emailer(queryResults, subjectLine, emailDestinations);
        }

        static void QueryRunner(object obj)
        {
            Tuple<string, DataSet, int> tuple = (Tuple<string, DataSet, int>)obj; //You have to pass in an object to thread this, so you have to cast as a Tuple if you want multiple parameters
            /*** Connecting to and retrieving data from the backoffice server ***/
            using (SqlConnection backofficeConnection = new SqlConnection(@"server=SERVERNAME;Trusted_Connection=yes;")) //Trusted_Connection=yes means it will use your windows credentials by default
            {
                SqlCommand command = new SqlCommand(tuple.Item1, backofficeConnection);
                command.Parameters.Add("@reportType", SqlDbType.Int);
                command.Parameters["@reportType"].Value = tuple.Item3;

                try
                {
                    command.Connection.Open();
                    command.CommandTimeout = 600; //Setting the timeout to 600 seconds, which is 10 minutes

                    command.ExecuteNonQuery();

                    SqlDataAdapter adapter = new SqlDataAdapter(command);
                    adapter.Fill(tuple.Item2);
                    adapter.Dispose();
                }
                catch (Exception ex)
                {
                    //Need to send failure email here
                    Console.WriteLine("SQL Error: " + ex.Message);
                    System.Environment.Exit(-1);
                }
            }
        }

        /*** This method is designed to pull the SQL script's path in the directory in which the project is housed ***/
        static string QueryFetcher(string scriptLocalPath)
        {
            return Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + scriptLocalPath;
        }

        static object DataFetcher(DataSet[] dataSet, TIMEPERIOD timePeriod, RESULTSET resultSet, LEADTYPE leadType, METRIC metric)
        {
            return dataSet[(int)timePeriod].Tables[(int)resultSet].Rows[(int)leadType][(int)metric];
        }

        static string FullBorderTableDataBuilder(int rowSpan, int columnSpan, int widthValue, int fontSize, string textAlign, string verticalAlign, string value)
        {
            string tableData = "";

            tableData += "<td rowspan=\"" + rowSpan.ToString()
                + "\" colspan=\"" + columnSpan.ToString()
                + "\" width=\"" + widthValue.ToString() + "%"
                + "\" style=\"font-size:" + fontSize.ToString()
                + "px;text-align:" + textAlign
                + ";vertical-align:" + verticalAlign
                + ";border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030"
                + ";\">" + value 
                + "</td>";

            return tableData;
        }

        static string NoLeftBorderTableDataBuilder(int rowSpan, int columnSpan, int widthValue, int fontSize, string textAlign, string verticalAlign, string value)
        {
            string tableData = "";

            tableData += "<td rowspan=\"" + rowSpan.ToString()
                + "\" colspan=\"" + columnSpan.ToString()
                + "\" width=\"" + widthValue.ToString() + "%"
                + "\" style=\"font-size:" + fontSize.ToString()
                + "px;text-align:" + textAlign
                + ";vertical-align:" + verticalAlign
                + ";border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; text-align: left"
                + ";\">" + value
                + "</td>";

            return tableData;
        }


        static string GetCSS()
        {
            string css = "<style>";

            css += ".emailTable {";
            css += "max-width:500px;";
            css += "width:100%;";
            css += "border-collapse:collapse;";
            css += "font-family:\"Helvetica Neue\", Helvetica, Arial, sans-serif;}";

            css += ".emailTable th, td{";
            css += "padding:3px;}";

            css += "</style>";
            return css;
        }

        static void Emailer(DataSet[] queryResults, string subjectLine, string[] emailDestinations)
        {
            //Images
            string urbanScienceLogo = @"http://focalpoint.urbanscience.net/MarComm/Logos/Urban%20Science/Stacked%20logo/us_logo_centered_tm_small.jpg";

            LinkedResource inline = new LinkedResource(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\us_logo_centered_tm_small.jpg", "image/jpeg");
            inline.ContentId = Guid.NewGuid().ToString();

            LinkedResource upArrow = new LinkedResource(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\UpArrow.png", "image/png");
            upArrow.ContentId = Guid.NewGuid().ToString();

            LinkedResource downArrow = new LinkedResource(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\DownArrow.png", "image/png");
            downArrow.ContentId = Guid.NewGuid().ToString();


            // Current Month Variables
            int cmOrgValids = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.organic, METRIC.valids);
            int cmOrgUniques = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.organic, METRIC.uniques);
            int cmOrgSales = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.organic, METRIC.sales);
            double cmOrgCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.organic, METRIC.closeRate);
            cmOrgCloseRate *= 100;

            int cmTPLValids = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.valids);
            int cmTPLUniques = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.uniques);
            int cmTPLSales = (int)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.sales);
            double cmTPLCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.closeRate);
            cmTPLCloseRate *= 100;

            double cmPercentOfRetail = (double)DataFetcher(queryResults, TIMEPERIOD.currentMonth, RESULTSET.second, LEADTYPE.organic, METRIC.valids);
            cmPercentOfRetail *= 100;


            // Previous Month Variables
            int pmOrgValids = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.organic, METRIC.valids);
            int pmOrgUniques = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.organic, METRIC.uniques);
            int pmOrgSales = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.organic, METRIC.sales);
            double pmOrgCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.organic, METRIC.closeRate);
            pmOrgCloseRate *= 100;

            int pmTPLValids = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.valids);
            int pmTPLUniques = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.uniques);
            int pmTPLSales = (int)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.sales);
            double pmTPLCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.first, LEADTYPE.TPL, METRIC.closeRate);
            pmTPLCloseRate *= 100;

            double pmPercentOfRetail = (double)DataFetcher(queryResults, TIMEPERIOD.previousMonth, RESULTSET.second, LEADTYPE.organic, METRIC.valids);
            pmPercentOfRetail *= 100;


            // FYTD Variables
            int fytdOrgValids = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.valids);
            int fytdOrgUniques = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.uniques);
            int fytdOrgSales = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.sales);
            double fytdOrgCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.closeRate);
            fytdOrgCloseRate *= 100;

            int fytdTPLValids = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.valids);
            int fytdTPLUniques = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.uniques);
            int fytdTPLSales = (int)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.sales);
            double fytdTPLCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.closeRate);
            fytdTPLCloseRate *= 100;

            double fytdPercentOfRetail = (double)DataFetcher(queryResults, TIMEPERIOD.currentFYTD, RESULTSET.second, LEADTYPE.organic, METRIC.valids);
            fytdPercentOfRetail *= 100;


            // Previous FYTD Variables
            int pfytdOrgValids = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.valids);
            int pfytdOrgUniques = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.uniques);
            int pfytdOrgSales = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.sales);
            double pfytdOrgCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.organic, METRIC.closeRate);
            pfytdOrgCloseRate *= 100;

            int pfytdTPLValids = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.valids);
            int pfytdTPLUniques = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.uniques);
            int pfytdTPLSales = (int)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.sales);
            double pfytdTPLCloseRate = (double)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.first, LEADTYPE.TPL, METRIC.closeRate);
            pfytdTPLCloseRate *= 100;

            double pfytdPercentOfRetail = (double)DataFetcher(queryResults, TIMEPERIOD.previousFYTD, RESULTSET.second, LEADTYPE.organic, METRIC.valids);
            pfytdPercentOfRetail *= 100;


            // Computed Variables
            double monthYoYPercentOfRetail = cmPercentOfRetail - pmPercentOfRetail;
            double fytdYoYPercentOfRetail = fytdPercentOfRetail - pfytdPercentOfRetail;

            double monthYoYOrgValids = (((double)cmOrgValids - (double)pmOrgValids) / (double)pmOrgValids) * 100;
            double monthYoYOrgUniques = (((double)cmOrgUniques - (double)pmOrgUniques) / (double)pmOrgUniques) * 100;
            double monthYoYOrgSales = (((double)cmOrgSales - (double)pmOrgSales) / (double)pmOrgSales) * 100;
            double monthYoYOrgCloseRate = cmOrgCloseRate - pmOrgCloseRate;

            double monthYoYTPLValids = (((double)cmTPLValids - (double)pmTPLValids) / (double)pmTPLValids) * 100;
            double monthYoYTPLUniques = (((double)cmTPLUniques - (double)pmTPLUniques) / (double)pmTPLUniques) * 100;
            double monthYoYTPLSales = (((double)cmTPLSales - (double)pmTPLSales) / (double)pmTPLSales) * 100;
            double monthYoYTPLCloseRate = cmTPLCloseRate - pmTPLCloseRate;

            double fytdYoYOrgValids = (((double)fytdOrgValids - (double)pfytdOrgValids) / (double)pfytdOrgValids) * 100;
            double fytdYoYOrgUniques = (((double)fytdOrgUniques - (double)pfytdOrgUniques) / (double)pfytdOrgUniques) * 100;
            double fytdYoYOrgSales = (((double)fytdOrgSales - (double)pfytdOrgSales) / (double)pfytdOrgSales) * 100;
            double fytdYoYOrgCloseRate = fytdOrgCloseRate - pfytdOrgCloseRate;

            double fytdYoYTPLValids = (((double)fytdTPLValids - (double)pfytdTPLValids) / (double)pfytdTPLValids) * 100;
            double fytdYoYTPLUniques = (((double)fytdTPLUniques - (double)pfytdTPLUniques) / (double)pfytdTPLUniques) * 100;
            double fytdYoYTPLSales = (((double)fytdTPLSales - (double)pfytdTPLSales) / (double)pfytdTPLSales) * 100;
            double fytdYoYTPLCloseRate = fytdTPLCloseRate - pfytdTPLCloseRate;


            //Setting up or down arrows
            string monthPoRArrow = (monthYoYPercentOfRetail > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdPoRArrow = (fytdYoYPercentOfRetail > 0) ? upArrow.ContentId : downArrow.ContentId;

            string monthOrgValidsArrow = (monthYoYOrgValids > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthOrgUniquesArrow = (monthYoYOrgUniques > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthOrgSalesArrow = (monthYoYOrgSales > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthOrgCloseRateArrow = (monthYoYOrgCloseRate > 0) ? upArrow.ContentId : downArrow.ContentId;

            string fytdOrgValidsArrow = (fytdYoYOrgValids > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdOrgUniquesArrow = (fytdYoYOrgUniques > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdOrgSalesArrow = (fytdYoYOrgSales > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdOrgCloseRateArrow = (fytdYoYOrgCloseRate > 0) ? upArrow.ContentId : downArrow.ContentId;

            string monthTPLValidsArrow = (monthYoYTPLValids > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthTPLUniquesArrow = (monthYoYTPLUniques > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthTPLSalesArrow = (monthYoYTPLSales > 0) ? upArrow.ContentId : downArrow.ContentId;
            string monthTPLCloseRateArrow = (monthYoYTPLCloseRate > 0) ? upArrow.ContentId : downArrow.ContentId;

            string fytdTPLValidsArrow = (fytdYoYTPLValids > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdTPLUniquesArrow = (fytdYoYTPLUniques > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdTPLSalesArrow = (fytdYoYTPLSales > 0) ? upArrow.ContentId : downArrow.ContentId;
            string fytdTPLCloseRateArrow = (fytdYoYTPLCloseRate > 0) ? upArrow.ContentId : downArrow.ContentId;


            string body = "<!DOCTYPE html>";
            body += "<html><head>";
            body += GetCSS();
            body += "</head><body>";

            // "Sales from Leads - New Vehicles Only" block
            body += "<div style=\"margin:10px;\"><table class=\"emailTable\">";
            body += "<tr><td rowspan=\"1\" colspan=\"10\" width=\"100%\" bgcolor=\"#303030\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030; color: white; font-size:20px;text-align:center;vertical-align:baseline;\">Sales from Leads - New Vehicles Only</td></tr>";

            //Image row 1
            body += "<tr><td colspan=2 rowspan=2 height=\"30\" width=\"60\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"middle\" style=\"display:block;\" height=\"50\" width=\"110\" src=\"cid:" + inline.ContentId + "\" alt=\"" + urbanScienceLogo + "\" /></td>";
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "MTD");
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "FYTD") + "</tr>";

            //Image row 2
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY") + "</tr>";

            //"% of Retail Sales" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "% of Retail Sales");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", (cmPercentOfRetail.ToString("0.0") + "%"));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthPoRArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYPercentOfRetail.ToString("0.0") + "pp");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", fytdPercentOfRetail.ToString("0.0") + "%");
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdPoRArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYPercentOfRetail.ToString("0.0") + "pp") + "</tr>";

            body += "<tr><td rowspan=\"1\" colspan=\"10\" width=\"100%\" height=\"10\"></td></tr>";


            // "Third Party Leads" block
            body += "<tr><td rowspan=\"1\" colspan=\"10\" width=\"100%\" bgcolor=\"#303030\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030; color: white; font-size:20px;text-align:center;vertical-align:baseline;\">Third Party Leads</td></tr>";

            //Image row 1
            body += "<tr><td colspan=2 rowspan=2 height=\"30\" width=\"60\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"middle\" style=\"display:block;\" height=\"50\" width=\"110\" src=\"cid:" + inline.ContentId + "\" alt=\"" + urbanScienceLogo + "\" /></td>";
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "MTD");
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "FYTD") + "</tr>";

            //Image row 2
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY") + "</tr>";

            //"Valid Leads" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Valid Leads");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmTPLValids));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthTPLValidsArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYTPLValids.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdTPLValids));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdTPLValidsArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYTPLValids.ToString("0.0") + "%") + "</tr>";

            //"Unique Leads" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Unique Leads");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmTPLUniques));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthTPLUniquesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYTPLUniques.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdTPLUniques));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdTPLUniquesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYTPLUniques.ToString("0.0") + "%") + "</tr>";

            //"Sales" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Sales");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmTPLSales));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthTPLSalesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYTPLSales.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdTPLSales));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdTPLSalesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYTPLSales.ToString("0.0") + "%") + "</tr>";

            //"Close Rate" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Close Rate");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", cmTPLCloseRate.ToString("0.0") + "%");
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthTPLCloseRateArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYTPLCloseRate.ToString("0.0") + "pp");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", fytdTPLCloseRate.ToString("0.0") + "%");
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdTPLCloseRateArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYTPLCloseRate.ToString("0.0") + "pp") + "</tr>";

            body += "<tr><td rowspan=\"1\" colspan=\"10\" width=\"100%\" height=\"10\"></td></tr>";


            // "Organic" block
            body += "<tr><td rowspan=\"1\" colspan=\"10\" width=\"100%\" bgcolor=\"#303030\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030; color: white; font-size:20px;text-align:center;vertical-align:baseline;\">Organic</td></tr>";

            //Image row 1
            body += "<tr><td colspan=2 rowspan=2 height=\"30\" width=\"60\" style=\"border-top: 1px solid #303030; border-right: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"middle\" style=\"display:block;\" height=\"50\" width=\"110\" src=\"cid:" + inline.ContentId + "\" alt=\"" + urbanScienceLogo + "\" /></td>";
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "MTD");
            body += FullBorderTableDataBuilder(1, 4, 40, 20, "center", "text-bottom", "FYTD") + "</tr>";

            //Image row 2
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "Results");
            body += FullBorderTableDataBuilder(1, 2, 20, 14, "center", "middle", "YoY") + "</tr>";

            //"Valid Leads" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Valid Leads");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmOrgValids));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthOrgValidsArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYOrgValids.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdOrgValids));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdOrgValidsArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYOrgValids.ToString("0.0") + "%") + "</tr>";

            //"Unique Leads" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Unique Leads");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmOrgUniques));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthOrgUniquesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYOrgUniques.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdOrgUniques));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdOrgUniquesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYOrgUniques.ToString("0.0") + "%") + "</tr>";

            //"Sales" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Sales");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", cmOrgSales));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthOrgSalesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYOrgSales.ToString("0.0") + "%");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", String.Format("{0:n0}", fytdOrgSales));
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdOrgSalesArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYOrgSales.ToString("0.0") + "%") + "</tr>";

            //"Close Rate" row
            body += "<tr>" + FullBorderTableDataBuilder(1, 2, 20, 14, "left", "middle", "Close Rate");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", cmOrgCloseRate.ToString("0.0") + "%");
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + monthOrgCloseRateArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", monthYoYOrgCloseRate.ToString("0.0") + "pp");
            body += FullBorderTableDataBuilder(1, 2, 20, 12, "center", "middle", fytdOrgCloseRate.ToString("0.0") + "%");
            body += "<td colspan=1 rowspan=1 height=\"10\" width=\"10\" style=\"border-top: 1px solid #303030; border-bottom: 1px solid #303030; border-left: 1px solid #303030\"><img align=\"right\" style=\"display:block;\" width=\"20\" height=\"20\" src=\"cid:" + fytdOrgCloseRateArrow + "\" /></td>";
            body += NoLeftBorderTableDataBuilder(1, 1, 10, 12, "center", "middle", fytdYoYOrgCloseRate.ToString("0.0") + "pp") + "</tr></table></div>";


            body += "<p>*Approximately 25% of sales for any given sales month are reported within the last three days of the sales month. Please consider this when comparing current period vs. previous year.</p></body></html>";

            AlternateView alternate = AlternateView.CreateAlternateViewFromString(body, null, MediaTypeNames.Text.Html);
            alternate.LinkedResources.Add(inline);

            SmtpClient smtp_server = new SmtpClient("SMTPCLIENT);
            MailMessage email = new MailMessage();
            email.AlternateViews.Add(alternate);

            Attachment attachment = new Attachment(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\us_logo_centered_tm_small.jpg");
            Attachment attachment2 = new Attachment(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\UpArrow.png");
            Attachment attachment3 = new Attachment(Directory.GetParent(Directory.GetCurrentDirectory()).Parent.FullName + @"\resources\DownArrow.png");

            alternate.LinkedResources.Add(inline);
            alternate.LinkedResources.Add(upArrow);
            alternate.LinkedResources.Add(downArrow);

            attachment.ContentDisposition.Inline = true;
            attachment2.ContentDisposition.Inline = true;
            attachment3.ContentDisposition.Inline = true;

            email.From = new MailAddress("FROM");

            foreach (var destination in emailDestinations)
            {
                email.To.Add(destination);
            }

            email.Subject = subjectLine;
            email.Body = "";
            email.IsBodyHtml = true;

            smtp_server.Port = 25;
            smtp_server.Credentials = new System.Net.NetworkCredential("USERNAME", "PASSWORD");
            smtp_server.EnableSsl = false;
            smtp_server.ServicePoint.MaxIdleTime = 1;
            smtp_server.ServicePoint.ConnectionLimit = 1;
            smtp_server.Timeout = 1000000;

            try
            {
                smtp_server.Send(email);
                smtp_server.Dispose();
            }
            catch (Exception ex)
            {
                smtp_server.Timeout = 1000000;
                smtp_server.Send(email);
                smtp_server.Dispose();

                Console.WriteLine("Exception caught in CreateMessageWithMultipleViews(): {0}",
                ex.ToString());
            }
        }
    }
}

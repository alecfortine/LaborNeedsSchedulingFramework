using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using WebMatrix.Data;
using System.Data.SqlClient;
using System.Data;
using System.Configuration;
using LaborNeedsSchedulingFramework.Models;
using System.Diagnostics;
using System.Text;
using System.IO;
using System.Xml;
using System.Xml.Serialization;

namespace LaborNeedsSchedulingFramework.Controllers
{
    public class HomeController : Controller
    {
        // create the datatables
        #region datatables
        DataTable ExcludedDates = new DataTable();
        DataTable WTGTrafficPercent = new DataTable();
        DataTable WeightedAverageTraffic_perWeek = new DataTable();
        DataTable PercentWeeklyTotal = new DataTable();
        DataTable PercentDailyTotal = new DataTable();
        DataTable AllocatedHours = new DataTable();
        DataTable PowerHourForecast = new DataTable();
        #endregion

        public ActionResult Index(DataInput input)
        {
            try
            {

                // create connection to source table
                #region connection stuff
                string strSQLCon = ConfigurationManager.ConnectionStrings["SqlServerConnection"].ToString();

                // open Connection
                SqlConnection sqlCon = new SqlConnection(ConfigurationManager.ConnectionStrings["SqlServerConnection"].ToString());
                sqlCon.Open();
                System.Diagnostics.Debug.WriteLine("State: {0}", sqlCon.State);
                System.Diagnostics.Debug.WriteLine("ConnectionString: {0}", sqlCon.ConnectionString);
                #endregion

                // sql select commands and weighting calculations
                #region calculate WTG

                // gets the date of the most recent sunday so weighting of previous weeks can begin from there
                var weekMarker = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Sunday);

                double[] defaultValues = { 0.36, 0.24, 0.16, 0.12, 0.08, 0.04 };

                double[] week = { input.week1Weight / 100, input.week2Weight / 100, input.week3Weight / 100, input.week4Weight / 100, input.week5Weight / 100, input.week6Weight / 100 };


                SqlCommand CalculateWTG6 = new SqlCommand();
                SqlCommand CalculateWTG5 = new SqlCommand();
                SqlCommand CalculateWTG4 = new SqlCommand();
                SqlCommand CalculateWTG3 = new SqlCommand();
                SqlCommand CalculateWTG2 = new SqlCommand();
                SqlCommand CalculateWTG1 = new SqlCommand();


                //if (week[0] > 0 || week[1] > 0 || week[2] > 0 || week[3] > 0 || week[4] > 0 || week[5] > 0)
                //{

                    // sets untouched inputs to default values
                    for (int i = 0; i < week.Length; i++)
                    {
                        if (week[i] == 0)
                        {
                            week[i] = defaultValues[i];
                        }
                    }

                    CalculateWTG6.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[0] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -5, '2015-05-10')";
                    CalculateWTG6.CommandType = CommandType.Text;
                    CalculateWTG5.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[1] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -4, '2015-05-10') AND[Date] >= dateadd(WEEK, -5, '2015-05-10')";
                    CalculateWTG5.CommandType = CommandType.Text;
                    CalculateWTG4.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[2] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -3, '2015-05-10') AND[Date] >= dateadd(WEEK, -4, '2015-05-10')";
                    CalculateWTG4.CommandType = CommandType.Text;
                    CalculateWTG3.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[3] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -2, '2015-05-10') AND[Date] >= dateadd(WEEK, -3, '2015-05-10')";
                    CalculateWTG3.CommandType = CommandType.Text;
                    CalculateWTG2.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[4] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -1, '2015-05-10') AND[Date] >= dateadd(WEEK, -2, '2015-05-10')";
                    CalculateWTG2.CommandType = CommandType.Text;
                    CalculateWTG1.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + week[5] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] >= dateadd(WEEK, -1, '2015-05-10')";
                    CalculateWTG1.CommandType = CommandType.Text;
                //}
                //else
                //{
                //    CalculateWTG6.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[0] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -5, '2015-05-10')";
                //    CalculateWTG6.CommandType = CommandType.Text;
                //    CalculateWTG5.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[1] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -4, '2015-05-10') AND[Date] >= dateadd(WEEK, -5, '2015-05-10')";
                //    CalculateWTG5.CommandType = CommandType.Text;
                //    CalculateWTG4.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[2] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -3, '2015-05-10') AND[Date] >= dateadd(WEEK, -4, '2015-05-10')";
                //    CalculateWTG4.CommandType = CommandType.Text;
                //    CalculateWTG3.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[3] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -2, '2015-05-10') AND[Date] >= dateadd(WEEK, -3, '2015-05-10')";
                //    CalculateWTG3.CommandType = CommandType.Text;
                //    CalculateWTG2.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[4] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -1, '2015-05-10') AND[Date] >= dateadd(WEEK, -2, '2015-05-10')";
                //    CalculateWTG2.CommandType = CommandType.Text;
                //    CalculateWTG1.CommandText = "Select FORMAT([Date], 'M/d/yyyy') as Date, [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * " + defaultValues[5] + "),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] >= dateadd(WEEK, -1, '2015-05-10')";
                //    CalculateWTG1.CommandType = CommandType.Text;
                //}

                // possible future version
                //SqlCommand CalculateWTG6 = new SqlCommand();
                //CalculateWTG6.CommandText = "Select[Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.04),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -5, weekMarker)";
                //CalculateWTG6.CommandType = CommandType.Text;
                //SqlCommand CalculateWTG5 = new SqlCommand();
                //CalculateWTG5.CommandText = "Select [Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.08),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -4, weekMarker) AND[Date] >= dateadd(WEEK, -5, weekMarker)";
                //CalculateWTG5.CommandType = CommandType.Text;
                //SqlCommand CalculateWTG4 = new SqlCommand();
                //CalculateWTG4.CommandText = "Select [Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.12),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -3, weekMarker) AND[Date] >= dateadd(WEEK, -4, weekMarker)";
                //CalculateWTG4.CommandType = CommandType.Text;
                //SqlCommand CalculateWTG3 = new SqlCommand();
                //CalculateWTG3.CommandText = "Select [Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.16),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -2, weekMarker) AND[Date] >= dateadd(WEEK, -3, weekMarker)";
                //CalculateWTG3.CommandType = CommandType.Text;
                //SqlCommand CalculateWTG2 = new SqlCommand();
                //CalculateWTG2.CommandText = "Select [Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.24),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] < dateadd(WEEK, -1, weekMarker) AND[Date] >= dateadd(WEEK, -2, weekMarker)";
                //CalculateWTG2.CommandType = CommandType.Text;
                //SqlCommand CalculateWTG1 = new SqlCommand();
                //CalculateWTG1.CommandText = "Select [Date], [WeekDay], HourofDay, TrafficOut, Round((TrafficOut * 0.36),2) as WTGTraffic From Backup_LastSixWeeksTraffic Where[Date] >= dateadd(WEEK, -1, weekMarker)";
                //CalculateWTG1.CommandType = CommandType.Text;

                #endregion

                // establish connection for commands and fill initial table
                #region fill initial table
                // establish connection for WTG commands
                CalculateWTG6.Connection = sqlCon;
                CalculateWTG5.Connection = sqlCon;
                CalculateWTG4.Connection = sqlCon;
                CalculateWTG3.Connection = sqlCon;
                CalculateWTG2.Connection = sqlCon;
                CalculateWTG1.Connection = sqlCon;

                SqlDataAdapter sixweeks = new SqlDataAdapter(CalculateWTG6);
                SqlDataAdapter fiveweeks = new SqlDataAdapter(CalculateWTG5);
                SqlDataAdapter fourweeks = new SqlDataAdapter(CalculateWTG4);
                SqlDataAdapter threeweeks = new SqlDataAdapter(CalculateWTG3);
                SqlDataAdapter twoweeks = new SqlDataAdapter(CalculateWTG2);
                SqlDataAdapter oneweeks = new SqlDataAdapter(CalculateWTG1);

                sixweeks.Fill(WTGTrafficPercent);
                fiveweeks.Fill(WTGTrafficPercent);
                fourweeks.Fill(WTGTrafficPercent);
                threeweeks.Fill(WTGTrafficPercent);
                twoweeks.Fill(WTGTrafficPercent);
                oneweeks.Fill(WTGTrafficPercent);
                #endregion

                // set the dates for each day in the past 6 weeks
                #region set dates

                // gets the current saturday
                DateTime weekLimitX = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Saturday);

                // temp value for testing with current database
                DateTime weekLimit = new DateTime(2015, 05, 16);

                // get the dates for this week
                DateTime saturday0String = weekLimitX.AddDays(0);
                DateTime friday0String = weekLimitX.AddDays(-1);
                DateTime thursday0String = weekLimitX.AddDays(-2);
                DateTime wednesday0String = weekLimitX.AddDays(-3);
                DateTime tuesday0String = weekLimitX.AddDays(-4);
                DateTime monday0String = weekLimitX.AddDays(-5);
                DateTime sunday0String = weekLimitX.AddDays(-6);

                // get the dates for the past 6 weeks
                DateTime saturday1String = weekLimit.AddDays(-7);
                DateTime friday1String = weekLimit.AddDays(-8);
                DateTime thursday1String = weekLimit.AddDays(-9);
                DateTime wednesday1String = weekLimit.AddDays(-10);
                DateTime tuesday1String = weekLimit.AddDays(-11);
                DateTime monday1String = weekLimit.AddDays(-12);
                DateTime sunday1String = weekLimit.AddDays(-13);

                DateTime saturday2String = weekLimit.AddDays(-14);
                DateTime friday2String = weekLimit.AddDays(-15);
                DateTime thursday2String = weekLimit.AddDays(-16);
                DateTime wednesday2String = weekLimit.AddDays(-17);
                DateTime tuesday2String = weekLimit.AddDays(-18);
                DateTime monday2String = weekLimit.AddDays(-19);
                DateTime sunday2String = weekLimit.AddDays(-20);

                DateTime saturday3String = weekLimit.AddDays(-21);
                DateTime friday3String = weekLimit.AddDays(-22);
                DateTime thursday3String = weekLimit.AddDays(-23);
                DateTime wednesday3String = weekLimit.AddDays(-24);
                DateTime tuesday3String = weekLimit.AddDays(-25);
                DateTime monday3String = weekLimit.AddDays(-26);
                DateTime sunday3String = weekLimit.AddDays(-27);

                DateTime saturday4String = weekLimit.AddDays(-28);
                DateTime friday4String = weekLimit.AddDays(-29);
                DateTime thursday4String = weekLimit.AddDays(-30);
                DateTime wednesday4String = weekLimit.AddDays(-31);
                DateTime tuesday4String = weekLimit.AddDays(-32);
                DateTime monday4String = weekLimit.AddDays(-33);
                DateTime sunday4String = weekLimit.AddDays(-34);

                DateTime saturday5String = weekLimit.AddDays(-35);
                DateTime friday5String = weekLimit.AddDays(-36);
                DateTime thursday5String = weekLimit.AddDays(-37);
                DateTime wednesday5String = weekLimit.AddDays(-38);
                DateTime tuesday5String = weekLimit.AddDays(-39);
                DateTime monday5String = weekLimit.AddDays(-40);
                DateTime sunday5String = weekLimit.AddDays(-41);

                DateTime saturday6String = weekLimit.AddDays(-42);
                DateTime friday6String = weekLimit.AddDays(-43);
                DateTime thursday6String = weekLimit.AddDays(-44);
                DateTime wednesday6String = weekLimit.AddDays(-45);
                DateTime tuesday6String = weekLimit.AddDays(-46);
                DateTime monday6String = weekLimit.AddDays(-47);
                DateTime sunday6String = weekLimit.AddDays(-48);

                // time strings
                string eightAM = "8AM-9AM";
                string nineAM = "9AM-10AM";
                string tenAM = "10AM-11AM";
                string elevenAM = "11AM-12PM";
                string twelvePM = "12PM-1PM";
                string onePM = "1PM-2PM";
                string twoPM = "2PM-3PM";
                string threePM = "3PM-4PM";
                string fourPM = "4PM-5PM";
                string fivePM = "5PM-6PM";
                string sixPM = "6PM-7PM";
                string sevenPM = "7PM-8PM";
                string eightPM = "8PM-9PM";
                string ninePM = "9PM-10PM";
                string tenPM = "10PM-11PM";


                // convert the dates to strings for correct date format
                DataInput dates = new DataInput
                {
                    saturday0 = saturday0String.ToString("M/d/yyyy"),
                    friday0 = friday0String.ToString("M/d/yyyy"),
                    thursday0 = thursday0String.ToString("M/d/yyyy"),
                    wednesday0 = wednesday0String.ToString("M/d/yyyy"),
                    tuesday0 = tuesday0String.ToString("M/d/yyyy"),
                    monday0 = monday0String.ToString("M/d/yyyy"),
                    sunday0 = sunday0String.ToString("M/d/yyyy"),

                    saturday1 = saturday1String.ToString("M/d/yyyy"),
                    friday1 = friday1String.ToString("M/d/yyyy"),
                    thursday1 = thursday1String.ToString("M/d/yyyy"),
                    wednesday1 = wednesday1String.ToString("M/d/yyyy"),
                    tuesday1 = tuesday1String.ToString("M/d/yyyy"),
                    monday1 = monday1String.ToString("M/d/yyyy"),
                    sunday1 = sunday1String.ToString("M/d/yyyy"),

                    saturday2 = saturday2String.ToString("M/d/yyyy"),
                    friday2 = friday2String.ToString("M/d/yyyy"),
                    thursday2 = thursday2String.ToString("M/d/yyyy"),
                    wednesday2 = wednesday2String.ToString("M/d/yyyy"),
                    tuesday2 = tuesday2String.ToString("M/d/yyyy"),
                    monday2 = monday2String.ToString("M/d/yyyy"),
                    sunday2 = sunday2String.ToString("M/d/yyyy"),

                    saturday3 = saturday3String.ToString("M/d/yyyy"),
                    friday3 = friday3String.ToString("M/d/yyyy"),
                    thursday3 = thursday3String.ToString("M/d/yyyy"),
                    wednesday3 = wednesday3String.ToString("M/d/yyyy"),
                    tuesday3 = tuesday3String.ToString("M/d/yyyy"),
                    monday3 = monday3String.ToString("M/d/yyyy"),
                    sunday3 = sunday3String.ToString("M/d/yyyy"),

                    saturday4 = saturday4String.ToString("M/d/yyyy"),
                    friday4 = friday4String.ToString("M/d/yyyy"),
                    thursday4 = thursday4String.ToString("M/d/yyyy"),
                    wednesday4 = wednesday4String.ToString("M/d/yyyy"),
                    tuesday4 = tuesday4String.ToString("M/d/yyyy"),
                    monday4 = monday4String.ToString("M/d/yyyy"),
                    sunday4 = sunday4String.ToString("M/d/yyyy"),

                    saturday5 = saturday5String.ToString("M/d/yyyy"),
                    friday5 = friday5String.ToString("M/d/yyyy"),
                    thursday5 = thursday5String.ToString("M/d/yyyy"),
                    wednesday5 = wednesday5String.ToString("M/d/yyyy"),
                    tuesday5 = tuesday5String.ToString("M/d/yyyy"),
                    monday5 = monday5String.ToString("M/d/yyyy"),
                    sunday5 = sunday5String.ToString("M/d/yyyy"),

                    saturday6 = saturday6String.ToString("M/d/yyyy"),
                    friday6 = friday6String.ToString("M/d/yyyy"),
                    thursday6 = thursday6String.ToString("M/d/yyyy"),
                    wednesday6 = wednesday6String.ToString("M/d/yyyy"),
                    tuesday6 = tuesday6String.ToString("M/d/yyyy"),
                    monday6 = monday6String.ToString("M/d/yyyy"),
                    sunday6 = sunday6String.ToString("M/d/yyyy"),

                };
                #endregion

                // set the model values
                DataInput table = new DataInput
                {
                    managerTable = AllocatedHours,
                    payRoll = input.payRoll,
                    minEmp = input.minEmp,
                    maxEmp = input.maxEmp,

                    week1Weight = input.week1Weight,
                    week2Weight = input.week2Weight,
                    week3Weight = input.week3Weight,
                    week4Weight = input.week4Weight,
                    week5Weight = input.week5Weight,
                    week6Weight = input.week6Weight,

                    monday1check = input.monday1check
                };

                // set the weightng inputs
                //DataInput weighting = new DataInput
                //{
                //    week1Weight = input.week1Weight,
                //    week2Weight = input.week2Weight,
                //    week3Weight = input.week3Weight,
                //    week4Weight = input.week4Weight,
                //    week5Weight = input.week5Weight,
                //    week6Weight = input.week6Weight
                //};

                // testing bool setting
                //DataInput bools = new DataInput
                //{
                //    monday1check = input.monday1check
                //};

                // run procedures
                Exclude_Dates(dates);
                Update_WeightedAverageTraffic(table);
                Calculate_PercentWeekly();
                Calculate_PercentDaily();
                Calculate_AllocatedHours(table);
                Calculate_PowerHourForecast();

                // dispose connections
                CalculateWTG6.Dispose();
                CalculateWTG5.Dispose();
                CalculateWTG4.Dispose();
                CalculateWTG3.Dispose();
                CalculateWTG2.Dispose();
                CalculateWTG1.Dispose();
                sqlCon.Close();
                sqlCon.Dispose();
                sixweeks.Dispose();
                fiveweeks.Dispose();
                fourweeks.Dispose();
                threeweeks.Dispose();
                twoweeks.Dispose();
                oneweeks.Dispose();

                // return table to the view
                return View(table);

            }
            catch (Exception ex)
            {
                Debug.WriteLine("Exception Message: " + ex.Message);
                return null;
            }
        }

        public DataTable Exclude_Dates(DataInput dates)
        {

            /*
             * Check which dates should be excluded
             * from calculations from previous weeks
             * 
             * Dates are set with a bool from the model
             * and will be set by manager input
             */

            #region table setup
            ExcludedDates.Columns.Add("Date", typeof(string));
            ExcludedDates.Columns.Add("WeekDay", typeof(string));
            ExcludedDates.Columns.Add("HourofDay", typeof(string));
            ExcludedDates.Columns.Add("TrafficOut", typeof(string));
            //ExcludedDates.Columns.Add("WTGPercent", typeof(string));
            ExcludedDates.Columns.Add("WTGTraffic", typeof(string));
            #endregion

            #region exclusion check
            // if date and exclusion check match, don't add row to the table
            foreach (DataRow row in WTGTrafficPercent.Rows)
            {

                if (row["Date"].ToString() == dates.sunday6 && dates.sunday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.sunday5 && dates.sunday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.sunday4 && dates.sunday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.sunday3 && dates.sunday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.sunday2 && dates.sunday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.sunday1 && dates.sunday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday6 && dates.monday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday5 && dates.monday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday4 && dates.monday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday3 && dates.monday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday2 && dates.monday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.monday1 && dates.monday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday6 && dates.tuesday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday5 && dates.tuesday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday4 && dates.tuesday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday3 && dates.tuesday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday2 && dates.tuesday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.tuesday1 && dates.tuesday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday6 && dates.wednesday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday5 && dates.wednesday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday4 && dates.wednesday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday3 && dates.wednesday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday2 && dates.wednesday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.wednesday1 && dates.wednesday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday6 && dates.thursday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday5 && dates.thursday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday4 && dates.thursday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday3 && dates.thursday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday2 && dates.thursday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.thursday1 && dates.thursday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday6 && dates.friday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday5 && dates.friday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday4 && dates.friday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday3 && dates.friday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday2 && dates.friday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.friday1 && dates.friday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday6 && dates.saturday6check == true)
                {
                    Debug.WriteLine(row["Date"].ToString());
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday5 && dates.saturday5check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday4 && dates.saturday4check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday3 && dates.saturday3check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday2 && dates.saturday2check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }
                else if (row["Date"].ToString() == dates.saturday1 && dates.saturday1check == true)
                {
                    Debug.WriteLine(row["Date"]);
                    continue;
                }

                ExcludedDates.Rows.Add(row.ItemArray);
            }
            #endregion

            return ExcludedDates;
        }

        public DataTable Update_WeightedAverageTraffic(DataInput input)
        {
            /*
             * Get the weighted average traffic from the
             * initial input table and prints it to a 
             * new table
             */

            // clean up later

            #region tablesetup
            WeightedAverageTraffic_perWeek.Columns.Add("HourOfDay", typeof(string));
            WeightedAverageTraffic_perWeek.Columns.Add("SunWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("MonWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("TueWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("WedWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("ThuWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("FriWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("SatWeightedAverage", typeof(double));
            WeightedAverageTraffic_perWeek.Columns.Add("Total", typeof(double));

            WeightedAverageTraffic_perWeek.Rows.Add("9AM-10AM");
            WeightedAverageTraffic_perWeek.Rows.Add("10AM-11AM");
            WeightedAverageTraffic_perWeek.Rows.Add("11AM-12PM");
            WeightedAverageTraffic_perWeek.Rows.Add("12PM-1PM");
            WeightedAverageTraffic_perWeek.Rows.Add("1PM-2PM");
            WeightedAverageTraffic_perWeek.Rows.Add("2PM-3PM");
            WeightedAverageTraffic_perWeek.Rows.Add("3PM-4PM");
            WeightedAverageTraffic_perWeek.Rows.Add("4PM-5PM");
            WeightedAverageTraffic_perWeek.Rows.Add("5PM-6PM");
            WeightedAverageTraffic_perWeek.Rows.Add("6PM-7PM");
            WeightedAverageTraffic_perWeek.Rows.Add("7PM-8PM");
            WeightedAverageTraffic_perWeek.Rows.Add("8PM-9PM");
            WeightedAverageTraffic_perWeek.Rows.Add("9PM-10PM");
            WeightedAverageTraffic_perWeek.Rows.Add("10PM-11PM");
            WeightedAverageTraffic_perWeek.Rows.Add("Total");
            #endregion

            // testing
            Debug.WriteLine(input.monday1check);

            #region Sunday
            string sun9AM = "WeekDay = 'Sun' AND HourOfDay = '9AM-10AM'";
            string sun10AM = "WeekDay = 'Sun' AND HourOfDay = '10AM-11AM'";
            string sun11AM = "WeekDay = 'Sun' AND HourOfDay = '11AM-12PM'";
            string sun12PM = "WeekDay = 'Sun' AND HourOfDay = '12PM-1PM'";
            string sun1PM = "WeekDay = 'Sun' AND HourOfDay = '1PM-2PM'";
            string sun2PM = "WeekDay = 'Sun' AND HourOfDay = '2PM-3PM'";
            string sun3PM = "WeekDay = 'Sun' AND HourOfDay = '3PM-4PM'";
            string sun4PM = "WeekDay = 'Sun' AND HourOfDay = '4PM-5PM'";
            string sun5PM = "WeekDay = 'Sun' AND HourOfDay = '5PM-6PM'";
            string sun6PM = "WeekDay = 'Sun' AND HourOfDay = '6PM-7PM'";
            string sun7PM = "WeekDay = 'Sun' AND HourOfDay = '7PM-8PM'";
            string sun8PM = "WeekDay = 'Sun' AND HourOfDay = '8PM-9PM'";
            string sun9PM = "WeekDay = 'Sun' AND HourOfDay = '9PM-10PM'";
            string sun10PM = "WeekDay = 'Sun' AND HourOfDay = '10PM-11PM'";


            //testing out a shorter way to do the procedure using for loops
            /*
            for (int n = 9; n <= 12; n++)
            {

                if (n < 11)
                {
                    sundayAM = "Weekday = 'Sun' AND HourOfDay = '" + n + "AM-" + (n + 1) + "AM'";
                }
                else if (n == 11)
                {
                    sundayAM = "Weekday = 'Sun' AND HourOfDay = '" + n + "PM-12PM'";
                }
                else
                {
                    sundayAM = "Weekday = 'Sun' AND HourOfDay = '" + n + "PM-1PM'";
                }
                //Debug.WriteLine(sundayAM);


                sundayTraffic = ExcludedDates.Select(sundayAM);

                for (int i = 0; i < sundayTraffic.Length; i++)
                {
                    if (i < sundayTraffic.Length - 1)
                    {
                        total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    }
                    else
                    {
                        total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                        Math.Round(total, MidpointRounding.AwayFromZero);
                        WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                        total = 0;
                        row++;
                    }
                }


            }

            for (int n = 1; n <= 10; n++)
            {
                sundayPM = "Weekday = 'Sun' AND HourOfDay = '" + n + "PM-" + (n + 1) + "PM'";
                //Debug.WriteLine(sundayPM);

                sundayTraffic = ExcludedDates.Select(sundayPM);

                for (int i = 0; i < sundayTraffic.Length; i++)
                {
                    if (i < sundayTraffic.Length - 1)
                    {
                        total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    }
                    else
                    {
                        total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                        Math.Round(total, MidpointRounding.AwayFromZero);
                        WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                        total = 0;
                        row++;
                    }
                }
            }
            */



            DataRow[] sundayTraffic;
            double total = 0;
            int col = 1;
            int row = 0;


            sundayTraffic = ExcludedDates.Select(sun9AM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun10AM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun11AM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun12PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun1PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun2PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun3PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun4PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun5PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun6PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun7PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun8PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun9PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            sundayTraffic = ExcludedDates.Select(sun10PM);

            for (int i = 0; i < sundayTraffic.Length; i++)
            {
                if (i < sundayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(sundayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }
            #endregion

            #region Monday
            string mon9AM = "WeekDay = 'Mon' AND HourOfDay = '9AM-10AM'";
            string mon10AM = "WeekDay = 'Mon' AND HourOfDay = '10AM-11AM'";
            string mon11AM = "WeekDay = 'Mon' AND HourOfDay = '11AM-12PM'";
            string mon12PM = "WeekDay = 'Mon' AND HourOfDay = '12PM-1PM'";
            string mon1PM = "WeekDay = 'Mon' AND HourOfDay = '1PM-2PM'";
            string mon2PM = "WeekDay = 'Mon' AND HourOfDay = '2PM-3PM'";
            string mon3PM = "WeekDay = 'Mon' AND HourOfDay = '3PM-4PM'";
            string mon4PM = "WeekDay = 'Mon' AND HourOfDay = '4PM-5PM'";
            string mon5PM = "WeekDay = 'Mon' AND HourOfDay = '5PM-6PM'";
            string mon6PM = "WeekDay = 'Mon' AND HourOfDay = '6PM-7PM'";
            string mon7PM = "WeekDay = 'Mon' AND HourOfDay = '7PM-8PM'";
            string mon8PM = "WeekDay = 'Mon' AND HourOfDay = '8PM-9PM'";
            string mon9PM = "WeekDay = 'Mon' AND HourOfDay = '9PM-10PM'";
            string mon10PM = "WeekDay = 'Mon' AND HourOfDay = '10PM-11PM'";

            DataRow[] mondayTraffic;


            mondayTraffic = ExcludedDates.Select(mon9AM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon10AM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon11AM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon12PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon1PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon2PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon3PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon4PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon5PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon6PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon7PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon8PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon9PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            mondayTraffic = ExcludedDates.Select(mon10PM);

            for (int i = 0; i < mondayTraffic.Length; i++)
            {
                if (i < mondayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(mondayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region Tuesday
            string tue9AM = "WeekDay = 'Tue' AND HourOfDay = '9AM-10AM'";
            string tue10AM = "WeekDay = 'Tue' AND HourOfDay = '10AM-11AM'";
            string tue11AM = "WeekDay = 'Tue' AND HourOfDay = '11AM-12PM'";
            string tue12PM = "WeekDay = 'Tue' AND HourOfDay = '12PM-1PM'";
            string tue1PM = "WeekDay = 'Tue' AND HourOfDay = '1PM-2PM'";
            string tue2PM = "WeekDay = 'Tue' AND HourOfDay = '2PM-3PM'";
            string tue3PM = "WeekDay = 'Tue' AND HourOfDay = '3PM-4PM'";
            string tue4PM = "WeekDay = 'Tue' AND HourOfDay = '4PM-5PM'";
            string tue5PM = "WeekDay = 'Tue' AND HourOfDay = '5PM-6PM'";
            string tue6PM = "WeekDay = 'Tue' AND HourOfDay = '6PM-7PM'";
            string tue7PM = "WeekDay = 'Tue' AND HourOfDay = '7PM-8PM'";
            string tue8PM = "WeekDay = 'Tue' AND HourOfDay = '8PM-9PM'";
            string tue9PM = "WeekDay = 'Tue' AND HourOfDay = '9PM-10PM'";
            string tue10PM = "WeekDay = 'Tue' AND HourOfDay = '10PM-11PM'";

            DataRow[] tuesdayTraffic;


            tuesdayTraffic = ExcludedDates.Select(tue9AM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue10AM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue11AM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue12PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue1PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue2PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue3PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue4PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue5PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue6PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue7PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue8PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue9PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            tuesdayTraffic = ExcludedDates.Select(tue10PM);

            for (int i = 0; i < tuesdayTraffic.Length; i++)
            {
                if (i < tuesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(tuesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region Wednesday
            string wed9AM = "WeekDay = 'Wed' AND HourOfDay = '9AM-10AM'";
            string wed10AM = "WeekDay = 'Wed' AND HourOfDay = '10AM-11AM'";
            string wed11AM = "WeekDay = 'Wed' AND HourOfDay = '11AM-12PM'";
            string wed12PM = "WeekDay = 'Wed' AND HourOfDay = '12PM-1PM'";
            string wed1PM = "WeekDay = 'Wed' AND HourOfDay = '1PM-2PM'";
            string wed2PM = "WeekDay = 'Wed' AND HourOfDay = '2PM-3PM'";
            string wed3PM = "WeekDay = 'Wed' AND HourOfDay = '3PM-4PM'";
            string wed4PM = "WeekDay = 'Wed' AND HourOfDay = '4PM-5PM'";
            string wed5PM = "WeekDay = 'Wed' AND HourOfDay = '5PM-6PM'";
            string wed6PM = "WeekDay = 'Wed' AND HourOfDay = '6PM-7PM'";
            string wed7PM = "WeekDay = 'Wed' AND HourOfDay = '7PM-8PM'";
            string wed8PM = "WeekDay = 'Wed' AND HourOfDay = '8PM-9PM'";
            string wed9PM = "WeekDay = 'Wed' AND HourOfDay = '9PM-10PM'";
            string wed10PM = "WeekDay = 'Wed' AND HourOfDay = '10PM-11PM'";

            DataRow[] wednesdayTraffic;


            wednesdayTraffic = ExcludedDates.Select(wed9AM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed10AM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed11AM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed12PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed1PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed2PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed3PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed4PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed5PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed6PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed7PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed8PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed9PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            wednesdayTraffic = ExcludedDates.Select(wed10PM);

            for (int i = 0; i < wednesdayTraffic.Length; i++)
            {
                if (i < wednesdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(wednesdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region Thursday
            string thu9AM = "WeekDay = 'Thu' AND HourOfDay = '9AM-10AM'";
            string thu10AM = "WeekDay = 'Thu' AND HourOfDay = '10AM-11AM'";
            string thu11AM = "WeekDay = 'Thu' AND HourOfDay = '11AM-12PM'";
            string thu12PM = "WeekDay = 'Thu' AND HourOfDay = '12PM-1PM'";
            string thu1PM = "WeekDay = 'Thu' AND HourOfDay = '1PM-2PM'";
            string thu2PM = "WeekDay = 'Thu' AND HourOfDay = '2PM-3PM'";
            string thu3PM = "WeekDay = 'Thu' AND HourOfDay = '3PM-4PM'";
            string thu4PM = "WeekDay = 'Thu' AND HourOfDay = '4PM-5PM'";
            string thu5PM = "WeekDay = 'Thu' AND HourOfDay = '5PM-6PM'";
            string thu6PM = "WeekDay = 'Thu' AND HourOfDay = '6PM-7PM'";
            string thu7PM = "WeekDay = 'Thu' AND HourOfDay = '7PM-8PM'";
            string thu8PM = "WeekDay = 'Thu' AND HourOfDay = '8PM-9PM'";
            string thu9PM = "WeekDay = 'Thu' AND HourOfDay = '9PM-10PM'";
            string thu10PM = "WeekDay = 'Thu' AND HourOfDay = '10PM-11PM'";

            DataRow[] thursdayTraffic;


            thursdayTraffic = ExcludedDates.Select(thu9AM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu10AM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu11AM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu12PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu1PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu2PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu3PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu4PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu5PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu6PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu7PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu8PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu9PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            thursdayTraffic = ExcludedDates.Select(thu10PM);

            for (int i = 0; i < thursdayTraffic.Length; i++)
            {
                if (i < thursdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(thursdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region Friday
            string fri9AM = "WeekDay = 'Fri' AND HourOfDay = '9AM-10AM'";
            string fri10AM = "WeekDay = 'Fri' AND HourOfDay = '10AM-11AM'";
            string fri11AM = "WeekDay = 'Fri' AND HourOfDay = '11AM-12PM'";
            string fri12PM = "WeekDay = 'Fri' AND HourOfDay = '12PM-1PM'";
            string fri1PM = "WeekDay = 'Fri' AND HourOfDay = '1PM-2PM'";
            string fri2PM = "WeekDay = 'Fri' AND HourOfDay = '2PM-3PM'";
            string fri3PM = "WeekDay = 'Fri' AND HourOfDay = '3PM-4PM'";
            string fri4PM = "WeekDay = 'Fri' AND HourOfDay = '4PM-5PM'";
            string fri5PM = "WeekDay = 'Fri' AND HourOfDay = '5PM-6PM'";
            string fri6PM = "WeekDay = 'Fri' AND HourOfDay = '6PM-7PM'";
            string fri7PM = "WeekDay = 'Fri' AND HourOfDay = '7PM-8PM'";
            string fri8PM = "WeekDay = 'Fri' AND HourOfDay = '8PM-9PM'";
            string fri9PM = "WeekDay = 'Fri' AND HourOfDay = '9PM-10PM'";
            string fri10PM = "WeekDay = 'Fri' AND HourOfDay = '10PM-11PM'";

            DataRow[] fridayTraffic;


            fridayTraffic = ExcludedDates.Select(fri9AM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri10AM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri11AM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri12PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri1PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri2PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri3PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri4PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri5PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri6PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri7PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri8PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri9PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            fridayTraffic = ExcludedDates.Select(fri10PM);

            for (int i = 0; i < fridayTraffic.Length; i++)
            {
                if (i < fridayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(fridayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region Saturday
            string sat9AM = "WeekDay = 'Sat' AND HourOfDay = '9AM-10AM'";
            string sat10AM = "WeekDay = 'Sat' AND HourOfDay = '10AM-11AM'";
            string sat11AM = "WeekDay = 'Sat' AND HourOfDay = '11AM-12PM'";
            string sat12PM = "WeekDay = 'Sat' AND HourOfDay = '12PM-1PM'";
            string sat1PM = "WeekDay = 'Sat' AND HourOfDay = '1PM-2PM'";
            string sat2PM = "WeekDay = 'Sat' AND HourOfDay = '2PM-3PM'";
            string sat3PM = "WeekDay = 'Sat' AND HourOfDay = '3PM-4PM'";
            string sat4PM = "WeekDay = 'Sat' AND HourOfDay = '4PM-5PM'";
            string sat5PM = "WeekDay = 'Sat' AND HourOfDay = '5PM-6PM'";
            string sat6PM = "WeekDay = 'Sat' AND HourOfDay = '6PM-7PM'";
            string sat7PM = "WeekDay = 'Sat' AND HourOfDay = '7PM-8PM'";
            string sat8PM = "WeekDay = 'Sat' AND HourOfDay = '8PM-9PM'";
            string sat9PM = "WeekDay = 'Sat' AND HourOfDay = '9PM-10PM'";
            string sat10PM = "WeekDay = 'Sat' AND HourOfDay = '10PM-11PM'";

            DataRow[] saturdayTraffic;


            saturdayTraffic = ExcludedDates.Select(sat9AM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat10AM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat11AM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat12PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat1PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat2PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat3PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat4PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat5PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat6PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat7PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat8PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat9PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row++;
                }
            }

            saturdayTraffic = ExcludedDates.Select(sat10PM);

            for (int i = 0; i < saturdayTraffic.Length; i++)
            {
                if (i < saturdayTraffic.Length - 1)
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                }
                else
                {
                    total += Convert.ToDouble(saturdayTraffic[i]["WTGTraffic"]);
                    Math.Round(total, MidpointRounding.AwayFromZero);
                    WeightedAverageTraffic_perWeek.Rows[row][col] = (Math.Round(total, MidpointRounding.AwayFromZero));
                    total = 0;
                    row = 0;
                    col++;
                }
            }

            #endregion

            #region totals
            int totalCol = 1;

            object sumSunday = WeightedAverageTraffic_perWeek.Compute("Sum(SunWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumSunday;
            totalCol++;
            object sumMonday = WeightedAverageTraffic_perWeek.Compute("Sum(MonWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumMonday;
            totalCol++;
            object sumTuesday = WeightedAverageTraffic_perWeek.Compute("Sum(TueWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumTuesday;
            totalCol++;
            object sumWednesday = WeightedAverageTraffic_perWeek.Compute("Sum(WedWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumWednesday;
            totalCol++;
            object sumThursday = WeightedAverageTraffic_perWeek.Compute("Sum(ThuWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumThursday;
            totalCol++;
            object sumFriday = WeightedAverageTraffic_perWeek.Compute("Sum(FriWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumFriday;
            totalCol++;
            object sumSaturday = WeightedAverageTraffic_perWeek.Compute("Sum(SatWeightedAverage)", "");
            WeightedAverageTraffic_perWeek.Rows[14][totalCol] = sumSaturday;

            foreach (DataRow trow in WeightedAverageTraffic_perWeek.Rows)
            {
                double rowSum = 0;
                foreach (DataColumn tcol in WeightedAverageTraffic_perWeek.Columns)
                {
                    if (!trow.IsNull(tcol))
                    {
                        string stringValue = trow[tcol].ToString();
                        double d;
                        if (double.TryParse(stringValue, out d))
                            rowSum += d;
                    }
                }
                trow.SetField("Total", rowSum);
            }

            #endregion

            // return the new table to the Index
            return WeightedAverageTraffic_perWeek;
        }

        public DataTable Calculate_PercentWeekly()
        {

            /*
             * Calculate each hour's percentage of the weekly
             * total from the weighted average traffic table
             */

            #region table setup
            PercentWeeklyTotal.Columns.Add("HourOfDay", typeof(string));
            PercentWeeklyTotal.Columns.Add("SunPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("MonPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("TuePercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("WedPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("ThuPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("FriPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("SatPercentOfWeeklyTotal", typeof(double));
            PercentWeeklyTotal.Columns.Add("Total", typeof(double));

            PercentWeeklyTotal.Rows.Add("9AM-10AM");
            PercentWeeklyTotal.Rows.Add("10AM-11AM");
            PercentWeeklyTotal.Rows.Add("11AM-12PM");
            PercentWeeklyTotal.Rows.Add("12PM-1PM");
            PercentWeeklyTotal.Rows.Add("1PM-2PM");
            PercentWeeklyTotal.Rows.Add("2PM-3PM");
            PercentWeeklyTotal.Rows.Add("3PM-4PM");
            PercentWeeklyTotal.Rows.Add("4PM-5PM");
            PercentWeeklyTotal.Rows.Add("5PM-6PM");
            PercentWeeklyTotal.Rows.Add("6PM-7PM");
            PercentWeeklyTotal.Rows.Add("7PM-8PM");
            PercentWeeklyTotal.Rows.Add("8PM-9PM");
            PercentWeeklyTotal.Rows.Add("9PM-10PM");
            PercentWeeklyTotal.Rows.Add("10PM-11PM");
            PercentWeeklyTotal.Rows.Add("Total");
            #endregion

            #region calculations
            double weeklyPercent = 0;
            double weeklyTotal = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[14]["Total"]);

            for (int col = 1; col < 8; col++)
            {
                for (int row = 0; row < 14; row++)
                {
                    double hourAvg = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row][col]);

                    weeklyPercent = (hourAvg / weeklyTotal) * 100;
                    PercentWeeklyTotal.Rows[row][col] = Math.Round(weeklyPercent, 1, MidpointRounding.AwayFromZero);
                }
            }
            #endregion

            #region totals
            int totalCol = 1;

            object sumSunday = PercentWeeklyTotal.Compute("Sum(SunPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumSunday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumMonday = PercentWeeklyTotal.Compute("Sum(MonPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumMonday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumTuesday = PercentWeeklyTotal.Compute("Sum(TuePercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumTuesday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumWednesday = PercentWeeklyTotal.Compute("Sum(WedPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumWednesday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumThursday = PercentWeeklyTotal.Compute("Sum(ThuPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumThursday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumFriday = PercentWeeklyTotal.Compute("Sum(FriPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumFriday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumSaturday = PercentWeeklyTotal.Compute("Sum(SatPercentOfWeeklyTotal)", "");
            PercentWeeklyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumSaturday), MidpointRounding.AwayFromZero);

            foreach (DataRow trow in PercentWeeklyTotal.Rows)
            {
                double rowSum = 0;
                foreach (DataColumn tcol in PercentWeeklyTotal.Columns)
                {
                    if (!trow.IsNull(tcol))
                    {
                        string stringValue = trow[tcol].ToString();
                        double d;
                        if (double.TryParse(stringValue, out d))
                            rowSum += d;
                    }
                }
                trow.SetField("Total", Math.Round(rowSum, MidpointRounding.AwayFromZero));
            }

            #endregion

            // return the new table to the Index
            return PercentWeeklyTotal;
        }

        public DataTable Calculate_PercentDaily()
        {

            /*
             * Calculate each hour's percentage of the daily
             * total from the weighted average traffic table
             */

            #region table setup
            PercentDailyTotal.Columns.Add("HourOfDay", typeof(string));
            PercentDailyTotal.Columns.Add("SunPercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("MonPercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("TuePercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("WedPercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("ThuPercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("FriPercentOfDailyTotal", typeof(double));
            PercentDailyTotal.Columns.Add("SatPercentOfDailyTotal", typeof(double));

            PercentDailyTotal.Rows.Add("9AM-10AM");
            PercentDailyTotal.Rows.Add("10AM-11AM");
            PercentDailyTotal.Rows.Add("11AM-12PM");
            PercentDailyTotal.Rows.Add("12PM-1PM");
            PercentDailyTotal.Rows.Add("1PM-2PM");
            PercentDailyTotal.Rows.Add("2PM-3PM");
            PercentDailyTotal.Rows.Add("3PM-4PM");
            PercentDailyTotal.Rows.Add("4PM-5PM");
            PercentDailyTotal.Rows.Add("5PM-6PM");
            PercentDailyTotal.Rows.Add("6PM-7PM");
            PercentDailyTotal.Rows.Add("7PM-8PM");
            PercentDailyTotal.Rows.Add("8PM-9PM");
            PercentDailyTotal.Rows.Add("9PM-10PM");
            PercentDailyTotal.Rows.Add("10PM-11PM");
            PercentDailyTotal.Rows.Add("Total");
            #endregion

            #region calculations
            double dailyPercent = 0;

            for (int col = 1; col < 8; col++)
            {
                double dailyTotal = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[14][col]);

                for (int row = 0; row < 14; row++)
                {
                    double hourAvg = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row][col]);

                    dailyPercent = (hourAvg / dailyTotal) * 100;
                    PercentDailyTotal.Rows[row][col] = Math.Round(dailyPercent, 1, MidpointRounding.AwayFromZero);
                }
            }
            #endregion

            #region Totals
            int totalCol = 1;

            object sumSunday = PercentDailyTotal.Compute("Sum(SunPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumSunday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumMonday = PercentDailyTotal.Compute("Sum(MonPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumMonday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumTuesday = PercentDailyTotal.Compute("Sum(TuePercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumTuesday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumWednesday = PercentDailyTotal.Compute("Sum(WedPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumWednesday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumThursday = PercentDailyTotal.Compute("Sum(ThuPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumThursday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumFriday = PercentDailyTotal.Compute("Sum(FriPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumFriday), MidpointRounding.AwayFromZero);
            totalCol++;
            object sumSaturday = PercentDailyTotal.Compute("Sum(SatPercentOfDailyTotal)", "");
            PercentDailyTotal.Rows[14][totalCol] = Math.Round(Convert.ToDouble(sumSaturday), MidpointRounding.AwayFromZero);

            /*
            foreach (DataRow trow in PercentDailyTotal.Rows)
            {
                double rowSum = 0;
                foreach (DataColumn tcol in PercentDailyTotal.Columns)
                {
                    if (!trow.IsNull(tcol))
                    {
                        string stringValue = trow[tcol].ToString();
                        double d;
                        if (double.TryParse(stringValue, out d))
                            rowSum += d;
                    }
                }
                trow.SetField("Total", Math.Round(rowSum, MidpointRounding.AwayFromZero));
            }
            */
            #endregion

            // return the new table to the Index
            return PercentDailyTotal;
        }

        public DataTable Calculate_AllocatedHours(DataInput input)
        {

            /* Calculate the amount of hours that should be 
             * allocated to each individual day from the 
             * total (payRoll) input and then calculates
             * the amount of employees that should be assigned
             * to each hour
             * 
             * If a minimum/maximum amount of employees is specified
             * the table is adjusted accordingly
             */

            #region table setup
            var sundayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Sunday);
            var mondayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Monday);
            var tuesdayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Tuesday);
            var wednesdayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Wednesday);
            var thursdayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Thursday);
            var fridayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Friday);
            var saturdayDate = DateTime.Today.AddDays(-(int)DateTime.Today.DayOfWeek + (int)DayOfWeek.Saturday);

            AllocatedHours.Columns.Add("HourOfDay", typeof(string));
            AllocatedHours.Columns.Add("Sunday " + Environment.NewLine + String.Format("{0:M/dd}", sundayDate), typeof(double));
            AllocatedHours.Columns.Add("Monday " + Environment.NewLine + String.Format("{0:M/dd}", mondayDate), typeof(double));
            AllocatedHours.Columns.Add("Tuesday " + Environment.NewLine + String.Format("{0:M/dd}", tuesdayDate), typeof(double));
            AllocatedHours.Columns.Add("Wednesday " + Environment.NewLine + String.Format("{0:M/dd}", wednesdayDate), typeof(double));
            AllocatedHours.Columns.Add("Thursday " + Environment.NewLine + String.Format("{0:M/dd}", thursdayDate), typeof(double));
            AllocatedHours.Columns.Add("Friday " + Environment.NewLine + String.Format("{0:M/dd}", fridayDate), typeof(double));
            AllocatedHours.Columns.Add("Saturday " + Environment.NewLine + String.Format("{0:M/dd}", saturdayDate), typeof(double));

            AllocatedHours.Rows.Add("9AM-10AM");
            AllocatedHours.Rows.Add("10AM-11AM");
            AllocatedHours.Rows.Add("11AM-12PM");
            AllocatedHours.Rows.Add("12PM-1PM");
            AllocatedHours.Rows.Add("1PM-2PM");
            AllocatedHours.Rows.Add("2PM-3PM");
            AllocatedHours.Rows.Add("3PM-4PM");
            AllocatedHours.Rows.Add("4PM-5PM");
            AllocatedHours.Rows.Add("5PM-6PM");
            AllocatedHours.Rows.Add("6PM-7PM");
            AllocatedHours.Rows.Add("7PM-8PM");
            AllocatedHours.Rows.Add("8PM-9PM");
            AllocatedHours.Rows.Add("9PM-10PM");
            AllocatedHours.Rows.Add("10PM-11PM");
            #endregion

            #region calculations
            //percent per hour * allocated hours
            double sunAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["SunPercentOfWeeklyTotal"]) / 100));
            double monAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["MonPercentOfWeeklyTotal"]) / 100));
            double tueAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["TuePercentOfWeeklyTotal"]) / 100));
            double wedAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["WedPercentOfWeeklyTotal"]) / 100));
            double thuAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["ThuPercentOfWeeklyTotal"]) / 100));
            double friAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["FriPercentOfWeeklyTotal"]) / 100));
            double satAllocatedHours = Math.Round(input.payRoll * (Convert.ToDouble(PercentWeeklyTotal.Rows[14]["SatPercentOfWeeklyTotal"]) / 100));

            // usually off by 1 hour
            // print to debug to see values for testing
            //System.Diagnostics.Debug.WriteLine(sunAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(monAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(tueAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(wedAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(thuAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(friAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(satAllocatedHours);
            //System.Diagnostics.Debug.WriteLine(sunAllocatedHours + monAllocatedHours + tueAllocatedHours + wedAllocatedHours + thuAllocatedHours +
            //                                   friAllocatedHours + satAllocatedHours);



            for (int col = 1; col < 8; col++)
            {
                if (col == 1)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * sunAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 2)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * monAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 3)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * tueAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 4)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * wedAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 5)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * thuAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 6)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * friAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
                else if (col == 7)
                {
                    for (int row = 0; row < 14; row++)
                    {
                        double hourPercentage = (Convert.ToDouble(PercentDailyTotal.Rows[row][col]) / 100);
                        double employees = hourPercentage * satAllocatedHours;

                        if (employees < input.minEmp)
                        {
                            employees = input.minEmp;
                        }
                        if (employees > input.maxEmp)
                        {
                            employees = input.maxEmp;
                        }
                        AllocatedHours.Rows[row][col] = Math.Round(employees, 0);
                    }
                }
            }
            #endregion

            return AllocatedHours;
        }

        public DataTable Calculate_PowerHourForecast()
        {

            /*
             * Determines the highest trafficked hours
             * for the day based on 3 hour periods
             */

            #region table setup
            PowerHourForecast.Columns.Add("HourOfDay", typeof(string));
            PowerHourForecast.Columns.Add("SunPowerHours", typeof(double));
            PowerHourForecast.Columns.Add("MonPowerHours", typeof(double));
            PowerHourForecast.Columns.Add("TuePowerHours", typeof(double));
            PowerHourForecast.Columns.Add("WedPowerHours", typeof(double));
            PowerHourForecast.Columns.Add("ThuPowerHours", typeof(double));
            PowerHourForecast.Columns.Add("FriPowerHours", typeof(double));
            PowerHourForecast.Columns.Add("SatPowerHours", typeof(double));

            PowerHourForecast.Rows.Add("9AM-10AM");
            PowerHourForecast.Rows.Add("10AM-11AM");
            PowerHourForecast.Rows.Add("11AM-12PM");
            PowerHourForecast.Rows.Add("12PM-1PM");
            PowerHourForecast.Rows.Add("1PM-2PM");
            PowerHourForecast.Rows.Add("2PM-3PM");
            PowerHourForecast.Rows.Add("3PM-4PM");
            PowerHourForecast.Rows.Add("4PM-5PM");
            PowerHourForecast.Rows.Add("5PM-6PM");
            PowerHourForecast.Rows.Add("6PM-7PM");
            PowerHourForecast.Rows.Add("7PM-8PM");
            PowerHourForecast.Rows.Add("8PM-9PM");
            PowerHourForecast.Rows.Add("9PM-10PM");
            PowerHourForecast.Rows.Add("10PM-11PM");
            #endregion

            #region calculations
            double forecast;
            double[] dayHighs = new double[7];
            double[] dayLows = new double[21];
            int n = 0;

            for (int col = 1; col < 8; col++)
            {
                double[] dayNumbers = new double[20];

                for (int row = 0; row < 14; row++)
                {

                    if (row < 12)
                    {
                        double currentHour = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row][col]);
                        forecast = currentHour + Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row + 1][col]) + Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row + 2][col]);

                        dayNumbers[row] = forecast;
                    }
                    else if (row == 12)
                    {
                        double currentHour = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row][col]);
                        forecast = currentHour + Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row + 1][col]);

                        dayNumbers[row] = forecast;
                    }
                    else
                    {
                        double currentHour = Convert.ToDouble(WeightedAverageTraffic_perWeek.Rows[row][col]);
                        forecast = currentHour;

                        dayNumbers[row] = forecast;
                    }

                    PowerHourForecast.Rows[row][col] = Math.Round(forecast, 0);

                }

                dayHighs[col - 1] = dayNumbers.Max();

                n += 3;
                int smallest = 9999;
                int second = 9999;
                int third = 9999;

                foreach (int i in dayNumbers)
                {
                    if (i < smallest && i != 0)
                    {
                        third = second;
                        second = smallest;
                        smallest = i;
                    }
                    else if (i < second && i != 0)
                    {
                        third = second;
                        second = i;
                    }
                    else if (i < third && i != 0)
                    {
                        third = i;
                    }
                }
                dayLows[n - 3] = smallest;
                dayLows[n - 2] = second;
                dayLows[n - 1] = third;

            }
            #endregion

            #region create variables for lows/highs
            // max hours for each day
            double sunHigh = dayHighs[0];
            double monHigh = dayHighs[1];
            double tueHigh = dayHighs[2];
            double wedHigh = dayHighs[3];
            double thuHigh = dayHighs[4];
            double friHigh = dayHighs[5];
            double satHigh = dayHighs[6];

            // min hours for each day (currently excluding zeros)
            double sunLow = dayLows[0];
            double sunLow2 = dayLows[1];
            double sunLow3 = dayLows[2];
            double monLow = dayLows[3];
            double monLow2 = dayLows[4];
            double monLow3 = dayLows[5];
            double tueLow = dayLows[6];
            double tueLow2 = dayLows[7];
            double tueLow3 = dayLows[8];
            double wedLow = dayLows[9];
            double wedLow2 = dayLows[10];
            double wedLow3 = dayLows[11];
            double thuLow = dayLows[12];
            double thuLow2 = dayLows[13];
            double thuLow3 = dayLows[14];
            double friLow = dayLows[15];
            double friLow2 = dayLows[16];
            double friLow3 = dayLows[17];
            double satLow = dayLows[18];
            double satLow2 = dayLows[19];
            double satLow3 = dayLows[20];

            //Debug.WriteLine(sunLow);
            //Debug.WriteLine(sunLow2);
            //Debug.WriteLine(sunLow3);
            //Debug.WriteLine(monLow);
            //Debug.WriteLine(monLow2);
            //Debug.WriteLine(monLow3);
            //Debug.WriteLine(tueLow);
            //Debug.WriteLine(tueLow2);
            //Debug.WriteLine(tueLow3);
            //Debug.WriteLine(wedLow);
            //Debug.WriteLine(wedLow2);
            //Debug.WriteLine(wedLow3);
            //Debug.WriteLine(thuLow);
            //Debug.WriteLine(thuLow2);
            //Debug.WriteLine(thuLow3);
            //Debug.WriteLine(friLow);
            //Debug.WriteLine(friLow2);
            //Debug.WriteLine(friLow3);
            //Debug.WriteLine(satLow);
            //Debug.WriteLine(satLow2);
            //Debug.WriteLine(satLow3);
            #endregion

            return PowerHourForecast;
        }


        // defaults
        public ActionResult Employee()
        {
            return View();
        }

        public ActionResult Contact()
        {
            return View();
        }
    }
}
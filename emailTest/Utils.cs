﻿using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

namespace Anko
{
    class Utils
    {
        private const string RESULTS_DIRECROTY_NAME = "TempResults";

        public static string resultsDirectoryPath = string.Empty;

        // function deletes 'old' results folder with all it's files and creates one new
        // results folder is needed for auto generated excel files for mail attachments
        public static void createResultsFolder()
        {
            // create results folder for the temp filtered excel files
            resultsDirectoryPath = Path.Combine(Directory.GetCurrentDirectory(), RESULTS_DIRECROTY_NAME);

            try
            {
                if (Directory.Exists(resultsDirectoryPath))
                {
                    Directory.Delete(resultsDirectoryPath, true);
                }

                Directory.CreateDirectory(resultsDirectoryPath);
            }
            catch (Exception e)
            {
                // fatal: cannot continue
                OrdersParser._Form.log(string.Format("Failed to create/delete results folder {0}. Error: {1}", resultsDirectoryPath, e.Message), OrdersParser.LogLevel.Error);
                return;
            }
        }

        // function kills all processes of a given name
        public static void killProcess(string processName)
        {
            foreach (var process in Process.GetProcessesByName(processName))
            {
                try
                {
                    process.Kill();
                }
                catch(Exception e)
                {
                    OrdersParser._Form.log(string.Format("Failed to kill process {0}. Error: {1}", processName, e.Message), OrdersParser.LogLevel.Error);
                }
            }
        }

        // function translates date as string into DateTime IL format
        public static DateTime getDateFromDynamicSheet(dynamic value)
        {
            DateTime val = DateTime.MinValue;

            try
            {
                // note: IL format
                val = DateTime.ParseExact(value, Common.DATE_FORMAT, CultureInfo.InvariantCulture);
            }
            catch (Exception e)
            {
                // empty cell or couldn't parse
                OrdersParser._Form.log(string.Format("Failed to parse date. val: {0}, error: {1}", value, e.Message), OrdersParser.LogLevel.Error);
            };

            return val;
        }

        // function parses returns string for any read value
        // in case of null, empty string is returned
        public static string getStringFromDynamicSheet(string value)
        {
            string val = string.Empty;

            try
            {
                val = value.Trim();
            }
            catch
            {
                // empty cell
            };

            return val;
        }

        // function converts column name (letter) to index
        // e.g. A -> 1, B -> 2, ..
        public static int getIndexFromColumnChar(char column)
        {
            return (Convert.ToInt32(column) - Convert.ToInt32('A') + 1);
        }

        // function generates table (for excel) from a given list of type
        public static object[,] generateObjectFromList<T>(List<T> resultList, out int rows, out int cols)
        {
            FieldInfo[] fields = typeof(T).GetFields(BindingFlags.Public | BindingFlags.Instance);

            // fill out parameters
            // table has one additional row for titles
            rows = resultList.Count + 1;
            cols = fields.Length;

            object[,] table = new object[rows, cols];
            string str = string.Empty;
            int rowOffset = 1;
            int row = 0;

            // fill the column names in the first row
            for (int col = 0; col < fields.Length; col++)
            {
                // setting column names as fields names
                table[0, col] = fields[col].Name;
            }

            // start from the second row
            foreach (T res in resultList)
            {
                for (int col = 0; col < fields.Length; col++)
                {
                    // for DateTime, leave only time in IL format
                    if (fields[col].FieldType.Equals(typeof(DateTime)))
                    {
                        // date format must be with dots, otherwise it gets messy in excel
                        table[row + rowOffset, col] = ((DateTime)(fields[col].GetValue(res))).ToString("dd.MM.yyyy");
                        continue;
                    }

                    // replace 'null' with empty string
                    if ((fields[col].FieldType.Equals(typeof(String))) && (string.IsNullOrEmpty((string)fields[col].GetValue(res)) == true))
                    {
                        table[row + rowOffset, col] = string.Empty;
                        continue;
                    }

                    // else
                    table[row + rowOffset, col] = fields[col].GetValue(res);
                }

                // go to next entry
                row++;
            }

            return table;
        }

        // function generates DataTable from a given list of type
        public static DataTable generateDataTableFromList<T>(List<T> resultList)
        {
            FieldInfo[] fields  = typeof(T).GetFields(BindingFlags.Public | BindingFlags.Instance);
            DataTable   table   = new DataTable();

            // fill the column names in the first row
            for (int col = 0; col < fields.Length; col++)
            {
                // setting column names as fields names
                table.Columns.Add(fields[col].Name);
            }

            // go over all the values in the list
            foreach (T res in resultList)
            {
                var values = new object[fields.Length];

                for (int col = 0; col < fields.Length; col++)
                {
                    // for DateTime, leave only time in IL format
                    if (fields[col].FieldType.Equals(typeof(DateTime)))
                    {
                        // date format must be with dots, otherwise it gets messy in excel
                        values[col] = ((DateTime)(fields[col].GetValue(res))).ToString("dd.MM.yyyy");
                        continue;
                    }

                    // replace 'null' with empty string
                    if ((fields[col].FieldType.Equals(typeof(String))) && (string.IsNullOrEmpty((string)fields[col].GetValue(res)) == true))
                    {
                        values[col] = string.Empty;
                        continue;
                    }

                    // else
                    values[col] = fields[col].GetValue(res);
                }

                // add the value
                table.Rows.Add(values);
            }

            return table;
        }

        // function gets a text and replaces all the parameters in the following format:
        // ({param1}, {param2} ..) by values from a dictionary
        // assumption: dictionary has values for all provided parameters
        public static string extractParameterFromDictionary(string text, Dictionary<string, string> bodyParameters)
        {
            return Regex.Replace(text,
                                 @"\{(\w+)\}", // replaces any text surrounded by { and }
                                 m =>
                                 {
                                     string value;
                                     return bodyParameters.TryGetValue(m.Groups[1].Value, out value) ? value : "null";
                                 });
        }

        // function return the embedded resource name based on the mail type
        // each mail type has body template saved as embedded resource
        public static string getResourceNameFromMailType(Common.MailType mailType)
        {
            switch (mailType)
            {
                case Common.MailType.Reports:
                    {
                        return Anko.Properties.Resources.OrdersMail;
                    }
                case Common.MailType.LoadingConfirmation:
                    {
                        return Anko.Properties.Resources.LoadingConfirmationMail;
                    }
                case Common.MailType.BookingConfirmation:
                    {
                        return Anko.Properties.Resources.BookingConfirmationMail;
                    }
                case Common.MailType.DocumentsReceipts:
                    {
                        return Anko.Properties.Resources.DocumentsReceipts;
                    }
                default:
                    OrdersParser._Form.log("Unrecognized mail type - cannot find mail template", OrdersParser.LogLevel.Error);
                    return string.Empty;
            }
        }

        // function checks whether text is HEB or EN
        // one char is enough for a text to be HEB
        public static bool isHebrewText(string text)
        {
            char firstHebChar = (char)1488; // א
            char lastHebChar  = (char)1514; // ת

            foreach (char c in text.ToCharArray())
            {
                if (c >= firstHebChar && c <= lastHebChar) return true;
            }

            return false;
        }

        // function capitalizes the first char - for better tables representation
        public static string uppercaseFirst(string text)
        {
            if (string.IsNullOrEmpty(text) == true)
            {
                return string.Empty;
            }

            // return char and concat substring
            return char.ToUpper(text[0]) + text.Substring(1);
        }

        // function extracts all tables from HTML page into a DataSet
        // courtesy of http://www.c-sharpcorner.com/code/3719/convert-html-tables-to-dataset-in-c-sharp.aspx
        public static DataSet convertHTMLTablesToDataSet(string HTML)
        {
            DataSet     ds                  = new DataSet();
            DataTable   dt                  = null;
            DataRow     dr                  = null;
            string      tableExpression     = "<TABLE[^>]*>(.*?)</TABLE>";
            string      headerExpression    = "<TH[^>]*>(.*?)</TH>";
            string      rowExpression       = "<TR[^>]*>(.*?)</TR>";
            string      columnExpression    = "<TD[^>]*>(.*?)</TD>";
            bool        bHeadersExist        = false;
            int         iCurrentColumn      = 0;
            int         iCurrentRow         = 0;
            string      val                 = string.Empty;

            // get a match for all the tables in the HTML
            MatchCollection Tables = Regex.Matches(HTML, tableExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

            // loop through each table element
            foreach (Match Table in Tables)
            {
                // reset the current row counter and the header flag
                iCurrentRow = 0;
                bHeadersExist = false;

                // add a new table to the DataSet
                dt = new DataTable();

                // create the relevant amount of columns for this table (use the headers if they exist, otherwise use default names)
                if (Table.Value.ToUpper().Contains("<TH"))
                {
                    // set the HeadersExist flag
                    bHeadersExist = true;

                    // get a match for all the rows in the table
                    MatchCollection Headers = Regex.Matches(Table.Value, headerExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                    // loop through each header element
                    foreach (Match Header in Headers)
                    {
                        // remove HTML formating, replace <br> by new line
                        val = Header.Groups[1].ToString();
                        val = Regex.Replace(val, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
                        val = Regex.Replace(val, "<.*?>", String.Empty);
                        dt.Columns.Add(val.Trim());
                    }
                }

                if (dt.Columns.Count == 0)
                {
                    // failed to find columns at all - move on
                    continue;
                }

                // get a match for all the rows in the table
                MatchCollection Rows = Regex.Matches(Table.Value, rowExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                // loop through each row element
                foreach (Match Row in Rows)
                {
                    // only loop through the row if it isn't a header row
                    if (!(iCurrentRow == 0 && bHeadersExist))
                    {
                        // create a new row and reset the current column counter
                        dr = dt.NewRow();
                        iCurrentColumn = 0;
                        
                        // get a match for all the columns in the row
                        MatchCollection Columns = Regex.Matches(Row.Value, columnExpression, RegexOptions.Multiline | RegexOptions.Singleline | RegexOptions.IgnoreCase);

                        // loop through each column element
                        foreach (Match Column in Columns)
                        {
                            // add the value to the DataRow
                            val = Column.Groups[1].ToString();
                            val = Regex.Replace(val, "<br>", Environment.NewLine, RegexOptions.IgnoreCase);
                            val = Regex.Replace(val, "<.*?>", String.Empty);
                            dr[iCurrentColumn] = val.Trim();

                            // increase the current column
                            iCurrentColumn++;
                        }

                        // add the DataRow to the DataTable
                        dt.Rows.Add(dr);
                    }

                    // increase the current row counter
                    iCurrentRow++;
                }

                // add the DataTable to the DataSet
                ds.Tables.Add(dt);
            }
            return ds;
        }

        // function verifies is there are arrivals to certain ports
        public static bool bArrivalsToPort(List<Common.Order> resultList, PortService.PortName portName)
        {
            return resultList.Any(x => x.toPlace.ToLower() == portName.ToString().ToLower());
        }

        // function returns true if orders excel file was parsed successfully
        public static bool bValidOrders()
        {
            return ((Common.orderList != null) && (Common.orderList.Count() > 0));
        }

        // function gets the desired destination port for a specific consignee
        // this is done by finding the consingee name in the customers internal DB
        // and fetching from there the required destination port
        public static PortService.PortName getDestinationPortByConsignee(string consignee)
        {
            // most of the customers having long names, therefore, 'contains' is enough
            // while, some customers having short name (e.g. bg), and 'contains' is useless
            // for customers with name of 2 chars, try 'starts with' or starts with dots e.g. b.g.
            foreach (Common.Customer customer in Common.customerList)
            {
                if (customer.name.Length > 2)
                {
                    if (consignee.ToLower().Contains(customer.name))
                    {
                        return customer.destinationPort;
                    }
                }
                else
                {
                    // generate the name with dots between the chars
                    string nameWithDots = string.Join(".", customer.name.ToCharArray()) + ".";

                    // try 'starts with'
                    if (consignee.ToLower().StartsWith(customer.name) || 
                        consignee.ToLower().StartsWith(nameWithDots))
                    {
                        return customer.destinationPort;
                    }
                }
            }

            return PortService.PortName.Unknown;
        }

        // function retuns true if strings equal (neglecting case sensitive and spaces)
        public static bool strCmp(string str1, string str2)
        {
            return (str1.Trim().Equals(str2.Trim(), StringComparison.InvariantCultureIgnoreCase));
        }
    }
}

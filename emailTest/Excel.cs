﻿using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using ExcelInerop = Microsoft.Office.Interop.Excel;

namespace Anko
{
    // class has no constructor and has static methods since should not have instances
    // use as singleton
    class Excel
    {
        // function initializes the class (should be called once at program init)
        public static void init()
        {
            // this can take several seconds
            //OrdersParser._Form.log("Create excel instances - please wait");

            //killExcel();

            // open once excel instance for all the project
            try
            {
                excelApp = new ExcelInerop.Application();
                // for performance optimization
                excelApp.ScreenUpdating = false;
            }
            catch (Exception e)
            {
                // fatal: cannot continue
                OrdersParser._Form.log(string.Format("Failed to open excel instance. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                OrdersParser._Form.log(string.Format("Try to restart your PC or application"), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }
        }

        // function destructs the class (should be called once at program halt)
        public static void dispose()
        {
            // to be on the safe side, function can be called several times
            // as error handling, therefore, some objects might be destroyed twice
            try
            {
                // this doesn't kill all COM objects
                // therefore, we have kill process afterwards to clean all resources
                excelApp.Quit();
                excelApp = null;
                excelSheet = null;
                workBook.Close();
                workBook = null;
            }
            catch { }

            killExcel();
        }

        // excel obj global variables
        private static ExcelInerop.Worksheet   excelSheet   = null;
        private static ExcelInerop.Workbook    workBook     = null;
        private static ExcelInerop.Application excelApp     = null;

        // internal excel parameters
        private const string CONFIGURATION_EXCEL_NAME   = "Configuration.xlsx";
        private const string CUSTOMERS_SHEET_NAME       = "Customers";
        private const string TANKO_DETAILS_SHEET_NAME   = "TankoExcelParams";
        private const string BOOKING_SHEE_NAME          = "Booking";
        private const string AGENTS_SHEE_NAME           = "Agents";

        // function kills all processes of excel
        private static void killExcel()
        {
            Utils.killProcess("excel");
        }

        // function parses internal excel DB and extracts all needed data from all sheets
        public static void getDetailsFromLocalDb()
        {
            string configurationExcelPath = Path.Combine(Directory.GetCurrentDirectory(), CONFIGURATION_EXCEL_NAME);

            try
            {
                workBook = excelApp.Workbooks.Open(configurationExcelPath);
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel file in: {0}. Error: {1}", configurationExcelPath, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // parse excel
            getCustomersDetails();
            getTankoExcelParameters();
            getShippingCompaniesDetais();
            getAgentsDetails();

            // close file instance
            workBook.Close();
            workBook = null;
        }

        // function extracts customer list from static DB (excel file)
        private static void getCustomersDetails()
        {
            // go to specific sheet
            try
            {
                excelSheet = workBook.Sheets[CUSTOMERS_SHEET_NAME];
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel sheet: {0}. Error: {1}", CUSTOMERS_SHEET_NAME, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // go over the whole table filling the customerLust
            // start from 3, since 
            // dynamic array starts from [1,1]
            // row 1 is full of nulls
            // row 2 has the column names
            string val = string.Empty;
            int col = 1;
            int effectiveDataOffset = 3;
            Common.Customer customer = new Common.Customer();
            Common.customerList = new List<Common.Customer>();
            int totalNumOfRows = excelSheet.UsedRange.Rows.Count - 2;
            dynamic sheet = excelSheet.UsedRange.Value2;


            try
            {
                for (int row = 0; row < totalNumOfRows; row++)
                {
                    // name
                    val = sheet[row + effectiveDataOffset, col];
                    if (string.IsNullOrEmpty(val) == false) customer.name = val;

                    // shipper/alias (if exists)
                    val = sheet[row + effectiveDataOffset, col+ 1];
                    if (string.IsNullOrEmpty(val) == false) customer.alias = val;

                    // to
                    val = sheet[row + effectiveDataOffset, col + 2];
                    if (string.IsNullOrEmpty(val) == false) customer.to.Add(val);

                    // cc
                    val = sheet[row + effectiveDataOffset, col + 3];
                    if (string.IsNullOrEmpty(val) == false) customer.cc.Add(val);

                    // report needed?
                    val = sheet[row + effectiveDataOffset, col + 4];
                    if (string.IsNullOrEmpty(val) == false)
                    {
                        if (val.Trim().ToLower().Equals("yes"))
                        customer.bSendReport = true;
                    }

                    // destination port
                    val = sheet[row + effectiveDataOffset, col + 5];
                    if (string.IsNullOrEmpty(val) == false)
                    {
                        // destination port found
                        if (val.Trim().ToLower().Equals("ashdod"))
                        {
                            customer.destinationPort = PortService.PortName.Ashdod;
                        }

                        if (val.Trim().ToLower().Equals("haifa"))
                        {
                            customer.destinationPort = PortService.PortName.Haifa;
                        }
                    }

                    // check that last entry for this customer (next should be not empty or out of bounds)
                    if ((row == (totalNumOfRows - 1)) || (string.IsNullOrEmpty(sheet[row + effectiveDataOffset + 1, col]) == false))
                    {
                        // this is the last customer, add to list and zero struct
                        Common.customerList.Add(customer);

                        // nullify and create new
                        customer = null;
                        customer = new Common.Customer();
                    }
                }
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed parsing private DB. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // success
            OrdersParser._Form.log(string.Format("Found {0} customers in the private DB", Common.customerList.Count));
            OrdersParser._Form.log(string.Format("Need to send reports to {0} customers", Common.customerList.Count(x => x.bSendReport == true)));

            // update global list outside of this APP domain
            ExcelRemote.RemoteExeclController.DataUpdater.updateCustomerList(Common.customerList);

            excelSheet = null;
        }

        // function extracts agents list from static DB (excel file)
        private static void getAgentsDetails()
        {
            // got to specific sheet
            try
            {
                excelSheet = workBook.Sheets[AGENTS_SHEE_NAME];
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel sheet: {0}. Error: {1}", AGENTS_SHEE_NAME, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // go over the whole table filling the customerLust
            // start from 3, since 
            // dynamic array starts from [1,1]
            // row 1 is full of nulls
            // row 2 has the column names
            string val = string.Empty;
            int col = 1;
            int effectiveDataOffset = 3;
            Common.Agent agent = new Common.Agent();
            Common.agentList = new List<Common.Agent>();
            int totalNumOfRows = excelSheet.UsedRange.Rows.Count - 2;
            dynamic sheet = excelSheet.UsedRange.Value2;

            try
            {
                for (int row = 0; row < totalNumOfRows; row++)
                {
                    // name
                    val = sheet[row + effectiveDataOffset, col];
                    if (string.IsNullOrEmpty(val) == false) agent.name = val;

                    // to
                    val = sheet[row + effectiveDataOffset, col + 1];
                    if (string.IsNullOrEmpty(val) == false) agent.to.Add(val);

                    // cc
                    val = sheet[row + effectiveDataOffset, col + 2];
                    if (string.IsNullOrEmpty(val) == false) agent.cc.Add(val);

                    // country
                    val = sheet[row + effectiveDataOffset, col + 3];
                    if (string.IsNullOrEmpty(val) == false) agent.countries.Add(val);

                    // check that last entry for this customer (next should be not empty or out of bounds)
                    if ((row == (totalNumOfRows - 1)) || (string.IsNullOrEmpty(sheet[row + effectiveDataOffset + 1, col]) == false))
                    {
                        // this is the last customer, add to list and zero struct
                        Common.agentList.Add(agent);

                        // nullify and create new
                        agent = null;
                        agent = new Common.Agent();
                    }
                }
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed parsing private DB. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // success
            OrdersParser._Form.log(string.Format("Found {0} agents in the private DB", Common.agentList.Count));

            // update global list outside of this APP domain
            ExcelRemote.RemoteExeclController.DataUpdater.updateAgentList(Common.agentList);

            excelSheet = null;
        }

        // function extracts shipping companies (booking) list from static DB (excel file)
        private static void getShippingCompaniesDetais()
        {
            // go to specific sheet
            try
            {
                excelSheet = workBook.Sheets[BOOKING_SHEE_NAME];
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel sheet: {0}. Error: {1}", BOOKING_SHEE_NAME, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // go over the whole table filling the customerLust
            // start from 3, since 
            // dynamic array starts from [1,1]
            // row 1 is full of nulls
            // row 2 has the column names
            string val = string.Empty;
            int col = 1;
            int effectiveDataOffset = 3;
            Common.ShippingCompany shippingCompany = new Common.ShippingCompany();
            Common.shippingCompanyList = new List<Common.ShippingCompany>();
            int totalNumOfRows = excelSheet.UsedRange.Rows.Count - 2;
            dynamic sheet = excelSheet.UsedRange.Value2;

            try
            {
                for (int row = 0; row < totalNumOfRows; row++)
                {
                    // shipping line
                    val = sheet[row + effectiveDataOffset, col];
                    if (string.IsNullOrEmpty(val) == false) shippingCompany.shippingLine = val;

                    // id
                    val = sheet[row + effectiveDataOffset, col + 1];
                    if (string.IsNullOrEmpty(val) == false) shippingCompany.id = val;

                    // agent name
                    val = sheet[row + effectiveDataOffset, col + 2];
                    if (string.IsNullOrEmpty(val) == false) shippingCompany.name = val;

                    // to
                    val = sheet[row + effectiveDataOffset, col + 3];
                    if (string.IsNullOrEmpty(val) == false) shippingCompany.to.Add(val);

                    // cc
                    val = sheet[row + effectiveDataOffset, col + 4];
                    if (string.IsNullOrEmpty(val) == false) shippingCompany.cc.Add(val);

                    // check that last entry for this customer (next should be not empty or out of bounds)
                    if ((row == (totalNumOfRows - 1)) || (string.IsNullOrEmpty(sheet[row + effectiveDataOffset + 1, col]) == false))
                    {
                        // this is the last customer, add to list and zero struct
                        Common.shippingCompanyList.Add(shippingCompany);

                        // nullify and create new
                        shippingCompany = null;
                        shippingCompany = new Common.ShippingCompany();
                    }
                }
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed parsing private DB. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // success
            OrdersParser._Form.log(string.Format("Found {0} shipping companies in the private DB", Common.shippingCompanyList.Count));

            // update global list outside of this APP domain
            ExcelRemote.RemoteExeclController.DataUpdater.updateShippingCompanyList(Common.shippingCompanyList);

            excelSheet = null;
        }

        // function extracts the Tanko excel parameters such as
        // file name and sender email
        private static void getTankoExcelParameters()
        {
            // go to specific sheet
            try
            {
                excelSheet = workBook.Sheets[TANKO_DETAILS_SHEET_NAME];
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel sheet: {0}. Error: {1}", TANKO_DETAILS_SHEET_NAME, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // 2D table with parameters and values
            // start from 3, since dynamic array starts from [1,1]
            // row 1 is full of nulls
            // row 2 has the column names
            string val = string.Empty;
            int col = 1;
            int effectiveDataOffset = 3;
            int totalNumOfRows = excelSheet.UsedRange.Rows.Count - 2;
            dynamic sheet = excelSheet.UsedRange.Value2;

            try
            {
                // sender email address
                val = sheet[effectiveDataOffset, col];
                if (string.IsNullOrEmpty(val) == false)
                {
                    if (val == "emailAddress") Outlook.tancoOrdersEmail = sheet[effectiveDataOffset, col + 1];
                }

                // excel Tanko file name
                val = sheet[effectiveDataOffset + 1, col];
                if (string.IsNullOrEmpty(val) == false)
                {
                    if (val == "fileName") Outlook.tancoOrdersFileName = sheet[effectiveDataOffset + 1, col + 1];
                }
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed parsing private DB. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // success
            OrdersParser._Form.log(string.Format("Tanko excel file name: {0}", Outlook.tancoOrdersFileName));
            OrdersParser._Form.log(string.Format("Sender email: {0}", Outlook.tancoOrdersEmail));

            excelSheet = null;
        }

        // function parses order details from a given DB (excel file)
        public static void getOrderDetails()
        {
            // sanity check that excel file was parsed successfully
            if (File.Exists(Common.plannedImportExcel) == false)
            {
                // file doesn't exist
                OrdersParser._Form.log(string.Format("Orders file is not found at: {0}", Common.plannedImportExcel), OrdersParser.LogLevel.Error);
                return;
            }

            try
            {
                workBook = excelApp.Workbooks.Open(Common.plannedImportExcel);
                excelSheet = workBook.ActiveSheet;
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open excel file in: {0}. Error: {1}", Common.plannedImportExcel, e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            OrdersParser._Form.log(string.Format("Excel file was successfully loaded {0}", Common.plannedImportExcel));

            Common.orderList = new List<Common.Order>();
            DateTime start = DateTime.Now;
            int totalNumOfRows = excelSheet.UsedRange.Rows.Count;
            dynamic sheet = excelSheet.UsedRange.Value2;

            //  need to skip the first 3 rows - hardcoded (static format)
            int effectiveDataStart = 4;

            try
            {
                // go over the whole table and fill orderList
                // no filtering at this point
                for (int row = effectiveDataStart; row <= totalNumOfRows; row++)
                {
                    Common.Order order = new Common.Order();

                    // integers
                    order.jobNo          = Convert.ToInt32(sheet[row, Utils.getIndexFromColumnChar('A')]);
                    order.consignmentNum = Convert.ToInt32(sheet[row, Utils.getIndexFromColumnChar('B')]);

                    // sanity check - job number must be valid
                    if (order.jobNo == 0)
                    {
                        // invalid - don't proceed with the parse
                        continue;
                    }

                    // strings
                    order.customer       = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('C')]);
                    order.shipper        = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('D')]);
                    order.consignee      = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('E')]);
                    order.customerRef    = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('F')]);
                    order.tankNum        = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('H')]);
                    order.activity       = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('I')]);
                    order.fromCountry    = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('K')]);
                    order.fromPlace      = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('M')]);
                    order.toCountry      = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('O')]);
                    order.toPlace        = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('P')]);
                    order.productName    = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('R')]);
                    order.vessel         = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('S')]);
                    order.voyage         = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('T')]);
                    order.MBL            = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('U')]);
                    order.arrivalStatus  = Utils.getStringFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('V')]);

                    // dates
                    order.loadingDate    = Utils.getDateFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('J')]);
                    order.sailingDate    = Utils.getDateFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('N')]);
                    order.arrivalDate    = Utils.getDateFromDynamicSheet(sheet[row, Utils.getIndexFromColumnChar('Q')]);

                    // order passed all criteria - add to list
                    Common.orderList.Add(order);
                    order = null;
                }
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to parse excel. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            OrdersParser._Form.log(string.Format("Parsed {0} records from excel", Common.orderList.Count));

            // update global list outside of this APP domain
            ExcelRemote.RemoteExeclController.DataUpdater.updateOrderList(Common.orderList);

            // close file instance
            excelSheet = null;
            workBook.Close();
            workBook = null;
        }

        // function generates excel file with specific customer orders
        public static void generateCustomerFile(dynamic valuesArray, int rows, int cols, Common.Customer customer, string outputFileName)
        {
            // can be null in the first load
            if (excelApp == null)
            {
                init();
            }

            try
            {
                workBook = excelApp.Workbooks.Add();
                excelSheet = workBook.ActiveSheet;
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to add new workbook into excel. Error: {0}", e.Message), OrdersParser.LogLevel.Error);
                dispose();
                return;
            }

            // alignment
            excelApp.DefaultSheetDirection = (int)ExcelInerop.Constants.xlLTR;
            excelSheet.DisplayRightToLeft = false;

            // fill data
            fillExcelValuesFromArray(valuesArray, rows, cols);

            // sheet design
            designSheet(rows, cols, customer);

            // save to file
            workBook.SaveAs(outputFileName,                                             // FileName
                            ExcelInerop.XlFileFormat.xlWorkbookDefault,                 // FileFormat
                            Type.Missing,                                               // Password
                            Type.Missing,                                               // WriteResPassword
                            false,                                                      // ReadOnlyRecommended
                            false,                                                      // CreateBackup
                            ExcelInerop.XlSaveAsAccessMode.xlNoChange,                  // AccessMode
                            ExcelInerop.XlSaveConflictResolution.xlLocalSessionChanges, // ConflictResolution
                            Type.Missing,                                               // AddToMru
                            Type.Missing,                                               // TextCodepage
                            Type.Missing,                                               // TextVisualLayout
                            false);                                                     // Local

            workBook.Close();
        }

        // function designs the sheet
        private static void designSheet(int rows, int cols, Common.Customer customer)
        {
            ExcelInerop.Range range = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rows, cols]];

            // borders
            ExcelInerop.Borders border = range.Borders;
            border.LineStyle = ExcelInerop.XlLineStyle.xlContinuous;
            border.Weight = 2d;

            // font
            range.Cells.Font.Name = "Calibri";
            range.Cells.Font.Size = "12";

            // sheet name
            excelSheet.Name = customer.name;
            
            // change range - take first row only
            range = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[1, cols]];
            range.Cells.Font.Color = Color.White;
            range.Cells.Interior.Color = Color.Red;
            range.Cells.Font.FontStyle = FontStyle.Bold;
            range.Cells.Font.Size = "14";

            // auto fit all columns
            range.Columns.EntireColumn.AutoFit();
        }

        // function fills excel range with data from dynamic values array
        private static void fillExcelValuesFromArray(dynamic valuesArray, int rows, int cols)
        {
            // define range to the table length
            ExcelInerop.Range range = excelSheet.Range[excelSheet.Cells[1, 1], excelSheet.Cells[rows, cols]];

            // get table into the range
            range.Value2 = valuesArray;
        }
    }
}

using System;
using System.Windows.Forms;
using System.Drawing;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Reflection;
using System.Globalization;

namespace emailTest
{
    public partial class OrdersParser : Form
    {
        public static OrdersParser _Form;

        public OrdersParser()
        {
            InitializeComponent();

            buttonsSetVisible(false);
            animateGif(true);
            arrivals_lbl.Visible = false;
            sails_lbl.Visible = false;

            // needed for logs from other classes
            _Form = this;

            log("Welcome!");

            // initialization is too long
            // start in new task
            Task.Factory.StartNew(() =>
                                        {
                                            // init office instances
                                            Excel.init();
                                            Outlook.init();

                                            // create temp results folder
                                            Utils.createResultsFolder();

                                            // parse local DB for customer details
                                            Excel.getCustomersDetails();

                                            // parse the Tanko excel params to load DB from outlook
                                            Excel.getTankoExcelParameters();

                                            // parse the local DB for agents details
                                            Excel.getAgentsDetails();

                                            // parse the local DB for shipping companies details
                                            Excel.getShippingCompaniesDetais();

                                            // fetch and save to file the most updated orders excel file
                                            Outlook.readLastOrdersFile();

                                            // parse the orders DB
                                            Excel.getOrderDetails();

                                            // today's arrivales
                                            updateArrivalsGrid();

                                            // yesterday's sails
                                            updateSailsGrid();
                                        })

                        // when done, call this CB
                        .ContinueWith(initCompleteCB);
        }

        // handler for "prepare mails" button click
        private void parse_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            if (string.IsNullOrEmpty(Common.plannedImportExcel) == true)
            {
                log("Failed to open excel with data - exiting", logLevel.error);
                cleanResources(false);
            }
            else
            {
                animateGif(true);

                // parsing is long, so in order not to block the GUI
                // start in new task
                Task.Factory.StartNew(() =>
                                            {
                                                // prepare mails for all customer based on DB data
                                                Outlook.prepareOrderMailsToAllCustomers();
                                            })
                            // when done, call this CB
                            .ContinueWith(mailCompleteCB);
            }
        }

        // function disposes all used classes
        private void cleanResources(bool bSuccess)
        {
            // dispose classes
            Excel.dispose();
            Outlook.dispose();

            animateGif(false);

            if (bSuccess == true)
            {
                log("Process done - mails are ready");
            }
        }

        // function starts/stops GIF animation
        // due to cross-threads operations, make sure to invoke when asked from different thread
        private void animateGif(bool bAnimate)
        {
            if (picbox.InvokeRequired == true)
            {
                picbox.Invoke(new MethodInvoker(delegate { picbox.Enabled = bAnimate; }));
            }
            else
            {
                picbox.Enabled = bAnimate;
            }
        }

        // function enables/disable form buttons
        // due to cross-threads operations, make sure to invoke when asked from different thread
        private void buttonsSetVisible(bool bVisible)
        {
            foreach (Button btn in new List<Button>() { reports_btn , loadConfirm_btn , bookConfirm_btn , docReceipts_btn})
            {
                if (btn.InvokeRequired == true)
                {
                    btn.Invoke(new MethodInvoker(delegate { btn.Enabled = bVisible; }));
                }
                else
                {
                    btn.Enabled = bVisible;
                }
            }
        }

        // CB called when mails are prepared (end of program)
        private void mailCompleteCB(Task obj)
        {
            cleanResources(true);
        }

        // CB called when init is complete
        private void initCompleteCB(Task obj)
        {
            animateGif(false);
            buttonsSetVisible(true);
            log(string.Format("Press on the '{0}' button to continue", reports_btn.Text));
        }

        public enum logLevel
        {
            info,
            error
        }

        // function prints log (basic is 'info')
        // since can be called from different processes
        // need to make sure that it can update GUI variables using invoke methods
        public void log(string msg, logLevel level = logLevel.info)
        {
            if (logTextBox.InvokeRequired == true)
            {
                logTextBox.Invoke(new MethodInvoker(delegate { logThreadSafe(msg, level); }));
            }
            else
            {
                logThreadSafe(msg, level);
            }
        }
        
        // function updates GUI, therefore, must be called on same thread
        private void logThreadSafe(string msg, logLevel level)
        {
            // add 'enter' only if not first
            if (string.IsNullOrEmpty(logTextBox.Text) == false) logTextBox.AppendText(Environment.NewLine);

            if (level == logLevel.error)
            {
                logTextBox.SelectionColor = Color.Red;
            }

            if (level == logLevel.info)
            {
                logTextBox.SelectionColor = Color.Black;
            }

            // needed for colors
            logTextBox.SelectionStart = logTextBox.TextLength;
            logTextBox.SelectionLength = 0;

            // add text
            logTextBox.AppendText(string.Format("{0}  {1}", DateTime.Now.ToString("HH:mm:ss.fff"), msg));

            logTextBox.SelectionColor = logTextBox.ForeColor;
            //logTextBox.ScrollToCaret();
            logTextBox.Refresh();
        }

        private void loadConfirm_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            // sanity check
            if (string.IsNullOrEmpty(Common.plannedImportExcel) == true)
            {
                log("Failed to open excel with data - exiting", logLevel.error);
                cleanResources(false);
            }
            else
            {
                animateGif(true);

                // parsing is long, so in order not to block the GUI
                // start in new task
                Task.Factory.StartNew(() =>
                                            {
                                                // prepare mails for all customer based on DB data
                                                Outlook.prepareLoadingMailsToAllAgents();
                                            })
                            // when done, call this CB
                            .ContinueWith(mailCompleteCB);
            }
        }

        // function updates today's arrivals data grid
        private void updateArrivalsGrid()
        {
            List<Common.Order>  resultList      = new List<Common.Order>();
            DateTime            now             = DateTime.Now;
            string              str             = string.Empty;

            // filter only needed customer (all the customers in the list)
            foreach (Common.Customer customer in Common.customerList)
            {
                resultList.AddRange(Outlook.filterCustomersByName(customer.name));
            }

            // filter only today's arrival dates
            // filter only loadings sent from the country of the agent
            // order by consignee
            resultList = resultList.Where(x => x.arrivalDate.Date == DateTime.Now.Date)
                                   .OrderBy(x => x.consignee)
                                   .ToList();

            // check if customer has orders
            if (resultList.Count == 0)
            {
                str = "No new arrivals totay";
                log(str);

                arrivals_lbl.Invoke(new MethodInvoker(delegate 
                                                                {
                                                                    arrivals_lbl.Text = str;
                                                                    arrivals_lbl.Visible = true;
                                                                    arrivals_lbl.Refresh();
                                                                }));

                return;
            }

            str = string.Format("{0} new arrivals today", resultList.Count);
            log(str);

            arrivals_lbl.Invoke(new MethodInvoker(delegate
                                                            {
                                                                arrivals_lbl.Text = str;
                                                                arrivals_lbl.Visible = true;
                                                                arrivals_lbl.Refresh();
                                                            }));

            // not all the columns are needed in the report - remove some
            List<Common.ArrivalsReport> targetResList = resultList.ConvertAll(x => new Common.ArrivalsReport
            {
                jobNo       = x.jobNo,
                consignee   = x.consignee,
                arrivalDate = x.arrivalDate,
            });

            // prepare DataTable to fill the grid
            DataTable table = Utils.generateDataTableFromList<Common.ArrivalsReport>(targetResList);

            arrivalsDataGrid.Invoke(new MethodInvoker(delegate 
                                                                {
                                                                    arrivalsDataGrid.AutoGenerateColumns = true;
                                                                    arrivalsDataGrid.DataSource = table;
                                                                    arrivalsDataGrid.AutoResizeColumns();
                                                                    arrivalsDataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                                                                    arrivalsDataGrid.Refresh();
                                                                }));
        }

        // function updates sails data grid in the last 3 days
        private void updateSailsGrid()
        {
            List<Common.Order>  resultList      = new List<Common.Order>();
            int                 sailingDays     = 3;
            string              str             = string.Empty;

            // filter only needed customer (all the customers in the list)
            foreach (Common.Customer customer in Common.customerList)
            {
                resultList.AddRange(Outlook.filterCustomersByName(customer.name));
            }

            // filter only yesterday's sailing dates
            // filter only loadings sent from the country of the agent
            // order by consignee
            resultList = resultList.Where(x => x.sailingDate.Date >= DateTime.Now.AddDays((-1) * (sailingDays)).Date &&
                                               x.sailingDate.Date <= DateTime.Now.AddDays(-1))
                                   .OrderByDescending(x => x.sailingDate)
                                   .ToList();

            // check if customer has orders
            if (resultList.Count == 0)
            {
                str = string.Format("No new sailings in the last {0} days", sailingDays);
                log(str);

                arrivals_lbl.Invoke(new MethodInvoker(delegate
                                                                {
                                                                    sails_lbl.Text = str;
                                                                    sails_lbl.Visible = true;
                                                                    sails_lbl.Refresh();
                                                                }));

                return;
            }

            str = string.Format("{0} new sailings in the last {1} days", resultList.Count, sailingDays);
            log(str);

            arrivals_lbl.Invoke(new MethodInvoker(delegate
                                                            {
                                                                sails_lbl.Text = str;
                                                                sails_lbl.Visible = true;
                                                                sails_lbl.Refresh();
                                                            }));

            // not all the columns are needed in the report - remove some
            List<Common.SailsReport> targetResList = resultList.ConvertAll(x => new Common.SailsReport
            {
                jobNo       = x.jobNo,
                shipper     = x.shipper,
                consignee   = x.consignee,
                tankNum     = x.tankNum,
                fromCountry = x.fromCountry,
                sailingDate = x.sailingDate,
            });

            // prepare DataTable to fill the grid
            DataTable table = Utils.generateDataTableFromList<Common.SailsReport>(targetResList);

            sailsDataGrid.Invoke(new MethodInvoker(delegate
                                                            {
                                                                sailsDataGrid.DataSource = table;
                                                                sailsDataGrid.DataBindingComplete += SailsDataGrid_DataBindingComplete;
                                                            }));
        }

        // handle for data load complete - colorize table
        private void SailsDataGrid_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            int lastColIndex = sailsDataGrid.Columns.Count - 1;

            foreach (DataGridViewRow row in sailsDataGrid.Rows)
            {
                DateTime sailingDate = DateTime.Parse(row.Cells[lastColIndex].Value.ToString());

                // colorize according to sailing data
                if (sailingDate.Date == DateTime.Now.AddDays(-3).Date)
                {
                    row.DefaultCellStyle.BackColor = Color.MediumSeaGreen;
                }

                if (sailingDate.Date == DateTime.Now.AddDays(-2).Date)
                {
                    row.DefaultCellStyle.BackColor = Color.DarkSeaGreen;
                }
            }

            //sailsDataGrid.DefaultCellStyle.Font = new Font(new FontFamily("Calibri"), 10f);
            sailsDataGrid.AutoGenerateColumns = true;
            sailsDataGrid.AutoResizeColumns();
            sailsDataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            sailsDataGrid.Refresh();
        }

        private void SailsDataGrid_DataBindingComplete(object sender, DataGridViewCellFormattingEventArgs e)
        {
            DataGridViewRow row = sailsDataGrid.Rows[e.RowIndex];// get you required index
                                                         // check the cell value under your specific column and then you can toggle your colors
            row.DefaultCellStyle.BackColor = Color.OldLace;
            log(DateTime.Now.ToShortTimeString());
        }

        public DataTable ToDataTable<T>(List<T> items)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);

            //Get all the properties
            PropertyInfo[] Props = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance);
            foreach (PropertyInfo prop in Props)
            {
                //Defining type of data column gives proper data table 
                var type = (prop.PropertyType.IsGenericType && prop.PropertyType.GetGenericTypeDefinition() == typeof(Nullable<>) ? Nullable.GetUnderlyingType(prop.PropertyType) : prop.PropertyType);
                //Setting column names as Property names
                dataTable.Columns.Add(prop.Name, type);
            }
            foreach (T item in items)
            {
                var values = new object[Props.Length];
                for (int i = 0; i < Props.Length; i++)
                {
                    //inserting property values to datatable rows
                    values[i] = Props[i].GetValue(item, null);
                }
                dataTable.Rows.Add(values);
            }
            //put a breakpoint here and check datatable
            return dataTable;
        }

        // function sends emails to all shipping companies with future bookings
        private void bookConfirm_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            // sanity check
            if (string.IsNullOrEmpty(Common.plannedImportExcel) == true)
            {
                log("Failed to open excel with data - exiting", logLevel.error);
                cleanResources(false);
            }
            else
            {
                animateGif(true);

                // parsing is long, so in order not to block the GUI
                // start in new task
                Task.Factory.StartNew(() =>
                                            {
                                                // prepare mails for all customer based on DB data
                                                Outlook.prepareBookingMailsToAllAgents();
                                            })
                            // when done, call this CB
                            .ContinueWith(mailCompleteCB);
            }
        }

        // function sends document recepts requests from agents
        private void docReceipts_btn_Click(object sender, EventArgs e)
        {
            List<Common.SailsReport> targetResList = new List<Common.SailsReport>();

            buttonsSetVisible(false);
            animateGif(true);

            // sanity check
            if (sailsDataGrid.SelectedRows.Count == 0)
            {
                log("Nothing was selected, you must select at least one sailing", logLevel.error);
                cleanResources(false);
                return;
            }

            // go over the selected rows and generate new List<T> for emails
            foreach (DataGridViewRow row in sailsDataGrid.SelectedRows)
            {
                Common.SailsReport report = new Common.SailsReport();

                report.jobNo        = Convert.ToInt32(row.Cells[0].Value);
                report.shipper      = row.Cells[1].Value.ToString();
                report.consignee    = row.Cells[2].Value.ToString();
                report.tankNum      = row.Cells[3].Value.ToString();
                report.fromCountry  = row.Cells[4].Value.ToString();
                report.sailingDate  = DateTime.Parse(row.Cells[5].Value.ToString());

                targetResList.Add(report);
            }


            // parsing is long, so in order not to block the GUI
            // start in new task
            Task.Factory.StartNew(() =>
                                        {
                                            // prepare mails for all customer based on DB data
                                            Outlook.prepareDocumentsReceiptesMailsToAllAgents(targetResList);
                                        })
                        // when done, call this CB
                        .ContinueWith(mailCompleteCB);

        }
    }
}

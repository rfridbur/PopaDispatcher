using System;
using System.Windows.Forms;
using System.Drawing;
using System.Threading.Tasks;
using System.Collections.Generic;
using System.Linq;
using System.Data;

namespace Anko
{
    public partial class OrdersParser : Form
    {
        public static OrdersParser _Form;

        // contractor
        public OrdersParser()
        {
            InitializeComponent();

            // form visualization init
            buttonsSetVisible(false);
            animateGif(true);
            arrivals_lbl.Text = string.Empty;
            sails_lbl.Text = string.Empty;
            ashdodLinkLbl.LinkClicked += AshdodLinkLbl_LinkClicked;
            haifaLinkLbl.LinkClicked += HaifaLinkLbl_LinkClicked;

            // needed to get form controls from other classed
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

                                            // parse local DB
                                            Excel.getDetailsFromLocalDb();

                                            // fetch and save to file the most updated orders excel file
                                            Outlook.readLastOrdersFile();

                                            // parse the orders DB
                                            Excel.getOrderDetails();

                                            // today's arrivals
                                            updateArrivalsGrid();

                                            // yesterday's sails
                                            updateSailsGrid();
                                        });
        }

        // function disposes all used classes
        private void cleanResources(bool bSuccess)
        {
            // dispose classes
            //Excel.dispose();
            //Outlook.dispose();

            animateGif(false);
            buttonsSetVisible(true);

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

        // CB called when mails are prepared (end of program)
        private void mailCompleteCB(Task obj)
        {
            cleanResources(true);
        }

        // CB called when init is complete
        private void initCompleteCB()
        {
            animateGif(false);
            buttonsSetVisible(true);
            log(string.Format("Init is complete"));
        }

        // function updates lbl.Text with input text
        public void updateLabel(Label lbl, string text)
        {
            if (lbl.InvokeRequired == true)
            {
                lbl.Invoke(new MethodInvoker(delegate { lbl.Text = text; lbl.Refresh(); }));
            }
            else
            {
                lbl.Text = text;
                lbl.Refresh();
            }
        }

        #region Log
        public enum LogLevel
        {
            Info,
            Error
        }

        // function prints log (basic is 'info')
        // since can be called from different processes
        // need to make sure that it can update GUI variables using invoke methods
        public void log(string msg, LogLevel level = LogLevel.Info)
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
        private void logThreadSafe(string msg, LogLevel level)
        {
            // add 'enter' only if not first
            if (string.IsNullOrEmpty(logTextBox.Text) == false) logTextBox.AppendText(Environment.NewLine);

            if (level == LogLevel.Error)
            {
                logTextBox.SelectionColor = Color.Red;
            }

            if (level == LogLevel.Info)
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
        #endregion

        #region Grids
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

#if OFFLINE
            // for testing purposes, since there might be no arrivals today, take several random arrivals
            resultList = resultList.Where(x => x.arrivalDate.Date >= DateTime.Now.Date)
                                   .Take(6)
                                   .OrderBy(x => x.consignee)
                                   .ToList();
#else
            // filter only today's arrival dates
            // filter only loadings sent from the country of the agent
            // order by consignee
            resultList = resultList.Where(x => x.arrivalDate.Date == DateTime.Now.Date)
                                   .OrderBy(x => x.consignee)
                                   .ToList();
#endif

            // check if customer has orders
            if (resultList.Count == 0)
            {
                str = "No new arrivals today";
                log(str);
                updateLabel(arrivals_lbl, str);
                initCompleteCB();
                return;
            }

            str = string.Format("{0} new arrivals today", resultList.Count);
            log(str);
            updateLabel(arrivals_lbl, str);

            // start async thread to get data from ports web
            // optimization: downloading data from web takes time, so do not do it
            // in case there are no arrivals today to this specific port
            if (Utils.bArrivalsToPort(resultList, PortService.PortName.Ashdod) == true)
            {
                PortService.getShipsFromPort(PortService.PortName.Ashdod);
            }

            if (Utils.bArrivalsToPort(resultList, PortService.PortName.Haifa) == true)
            {
                PortService.getShipsFromPort(PortService.PortName.Haifa);
            }

            // not all the columns are needed in the report - remove some
            List<Common.ArrivalsReport> targetResList = resultList.ConvertAll(x => new Common.ArrivalsReport
            {
                jobNo       = x.jobNo,
                consignee   = x.consignee,
                toPlace     = x.toPlace,
                vessel      = x.vessel,
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
                updateLabel(sails_lbl, str);

                return;
            }

            str = string.Format("{0} new sailings in the last {1} days", resultList.Count, sailingDays);
            log(str);
            updateLabel(sails_lbl, str);

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
                                                                sailsDataGrid.DataBindingComplete += sailsDataGrid_DataBindingComplete;
                                                            }));
        }

        // handler for data load complete - colorize table
        private void sailsDataGrid_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            int lastColIndex = sailsDataGrid.Columns.Count - 1;

            foreach (DataGridViewRow row in sailsDataGrid.Rows)
            {
                DateTime sailingDate = DateTime.Parse(row.Cells[lastColIndex].Value.ToString());

                // colorize according to sailing data
                if (sailingDate.Date == DateTime.Now.AddDays(-3).Date)
                {
                    row.DefaultCellStyle.BackColor = Color.HotPink;
                }

                if (sailingDate.Date == DateTime.Now.AddDays(-2).Date)
                {
                    row.DefaultCellStyle.BackColor = Color.Pink;
                }
            }

            //sailsDataGrid.DefaultCellStyle.Font = new Font(new FontFamily("Calibri"), 10f);
            sailsDataGrid.AutoGenerateColumns = true;
            sailsDataGrid.AutoResizeColumns();
            sailsDataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
            sailsDataGrid.Refresh();
        }

        // function updates arrivals data grid with data taken from port websites
        // function should be called when data reading from all ports is complete
        // and the following lists are filled with data:
        // Common.ashdodAnchoringList
        // Common.haifaAnchoringList
        // function colors rows according to arrival status and adds tooltips with additional data
        public void arrivalsDataGrid_WebSyncComplete()
        {
            // get columns based on Common.ArrivalsReport
            PortService.PortName    portName        = PortService.PortName.Unknown;
            int                     toPlaceIndex    = arrivalsDataGrid.Columns["toPlace"].Index;
            int                     vesselIndex     = arrivalsDataGrid.Columns["vessel"].Index;
            string                  portNameStr     = string.Empty;
            string                  toolTipStr      = string.Empty;

            // go over each rows
            foreach (DataGridViewRow row in arrivalsDataGrid.Rows)
            {
                string vesselName = string.Empty;
                portNameStr = row.Cells[toPlaceIndex].Value.ToString();

                if (portNameStr.ToLower() == "ashdod")
                {
                    portName = PortService.PortName.Ashdod;
                }

                if (portNameStr.ToLower() == "haifa")
                {
                    portName = PortService.PortName.Haifa;
                }

                vesselName = row.Cells[vesselIndex].Value.ToString();

                if (string.IsNullOrEmpty(vesselName) == false)
                {
                    // colorize according to arrival status
                    // get tool tip as well during parsing
                    switch (PortService.shipStatusInPort(vesselName, portName, out toolTipStr))
                    {
                        case PortService.ShipStatus.Arrived:
                            row.DefaultCellStyle.BackColor = Color.LightGreen;
                            break;
                        case PortService.ShipStatus.Expected:
                            row.DefaultCellStyle.BackColor = Color.LightPink;
                            break;
                        case PortService.ShipStatus.Unknown:
                            row.DefaultCellStyle.BackColor = Color.LightYellow;
                            break;
                        default:
                            break;
                    }
                }
                else
                {
                    // vessel name is empty, this can happen if excel has not been updated yet
                    // in such case, mark as unknown
                    row.DefaultCellStyle.BackColor = Color.LightYellow;
                    toolTipStr = "Vessel not found";
                }

                // add tool tip with additional data about this ship
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.ToolTipText = toolTipStr;
                }
            }

            //sailsDataGrid.DefaultCellStyle.Font = new Font(new FontFamily("Calibri"), 10f);
            arrivalsDataGrid.AutoGenerateColumns = true;
            arrivalsDataGrid.AutoResizeColumns();
            arrivalsDataGrid.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            arrivalsDataGrid.Invoke(new MethodInvoker(delegate
                                                            {
                                                                arrivalsDataGrid.Refresh();
                                                            }));

            initCompleteCB();
        }
        #endregion

        #region Buttons
        // handler for "prepare mails" button click
        private void parse_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            // sanity check that excel file was parsed successfully
            if (Utils.bValidOrders() == false)
            {
                // file doesn't exist
                log("Orders are not valid", LogLevel.Error);
                return;
            }

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

        // function sends mails to agents asking for loading confirmation
        private void loadConfirm_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            // sanity check that excel file was parsed successfully
            if (Utils.bValidOrders() == false)
            {
                // file doesn't exist
                log("Orders are not valid", LogLevel.Error);
                return;
            }

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

        // function sends emails to all shipping companies with future bookings
        private void bookConfirm_btn_Click(object sender, EventArgs e)
        {
            buttonsSetVisible(false);

            // sanity check that excel file was parsed successfully
            if (Utils.bValidOrders() == false)
            {
                // file doesn't exist
                log("Orders are not valid", LogLevel.Error);
                return;
            }

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

        // function sends document receipts requests from agents
        private void docReceipts_btn_Click(object sender, EventArgs e)
        {
            List<Common.SailsReport> targetResList = new List<Common.SailsReport>();

            buttonsSetVisible(false);
            animateGif(true);

            // sanity check
            if (sailsDataGrid.SelectedRows.Count == 0)
            {
                log("Nothing was selected, you must select at least one sailing", LogLevel.Error);
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
        #endregion

        #region Hyperlinks
        // hyperlink for Haifa port
        private void HaifaLinkLbl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // specify that the link was visited
            if (haifaLinkLbl.InvokeRequired == true)
            {
                haifaLinkLbl.Invoke(new MethodInvoker(delegate { haifaLinkLbl.LinkVisited = true; }));
            }
            else
            {
                haifaLinkLbl.LinkVisited = true;
            }

            // navigate to a URL
            System.Diagnostics.Process.Start(PortService.HAIFA_URL);
        }

        //hyperlink for port ashdod port
        private void AshdodLinkLbl_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            // specify that the link was visited
            if (ashdodLinkLbl.InvokeRequired == true)
            {
                ashdodLinkLbl.Invoke(new MethodInvoker(delegate { ashdodLinkLbl.LinkVisited = true; }));
            }
            else
            {
                ashdodLinkLbl.LinkVisited = true;
            }

            // navigate to a URL
            System.Diagnostics.Process.Start(PortService.ASHDOD_URL);
        }
        #endregion
    }
}
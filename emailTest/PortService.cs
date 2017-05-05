using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Threading;
using System.Windows.Forms;
using System.Linq;

namespace emailTest
{
    class PortService
    {
        public static string ASHDOD_URL = "https://www.ashdodport.co.il/english/onlineservices/pages/apc-eng-port_shipsgeneral.aspx";
        public static string HAIFA_URL = "http://www.port2port.co.il/ports/haifa/MoorPlanHaifaPort.html";
        private const string AT_BERTH_HEB = "ברציף";
        private static string[] expectedStrList = { "anchor", "sailed", "expected", "צפויה", "מחוץ לנמל" };
        private static int webDataReadyCounter = 0;

        public static void getShipsFromPort(PortName portName)
        {
            string                                  url     = string.Empty;
            WebBrowserDocumentCompletedEventHandler cbFunc  = null;

            // sanity check
            if (portName == PortName.Unknown)
            {
                OrdersParser._Form.log(string.Format("Unknown port: {0}", portName.ToString()), OrdersParser.LogLevel.Error);
                return;
            }

            OrdersParser._Form.log(string.Format("Fetching data from {0} port", portName.ToString()));

            // increase counter to inform on web activity
            Interlocked.Increment(ref webDataReadyCounter);

            if (portName == PortName.Ashdod)
            {
                url = ASHDOD_URL;
                cbFunc = browser_DocumentCompletedAshdod;
            }

            if (portName == PortName.Haifa)
            {
                url = HAIFA_URL;
                cbFunc = browser_DocumentCompletedHaifa;
            }

            var th = new Thread(() =>   {
                                            var br = new WebBrowser();
                                            br.ScriptErrorsSuppressed = true;
                                            br.DocumentCompleted += cbFunc;
                                            br.Navigate(url);
                                            Application.Run();
                                        });

            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        private static void browser_DocumentCompletedAshdod(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // document complete can fire multiple times, wait for the last one
            if (e.Url.AbsolutePath != ((WebBrowser)sender).Url.AbsolutePath)
            {
                return;
            }

            var br = sender as WebBrowser;
            string pageHtml = br.Document.GetElementsByTagName("HTML")[0].OuterHtml;
            DataSet ashdodTable = Utils.convertHTMLTablesToDataSet(pageHtml);
            Common.ashdodAnchoringList = generateTableForAshdodPort(ashdodTable);

            pageLoadCompleter();
        }

        private static void browser_DocumentCompletedHaifa(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            // document complete can fire multiple times, wait for the last one
            if (e.Url.AbsolutePath != ((WebBrowser)sender).Url.AbsolutePath)
            {
                return;
            }

            var br = sender as WebBrowser;
            string pageHtml = br.Document.GetElementsByTagName("HTML")[0].OuterHtml;
            DataSet haifaTable = Utils.convertHTMLTablesToDataSet(pageHtml);
            Common.haifaAnchoringList = generateTableForHaifaPort(haifaTable);

            pageLoadCompleter();
        }

        // function completes the async web data reading and returns to form
        private static void pageLoadCompleter()
        {
            // return the decremented value
            if (Interlocked.Decrement(ref webDataReadyCounter) > 0)
            {
                // still loading, fold
                return;
            }

            OrdersParser._Form.log("Data fetching from port websites is complete");

            // load is complete, return to form and design table
            OrdersParser._Form.arrivalsDataGrid_WebSyncComplete();
        }

        private static List<Common.Anchoring> generateTableForHaifaPort(DataSet dataSet)
        {
            string      newLineStr          = Environment.NewLine;
            string[]    stringSeparators    = new string[] { newLineStr };
            string      dateParseFormat     = "dd/MM/yy HH:mm";
            string[]    tmpArr              = { };

            List<Common.Anchoring> anchoringPortList = new List<Common.Anchoring>();

            // we assume only one table in the page - take the first one
            DataTable dt = dataSet.Tables[0];

            // add all the records from a given table into AnchoringPortList
            // usie Trim() only in case of splits, since all atomic variables are already trimmed
            foreach (DataRow row in dt.Rows)
            {
                Common.Anchoring anchoringShip = new Common.Anchoring();

                // col[0] manifests
                // import and export manifests are in same column separated by new line
                // need to remove the '/' from manifest to cast into int
                string val = row[0].ToString().Replace("/", string.Empty);
                tmpArr = val.Split(stringSeparators, StringSplitOptions.None);
                if (tmpArr.Length == 2)
                {
                    anchoringShip.importManifest = Convert.ToInt32(tmpArr[0].Trim());
                    anchoringShip.exportManifest = Convert.ToInt32(tmpArr[0].Trim());
                }

                // col[1] ship name
                anchoringShip.shipName = row[1].ToString().Replace(newLineStr, string.Empty);

                // col[2] cargo
                // import and export cargo are in same column separated by new line
                tmpArr = row[2].ToString().Split(stringSeparators, StringSplitOptions.None);
                if (tmpArr.Length == 2)
                {
                    anchoringShip.importCargo = tmpArr[0].Trim();
                    anchoringShip.exportCargo = tmpArr[0].Trim();
                }

                // col[3] status
                anchoringShip.status = row[3].ToString();

                // col[4] platform
                anchoringShip.platform = row[4].ToString();

                // col[5] operator
                anchoringShip.operatingAgent = row[5].ToString().Replace(newLineStr, string.Empty);

                // col[6] line
                anchoringShip.serviceLinePorts = row[6].ToString().Replace(newLineStr, string.Empty);

                // col[7] partners
                anchoringShip.partners = row[7].ToString().Replace(newLineStr, string.Empty);

                // col[8] export start time
                if (string.IsNullOrEmpty(row[8].ToString()) == false)
                {
                    anchoringShip.exportStartTime = DateTime.ParseExact(row[8].ToString(), dateParseFormat, CultureInfo.InvariantCulture);
                }

                // col[9] export end time
                if (string.IsNullOrEmpty(row[9].ToString()) == false)
                {
                    anchoringShip.exportEndTime = DateTime.ParseExact(row[9].ToString(), dateParseFormat, CultureInfo.InvariantCulture);
                }

                // col[10] arrival date
                if (string.IsNullOrEmpty(row[10].ToString()) == false)
                {
                    anchoringShip.arrivalDate = DateTime.ParseExact(row[10].ToString(), dateParseFormat, CultureInfo.InvariantCulture);
                }

                // row is ready, add to list
                anchoringPortList.Add(anchoringShip);
                anchoringShip = null;
            }

            return anchoringPortList;
        }

        private static List<Common.Anchoring> generateTableForAshdodPort(DataSet dataSet)
        {
            string      newLineStr          = Environment.NewLine;
            string[]    stringSeparators    = new string[] { newLineStr };
            string      dateParseFormat     = "yyyy-MM-ddHH:mm";
            string      value               = string.Empty;
            string[]    tmpArr              = { };

            List<Common.Anchoring> anchoringPortList = new List<Common.Anchoring>();

            // there might be several tables, we assume that they have same format
            foreach (DataTable dt in dataSet.Tables)
            {
                // add all the records from a given table into AnchoringPortList
                // usie Trim() only in case of splits, since all atomic variables are already trimmed
                foreach (DataRow row in dt.Rows)
                {
                    Common.Anchoring anchoringShip = new Common.Anchoring();

                    // col[0] last port
                    anchoringShip.lastPort = row[0].ToString();

                    // col[1] flag
                    anchoringShip.flag = row[1].ToString();

                    // col[2] arrival date
                    value = row[2].ToString().Replace(newLineStr, string.Empty);
                    if (string.IsNullOrEmpty(value) == false)
                    {
                        anchoringShip.arrivalDate = DateTime.ParseExact(value, dateParseFormat, CultureInfo.InvariantCulture);
                    }

                    // col[3] manifest
                    tmpArr = row[3].ToString().Split(stringSeparators, StringSplitOptions.None);
                    if (tmpArr.Length == 2)
                    {
                        anchoringShip.importManifest = Convert.ToInt32(tmpArr[0].Trim());
                        anchoringShip.exportManifest = Convert.ToInt32(tmpArr[1].Trim());
                    }

                    // col[4] cargo type
                    tmpArr = row[4].ToString().Split(stringSeparators, StringSplitOptions.None);
                    if (tmpArr.Length == 2)
                    {
                        anchoringShip.importCargo = tmpArr[0].Trim();
                        anchoringShip.exportCargo = tmpArr[1].Trim();
                    }

                    // col[5] place code
                    anchoringShip.status = row[5].ToString().Trim();

                    // col[6] yard status
                    anchoringShip.yardStatus = row[6].ToString().Trim();

                    // col[7] agent
                    anchoringShip.operatingAgent = row[7].ToString().Replace(newLineStr, string.Empty).Trim();

                    // col[8] shceduled
                    anchoringShip.bScheduled = false;
                    if (row[8].ToString().ToLower() == "yes")
                    {
                        anchoringShip.bScheduled = true;
                    }

                    // col[9] ship name
                    anchoringShip.shipName = row[9].ToString().Replace(newLineStr, string.Empty).Trim();

                    // row is ready, add to list
                    anchoringPortList.Add(anchoringShip);
                    anchoringShip = null;
                }
            }

            return anchoringPortList;
        }

        // enum for ship status in certain port
        public enum ShipStatus
        {
            Unknown,
            Arrived,
            Expected
        }

        // enum for all available ports
        public enum PortName
        {
            Unknown,
            Ashdod,
            Haifa
        }

        // function returns ship's arrival status to port
        public static ShipStatus shipStatusInPort(string vesselName, PortName portName, out string summaryStr)
        {
            List<Common.Anchoring> templist         = new List<Common.Anchoring>();
            string                 status           = string.Empty;
            string                 arrivalDateStr   = string.Empty;
            string                 dateFormat       = "dd/MM/yy HH:mm";

            summaryStr = string.Empty;

            // sanity check
            if (portName == PortName.Unknown)
            {
                summaryStr = "Unknown port";
                return ShipStatus.Unknown;
            }

            if (portName == PortName.Ashdod)
            {
                templist = Common.ashdodAnchoringList.Where(x => x.shipName.ToLower().Contains(vesselName.ToLower()))
                                                     .OrderBy(x => x.importManifest)
                                                     .ToList();

                if (templist.Count > 0)
                {
                    // haifa port has the following statuses:
                    // integer, means that ship is at berth
                    // anchor, sailed, expected, means that ship is not at berth
                    status = templist.FirstOrDefault().status;
                    arrivalDateStr = templist.FirstOrDefault().arrivalDate.ToString(dateFormat);
                    int res;

                    if (int.TryParse(status, out res) == true)
                    {
                        // the value does't matter, but ship has arrived
                        summaryStr = string.Format("Status: Arrived\r\nArrival Date: {0}", arrivalDateStr);
                        return ShipStatus.Arrived;
                    }
                }
            }

            if (portName == PortName.Haifa)
            {
                templist = Common.haifaAnchoringList.Where(x => x.shipName.ToLower().Contains(vesselName.ToLower()))
                                                    .OrderBy(x => x.importManifest)
                                                    .ToList();

                if (templist.Count > 0)
                {
                    // ashdod port has only 3 possible hebrew string statuses
                    status = templist.FirstOrDefault().status;
                    arrivalDateStr = templist.FirstOrDefault().arrivalDate.ToString(dateFormat);

                    // check if arrived
                    if (status == AT_BERTH_HEB)
                    {
                        summaryStr = string.Format("Status: Arrived\r\nArrival Date: {0}", arrivalDateStr);
                        return ShipStatus.Arrived;
                    }
                }
            }

            // check in the expected list keywords
            if (expectedStrList.Contains(status.ToLower()))
            {
                summaryStr = string.Format("Status: Expected\r\nArrival Date: {0}", arrivalDateStr);
                return ShipStatus.Expected;
            }

            // error: didn't find anything - should not happen
            summaryStr = "Ship is not found in port";
            return ShipStatus.Unknown;
        }
    }
}
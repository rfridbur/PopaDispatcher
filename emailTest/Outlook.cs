using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading;
using System.Windows.Forms;
using OutlookInterop = Microsoft.Office.Interop.Outlook;

namespace emailTest
{
    // class has no constructor and has static methods since should not have instances
    // use as singleton
    class Outlook
    {
        // function initializes the class (should be called once at program init)
        public static void init()
        {
            // this can take several seconds
            OrdersParser._Form.log("Create outlook instances - please wait");

            // open once outlook instance for all project
            try
            {
                outlookApp = new OutlookInterop.Application();
            }
            catch (Exception e)
            {
                // fatal: cannot continue
                OrdersParser._Form.log(string.Format("Failed to open outlook instance. Error: {0}", e.Message), OrdersParser.logLevel.error);
                OrdersParser._Form.log(string.Format("Try to restart your PC or application"), OrdersParser.logLevel.error);
                return;
            }
        }

        // function destructs the class (should be called once at program halt)
        public static void dispose()
        {
            outlookApp = null;
        }

        private static  OutlookInterop.Application  outlookApp          = null;
        private const   string                      SIGNATURE_LOGO_NAME = "logo.jpg";
        public  static  string                      tancoOrdersFileName = string.Empty;
        public  static  string                      tancoOrdersEmail    = string.Empty;

        // function adds HTML line in the middle of the file
        // must be called between after prefix and prior to suffix
        private static string addHtmlLine(HorizontalAlignment   alignment, 
                                          string                color, 
                                          string                msg, 
                                          int                   fontSize,
                                          string                fontName = "calibri")
        {
            string direction = string.Empty;

            // in case of empty line - add <br>
            if (string.IsNullOrEmpty(msg))
            {
                return "<br>";
            }

            // for tables, do not add any addition, just leave as is
            // when table is created, all the relevant modifications must be done there
            if (msg.Contains("<table"))
            {
                return msg;
            }

            if (alignment == HorizontalAlignment.Left) direction = "ltr";
            if (alignment == HorizontalAlignment.Right) direction = "rtl";

            return string.Format("<div align={0}; dir={1}><span style=\"color:{2};\"><font face=\"{3}\" size=\"{4}\">{5}</font></span></div>",
                                  alignment.ToString(), 
                                  direction, 
                                  color, 
                                  fontName, 
                                  fontSize.ToString(),
                                  msg);
        }

        // function creates a table from list in HTML format
        private static string addHtmlTable<T>(List<T> dataList)
        {
            string  htmlTableStr    = string.Empty;
            int     rows            = 0;
            int     cols            = 0;

            object[,] valuesArray = Utils.generateObjectFromList<T>(dataList, out rows, out cols);

            htmlTableStr += @"<table border=2 width=\""500\"" rules=all cellpadding=3 bgcolor=AntiqueWhite>";

            for (int i = 0; i < rows; i++)
            {
                htmlTableStr += "<tr align=left";

                // apply only to first row
                if (i == 0)
                {
                    // doesn't work in outlook, but a correct HTML syntax - need to fix
                    htmlTableStr += string.Format(" style=\"{0};\"", "border-bottom: solid");
                }

                htmlTableStr += ">";

                for (int j = 0; j < cols; j++)
                {
                    // capitylize the titles (first row)
                    if (i == 0)
                    {
                        valuesArray[0, j] = Utils.uppercaseFirst(valuesArray[0,j].ToString());
                    }

                    htmlTableStr += string.Format("<th>{0}</th>", valuesArray[i,j]);
                }

                htmlTableStr += "</tr>";
            }

            htmlTableStr += "</table>";

            return htmlTableStr;
        }

        // function adds HTML prefix
        private static string addHtmlPreffix()
        {
            string tempMsg = string.Empty;

            tempMsg += "<!DOCTYPE HTML>";
            tempMsg += "<html>";
            tempMsg += "<body>";

            return tempMsg;
        }

        // function adds HTML suffix
        private static string addHtmlSuffix()
        {
            string tempMsg = string.Empty;

            tempMsg += "</body>";
            tempMsg += "</html>";

            return tempMsg;
        }

        // function prepares mails to all customers according to the mail type
        // certain mails are sent to chosen customers
        // possible optimization: parallel foreach
        public static void prepareOrderMailsToAllCustomers()
        {
            // send mails only to selected customers having bSendReport = true
            foreach (Common.Customer customer in Common.customerList.Where(x => x.bSendReport == true))
            {
                prepareOrdersMailToCustomer(customer);
            }
        }

        // function prepares mail per customer in case it has any orders
        private static void prepareOrdersMailToCustomer(Common.Customer customer)
        {
            DateTime    now             = DateTime.Now;
            string      outputFileName  = string.Empty;
            int         rows            = 0;
            int         cols            = 0;

            // filter the list for certain customer
            // filter out past arrivals
            // order by arrival date
            List<Common.Order> resultList = filterCustomersByName(customer.name).Where(x => x.arrivalDate > now)
                                                                                .OrderBy(x => x.arrivalDate)
                                                                                .ToList();

            // check if customer has orders
            if (resultList.Count == 0)
            {
                OrdersParser._Form.log(string.Format("{0}: no new orders for this customer - mail won't be sent", customer.name));
                return;
            }

            OrdersParser._Form.log(string.Format("{0}: {1} new orders found", customer.name, resultList.Count));

            // not all the columns are needed in the report - remove some
            List<Common.OrderReport> targetResList = resultList.ConvertAll(x => new Common.OrderReport
            {
                jobNo       = x.jobNo,
                shipper     = x.shipper,
                consignee   = x.consignee,
                customerRef = x.customerRef,
                tankNum     = x.tankNum,
                activity    = x.activity,
                loadingDate = x.loadingDate,
                fromCountry = x.fromCountry,
                fromPlace   = x.fromPlace,
                sailingDate = x.sailingDate,
                toCountry   = x.toCountry,
                toPlace     = x.toPlace,
                arrivalDate = x.arrivalDate,
                productName = x.productName,
                vessel      = x.vessel,
                voyage      = x.voyage,
            });

            // prepare 2d array for excel print
            object[,] valuesArray = Utils.generateObjectFromList<Common.OrderReport>(targetResList, out rows, out cols);

            // create new excel file with this data
            outputFileName = Path.Combine(Utils.resultsDirectoryPath, string.Format("{0}_{1}.{2}", customer.name, now.ToString(Common.DATE_FORMAT), "xlsx"));
            Excel.generateCustomerFile(valuesArray, rows, cols, customer, outputFileName);

            // send mail to customer
            MailDetails mailDetails         = new MailDetails();
            mailDetails.mailRecepient       = customer;
            mailDetails.mailType            = Common.MailType.Reports;
            mailDetails.bodyParameters      = new Dictionary<string, string>()
            {
                { "arrivalDate"         , resultList.FirstOrDefault().arrivalDate.ToString(Common.DATE_FORMAT) },
                { "totalNumOfOrders"    , resultList.Count.ToString() }
            };
            mailDetails.subject             = "דו'ח סטטוס הזמנות";
            mailDetails.attachments         = new List<string>() { outputFileName };
            mailDetails.bHighImportance     = true;

            // send mail
            sendMail(mailDetails);
        }

        // function prepares mails of future loadings to all agents
        // possible optimization: parallel foreach
        public static void prepareLoadingMailsToAllAgents()
        {
            foreach (Common.Agent agent in Common.agentList)
            {
                prepareLoadingConfirmationMailToCustomer(agent);
            }
        }

        private static void prepareLoadingConfirmationMailToCustomer(Common.Agent agent)
        {
            List<Common.Order>  resultList      = new List<Common.Order>();
            string              outputFileName  = string.Empty;

            // filter only needed customer (all the customers in the list)
            foreach (Common.Customer customer in Common.customerList)
            {
                resultList.AddRange(filterCustomersByName(customer.name));
            }

            // filter only tomorrow's loading dates
            // filter only loadings sent from the country of the agent
            // order by consignee
            resultList = resultList.Where(x => agent.countries.Contains(x.fromCountry) && x.loadingDate.Date == DateTime.Now.AddDays(1).Date)
                                   .OrderBy(x => x.consignee)
                                   .ToList();

            // check if customer has orders
            if (resultList.Count == 0)
            {
                OrdersParser._Form.log(string.Format("no new loadings tomorrow for {0}", agent.name));
                return;
            }

            OrdersParser._Form.log(string.Format("{0} loading for tomorrow found for {1}", resultList.Count, agent.name));

            // not all the columns are needed in the report - remove some
            List<Common.LoadingReport> targetResList = resultList.ConvertAll(x => new Common.LoadingReport
            {
                jobNo       = x.jobNo,
                consignee   = x.consignee,
                loadingDate = x.loadingDate,
                fromCountry = x.fromCountry,
            });

            string htmlTableStr = addHtmlTable<Common.LoadingReport>(targetResList);

            // send mail to customer
            MailDetails mailDetails = new MailDetails();
            mailDetails.mailRecepient = agent;
            mailDetails.mailType = Common.MailType.LoadingConfirmation;
            mailDetails.bodyParameters = new Dictionary<string, string>()
            {
                { "table"       , htmlTableStr },
                { "date"        , DateTime.Now.AddDays(1).ToString(Common.DATE_FORMAT) },
                { "day"         , DateTime.Now.AddDays(1).DayOfWeek.ToString() },
                { "agent"       , agent.name},
            };
            mailDetails.subject = "Loading confirmation for tomorrow";
            mailDetails.attachments = new List<string>() { };
            mailDetails.bHighImportance = false;

            // send mail
            sendMail(mailDetails);
        }

        public static void prepareDocumentsReceiptesMailsToAllAgents(List<Common.SailsReport> report)
        {
            foreach (Common.Agent agent in Common.agentList)
            {
                prepareDocumentsReceiptsMailToAgent(agent, report);
            }
        }

        // function goes over the selevted rows from the data grid and sends mails to agents
        private static void prepareDocumentsReceiptsMailToAgent(Common.Agent agent, List<Common.SailsReport> report)
        {
            string outputFileName = string.Empty;

            // filter only shipping sent from the country of the agent
            // order by consignee
            report = report.Where(x => agent.countries.Contains(x.fromCountry))
                           .OrderBy(x => x.sailingDate)
                           .ToList();

            // check if customer has orders
            if (report.Count == 0)
            {
                return;
            }

            OrdersParser._Form.log(string.Format("Documents receipts are needed from {0} on {1} sailings", agent.name, report.Count));

            // not all the columns are needed in the report - remove some
            List<Common.SailsReport> targetResList = report.ConvertAll(x => new Common.SailsReport
            {
                jobNo       = x.jobNo,
                shipper     = x.shipper,
                consignee   = x.consignee,
                tankNum     = x.tankNum,
                fromCountry = x.fromCountry,
                sailingDate = x.sailingDate,
            });

            string htmlTableStr = addHtmlTable<Common.SailsReport>(targetResList);

            // send mail to customer
            MailDetails mailDetails = new MailDetails();
            mailDetails.mailRecepient = agent;
            mailDetails.mailType = Common.MailType.DocumentsReceipts;
            mailDetails.bodyParameters = new Dictionary<string, string>()
            {
                { "table"       , htmlTableStr },
                { "agent"       , agent.name},
            };
            mailDetails.subject = "BL is missing";
            mailDetails.attachments = new List<string>() { };
            mailDetails.bHighImportance = false;

            // send mail
            sendMail(mailDetails);
        }

        // function prepares mails of future boakings to all shipping companies
        // possible optimization: parallel foreach
        public static void prepareBookingMailsToAllAgents()
        {
            // create a mail per shipping company
            foreach (Common.ShippingCompany company in Common.shippingCompanyList)
            {
                prepareBookingMailToAgent(company);
            }
        }

        private static void prepareBookingMailToAgent(Common.ShippingCompany shippingCompany)
        {
            List<Common.Order>  resultList      = new List<Common.Order>();
            string              outputFileName  = string.Empty;
            int                 bookingDays     = 5;

            // filter only needed customer (all the customers in the list)
            foreach (Common.Customer customer in Common.customerList)
            {
                resultList.AddRange(filterCustomersByName(customer.name));
            }

            // filter only tomorrow's loading dates
            // filter only loadings sent from the country of the agent
            // order by consignee
            resultList = resultList.Where(x => x.sailingDate.Date == DateTime.Now.AddDays(bookingDays).Date && 
                                               x.MBL.ToLower().StartsWith(shippingCompany.id.ToLower()))
                                   .OrderBy(x => x.sailingDate)
                                   .ToList();

            // check if customer has orders
            if (resultList.Count == 0)
            {
                OrdersParser._Form.log(string.Format("no new bookings in {0} days for {1}", bookingDays, shippingCompany.shippingLine));
                return;
            }

            OrdersParser._Form.log(string.Format("{0} bookings in {1} days found for {2}", resultList.Count, bookingDays, shippingCompany.shippingLine));

            // not all the columns are needed in the report - remove some
            List<Common.BookingsReport> targetResList = resultList.ConvertAll(x => new Common.BookingsReport
            {
                jobNo       = x.jobNo,
                fromCountry = x.fromCountry,
                sailingDate = x.sailingDate,
                toCountry   = x.toCountry,
                toPlace     = x.toPlace,
                vessel      = x.vessel,
                voyage      = x.voyage,
                MBL         = x.MBL,
            });

            string htmlTableStr = addHtmlTable<Common.BookingsReport>(targetResList);

            // send mail to customer
            MailDetails mailDetails = new MailDetails();
            mailDetails.mailRecepient = shippingCompany;
            mailDetails.mailType = Common.MailType.BookingConfirmation;
            mailDetails.bodyParameters = new Dictionary<string, string>()
            {
                { "table"       , htmlTableStr },
                { "agent"       , shippingCompany.name},
            };
            mailDetails.subject = string.Format("Booking confirmation for the upcoming {0} days", bookingDays);
            mailDetails.attachments = new List<string>() { };
            mailDetails.bHighImportance = false;

            // send mail
            sendMail(mailDetails);
        }

        // function filters full order list by customer name
        // and returns only the relevant customers matching the name
        // TODO: move to utils class
        public static List<Common.Order> filterCustomersByName(string customerName)
        {
            string name = customerName.ToLower();
            List<Common.Order> res = new List<Common.Order>();

            // most of the customers having long names, therefore, 'contains' is enough
            // while, some customers having short name (e.g. bg), and 'contains' is useless
            // for customers with name of 2 chars, try 'starts with' or starts with dots e.g. b.g.
            if (customerName.Length > 2)
            {
                res = Common.orderList.Where(x => x.consignee.ToLower().Contains(name)).ToList();
            }
            else
            {
                // generate the name with dots between the chars
                string nameWithDots = string.Join(".", name.ToCharArray()) + ".";

                // try 'starts with'
                res = Common.orderList.Where(x => x.consignee.ToLower().StartsWith(name) ||
                                                  x.consignee.ToLower().StartsWith(nameWithDots)).ToList();
            }

            // change customer name (formatting)
            res.ForEach(x => x.consignee = name.ToUpper());

            return res;
        }

        // function goes into the inbox and looks for the orders excel file
        // this file is sent on a daily basis, so need to load the latest one
        public static void readLastOrdersFile()
        {
            OutlookInterop.MAPIFolder SentMail = null;

            try
            {
                SentMail = outlookApp.ActiveExplorer().Session.GetDefaultFolder(OutlookInterop.OlDefaultFolders.olFolderInbox);
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to open outlook inbox folder. Error: {0}", e.Message), OrdersParser.logLevel.error);
                return;
            }

            // get all the items in inbox and sort by received date in descending form
            OutlookInterop.Items inboxItems = SentMail.Items;
            inboxItems.Sort("[ReceivedTime]", true);

            // loop over each mail having attachments and look for Tanco_Planned_Import_to_IL file
            foreach (OutlookInterop.MailItem newEmail in inboxItems)
            {
                if (newEmail != null)
                {
                    // filter out all senders except the one I'm looking for
                    if (newEmail.Sender.Address != tancoOrdersEmail) continue;

                    // loop over all the attachments
                    foreach (OutlookInterop.Attachment item in newEmail.Attachments)
                    {
                        var test = item.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E");
                        if (string.IsNullOrEmpty((string)test))
                        {
                            // real attachment (not embedded pic/logo)
                            // check if attachment matches, since sende might send other mails as well
                            if (item.DisplayName.ToLower().Contains(tancoOrdersFileName.ToLower()))
                            {
                                // match - save this attachment into file and bail out
                                OrdersParser._Form.log("Found Tanco shipping to IL file");
                                OrdersParser._Form.log(string.Format("Received date: {0}, FileName: {1}, Sender: {2}", 
                                                                      newEmail.ReceivedTime.ToString(Common.DATE_FORMAT),
                                                                      item.DisplayName,
                                                                      newEmail.Sender.Address));

                                // update global var
                                Common.plannedImportExcel = Path.Combine(Utils.resultsDirectoryPath, item.FileName);

                                // save file into temp folder
                                item.SaveAsFile(Common.plannedImportExcel);
                                return;
                            }
                        }
                    }
                }
            }

            // no file was found
            OrdersParser._Form.log(string.Format("Failed to find any Tannco excel file. Searched for {0} mails in inbox", 
                                                  inboxItems.Count), OrdersParser.logLevel.error);

            // plannedImportExcel should be automatically extracted from outlook
            // if file is empty, it means that extraction failed therefore,
            // user is alsked to provide a file 
            if (string.IsNullOrEmpty(Common.plannedImportExcel) == true)
            {
                OrdersParser._Form.log("Choose Tanco orders excel file");
#if OFFLINE
                Common.plannedImportExcel = @"C:\Users\rfridbur\Downloads\Tanco_Planned_Import_to_IL.xls";
#else
                // displays an OpenFileDialog so the user can select excel file
                Thread t = new Thread((ThreadStart)(() => 
                {
                    OpenFileDialog fileDialog = new OpenFileDialog();
                    fileDialog.Filter = "All Excel Files|*.xls*";
                    fileDialog.Title = "Select Excel File";
                    
                    // show dialog and import file
                    if (fileDialog.ShowDialog() == DialogResult.OK)
                    {
                        Common.plannedImportExcel = fileDialog.FileName;
                    }
                }));

                t.SetApartmentState(ApartmentState.STA);
                t.Start();
                t.Join();
#endif
            }
        }

        // basic template function to send mail to specific customer
        private static void sendMail(MailDetails mailDetails)
        {
            OutlookInterop._MailItem    oMailItem           = null;
            OutlookInterop.Attachment   attachment          = null;
            string                      bodyMsg             = string.Empty;
            string                      signatureName       = SIGNATURE_LOGO_NAME;
            string                      signatureFilePath   = Path.Combine(Directory.GetCurrentDirectory(), signatureName);
            string                      tempMsg             = string.Empty;

            // generate logo file from embeded resource (image)
            if (File.Exists(signatureFilePath) == false)
            {
                ImageConverter converter = new ImageConverter();
                byte[] tempByteArr = (byte[])converter.ConvertTo(Anko.Properties.Resources.logo, typeof(byte[]));
                File.WriteAllBytes(signatureFilePath, tempByteArr);
            }

            try
            {
                // initiate outlook parameters
                oMailItem = (OutlookInterop._MailItem)outlookApp.CreateItem(OutlookInterop.OlItemType.olMailItem);
            }
            catch (Exception e)
            {
                OrdersParser._Form.log(string.Format("Failed to create mail. Error: {0}", e.Message), OrdersParser.logLevel.error);
                dispose();
                return;
            }

            OutlookInterop.Recipient mailTo;
            OutlookInterop.Recipient mailCc;
            OutlookInterop.Recipients toList = oMailItem.Recipients;
            OutlookInterop.Recipients ccList = oMailItem.Recipients;

            // TO
            foreach (String recipient in mailDetails.mailRecepient.to)
            {
                mailTo = toList.Add(recipient);
                mailTo.Type = (int)OutlookInterop.OlMailRecipientType.olTo;
                mailTo.Resolve();
            }

            // CC
            foreach (String recipient in mailDetails.mailRecepient.cc)
            {
                mailCc = ccList.Add(recipient);
                mailCc.Type = (int)OutlookInterop.OlMailRecipientType.olCC;
                mailCc.Resolve();
            }

            // add Subject
            oMailItem.Subject = mailDetails.subject;

            // add attachment excel file
            foreach (string filePath in mailDetails.attachments)
            {
                oMailItem.Attachments.Add(filePath, 
                                          OutlookInterop.OlAttachmentType.olByValue, 
                                          1, 
                                          Path.GetFileName(filePath));
            }

            // verify that sifnature file exists
            if (File.Exists(signatureFilePath) == true)
            {
                // prepare body
                attachment = oMailItem.Attachments.Add(signatureFilePath,
                                                       OutlookInterop.OlAttachmentType.olEmbeddeditem,
                                                       null,
                                                       string.Empty);

                attachment.PropertyAccessor.SetProperty("http://schemas.microsoft.com/mapi/proptag/0x3712001E", signatureName);
            }
            else
            {
                OrdersParser._Form.log(string.Format("Cannot find signatue logo in {0}", signatureFilePath), OrdersParser.logLevel.error);
            }

            bodyMsg += addHtmlPreffix();

            // message (based on templates from mailType)
            bodyMsg += addMailBodyMsg(mailDetails.mailType, mailDetails.bodyParameters);

            // signature
            bodyMsg += addMailSignature(signatureName);

            oMailItem.BodyFormat = OutlookInterop.OlBodyFormat.olFormatHTML;
            oMailItem.HTMLBody = bodyMsg;

            // mark as high importance
            if (mailDetails.bHighImportance)
            {
                oMailItem.Importance = OutlookInterop.OlImportance.olImportanceHigh;
            }

            // async - display mail and proceed
            oMailItem.Display(false);
        }

        // function adds mail body into HTML format, based on the following parameters:
        // * mailType:       needed to decide which template from embeded resource to use
        // * bodyParameters: the template might have variabls, which are needed to be replaced from this dic
        private static string addMailBodyMsg(Common.MailType mailType, Dictionary<string, string> bodyParameters)
        {
            string              embededResource = Utils.getResourceNameFromMailType(mailType);
            string              msg             = string.Empty;
            string              modifiedText    = string.Empty;
            string              font            = "calibri";
            HorizontalAlignment alignment       = HorizontalAlignment.Left;

            // parse the mail body from embeded file into HTML format
            // replace all the parameter (if exist) by values from disctionary
            modifiedText = Utils.extractParameterFromDictionary(embededResource, bodyParameters);

            // determine the language to know which HTML params to set
            if (Utils.isHebrewText(modifiedText))
            {
                font = "arial";
                alignment = HorizontalAlignment.Right;
            }

            // add into HTML line by line
            foreach (string line in modifiedText.Split(new string[] { Environment.NewLine }, StringSplitOptions.None))
            {
                msg += addHtmlLine(alignment, "black", line, 3, font);
            }

            return msg;
        }

        // function adds signature into HTML format
        // assumption: signatureName (logo) is already added as embeded attachement
        //             otherwise, this addition will not do much
        private static string addMailSignature(string signatureName)
        {
            string text = string.Empty;
            string msg  = string.Empty;

            // parse the signature from embeded file into HTML format
            text = Anko.Properties.Resources.SignatureText;
            foreach (string line in text.Split(new string[] { Environment.NewLine }, StringSplitOptions.None))
            {
                msg += addHtmlLine(HorizontalAlignment.Left, "blue", line, 3);
            }

            // add the signature logo
            // (added as embedede atachment outside of this func - here we just add into HTML)
            msg += string.Format("<div align=left; dir=ltr><img src=\"cid:{0}\" width=246 height=213></div>", signatureName);

            return msg;
        }

        // mail template
        class MailDetails
        {
            public Common.MailingRecepient      mailRecepient;
            public Common.MailType              mailType;
            public Dictionary<string, string>   bodyParameters;
            public string                       subject;
            public List<string>                 attachments;
            public bool                         bHighImportance;
        }
    }
}

﻿
using System;
using System.Collections.Generic;

namespace emailTest
{
    class Common
    {
        public static string DATE_FORMAT = "dd-MM-yyyy";
        public static string plannedImportExcel = string.Empty;

        // customer class holds all customer details
        // customers purchase products from the company
        public class Customer : MailingRecepient
        {
            public bool bSendReport;
        }

        // agent class holds all agent details
        // agents are internal works
        public class Agent : MailingRecepient
        {
            // Ctor
            // * initialize 'countries' list
            public Agent()
            {
                countries = new List<string>();
            }

            // consider using CultureInfo
            public List<string> countries;
        }

        // shippingCompany class holds all shipping companies and agenst
        // e.g. ZIM, MSC
        public class ShippingCompany : MailingRecepient
        {
            public string shippingLine;
            public string id;
        }

        // mail recepient holds all mailing details
        public class MailingRecepient
        {
            // Ctor
            // * initialize 'to' and 'cc' lists
            public MailingRecepient()
            {
                to = new List<string>();
                cc = new List<string>();
            }

            public string name;
            public List<string> to;
            public List<string> cc;
        }

        public static List<Customer> customerList;

        public static List<Agent> agentList;

        public static List<ShippingCompany> shippingCompanyList;

        // order class according to Tanko import excel
        // columns are hardcoded and static
        public class Order
        {
            public int      jobNo;          // A
            public int      consignmentNum; // B
            public string   customer;       // C
            public string   shipper;        // D
            public string   consignee;      // E
            public string   customerRef;    // F
                                            // G (empty column)
            public string   tankNum;        // H
            public string   activity;       // I
            public DateTime loadingDate;    // J
            public string   fromCountry;    // K
                                            // L (empty column)
            public string   fromPlace;      // M
            public DateTime sailingDate;    // N
            public string   toCountry;      // O
            public string   toPlace;        // P
            public DateTime arrivalDate;    // Q
            public string   productName;    // R
            public string   vessel;         // S
            public string   voyage;         // T
            public string   MBL;            // U
            public string   arrivalStatus;  // V
        }

        public static List<Order> orderList;

        // mail types and formats existing in the system
        public enum MailType
        {
            Reports,
            LoadingConfirmation,
            BookingConfirmation,
            DocumentsReceipts
        }

        // subclass of Order
        public class OrderReport
        {
            public int      jobNo;          // A
            public string   shipper;        // D
            public string   consignee;      // E
            public string   customerRef;    // F
            public string   tankNum;        // H
            public string   activity;       // I
            public DateTime loadingDate;    // J
            public string   fromCountry;    // K
            public string   fromPlace;      // M
            public DateTime sailingDate;    // N
            public string   toCountry;      // O
            public string   toPlace;        // P
            public DateTime arrivalDate;    // Q
            public string   productName;    // R
            public string   vessel;         // S
            public string   voyage;         // T
        }

        // subclass of Order
        public class LoadingReport
        {
            public int      jobNo;          // A
            public string   consignee;      // E
            public DateTime loadingDate;    // J
            public string   fromCountry;    // K
        }

        // subclass of Order
        public class ArrivalsReport
        {
            public int      jobNo;          // A
            public string   consignee;      // E
            public DateTime arrivalDate;    // Q
        }

        // subclass of Order
        public class SailsReport
        {
            public int      jobNo;          // A
            public string   shipper;        // D
            public string   consignee;      // E
            public string   tankNum;        // H
            public string   fromCountry;    // K
            public DateTime sailingDate;    // N
        }

        // subclass of Order
        public class BookingsReport
        {
            public int      jobNo;          // A
            public string   fromCountry;    // K
            public DateTime sailingDate;    // N
            public string   toCountry;      // O
            public string   toPlace;        // P
            public string   vessel;         // S
            public string   voyage;         // T
            public string   MBL;            // U
        }
    }
}
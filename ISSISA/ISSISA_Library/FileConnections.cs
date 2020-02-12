using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Data.OleDb;
using System.Data;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;
using System.ComponentModel;
using System.Threading.Tasks;
using ISSIAS_Library;
using OfficeOpenXml;

namespace ISSISA_Library
{
    public class FileConnections
    {
        //This is the dbContext using Entity Framework to grab data
        private readonly FATSContext _db;

        //This is used for exporting excel document
        private Excel.Application xlApp = new Excel.Application();

        //these lists contain the file properties that are being imported or exported
        public List<fileNaming> files = new List<fileNaming>();
        public List<fileNaming> finished_files = new List<fileNaming>();

        //these lists contain all the assets in the fiscal book, imported files, or assets that were found
        //as a result of comparting the previous two lists
        public BindingList<asset> fb_assets = new BindingList<asset>();
        public BindingList<asset> imported_devices = new BindingList<asset>();
        public BindingList<asset> found_devices = new BindingList<asset>();
        public BindingList<asset> missing_devices = new BindingList<asset>();

        public IEnumerable<asset> MissingAssets => fb_assets.Where(asset =>
            found_devices.All(found => found.asset_number != asset.asset_number)).ToList();


        public BindingList<asset> locationValidate_devices = new BindingList<asset>();
        public BindingList<asset> serialValidate_devices = new BindingList<asset>();
        public BindingList<asset> roomValidate_devices = new BindingList<asset>();
        public BindingList<asset> locationRoomValidate_devices = new BindingList<asset>();
        public BindingList<asset> locationSerialValidate_devices = new BindingList<asset>();
        public BindingList<asset> serialRoomValidate_devices = new BindingList<asset>();
        public BindingList<asset> locationRoomSerialValidate_devices = new BindingList<asset>();

        //this private member holds the file properties for the fiscal book
        private fileNaming _fiscal_book_address = new fileNaming("No File Selected!");

        //this public member is used to make sure the fiscal book is in fact correctly chosen. 
        public string fiscal_book_address
        {
            get => _fiscal_book_address.name;
            set
            {
                if (Path.GetExtension(value) == ".xlsx")
                {
                    _fiscal_book_address.name = Path.GetFileNameWithoutExtension(value);
                    _fiscal_book_address.path = value;
                    _fiscal_book_address.type = ".xlsx";
                }
                else if (value == null)
                {
                    _fiscal_book_address.name = "No File Selected!";
                }
                else
                {
                    throw new FileLoadException();
                }
            }
        }


        public FileConnections()
        {
            _db = new FATSContext();
            //_db.Database.Connection.Open();
        }

        //fb example: FY 2016 20160114
        //Sheet exists that must be called ISS Assets Inventory + year
        //calling funcitons: import_data       
        private void open_fiscal_book()
        {
            var year = fiscal_book_address.Substring(fiscal_book_address.IndexOf("FY ", StringComparison.Ordinal) + 3,
                4);
            var sheetName = "ISS Asset Inventory " + year;
            var fi = new FileInfo(_fiscal_book_address.path);
            using (var xlPackage = new ExcelPackage(fi))
            {
                var ws = xlPackage.Workbook.Worksheets[sheetName];
                var ignore = true;

                for (var i = 1; i < ws.Dimension.End.Row; i++)
                {
                    var cells = ws.Cells[i, 1, i, ws.Dimension.End.Column];

                    var result = cells.Where(c => c.Value != null)
                        .Where(n => n.Value.ToString() == "Asset #").ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    if (ignore)
                        ignore = false;
                    else
                    {
                        var multi = (cells.Value as object[,]);

                        var b = new asset(multi[0, 0],
                            multi[0, 1],
                            multi[0, 2],
                            multi[0, 3],
                            multi[0, 4],
                            multi[0, 5],
                            multi[0, 6],
                            multi[0, 7],
                            multi[0, 8],
                            multi[0, 9],
                            multi[0, 10],
                            multi[0, 11],
                            multi[0, 12],
                            multi[0, 13],
                            multi[0, 14],
                            multi[0, 15]);

                        if (b.serial_number == null && b.fats_serial_number == null)
                            missing_devices.Add(b);
                        else
                            fb_assets.Add(b);
                    }
                }
            }
        }


        //this function takes a string that is a path to a file to be imported. 
        //that string is broken up into its elements of full path, extension, and name
        //calling funcitons: add file button on form.
        public void add_file(string x)
        {
            if (IndexOf(x, files) != 1) return;
            var myFile = new fileNaming
            {
                path = x,
                name = Path.GetFileNameWithoutExtension(x),
                type = Path.GetExtension(x)
            };


            files.Add(myFile);
        }

        //removes the file from the import file list
        public void remove_file(fileNaming x)
        {
            files.RemoveAt(files.IndexOf(x));
        }


        //this function handles the importing of data from supported type csv files.
        //calling funcitons: open_file
        //csvType defines how data is going to be removed from the csv
        private void open_csv_file(fileNaming x, FileType csvType, string skipUntil = null, string breakAt = null)
        {
            bool hasSkip = false, hasBreak = false, ignore = true;
            string serial;

            if (skipUntil != null)
                hasSkip = true;
            if (breakAt != null)
                hasBreak = true;

            using (var csvParser = new TextFieldParser(x.path))
            {
                csvParser.TextFieldType = FieldType.Delimited;
                csvParser.SetDelimiters(",");
                csvParser.TrimWhiteSpace = true;
                csvParser.HasFieldsEnclosedInQuotes = true;
                while (!csvParser.EndOfData)
                {
                    var parts = csvParser.ReadFields().ToList();
                    if (hasSkip && ignore &&
                        !parts.Where(n => n.ToString().Replace("\"", "") == skipUntil).ToList().Any())
                        continue;
                    if (hasBreak && parts.Where(n => n.ToString() == breakAt).ToList().Any())
                        break;
                    if (ignore)
                        ignore = false;
                    else
                    {
                        var a = new asset();
                        switch (csvType)
                        {
                            //Tropos
                            case FileType.Tropos:
                                a.serial_number = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.status = parts.ElementAt(2);
                                a.physical_location = parts.ElementAt(3);
                                break;
                            //Wireless_Controllers
                            //File Name conflict in 2016 & 2017. different column names and locations. original code in comment
                            case FileType.WirelessControllers:
                                a.controller_name = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.physical_location = parts.ElementAt(2);
                                a.status = parts.ElementAt(3);
                                a.serial_number = parts.ElementAt(4);
                                a.model = parts.ElementAt(5);
                                break;
                            //Cisco Wireless Controllers
                            case FileType.CiscoWirelessControllers:
                                a.model = parts.ElementAt(0);
                                a.device_name = parts.ElementAt(1);
                                a.description = parts.ElementAt(3);
                                a.physical_location = parts.ElementAt(4);
                                a.contact = parts.ElementAt(5);
                                a.serial_number = parts.ElementAt(6);
                                break;
                            //aps_wireless
                            case FileType.ApsWireless:
                                a.device_name = parts.ElementAt(0);
                                a.mac_address = parts.ElementAt(1);
                                a.ip_address = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(3);
                                a.model = parts.ElementAt(4);
                                a.physical_location = parts.ElementAt(5);
                                a.controller_name = parts.ElementAt(6);
                                break;
                            //device type - UPS
                            case FileType.DeviceTypeUps:
                                if (parts.ToArray().Length > 6)
                                    a.serial_number = parts.ElementAt(3).Replace("\"", "");
                                else
                                    continue;
                                if (string.IsNullOrEmpty(a.serial_number) || a.serial_number.Contains("Serial Number"))
                                    continue;
                                a.ip_address = parts.ElementAt(0).Replace("\"", "");
                                a.hostname = parts.ElementAt(1).Replace("\"", "");
                                a.model = parts.ElementAt(2).Replace("\"", "");
                                a.firmware = parts.ElementAt(4).Replace("\"", "");
                                a.physical_location = parts.ElementAt(5).Replace("\"", "") + " " +
                                                      parts.ElementAt(6).Replace("\"", "");
                                break;
                            //Brocade switch
                            case FileType.BrocadeSwitch:

                                serial = parts.ElementAt(6).Replace("\"", "");
                                var serialList = serial.Split(';').ToList();
                                if (serialList.Last() == "")
                                    serialList.RemoveAt(serialList.Count - 1);
                                var r = new Regex(@"^Unit\s\d+\s-\s");
                                a.status = parts.ElementAt(0).Replace("\"", "");
                                a.device_name = parts.ElementAt(1).Replace("\"", "");
                                a.ip_address = parts.ElementAt(4).Replace("\"", "");
                                a.model = parts.ElementAt(8).Replace("\"", "");
                                a.firmware = parts.ElementAt(9).Replace("\"", "");
                                a.contact = parts.ElementAt(10).Replace("\"", "");
                                a.physical_location = parts.ElementAt(11).Replace("\"", "");
                                a.last_scanned = parts.ElementAt(12).Replace("\"", "");
                                a.source = x.name;
                                foreach (var s in serialList)
                                {
                                    var b = new asset(a);
                                    serial = s;
                                    if (s.Contains("Unit"))
                                    {
                                        serial = r.Replace(s, "");
                                    }

                                    b.serial_number = serial;

                                    imported_devices.Add(b);
                                }

                                break;
                            //Device List- total-assest
                            case FileType.DeviceListTotalAssest:

                                a.description = parts.ElementAt(0);
                                a.model = parts.ElementAt(1);
                                a.hostname = parts.ElementAt(2);
                                a.device_name = parts.ElementAt(3);
                                a.serial_number = parts.ElementAt(4);
                                a.ip_address = parts.ElementAt(5);
                                a.physical_location = parts.ElementAt(6).Replace("\"", "");
                                a.mac_address = parts.ElementAt(7);
                                a.contact = parts.ElementAt(8);
                                a.room_number = parts.ElementAt(9);
                                a.location = parts.ElementAt(10);

                                break;
                            //UPS yyyy Device List
                            case FileType.UpsDeviceList:
                                a.model = parts.ElementAt(0);
                                a.hostname = parts.ElementAt(1);
                                a.serial_number = parts.ElementAt(2);
                                a.ip_address = parts.ElementAt(3);
                                a.physical_location = parts.ElementAt(4);
                                a.device_name = parts.ElementAt(5);
                                a.contact = parts.ElementAt(6);
                                a.room_number = parts.ElementAt(7);
                                a.location = parts.ElementAt(8);
                                break;
                            //GBIC Transceiver Report
                            case FileType.GbicTransceiver:
                                a.ip_address = parts.ElementAt(0);
                                a.device_name = parts.ElementAt(1);
                                a.description = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(5);
                                break;
                            //LMS switch and Router report
                            case FileType.LmsSwitchAndRouterReport:
                                if (parts.ElementAt(0) == "")
                                    continue;
                                serial = parts.ElementAt(5);

                                a = imported_devices.FirstOrDefault(p => p.serial_number == serial);
                                var childSN = parts.ElementAt(7);
                                if (a == null)
                                {
                                    a = new asset();
                                    a.device_name = parts.ElementAt(0);
                                    a.physical_location = parts.ElementAt(1);
                                    a.contact = parts.ElementAt(2);
                                    a.model = parts.ElementAt(3);
                                    a.serial_number = serial;
                                    if (!string.IsNullOrEmpty(childSN) && childSN != "N/A")
                                    {
                                        AddChild(a, childSN, x);
                                    }

                                    imported_devices.Add(a);
                                }
                                else if (!string.IsNullOrEmpty(childSN) && childSN != "N/A")
                                {
                                    AddChild(a, childSN, x);
                                }

                                break;
                            //Detailed_Router_Report_-_Yearly_Inventory                        
                            case FileType.DetailedRouterReportYearlyInventory:
                                a.model = parts.ElementAt(0);
                                a.device_name = parts.ElementAt(1);
                                var index = parts.ElementAt(3).IndexOf("RELEASE SOFTWARE");
                                a.description = index != -1
                                    ? parts.ElementAt(3).Substring(0, index - 2)
                                    : parts.ElementAt(3);
                                a.description.Replace(Environment.NewLine, "");
                                a.physical_location = parts.ElementAt(5);
                                a.contact = parts.ElementAt(6);
                                a.serial_number = parts.ElementAt(7);
                                break;
                            //Switchrouterinventoryserialnumber
                            case FileType.SwitchRouterInventorySerialNumber:
                                a.device_name = parts.ElementAt(0);
                                a.physical_location = parts.ElementAt(1);
                                a.contact = parts.ElementAt(2);
                                a.model = parts.ElementAt(3);
                                a.serial_number = parts.ElementAt(5);
                                a.children = new List<string> {parts.ElementAt(7)};
                                break;
                            case FileType.SwitchSerialNoReportForInventory:
                                a.ip_address = parts.ElementAt(0);
                                a.device_name = parts.ElementAt(1);
                                a.description = parts.ElementAt(2);
                                a.status = parts.ElementAt(4);
                                a.serial_number = parts.ElementAt(5);
                                break;
                            //Wireless_APs_Yearly_Inventory_Report
                            case FileType.WirelessAPsYearlyInventoryReport:
                                a.device_name = parts.ElementAt(0);
                                a.model = parts.ElementAt(1);
                                a.controller_name = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(4);
                                break;
                            case FileType.BrocadeModuleReport:
                                a.description = parts.ElementAt(0);
                                a.serial_number = parts.ElementAt(1);
                                a.device_name = parts.ElementAt(2);
                                break;
                        }

                        //add source file value from what file it came from
                        a.source = x.name;
                        //this is because i did a deep copy from asset a to b. L because special case that has children SN 
                        if (a.serial_number != "" && csvType != FileType.BrocadeSwitch &&
                            csvType != FileType.LmsSwitchAndRouterReport)
                        {
                            imported_devices.Add(a);
                        }
                    }
                }
            }
        }


        //This is going to handle adding a child asset to the device list.
        private void AddChild(asset a, string childSN, fileNaming x)
        {
            a.children.Add(childSN);
            var b = new asset(a) {serial_number = childSN};
            b.children.Clear();
            b.master = a.serial_number;
            b.source = x.name;
            imported_devices.Add(b);
        }


        public enum FileType
        {
            Tropos,
            ApsWireless,
            CiscoWirelessControllers,
            WirelessControllers,
            BrocadeSwitch,
            DeviceListTotalAssest,
            UpsDeviceList,
            LmsSwitchAndRouterReport,
            DeviceTypeUps,
            DetailedRouterReportYearlyInventory,
            GbicTransceiver,
            SwitchRouterInventorySerialNumber,
            WirelessAPsYearlyInventoryReport,
            SwitchSerialNoReportForInventory,
            BrocadeModuleReport,
            BrocadeWired,
            TMSInventory
        }


        private void open_file(fileNaming x)
        {
            switch (x.type)
            {
                case ".xlsx":
                    if (x.name.Contains("Brocade Wired"))
                        open_xlsx_file(x, FileType.BrocadeWired, "Product Status");
                    else if (x.name.Contains("Wireless APs"))
                        open_xlsx_file(x, FileType.WirelessAPsYearlyInventoryReport, "AP Name");
                    else if (x.name.Contains("Wireless_APs"))
                        open_xlsx_file(x, FileType.WirelessAPsYearlyInventoryReport, "AP Name");
                    else if (x.name.Contains("UPS"))
                        open_xlsx_file(x, FileType.DeviceTypeUps);
                    break;
                case ".xls":
                    if (x.name.Contains("TMS-Inventory"))
                        open_xls_file(x, FileType.TMSInventory);
                    else if (x.name.Contains("GBIC"))
                        open_xls_file(x, FileType.GbicTransceiver, "DeviceIP Address");
                    else if (x.name.Contains("Detailed_Router_Report_-_Yearly_Inventory") ||
                             x.name.Contains("Detailed_Switch_Report_-_Yearly_Inventory") ||
                             x.name.Contains("Detailed Router") ||
                             x.name.Contains("Detailed Switch"))
                        open_xls_file(x, FileType.DetailedRouterReportYearlyInventory, "Product Series");
                    break;
                case ".csv":
                    if (x.name.Contains("Tropos"))
                        open_csv_file(x, FileType.Tropos, "InventorySerialNumber");
                    else if (x.name.Contains("aps_wireless"))
                        open_csv_file(x, FileType.ApsWireless, "AP Name", "Disassociated AP(s)");
                    else if (x.name.Contains("Cisco_Wireless_Controllers"))
                        open_csv_file(x, FileType.CiscoWirelessControllers, "Product Series");
                    else if (x.name.Contains("Wireless_Controllers"))
                        open_csv_file(x, FileType.WirelessControllers, "Controller Name");
                    //open_csv_file(x, 'W', "Controller Name");
                    else if (x.name.Contains("Brocade switch"))
                        open_csv_file(x, FileType.BrocadeSwitch, "Product Status");
                    else if (x.name.Contains("Device List- total-assest"))
                        open_csv_file(x, FileType.DeviceListTotalAssest, "Type");
                    else if (x.name.Contains("UPS") && x.name.Contains("Device List"))
                        open_csv_file(x, FileType.UpsDeviceList, "Model");
                    else if (x.name.Contains("LMS switch and Router report"))
                        open_csv_file(x, FileType.LmsSwitchAndRouterReport, "Device Name");
                    else if (x.name.Contains("device type - UPS"))
                        open_csv_file(x, FileType.DeviceTypeUps);
                    else if (x.name.Contains("Detailed_Router_Report_-_Yearly_Inventory") ||
                             x.name.Contains("Detailed_Switch_Report_-_Yearly_Inventory"))
                        open_csv_file(x, FileType.DetailedRouterReportYearlyInventory, "Product Series");
                    else if (x.name.Contains("GBIC"))
                        open_csv_file(x, FileType.GbicTransceiver, "DeviceIP Address");
                    else if (x.name.Contains("Switchrouterinventoryserialnumber"))
                        open_csv_file(x, FileType.SwitchRouterInventorySerialNumber, "Device Name");
                    else if (x.name.Contains("Switch_Serial_No._report_for_Inventory"))
                        open_csv_file(x, FileType.SwitchSerialNoReportForInventory, "DeviceIP Address");
                    else if (x.name.Contains("Wireless APs"))
                        open_csv_file(x, FileType.WirelessAPsYearlyInventoryReport, "AP Name");
                    else if (x.name.Contains("Brocade Module Report"))
                        open_csv_file(x, FileType.BrocadeModuleReport, "Description");
                    else
                        throw new NotSupportedException();
                    break;
                case ".txt":
                    if (x.name.Contains("Brocade transceivers"))
                        open_txt_brocade_transceivers(x);
                    else if (x.name.Contains("netscraper"))
                        OpenNetscraperDump(x);
                    else if (x.name.Contains("Cisco Dump"))
                        open_text_dump(x);
                    else if (x.name.Contains("failed inventory collectionLMS"))
                        open_text_failedLMS(x);
                    else if (x.name.Contains("Brocade"))
                        open_text_Brocade(x);
                    else if (x.name.ToLower().Contains("inv raw"))
                        open_text_LMS_show_inv(x);
                    break;
                default:
                    throw new NotSupportedException();
            }
        }

        private void open_text_LMS_show_inv(fileNaming x)
        {
            var strings = File.ReadAllLines(x.path).ToList();
            string device = "", description = "";
            var splitter = new[] {"SN:"};
            foreach (var element in strings)
            {
                if (element.Contains("Device Name"))
                    device = element.Split(':')[1].Trim();
                else if (element.Contains("DESCR:"))
                    description = element.Substring(element.IndexOf("DESCR:") + 6).Trim().Trim('"');
                else if (element.Contains("SN:"))
                {
                    var splitStr = element.Split(splitter, StringSplitOptions.None);
                    if (splitStr.Length <= 1) continue;
                    var serial = splitStr[1].Trim();
                    if (string.IsNullOrEmpty(serial)) continue;
                    if (imported_devices.Any(im_device => im_device.serial_number.Equals(serial))) continue;
                    var a = new asset
                    {
                        device_name = device,
                        description = description,
                        serial_number = serial,
                        source = x.name
                    };
                    imported_devices.Add(a);
                }
            }
        }

        private void open_text_failedLMS(fileNaming x)
        {
            string device = "", description = "";
            using (var sr = new StreamReader(x.path))
            {
                //this loops through each line in the file
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    if (line.Contains("#"))
                        device = line.Substring(0, line.IndexOf('#')).Trim();
                    else if (line.Contains("DESCR:"))
                    {
                        var desc = line.Substring(line.IndexOf("DESCR:") + 6).Trim().Trim('"');
                        if (!string.IsNullOrEmpty(desc))
                            description = desc;
                    }
                    else if (line.Contains("SN:"))
                    {
                        var serial = line.Substring(line.IndexOf("SN:") + 3).Trim();
                        if (string.IsNullOrEmpty(serial)) continue;
                        if (imported_devices.Any(im_device => im_device.serial_number.Equals(serial))) continue;
                        var a = new asset
                        {
                            device_name = device,
                            description = description,
                            serial_number = serial,
                            source = x.name
                        };
                        imported_devices.Add(a);
                        description = "";
                    }
                }

                sr.Close();
            }
        }

        private void open_text_Brocade(fileNaming x)
        {
            var device = "";
            using (var sr = new StreamReader(x.path))
            {
                //this loops through each line in the file
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim();
                    if (string.IsNullOrEmpty(line)) continue;
                    if (line.Contains("SSH@"))
                        device = line.Split('#')[0].Substring(4).Trim();
                    else if (line.Contains("Serial#:"))
                    {
                        var serial = line.Substring(line.IndexOf("Serial#:") + 8).Trim();
                        if (string.IsNullOrEmpty(serial)) continue;
                        if (imported_devices.Count(im_device => im_device.serial_number.Equals(serial)) !=
                            0) continue;
                        var a = new asset
                        {
                            device_name = device,
                            serial_number = serial,
                            source = x.name
                        };
                        imported_devices.Add(a);
                    }
                }

                sr.Close();
            }
        }

        private void open_xls_file(fileNaming x, FileType xlsType, string skipUntil = null, string breakAt = null)
        {
            var ignore = !(skipUntil == null && breakAt == null);

            string sheetName;

            var con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + x.path +
                                          @";Extended Properties=""Excel 8.0;HDR=YES;""");
            con.Open();
            var dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
                new object[] {null, null, null, "TABLE"});
            switch (xlsType)
            {
                case FileType.TMSInventory:
                    sheetName = "TMS Excel Export 1$";
                    break;
                case FileType.GbicTransceiver:

                    sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                    break;
                case FileType.DetailedRouterReportYearlyInventory:

                    sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
                    break;
                default:
                    sheetName = "Sheet1$";
                    break;
            }


            try
            {
                //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference
                var myDataSet = new DataSet();
                var myCommand = new OleDbDataAdapter(" SELECT * from [" + sheetName + "]", con);
                myCommand.Fill(myDataSet);
                con.Close();

                //Travers through each row in the dataset
                foreach (DataRow myDataRow in myDataSet.Tables[0].Rows)
                {
                    //Stores info in Datarow into an array
                    var cells = myDataRow.ItemArray;
                    //Traverse through each array and put into object cellContent as type Object
                    //Using Object as for some reason the Dataset reads some blank value which
                    //causes a hissy fit when trying to read. By using object I can convert to
                    //String at a later point.
                    var result = cells.Where(n => n.ToString() == skipUntil).ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    if (ignore)
                        ignore = false;
                    else
                    {
                        var a = new asset();
                        switch (xlsType)
                        {
                            case FileType.TMSInventory:
                                a.device_name = Convert.ToString(cells[0]);
                                a.ip_address = Convert.ToString(cells[1]);
                                a.description = Convert.ToString(cells[2]) + " (" + Convert.ToString(cells[4]) + ")";
                                a.serial_number = Convert.ToString(cells[3]);
                                break;
                            case FileType.GbicTransceiver:
                                a.ip_address = Convert.ToString(cells[0]);
                                a.device_name = Convert.ToString(cells[1]);
                                a.description = Convert.ToString(cells[2]);
                                a.serial_number = Convert.ToString(cells[5]);
                                break;
                            case FileType.DetailedRouterReportYearlyInventory:
                                a.model = Convert.ToString(cells[0]);
                                a.device_name = Convert.ToString(cells[1]);
                                a.ip_address = Convert.ToString(cells[2]);
                                var index = Convert.ToString(cells[4])
                                    .IndexOf("RELEASE SOFTWARE", StringComparison.Ordinal);
                                a.description = index != -1
                                    ? Convert.ToString(cells[4]).Substring(0, index - 2)
                                    : Convert.ToString(cells[4]);
                                a.description.Replace(Environment.NewLine, "");
                                a.physical_location = Convert.ToString(cells[5]);
                                a.contact = Convert.ToString(cells[6]);
                                a.serial_number = Convert.ToString(cells[7]);
                                break;
                        }

                        a.source = x.name;
                        imported_devices.Add(a);
                    }
                }
            }
            finally
            {
                con.Close();
            }
        }

        private void open_xlsx_file(fileNaming x, FileType xlsType, string skipUntil = null, string breakAt = null)
        {
            var ignore = !(skipUntil == null && breakAt == null);
            //string sheetName;
            //using (var con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + x.path +
            //                                     @";Extended Properties=""Excel 8.0;HDR=YES;"""))
            //{
            //    con.Open();
            //    var dtSchema = con.GetOleDbSchemaTable(OleDbSchemaGuid.Tables,
            //        new object[] {null, null, null, "TABLE"});
            //    sheetName = dtSchema.Rows[0].Field<string>("TABLE_NAME");
            //    con.Close();
            //}

            var fi = new FileInfo(x.path);
            using (var xlPackage = new ExcelPackage(fi))
            {
                var ws = xlPackage.Workbook.Worksheets[xlPackage.Workbook.Worksheets.First().Name];

                for (var i = 1; i < ws.Dimension.End.Row; i++)
                {
                    var cells = ws.Cells[i, 1, i, ws.Dimension.End.Column];

                    var result = cells.Where(c => c.Value != null)
                        .Where(n => n.Value.ToString() == skipUntil).ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    if (ignore)
                        ignore = false;
                    else
                    {
                        var multi = (cells.Value as object[,]);
                        if (multi == null) continue;
                        string serial;
                        switch (xlsType)
                        {
                            case FileType.BrocadeWired:
                                serial = multi[0, 6] as string;
                                if (string.IsNullOrEmpty(serial)) continue;
                                var serials = Regex.Replace(serial, @"Unit \d - ", "");
                                foreach (var s in serials.Split(new[] {';'}, StringSplitOptions.RemoveEmptyEntries))
                                {
                                    var brocadeAsset = new asset
                                    {
                                        device_name = multi[0, 1] as string,
                                        ip_address = multi[0, 4] as string,
                                        asset_type = multi[0, 5] as string,
                                        serial_number = s,
                                        status = multi[0, 7] as string,
                                        model = multi[0, 8] as string,
                                        firmware = multi[0, 9] as string,
                                        contact = multi[0, 10] as string,
                                        location = multi[0, 11] as string,
                                        last_scanned = multi[0, 12] as string,
                                        source = x.name
                                    };

                                    imported_devices.Add(brocadeAsset);
                                }

                                break;
                            case FileType.WirelessAPsYearlyInventoryReport:
                                var apAsset = new asset
                                {
                                    device_name = multi[0, 0] as string,
                                    model = multi[0, 1] as string,
                                    controller_name = multi[0, 2] as string,
                                    serial_number = multi[0, 4] as string,
                                    source = x.name
                                };

                                imported_devices.Add(apAsset);

                                break;
                            case FileType.DeviceTypeUps:
                                serial = multi[0, 3] as string;
                                if (string.IsNullOrEmpty(serial) || serial.Contains("Serial Number"))
                                    continue;
                                var upsAsset = new asset
                                {
                                    serial_number = serial,
                                    ip_address = multi[0, 0] as string,
                                    hostname = multi[0, 1] as string,
                                    model = multi[0, 2] as string,
                                    firmware = multi[0, 4] as string,
                                    physical_location = $"{multi[0, 5] as string} {multi[0, 6] as string}",
                                    source = x.name
                                };
                                imported_devices.Add(upsAsset);
                                break;
                        }
                    }
                }
            }
        }


        //this function handles importing data from a text file
        //calling functions: open_file
        private void open_text_dump(fileNaming x)
        {
            using (var sr = new StreamReader(x.path))
            {
                //this loops through each line in the file
                string line;
                while ((line = sr.ReadLine()) != null)
                {
                    //this attempts to see if the string exists in the line if it does continue
                    var index = line.IndexOf("Device Name");

                    if (index == -1) continue;
                    //grab the device name for all the corresponding serials
                    var deviceName = line.Substring(15);
                    //string masterSN = "";
                    var serials = new List<string>();

                    //this loops through each line after device has been found until it hits #.
                    while ((line = sr.ReadLine()) != null && (line.IndexOf("#") == -1 || line.IndexOf("Serial#") > 1))
                    {
                        //logical error checking for stupid data
                        if (line.IndexOf("SN: 0x") != -1 || line.IndexOf("SN: N/A") != -1 ||
                            line.IndexOf("NAME:") != -1) continue;
                        //if a serial number exists grab the data
                        if ((index = line.IndexOf("SN:")) <= 1) continue;
                        line = line.Substring(index + 4);
                        //if serial is not null
                        if (line != "")
                        {
                            //if it is the first serial number make it the master
                            //if (masterSN == "")
                            //    masterSN = line;
                            serials.Add(line);
                        }
                    }

                    foreach (var sn in serials)
                    {
                        var a = new asset
                        {
                            serial_number = sn,
                            device_name = deviceName
                        };
                        if (serials.ElementAt(0) == sn && serials.Count > 1)
                            a.children = serials.Skip(1).ToList();
                        else
                            a.master = serials.ElementAt(0);
                        a.source = x.name;
                        imported_devices.Add(a);
                    }

                    //foreach (string sn in serials)
                    //{
                    //    asset a = new asset();
                    //    a.serial_number = sn;
                    //    a.device_name = deviceName;
                    //    if (masterSN == sn)
                    //        a.description = "Master SN";
                    //    a.source = x.name;
                    //    imported_devices.Add(a);
                    //}
                }

                sr.Close();
            }
        }


        private void OpenNetscraperDump(fileNaming x)
        {
            var validSn = new[]
            {
                "SN:",
                "Serial Number:",
                "Serial Number :",
                "Serial#:",
                "Daughterboard serial number     :"
            };


            using (var sr = new StreamReader(x.path))
            {
                string line;
                var deviceName = "";
                var description = "";
                while ((line = sr.ReadLine()) != null)
                {
                    if (string.IsNullOrEmpty(line)) continue;


                    if (line.Last() == '#' && line.Count(c => c == '#') == 1)
                    {
                        deviceName = line.Split('#')[0];
                        description = "";

                        continue;
                    }

                    if (line.Contains("DESCR:"))
                    {
                        var indexOf = line.IndexOf("DESCR:", StringComparison.Ordinal) + 6;
                        var desc = line.Substring(indexOf).Trim().Trim('"');
                        if (string.IsNullOrEmpty(desc)) continue;
                        description = desc;
                    }

                    var sn = "";
                    var notes = "";
                    foreach (var s in validSn)
                    {
                        if (!line.Contains(s)) continue;
                        var indexOf = line.IndexOf(s, StringComparison.Ordinal) + s.Length;
                        sn = line.Substring(indexOf).Trim();
                        var trimIndex = sn.IndexOf(' ');
                        if (trimIndex > 0)
                        {
                            notes = sn.Substring(trimIndex).Trim();
                            sn = sn.Substring(0, trimIndex);
                        }

                        break;
                    }


                    if (string.IsNullOrEmpty(sn) || sn == "N/A" || sn == "0x") continue;
                    if (imported_devices.Any(serial => serial.serial_number == sn)) continue;


                    var asset = new asset
                    {
                        serial_number = sn,
                        description = description,
                        device_name = deviceName,
                        source = x.name,
                        notes = notes
                    };

                    imported_devices.Add(asset);
                }

                sr.Close();
            }
        }

        private void open_txt_brocade_transceivers(fileNaming x)
        {
            using (var sr = new StreamReader(x.path))
            {
                var lines = new List<string>();
                //this loops through each line in the file
                string ip = "", device = "";
                var getNextDevice = false;
                var line = "";
                while ((line = sr.ReadLine()) != null)
                {
                    line = line.Trim(' ');
                    if (line != "")
                    {
                        var result = Regex.Match(line, @"^(?:[0-9]{1,3}\.){3}[0-9]{1,3}$").Value;
                        if (!string.IsNullOrEmpty(result))
                        {
                            ip = result;
                            if (lines.Last() != "")
                                device = lines.Last();
                            else
                                getNextDevice = true;
                        }
                        else if (line.Contains("Serial#"))
                        {
                            var serial = line.Substring(line.LastIndexOf(':') + 1).Trim(' ');
                            if (serial != null)
                            {
                                var a = new asset
                                {
                                    device_name = device,
                                    ip_address = ip,
                                    serial_number = serial,
                                    source = x.name
                                };
                                imported_devices.Add(a);
                                lines.Clear();
                            }
                        }
                        else if (getNextDevice)
                        {
                            getNextDevice = false;
                            device = line;
                        }
                    }

                    lines.Add(line);
                }

                sr.Close();
            }
        }

        //this function handles writing the compared data to excel file.
        //calling functions: save button on form.
        public void write_to_excel(string x, BindingList<asset> exportList)
        {
            if (xlApp == null)
            {
                Console.WriteLine(
                    "EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = false;
            var wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            var ws = (Excel.Worksheet) wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine(
                    "Worksheet could not be created. Check that your office installation and project references are correct.");
            }

            int rows = 1, columns = 1;

            //write the headers           

            ws.Cells[rows, columns++] = "Asset #";
            ws.Cells[rows, columns++] = "Missing " + fiscal_book_address.Substring(3, 4);
            ws.Cells[rows, columns++] = "ISS Divison";
            ws.Cells[rows, columns++] = "Description";
            ws.Cells[rows, columns++] = "Model";
            ws.Cells[rows, columns++] = "Asset Type";
            ws.Cells[rows, columns++] = "Location";
            ws.Cells[rows, columns++] = "Physical Location";
            ws.Cells[rows, columns++] = "Room Per Advantage #";
            ws.Cells[rows, columns++] = "Room Per FATS";
            ws.Cells[rows, columns++] = "Room Number";
            ws.Cells[rows, columns++] = "Cost";
            ws.Cells[rows, columns++] = "Last Inv";
            ws.Cells[rows, columns++] = "Serial #";
            ws.Cells[rows, columns++] = "Serial # Per FATS";
            ws.Cells[rows, columns++] = "Master SN";
            ws.Cells[rows, columns++] = "Children SN";
            ws.Cells[rows, columns++] = "FATS Owner";
            ws.Cells[rows, columns++] = "Notes";
            ws.Cells[rows, columns++] = "Status";
            ws.Cells[rows, columns++] = "Device Name";
            ws.Cells[rows, columns++] = "Mac Address";
            ws.Cells[rows, columns++] = "IP Address";
            ws.Cells[rows, columns++] = "Hostname";
            ws.Cells[rows, columns++] = "Controller Name";
            ws.Cells[rows, columns++] = "Firmware";
            ws.Cells[rows, columns++] = "Contact";
            ws.Cells[rows, columns++] = "Last Scanned";
            ws.Cells[rows++, columns] = "Source";

            columns = 1;
            //write the data
            //Parallel.ForEach(exportList, a =>
            //{

            foreach (var a in exportList.AsParallel())
            {
                ws.Cells[rows, columns++] = a.asset_number;
                ws.Cells[rows, columns++] = a.missing.ToString();
                ws.Cells[rows, columns++] = a.iss_division;
                ws.Cells[rows, columns++] = a.description;
                ws.Cells[rows, columns++] = a.model;
                ws.Cells[rows, columns++] = a.asset_type;
                ws.Cells[rows, columns++] = a.location;
                ws.Cells[rows, columns++] = a.physical_location;
                ws.Cells[rows, columns++] = a.room_per_advantage;
                ws.Cells[rows, columns++] = a.room_per_fats;
                ws.Cells[rows, columns++] = a.room_number;
                ws.Cells[rows, columns++] = a.cost;
                ws.Cells[rows, columns++] = a.last_inv.ToString();
                ws.Cells[rows, columns++] = a.serial_number;
                ws.Cells[rows, columns++] = a.fats_serial_number;
                ws.Cells[rows, columns++] = a.master;
                ws.Cells[rows, columns] = "";
                foreach (var child in a.children.AsParallel())
                {
                    var cellValue = (string) (ws.Cells[rows, columns] as Excel.Range).Value;
                    ws.Cells[rows, columns] = cellValue + child + ";";
                }

                columns++;
                ws.Cells[rows, columns++] = a.fats_owner;
                ws.Cells[rows, columns++] = a.notes;
                ws.Cells[rows, columns++] = a.status;
                ws.Cells[rows, columns++] = a.device_name;
                ws.Cells[rows, columns++] = a.mac_address;
                ws.Cells[rows, columns++] = a.ip_address;
                ws.Cells[rows, columns++] = a.hostname;
                ws.Cells[rows, columns++] = a.controller_name;
                ws.Cells[rows, columns++] = a.firmware;
                ws.Cells[rows, columns++] = a.contact;
                ws.Cells[rows, columns++] = a.last_scanned;
                ws.Cells[rows++, columns] = a.source;
                columns = 1;
            }

            //});
            //save the file using the param as the path and name
            ws.Columns.AutoFit();
            wb.SaveAs(x);

            //open the saved file

            open_excel_file(x);
        }

        public void WriteToExcel(string x, IEnumerable<asset> exportList)
        {
            byte[] file;

            using (var xlPackage = new ExcelPackage())
            {
                xlPackage.Workbook.Properties.Title = x;
                var ws = xlPackage.Workbook.Worksheets.Add("Sheet1");
                int rows = 1, columns = 1;
                ws.Cells[rows, columns++].Value = "Asset #";
                ws.Cells[rows, columns++].Value = "Missing " + fiscal_book_address.Substring(3, 4);
                ws.Cells[rows, columns++].Value = "ISS Divison";
                ws.Cells[rows, columns++].Value = "Description";
                ws.Cells[rows, columns++].Value = "Model";
                ws.Cells[rows, columns++].Value = "Asset Type";
                ws.Cells[rows, columns++].Value = "Location";
                ws.Cells[rows, columns++].Value = "Physical Location";
                ws.Cells[rows, columns++].Value = "Room Per Advantage #";
                ws.Cells[rows, columns++].Value = "Room Per FATS";
                ws.Cells[rows, columns++].Value = "Room Number";
                ws.Cells[rows, columns++].Value = "Cost";
                ws.Cells[rows, columns++].Value = "Last Inv";
                ws.Cells[rows, columns++].Value = "Serial #";
                ws.Cells[rows, columns++].Value = "Serial # Per FATS";
                ws.Cells[rows, columns++].Value = "Master SN";
                ws.Cells[rows, columns++].Value = "Children SN";
                ws.Cells[rows, columns++].Value = "FATS Owner";
                ws.Cells[rows, columns++].Value = "Notes";
                ws.Cells[rows, columns++].Value = "Status";
                ws.Cells[rows, columns++].Value = "Device Name";
                ws.Cells[rows, columns++].Value = "Mac Address";
                ws.Cells[rows, columns++].Value = "IP Address";
                ws.Cells[rows, columns++].Value = "Hostname";
                ws.Cells[rows, columns++].Value = "Controller Name";
                ws.Cells[rows, columns++].Value = "Firmware";
                ws.Cells[rows, columns++].Value = "Contact";
                ws.Cells[rows, columns++].Value = "Last Scanned";
                ws.Cells[rows++, columns].Value = "Source";

                columns = 1;
                //write the data
                //Parallel.ForEach(exportList, a =>
                //{

                foreach (var a in exportList.AsParallel())
                {
                    ws.Cells[rows, columns++].Value = a.asset_number;
                    ws.Cells[rows, columns++].Value = a.missing.ToString();
                    ws.Cells[rows, columns++].Value = a.iss_division;
                    ws.Cells[rows, columns++].Value = a.description;
                    ws.Cells[rows, columns++].Value = a.model;
                    ws.Cells[rows, columns++].Value = a.asset_type;
                    ws.Cells[rows, columns++].Value = a.location;
                    ws.Cells[rows, columns++].Value = a.physical_location;
                    ws.Cells[rows, columns++].Value = a.room_per_advantage;
                    ws.Cells[rows, columns++].Value = a.room_per_fats;
                    ws.Cells[rows, columns++].Value = a.room_number;
                    ws.Cells[rows, columns++].Value = a.cost;
                    ws.Cells[rows, columns++].Value = a.last_inv.ToString();
                    ws.Cells[rows, columns++].Value = a.serial_number;
                    ws.Cells[rows, columns++].Value = a.fats_serial_number;
                    ws.Cells[rows, columns++].Value = a.master;
                    ws.Cells[rows, columns].Value = "";
                    foreach (var child in a.children.AsParallel())
                    {
                        var cellValue = (string) (ws.Cells[rows, columns] as Excel.Range).Value;
                        ws.Cells[rows, columns].Value = cellValue + child + ";";
                    }

                    columns++;
                    ws.Cells[rows, columns++].Value = a.fats_owner;
                    ws.Cells[rows, columns++].Value = a.notes;
                    ws.Cells[rows, columns++].Value = a.status;
                    ws.Cells[rows, columns++].Value = a.device_name;
                    ws.Cells[rows, columns++].Value = a.mac_address;
                    ws.Cells[rows, columns++].Value = a.ip_address;
                    ws.Cells[rows, columns++].Value = a.hostname;
                    ws.Cells[rows, columns++].Value = a.controller_name;
                    ws.Cells[rows, columns++].Value = a.firmware;
                    ws.Cells[rows, columns++].Value = a.contact;
                    ws.Cells[rows, columns++].Value = a.last_scanned;
                    ws.Cells[rows++, columns].Value = a.source;
                    columns = 1;
                }

                xlPackage.SaveAs(new FileInfo(x));
            }
        }

        //this function handles the comparison between imported devices and assets from fiscal book
        //calling functions: import_data
        private void Compare()
        {
            //every asset in fical book is placed in found devices list

            //found_devices = fb_assets;

            //loops through each asset in list
            var serialsOnly = fb_assets.Where(f => !string.IsNullOrEmpty(f.serial_number)).ToList();
            var fatsSerials = fb_assets.Where(f => !string.IsNullOrEmpty(f.fats_serial_number)).ToList();

            foreach (var a in imported_devices)
            {
                //grab the asset from fiscal book that has matching serials from imported device data


                var existingAsset = serialsOnly.AsParallel()
                                        .FirstOrDefault(x => x.serial_number.Contains(a.serial_number))
                                    ?? fatsSerials.AsParallel()
                                        .FirstOrDefault(x => x.fats_serial_number.Contains(a.serial_number));

                if (existingAsset != null)
                {
                    var found = new asset(existingAsset);
                    UpdateFoundAsset(found, a);
                    found_devices.Add(found);
                }
                else
                {
                    //doesnt exist
                    var missing = new asset(a);
                    missing_devices.Add(missing);
                }
            }
        }

        private static void UpdateFoundAsset(asset a, asset b)
        {
            //update some of the found device fields by combining data
            if (!string.IsNullOrEmpty(b.description))
                a.description = a.description + " (" + b.description + ")";
            if (!string.IsNullOrEmpty(b.model))
                a.model = a.model + " (" + b.model + ")";
            if (!string.IsNullOrEmpty(b.physical_location))
                a.physical_location = a.physical_location + " (" + b.physical_location + ")";
            if (!string.IsNullOrEmpty(b.location) && !a.location.Contains(b.location))
                a.location = a.location + " (" + b.location + ")";
            a.serial_number = b.serial_number;
            a.status = b.status;
            a.device_name = b.device_name;
            a.mac_address = b.mac_address;
            a.ip_address = b.ip_address;
            a.controller_name = b.controller_name;
            a.source = b.source;
            a.hostname = b.hostname;
            a.firmware = b.firmware;
            //this field is used to show that a match has been found between data sets.
            a.found = true;
            b.found = true;
            a.children = b.children;
            a.master = b.master;
        }


        //used for debugging only
        public void show_imported_devices()
        {
            using (var sw = new StreamWriter("custs.txt"))
            {
                foreach (var a in imported_devices)
                {
                    sw.WriteLine(a.output());
                }
            }
        }

        //this is simply used to say if the file to be imported is alreay in the list or not
        //if 1 then it is not in the list. if 0 then the file is already selected to be imported
        //calling functions: add_file
        private int IndexOf(string x, IEnumerable<fileNaming> a)
        {
            x = Path.GetFileNameWithoutExtension(x);
            return a.Any(y => x != null && x.Equals(y.name)) ? 0 : 1;
        }

        //This clears all list data.
        //calling functions: run button on form -> run_button_Click 
        public void clear_data()
        {
            imported_devices.Clear();
            fb_assets.Clear();
            finished_files.Clear();
            found_devices.Clear();
            missing_devices.Clear();
            locationValidate_devices.Clear();
            serialValidate_devices.Clear();
            roomValidate_devices.Clear();
        }

        //this function makes the necessay calls to import all data from selected files the compare the data
        //calling functions: run button on form -> run_button_Click
        public void import_data()
        {
            //open fiscal book first
            open_fiscal_book();

            //open each file in the list
            foreach (var file in files)
            {
                open_file(file);
            }

            //then compare the data
            Compare();

            //compare Fats to Fiscal Book
            //compareFats();
            //get the current date
            var date = DateTime.Now.ToString("yyyyMMdd");

            //add report to finished files list
            finished_files.Add(new fileNaming(date + " Inventory Report "));
            finished_files.Add(new fileNaming(date + " Missing Inventory Report "));
            finished_files.Add(new fileNaming(date + " Missing Asset Report "));
        }


        /// <summary>
        /// work on getting functionality for fats working
        /// </summary>
        private void compareFats()
        {
            var wrongFatsData = new List<asset>();
            foreach (var fbAsset in fb_assets.AsParallel().Where(x => x.iss_division == "NETWORK"))
            {
                if (_db.FatsAsset.Any(x => x.AssetNumber == fbAsset.asset_number
                                           && (x.LocationCode != fbAsset.location
                                               || x.Room != fbAsset.room_per_fats)))
                    wrongFatsData.Add(fbAsset);
            }
        }


        //this function makes the necessay calls to import the data only from the inventory review file
        //calling functions: run button on form ->run_button_Click
        public void import_review_data()
        {
            open_review_book();


            //then validate the data
            Validate();
            //get the current date
            var date = DateTime.Now.ToString("yyyyMMdd");

            //add report to finished files list
            finished_files.Add(new fileNaming(date + " Inventory Report "));
        }

        private void open_review_book()
        {
            const string sheetName = "Advantage_FATS";
            var con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fiscal_book_address.path +
                                          ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"");
            con.Open();

            try
            {
                //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference
                var myDataSet = new DataSet();
                var myCommand = new OleDbDataAdapter(" SELECT * from [" + sheetName + "$]", con);
                var ignore = true;
                myCommand.Fill(myDataSet);
                con.Close();

                //Travers through each row in the dataset
                foreach (DataRow myDataRow in myDataSet.Tables[0].Rows)
                {
                    //Stores info in Datarow into an array
                    var cells = myDataRow.ItemArray;
                    //Traverse through each array and put into object cellContent as type Object
                    //Using Object as for some reason the Dataset reads some blank value which
                    //causes a hissy fit when trying to read. By using object I can convert to
                    //String at a later point.
                    var result = cells.Where(n => n.ToString() == "ASSET_NUMBER").ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    if (ignore)
                        ignore = false;
                    else
                    {
                        var b = new asset(cells[0],
                            cells[2],
                            cells[4],
                            cells[6],
                            cells[7],
                            cells[8],
                            cells[11],
                            cells[15],
                            cells[17],
                            cells[18],
                            cells[19],
                            cells[20],
                            cells[22],
                            cells[23]);
                        fb_assets.Add(b);
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                con.Close();
            }
        }

        private void Validate()
        {
            locationValidate_devices = new BindingList<asset>(
                fb_assets.Where(a => !a.location.Equals(a.physical_location))
                    .ToList());
            serialValidate_devices = new BindingList<asset>(
                fb_assets.Where(a =>
                        !(a.serial_number == "" && (a.fats_serial_number == "#N/A" || a.fats_serial_number == "")))
                    .Where(a => !Convert.ToString(a.serial_number).Equals(a.fats_serial_number))
                    .ToList());
            roomValidate_devices = new BindingList<asset>(
                fb_assets
                    .Where(a => !(a.room_per_advantage == "" && (a.room_per_fats == "#N/A" || a.room_per_fats == "")))
                    .Where(a => !a.room_per_advantage.Equals(a.room_per_fats))
                    .ToList());
            locationRoomValidate_devices = new BindingList<asset>((from loc in locationValidate_devices
                join room in roomValidate_devices on loc.asset_number equals room.asset_number
                select loc).ToList());
            locationSerialValidate_devices = new BindingList<asset>((from loc in locationValidate_devices
                join serial in serialValidate_devices on loc.asset_number equals serial.asset_number
                select loc).ToList());
            serialRoomValidate_devices = new BindingList<asset>((from serial in serialValidate_devices
                join room in roomValidate_devices on serial.asset_number equals room.asset_number
                select serial).ToList());
            locationRoomSerialValidate_devices = new BindingList<asset>((from locRoom in locationRoomValidate_devices
                join serialRoom in serialRoomValidate_devices on locRoom.asset_number equals serialRoom.asset_number
                select locRoom).ToList());

            locationValidate_devices = new BindingList<asset>(locationValidate_devices
                .Except(locationRoomValidate_devices, new assetEqualityComparer()).ToList());
            locationValidate_devices = new BindingList<asset>(locationValidate_devices
                .Except(locationSerialValidate_devices, new assetEqualityComparer()).ToList());

            serialValidate_devices = new BindingList<asset>(serialValidate_devices
                .Except(serialRoomValidate_devices, new assetEqualityComparer()).ToList());

            //ALL serialValidates are in locationSerial due to all records being included
            //serialValidate_devices = new BindingList<asset>(serialValidate_devices.Except(locationSerialValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());

            roomValidate_devices = new BindingList<asset>(roomValidate_devices
                .Except(serialRoomValidate_devices, new assetEqualityComparer()).ToList());
            //roomValidate_devices = new BindingList<asset>(roomValidate_devices.Except(locationRoomValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());
        }

        //this function handles writing the compared data to excel file.
        //calling functions: save button on form.
        public void write_validate_to_excel(string x)
        {
            ICollection<KeyValuePair<string, BindingList<asset>>> worksheets =
                new Dictionary<string, BindingList<asset>>();
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations", locationValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Serials", serialValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Rooms", roomValidate_devices));
            worksheets.Add(
                new KeyValuePair<string, BindingList<asset>>("Locations & Rooms", locationRoomValidate_devices));
            worksheets.Add(
                new KeyValuePair<string, BindingList<asset>>("Locations & Serials", locationSerialValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Serials & Rooms", serialRoomValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations & Rooms & Serials",
                locationRoomSerialValidate_devices));

            if (xlApp == null)
            {
                Console.WriteLine(
                    "EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = false;
            var wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);


            foreach (var worksheet in worksheets)
            {
                //Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[worksheet.Key];
                var ws = wb.Sheets.Count == 0
                    ? (Excel.Worksheet) wb.Sheets.Add()
                    : (Excel.Worksheet) wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                ws.Name = worksheet.Key;
                if (ws == null)
                {
                    Console.WriteLine(
                        "Worksheet could not be created. Check that your office installation and project references are correct.");
                }

                int rows = 1, columns = 1;

                //write the headers           

                ws.Cells[rows, columns++] = "Asset #";
                //ws.Cells[rows, columns++] = "Missing " + fiscal_book_address.Substring(3, 4);
                //ws.Cells[rows, columns++] = "ISS Divison";
                ws.Cells[rows, columns++] = "Description";
                ws.Cells[rows, columns++] = "Model";
                ws.Cells[rows, columns++] = "Asset Type";
                ws.Cells[rows, columns++] = "Advantage Location";
                ws.Cells[rows, columns++] = "FATS Location";
                ws.Cells[rows, columns++] = "Room Per Advantage #";
                ws.Cells[rows, columns++] = "Room Per FATS";
                //ws.Cells[rows, columns++] = "Room Number";
                //ws.Cells[rows, columns++] = "Cost";
                //ws.Cells[rows, columns++] = "Last Inv";
                ws.Cells[rows, columns++] = "Advantage Serial #";
                ws.Cells[rows, columns++] = "FATS Serial #";
                //ws.Cells[rows, columns++] = "Children SN";
                ws.Cells[rows, columns++] = "FATS Owner";
                //ws.Cells[rows, columns++] = "Notes";
                //ws.Cells[rows, columns++] = "Status";
                //ws.Cells[rows, columns++] = "Device Name";
                //ws.Cells[rows, columns++] = "Mac Address";
                ws.Cells[rows, columns++] = "IP Address";
                ws.Cells[rows++, columns] = "Hostname";
                //ws.Cells[rows, columns++] = "Controller Name";
                //ws.Cells[rows, columns++] = "Firmware";
                //ws.Cells[rows, columns++] = "Contact";
                //ws.Cells[rows, columns++] = "Last Scanned";
                //ws.Cells[rows++, columns] = "Source";

                columns = 1;
                //write the data
                foreach (var a in worksheet.Value)
                {
                    ws.Cells[rows, columns++] = a.asset_number;
                    //ws.Cells[rows, columns++] = a.missing.ToString();
                    //ws.Cells[rows, columns++] = a.iss_division;
                    ws.Cells[rows, columns++] = a.description;
                    ws.Cells[rows, columns++] = a.model;
                    ws.Cells[rows, columns++] = a.asset_type;
                    ws.Cells[rows, columns++] = a.location;
                    ws.Cells[rows, columns++] = a.physical_location;
                    ws.Cells[rows, columns++] = a.room_per_advantage;
                    ws.Cells[rows, columns++] = a.room_per_fats;
                    //ws.Cells[rows, columns++] = a.room_number;
                    //ws.Cells[rows, columns++] = a.cost;
                    //ws.Cells[rows, columns++] = a.last_inv.ToString();
                    ws.Cells[rows, columns++] = a.serial_number;
                    ws.Cells[rows, columns++] = a.fats_serial_number;
                    //ws.Cells[rows, columns] = "";
                    //foreach (string child in a.children)
                    //{
                    //    var cellValue = (string)(ws.Cells[rows, columns] as Excel.Range).Value;
                    //    ws.Cells[rows, columns] = cellValue + child + ";";
                    //}
                    //columns++;
                    ws.Cells[rows, columns++] = a.fats_owner;
                    //ws.Cells[rows, columns++] = a.notes;
                    //ws.Cells[rows, columns++] = a.status;
                    //ws.Cells[rows, columns++] = a.device_name;
                    //ws.Cells[rows, columns++] = a.mac_address;
                    ws.Cells[rows, columns++] = a.ip_address;
                    ws.Cells[rows++, columns] = a.hostname;
                    //ws.Cells[rows, columns++] = a.controller_name;
                    //ws.Cells[rows, columns++] = a.firmware;
                    //ws.Cells[rows, columns++] = a.contact;
                    //ws.Cells[rows, columns++] = a.last_scanned;
                    //ws.Cells[rows++, columns] = a.source;
                    columns = 1;
                }

                //save the file using the param as the path and name
                ws.Columns.AutoFit();
            }

            wb.SaveAs(x);

            //open the saved file

            open_excel_file(x);
        }


        //calling functions: write_to_excel
        //this funciton is to open the file that is at path x
        public void open_excel_file(string x)
        {
            if (xlApp == null)
            {
                Console.WriteLine(
                    "EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }

            xlApp.Visible = true;
            var openwb = xlApp.Workbooks;
            openwb.Open(x);
        }
    }

    //this contains the path, extension and name of the file.
    public class fileNaming
    {
        public fileNaming(string x)
        {
            name = x;
        }

        public fileNaming()
        {
        }

        public string name { get; set; }
        public string path { get; set; }
        public string type { get; set; }
    }
}
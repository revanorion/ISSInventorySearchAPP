using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data.OleDb;
using System.Data;
using System.Text.RegularExpressions;
using Microsoft.VisualBasic.FileIO;

namespace ISSISA_Library
{
    public class FileConnections
    {

        //This is used for exporting excel document
        public Excel.Application xlApp = new Excel.Application();

        //these lists contain the file properties that are being imported or exported
        public List<fileNaming> files = new List<fileNaming>();
        public List<fileNaming> finished_files = new List<fileNaming>();

        //these lists contain all the assets in the fiscal book, imported files, or assets that were found
        //as a result of comparting the previous two lists
        public List<asset> fb_assets = new List<asset>();
        public List<asset> imported_devices = new List<asset>();
        public List<asset> found_devices;

        //this private member holds the file properties for the fiscal book
        private fileNaming _fiscal_book_address = new fileNaming("No File Selected!");
        //this public member is used to make sure the fiscal book is in fact correctly chosen. 
        public string fiscal_book_address
        {
            get { return _fiscal_book_address.name; }
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

        }

        //fb example: FY 2016 20160114
        //Sheet exists that must be called ISS Assets Inventory + year
        //calling funcitons: import_data       
        public void open_fiscal_book()
        {
            string year = fiscal_book_address.Substring(3, 4);
            string sheetName = "ISS Assets Inventory " + year;
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fiscal_book_address.path + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"");
            con.Open();

            try
            {
                //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference
                DataSet myDataSet = new DataSet();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(" SELECT * from [" + sheetName + "$]", con);
                bool ignore = true;
                myCommand.Fill(myDataSet);
                DataTable d;
                // d = myDataSet.Tables[0];
                // d.WriteXml("c:\\myxml.xml");
                con.Close();

                //Travers through each row in the dataset
                foreach (DataRow myDataRow in myDataSet.Tables[0].Rows)
                {
                    //Stores info in Datarow into an array
                    Object[] cells = myDataRow.ItemArray;
                    //Traverse through each array and put into object cellContent as type Object
                    //Using Object as for some reason the Dataset reads some blank value which
                    //causes a hissy fit when trying to read. By using object I can convert to
                    //String at a later point.
                    var result = cells.Where(n => n.ToString() == "Asset #").ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    else if (ignore)
                        ignore = false;
                    else
                    {
                        asset b = new asset(cells[0],
                            cells[1],
                            cells[2],
                            cells[3],
                            cells[4],
                            cells[5],
                            cells[6],
                            cells[7],
                            cells[8],
                            cells[9],
                            cells[10],
                            cells[11],
                            cells[12],
                            cells[13],
                            cells[14]);
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


        //this function takes a string that is a path to a file to be imported. 
        //that string is broken up into its elements of full path, extension, and name
        //calling funcitons: add file button on form.
        public void add_file(string x)
        {
            if (IndexOf(x, files) == 1)
            {
                fileNaming myFile = new fileNaming();
                myFile.path = x;
                myFile.name = Path.GetFileNameWithoutExtension(x);
                myFile.type = Path.GetExtension(x);


                files.Add(myFile);

            }
        }

        //removes the file from the import file list
        public void remove_file(fileNaming x)
        {
            files.RemoveAt(files.IndexOf(x));
        }


        //this function handles the importing of data from supported type csv files.
        //calling funcitons: open_file
        public void open_csv_file(fileNaming x, string skipUntil = null, string breakAt = null)
        {
            bool hasSkip = false, hasBreak = false, ignore = true;
            string serial;

            if (skipUntil != null)
                hasSkip = true;
            if (breakAt != null)
                hasBreak = true;
            char csvType;
            if (x.name.Contains("Tropos Export Data"))
                csvType = 'T';
            else if (x.name.Contains("Wireless_Controllers"))
                csvType = 'W';
            else if (x.name.Contains("aps_wireless"))
                csvType = 'A';
            else if (x.name.Contains("device type - UPS"))
                csvType = 'S';
            else if (x.name.Contains("UPS"))
                csvType = 'U';
            else if (x.name.Contains("Brocade switch"))
                csvType = 'B';
            else if (x.name.Contains("Device List- total-assest"))
                csvType = 'D';
            else if (x.name.Contains("LMS switch and Router report"))
                csvType = 'L';
            else
                throw new NotSupportedException();

            using (var csvParser = new TextFieldParser(x.path))
            {
                csvParser.TextFieldType = FieldType.Delimited;
                csvParser.SetDelimiters(",");
                csvParser.TrimWhiteSpace = true;
                csvParser.HasFieldsEnclosedInQuotes = true;
                while (!csvParser.EndOfData)
                {
                    List<string> parts = csvParser.ReadFields().ToList();
                    if (hasSkip && ignore && parts.Where(n => n.ToString().Replace("\"", "") == skipUntil).ToList().Count() == 0)
                        continue;
                    else if (hasBreak && parts.Where(n => n.ToString() == breakAt).ToList().Count() > 0)
                        break;
                    else if (ignore)
                        ignore = false;
                    else
                    {
                        asset a = new asset();
                        switch (csvType)
                        {
                            case 'T':
                                a.serial_number = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.status = parts.ElementAt(2);
                                a.physical_location = parts.ElementAt(3);
                                break;
                            case 'W':
                                a.controller_name = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.physical_location = parts.ElementAt(2);
                                a.status = parts.ElementAt(3);
                                a.serial_number = parts.ElementAt(4);
                                a.model = parts.ElementAt(5);
                                break;
                            case 'A':
                                a.device_name = parts.ElementAt(0);
                                a.mac_address = parts.ElementAt(1);
                                a.ip_address = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(3);
                                a.model = parts.ElementAt(4);
                                a.physical_location = parts.ElementAt(5);
                                a.controller_name = parts.ElementAt(6);
                                break;
                            case 'S':
                                if (parts.ToArray().Length > 6)
                                    a.serial_number = parts.ElementAt(3).Replace("\"", "");
                                else
                                    continue;
                                if (a.serial_number == null || a.serial_number == "" || a.serial_number.Contains("Serial Number"))
                                    continue;
                                a.ip_address = parts.ElementAt(0).Replace("\"", "");
                                a.hostname = parts.ElementAt(1).Replace("\"", "");
                                a.model = parts.ElementAt(2).Replace("\"", "");
                                a.firmware = parts.ElementAt(4).Replace("\"", "");
                                a.physical_location = parts.ElementAt(5).Replace("\"", "") + " " + parts.ElementAt(6).Replace("\"", "");
                                break;
                            case 'B':

                                serial = parts.ElementAt(6).Replace("\"", "");
                                List<string> serialList = serial.Split(';').ToList();
                                if (serialList.Last() == "")
                                    serialList.RemoveAt(serialList.Count() - 1);
                                Regex r = new Regex(@"^Unit\s\d+\s-\s");
                                a.status = parts.ElementAt(0).Replace("\"", "");
                                a.device_name = parts.ElementAt(1).Replace("\"", "");
                                a.ip_address = parts.ElementAt(4).Replace("\"", "");
                                a.model = parts.ElementAt(8).Replace("\"", "");
                                a.firmware = parts.ElementAt(9).Replace("\"", "");
                                a.contact = parts.ElementAt(10).Replace("\"", "");
                                a.physical_location = parts.ElementAt(11).Replace("\"", "");
                                a.last_scanned = parts.ElementAt(12).Replace("\"", "");
                                a.source = x.name;
                                foreach (string s in serialList)
                                {
                                    asset b = new asset(a);
                                    serial = s;
                                    if (s.Contains("Unit"))
                                    {
                                        serial = r.Replace(s, "");
                                    }
                                    b.serial_number = serial;

                                    imported_devices.Add(b);
                                }
                                break;
                            case 'D':

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
                            case 'U':
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
                            case 'L':
                                if (parts.ElementAt(0) == "")
                                    continue;
                                serial = parts.ElementAt(5);

                                a = imported_devices.FirstOrDefault(p => p.serial_number == serial);
                                string childSN = parts.ElementAt(7);
                                if (a == null)
                                {
                                    a = new asset();
                                    a.device_name = parts.ElementAt(0);
                                    a.physical_location = parts.ElementAt(1);
                                    a.contact = parts.ElementAt(2);
                                    a.model = parts.ElementAt(3);
                                    a.serial_number = serial;
                                    if (childSN != null && childSN != "" && childSN != "N/A")
                                    {
                                        a.children.Add(childSN);
                                    }
                                    imported_devices.Add(a);
                                }
                                else if (childSN != null && childSN != "" && childSN != "N/A")
                                {
                                    a.children.Add(childSN);
                                }

                                break;
                        }
                        a.source = x.name;
                        //this is because i did a deep copy from asset a to b. L because special case that has children SN 
                        if (a.serial_number != "" && csvType != 'B' && csvType != 'L')
                        {
                            imported_devices.Add(a);
                        }
                        else if (csvType != 'B')
                        {
                            int i;
                        }
                    }
                }
            }
        }



        //public void test(fileNaming x)
        //{            
        //    using (var csvParser = new TextFieldParser(x.path))
        //    {
        //        csvParser.TextFieldType = FieldType.Delimited;
        //        csvParser.SetDelimiters(",");
        //        csvParser.TrimWhiteSpace = true;
        //        csvParser.HasFieldsEnclosedInQuotes = true;
        //        while (!csvParser.EndOfData)
        //        {
        //            List<string>fieldRow = csvParser.ReadFields().ToList();

        //            foreach (string fieldRowCell in fieldRow)
        //            {
        //                // todo
        //            }
        //        }
        //    }
        //}


        //public void open_csv_file(fileNaming x, string skipUntil = null, string breakAt = null)
        //{
        //    test(x);


        //    bool hasSkip = false, hasBreak = false, ignore = true;
        //    char[] delimiters = new char[] { ',' };

        //    if (skipUntil != null)
        //        hasSkip = true;
        //    if (breakAt != null)
        //        hasBreak = true;
        //    char csvType;
        //    if (x.name.Contains("Tropos Export Data"))
        //        csvType = 'T';
        //    else if (x.name.Contains("Wireless_Controllers"))
        //        csvType = 'W';
        //    else if (x.name.Contains("aps_wireless"))
        //        csvType = 'A';
        //    else if (x.name.Contains("UPS"))
        //        csvType = 'U';
        //    else if (x.name.Contains("Brocade switch"))
        //        csvType = 'B';
        //    else if (x.name.Contains("Device List- total-assest"))
        //        csvType = 'D';
        //    else
        //        throw new NotSupportedException();

        //    using (StreamReader reader = new StreamReader(x.path))
        //    {
        //        while (true)
        //        {
        //            string line = reader.ReadLine();
        //            if (line == null)
        //                break;

        //            List<string> parts = line.Split(delimiters, StringSplitOptions.None).ToList();
        //            if (hasSkip && ignore && parts.Where(n => n.ToString().Replace("\"", "") == skipUntil).ToList().Count() == 0)
        //                continue;
        //            else if (hasBreak && parts.Where(n => n.ToString() == breakAt).ToList().Count() > 0)
        //                break;
        //            else if (ignore)
        //                ignore = false;
        //            else
        //            {
        //                asset a = new asset();
        //                switch (csvType)
        //                {
        //                    case 'T':
        //                        a.serial_number = parts.ElementAt(0);
        //                        a.ip_address = parts.ElementAt(1);
        //                        a.status = parts.ElementAt(2);
        //                        a.physical_location = parts.ElementAt(3);
        //                        break;
        //                    case 'W':
        //                        a.controller_name = parts.ElementAt(0);
        //                        a.ip_address = parts.ElementAt(1);
        //                        a.physical_location = parts.ElementAt(2);
        //                        a.status = parts.ElementAt(3);
        //                        a.serial_number = parts.ElementAt(4);
        //                        a.model = parts.ElementAt(5);
        //                        break;
        //                    case 'A':
        //                        a.device_name = parts.ElementAt(0);
        //                        a.mac_address = parts.ElementAt(1);
        //                        a.ip_address = parts.ElementAt(2);
        //                        a.serial_number = parts.ElementAt(3);
        //                        a.model = parts.ElementAt(4);
        //                        a.physical_location = parts.ElementAt(5);
        //                        a.controller_name = parts.ElementAt(6);
        //                        break;
        //                    case 'U':
        //                        if (parts.ToArray().Length > 6)
        //                            a.serial_number = parts.ElementAt(3).Replace("\"", "");
        //                        else
        //                            continue;
        //                        if (a.serial_number == null || a.serial_number == "" || a.serial_number.Contains("Serial Number"))
        //                            continue;
        //                        a.ip_address = parts.ElementAt(0).Replace("\"", "");
        //                        a.hostname = parts.ElementAt(1).Replace("\"", "");
        //                        a.model = parts.ElementAt(2).Replace("\"", "");
        //                        a.firmware = parts.ElementAt(4).Replace("\"", "");
        //                        a.physical_location = parts.ElementAt(5).Replace("\"", "") + " " + parts.ElementAt(6).Replace("\"", "");
        //                        break;
        //                    case 'B':

        //                        string serial = parts.ElementAt(6).Replace("\"", "");
        //                        List<string> serialList = serial.Split(';').ToList();
        //                        if (serialList.Last() == "")
        //                            serialList.RemoveAt(serialList.Count() - 1);
        //                        Regex r = new Regex(@"^Unit\s\d+\s-\s");
        //                        a.status = parts.ElementAt(0).Replace("\"", "");
        //                        a.device_name = parts.ElementAt(1).Replace("\"", "");
        //                        a.ip_address = parts.ElementAt(4).Replace("\"", "");
        //                        a.model = parts.ElementAt(8).Replace("\"", "");
        //                        a.firmware = parts.ElementAt(9).Replace("\"", "");
        //                        a.contact = parts.ElementAt(10).Replace("\"", "");
        //                        a.physical_location = parts.ElementAt(11).Replace("\"", "");
        //                        a.last_scanned = parts.ElementAt(12).Replace("\"", "");
        //                        a.source = x.name;
        //                        foreach (string s in serialList)
        //                        {
        //                            asset b = new asset(a);
        //                            serial = s;
        //                            if (s.Contains("Unit"))
        //                            {
        //                                serial = r.Replace(s, "");
        //                            }
        //                            b.serial_number = serial;

        //                            imported_devices.Add(b);
        //                        }
        //                        break;
        //                    case 'D':
        //                        a.description = parts.ElementAt(0);
        //                        a.model = parts.ElementAt(1);
        //                        a.hostname = parts.ElementAt(2);
        //                        a.device_name = parts.ElementAt(3);
        //                        a.serial_number = parts.ElementAt(4);
        //                        a.ip_address = parts.ElementAt(5);
        //                        a.physical_location = parts.ElementAt(6).Replace("\"", "");
        //                        a.mac_address = parts.ElementAt(7);
        //                        a.contact = parts.ElementAt(8);
        //                        a.room_number = parts.ElementAt(9);
        //                        if (parts.ElementAt(10) != "")
        //                        {
        //                            a.location = Convert.ToInt16(parts.ElementAt(10));

        //                        }
        //                        break;
        //                }

        //                if (a.serial_number != "")
        //                {
        //                    a.source = x.name;
        //                    imported_devices.Add(a);
        //                }
        //            }
        //        }
        //        reader.Close();

        //    }
        //}
        //this function handles the type of file to be opened.
        //this is based both of file extension as well if the file contains the specified name
        //calling functions: import_data


        public void open_file(fileNaming x)
        {
            switch (x.type)
            {
                case ".xlsx":
                    break;
                case ".xls":
                    if (x.name.Contains("TMS-Inventory"))
                        open_xls_file(x);
                    break;
                case ".csv":
                    if (x.name.Contains("Tropos"))
                        open_csv_file(x, "InventorySerialNumber");
                    else if (x.name.Contains("aps_wireless"))
                        open_csv_file(x, "AP Name", "Disassociated AP(s)");
                    else if (x.name.Contains("Wireless_Controllers"))
                        open_csv_file(x, "Controller Name");
                    else if (x.name.Contains("Brocade switch"))
                        open_csv_file(x, "Product Status");
                    else if (x.name.Contains("Device List- total-assest"))
                        open_csv_file(x, "Type");
                    else if (x.name.Contains("UPS") && x.name.Contains("Device List"))
                        open_csv_file(x, "Model");
                    else if (x.name.Contains("LMS switch and Router report"))
                        open_csv_file(x, "Device Name");
                    else if (x.name.Contains("device type - UPS"))
                        open_csv_file(x);
                    else
                        open_csv_file(x);
                    break;
                case ".txt":
                    open_text_dump(x);
                    break;
                default:
                    throw new NotSupportedException();
            }
        }

        private void open_xls_file(fileNaming x, string skipUntil = null, string breakAt = null)
        {
            bool hasSkip = false, hasBreak = false, ignore = true;

            if (skipUntil == null && breakAt == null)
                ignore = false;
            if (skipUntil != null)
                hasSkip = true;
            if (breakAt != null)
                hasBreak = true;
            char xlsType;
            if (x.name.Contains("TMS-Inventory"))
                xlsType = 'T';
            else
                throw new NotSupportedException();
            string sheetName = "";
            switch (xlsType)
            {
                case 'T':
                    sheetName = "TMS Excel Export 1";
                    break;
                default:
                    sheetName = "Sheet1";
                    break;
            }
            OleDbConnection con = new OleDbConnection(@"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + x.path + @";Extended Properties=""Excel 8.0;HDR=YES;""");
            con.Open();

            try
            {
                //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference
                DataSet myDataSet = new DataSet();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(" SELECT * from [" + sheetName + "$]", con);
                myCommand.Fill(myDataSet);
                DataTable d;
                con.Close();

                //Travers through each row in the dataset
                foreach (DataRow myDataRow in myDataSet.Tables[0].Rows)
                {
                    //Stores info in Datarow into an array
                    Object[] cells = myDataRow.ItemArray;
                    //Traverse through each array and put into object cellContent as type Object
                    //Using Object as for some reason the Dataset reads some blank value which
                    //causes a hissy fit when trying to read. By using object I can convert to
                    //String at a later point.
                    var result = cells.Where(n => n.ToString() == skipUntil).ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    else if (ignore)
                        ignore = false;
                    else
                    {
                        asset a = new asset();
                        switch (xlsType)
                        {
                            case 'T':
                                a.device_name = Convert.ToString(cells[0]);
                                a.ip_address = Convert.ToString(cells[1]);
                                a.description = Convert.ToString(cells[2]) + " (" + Convert.ToString(cells[4]) + ")";
                                a.serial_number = Convert.ToString(cells[3]);
                                break;
                        }
                        a.source = x.name;
                        imported_devices.Add(a);
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

        //this function handles importing data from a text file
        //calling functions: open_file
        private void open_text_dump(fileNaming x)
        {
            string line = "";


            using (StreamReader sr = new StreamReader(x.path))
            {
                //this loops through each line in the file
                while ((line = sr.ReadLine()) != null)
                {

                    //this attempts to see if the string exists in the line if it does continue
                    int index = line.IndexOf("Device Name");

                    if (index != -1)
                    {

                        //grab the device name for all the corresponding serials
                        string deviceName = line.Substring(15);
                        string masterSN = "";
                        List<string> serials = new List<string>();

                        //this loops through each line after device has been found until it hits #.
                        while ((line = sr.ReadLine()) != null && (line.IndexOf("#") == -1 || line.IndexOf("Serial#") > 1))
                        {
                            //logical error checking for stupid data
                            if (line.IndexOf("SN: 0x") == -1 && line.IndexOf("SN: N/A") == -1 && line.IndexOf("NAME:") == -1)
                            {
                                //if a serial number exists grab the data
                                if ((index = line.IndexOf("SN:")) > 1)
                                {
                                    line = line.Substring(index + 4);
                                    //if serial is not null
                                    if (line != "")
                                    {
                                        //if it is the first serial number make it the master
                                        if (masterSN == "")
                                            masterSN = line;
                                        serials.Add(line);
                                    }
                                }
                            }
                        }
                        foreach (string sn in serials)
                        {
                            asset a = new asset();
                            a.serial_number = sn;
                            a.device_name = deviceName;
                            if (masterSN == sn)
                                a.description = "Master SN";
                            a.source = x.name;
                            imported_devices.Add(a);
                        }
                    }
                }
                sr.Close();
            }
        }



        public void write_missing_to_excel(string x)
        {
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = false;
            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
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
            foreach (asset a in imported_devices.Where(n => !n.found))
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
                foreach (string child in a.children)
                    ws.Cells[rows, columns] += child + ";";
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
                ws.Cells[rows, columns++] = a.last_scanned.ToString();
                ws.Cells[rows++, columns] = a.source;
                columns = 1;
            }

            //save the file using the param as the path and name
            ws.Columns.AutoFit();
            wb.SaveAs(x);
            //open the saved file

            open_excel_file(x);

        }



        //this function handles writing the compared data to excel file.
        //calling functions: save button on form.
        public void write_to_excel(string x)
        {
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = false;
            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
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
            foreach (asset a in found_devices.Where(n => n.found))
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
                foreach (string child in a.children)
                    ws.Cells[rows, columns] += child + ";";
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

            //save the file using the param as the path and name
            ws.Columns.AutoFit();
            wb.SaveAs(x);

            //open the saved file

            open_excel_file(x);
        }

        //this function handles the comparison between imported devices and assets from fiscal book
        //calling functions: import_data
        public void compare()
        {

            //every asset in fical book is placed in found devices list
            found_devices = fb_assets.ToList();

            //loops through each asset in list
            foreach (asset a in imported_devices)
            {
                //grab the asset from fiscal book that has matching serials from imported device data
                asset existingAsset = found_devices.FirstOrDefault(x => x.serial_number.Contains(a.serial_number));
                if (existingAsset != null)
                {
                    updateFoundAsset(existingAsset, a);

                    foreach (string child in a.children)
                    {
                        asset existingAsset2 = found_devices.FirstOrDefault(x => x.serial_number.Contains(child));
                        updateFoundAsset(existingAsset2, a, false);
                    }
                }
                else
                {
                    //doesnt exist
                }
            }
        }

        public void updateFoundAsset(asset a, asset b, bool parent = true)
        {
            //update some of the found device fields by combining data
            if (!String.IsNullOrEmpty(b.description))
                a.description = a.description + " (" + b.description + ")";
            if (!String.IsNullOrEmpty(b.model))
                a.model = a.model + " (" + b.model + ")";
            if (!String.IsNullOrEmpty(b.physical_location))
                a.physical_location = a.physical_location + " (" + b.physical_location + ")";
            if (!String.IsNullOrEmpty(b.location) && !a.location.Contains(b.location))
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
            if (parent)
                a.children = b.children;

        }


        //used for debugging only
        public void show_imported_devices()
        {

            using (StreamWriter sw = new StreamWriter("custs.txt"))
            {
                foreach (asset a in imported_devices)
                {
                    sw.WriteLine(a.output());
                }
            }
        }

        //this is simply used to say if the file to be imported is alreay in the list or not
        //if 1 then it is not in the list. if 0 then the file is already selected to be imported
        //calling functions: add_file
        public int IndexOf(string x, List<fileNaming> a)
        {
            x = Path.GetFileNameWithoutExtension(x);
            foreach (fileNaming y in a)
            {
                if (x.Equals(y.name))
                {
                    return 0;
                }
            }
            return 1;
        }

        //this function makes the necessay calls to import all data from selected files the compare the data
        //calling functions: run button on form.
        public void import_data()
        {
            //open fiscal book first
            open_fiscal_book();

            //open each file in the list
            foreach (fileNaming file in files)
            {
                open_file(file);
            }

            //then compare the data
            compare();
            //get the current date
            string date = DateTime.Now.ToShortDateString().Replace("/", "");
            //add report to finished files list
            finished_files.Add(new fileNaming(date + " Inventory Report "));
            finished_files.Add(new fileNaming(date + " Missing Inventory Report "));
        }

        //calling functions: write_to_excel
        //this funciton is to open the file that is at path x
        public void open_excel_file(string x)
        {
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = true;
            Excel.Workbooks openwb = xlApp.Workbooks;
            openwb.Open(x);
        }

    }

    //this contains the path, extension and name of the file.
    public class fileNaming
    {
        public fileNaming(string x) { name = x; }
        public fileNaming() { }
        public string name { get; set; }
        public string path { get; set; }
        public string type { get; set; }
    }
}

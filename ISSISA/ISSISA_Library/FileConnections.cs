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
using System.ComponentModel;
using System.Collections;

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
        public BindingList<asset> fb_assets = new BindingList<asset>();
        public BindingList<asset> imported_devices = new BindingList<asset>();
        public BindingList<asset> found_devices = new BindingList<asset>();
        public BindingList<asset> missing_devices = new BindingList<asset>();

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
        //csvType defines how data is going to be removed from the csv
        public void open_csv_file(fileNaming x, char csvType, string skipUntil = null, string breakAt = null)
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
                            //Tropos
                            case 'T':
                                a.serial_number = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.status = parts.ElementAt(2);
                                a.physical_location = parts.ElementAt(3);
                                break;
                            //Wireless_Controllers
                            case 'W':
                                a.controller_name = parts.ElementAt(0);
                                a.ip_address = parts.ElementAt(1);
                                a.physical_location = parts.ElementAt(2);
                                a.status = parts.ElementAt(3);
                                a.serial_number = parts.ElementAt(4);
                                a.model = parts.ElementAt(5);
                                break;
                            //aps_wireless
                            case 'A':
                                a.device_name = parts.ElementAt(0);
                                a.mac_address = parts.ElementAt(1);
                                a.ip_address = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(3);
                                a.model = parts.ElementAt(4);
                                a.physical_location = parts.ElementAt(5);
                                a.controller_name = parts.ElementAt(6);
                                break;
                            //device type - UPS
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
                            //Brocade switch
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
                            //Device List- total-assest
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
                            //UPS yyyy Device List
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
                            //LMS switch and Router report
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
                                        addChild(a, childSN, x);
                                    }
                                    imported_devices.Add(a);
                                }
                                else if (childSN != null && childSN != "" && childSN != "N/A")
                                {
                                    addChild(a, childSN, x);
                                }

                                break;
                            //Detailed_Router_Report_-_Yearly_Inventory                        
                            case 'R':
                                a.model = parts.ElementAt(0);
                                a.device_name = parts.ElementAt(1);
                                int index = parts.ElementAt(3).IndexOf("RELEASE SOFTWARE");
                                if (index != -1)
                                    a.description = parts.ElementAt(3).Substring(0, index - 2);
                                else
                                    a.description = parts.ElementAt(3);
                                a.description.Replace(System.Environment.NewLine, "");
                                a.physical_location = parts.ElementAt(4);
                                a.contact = parts.ElementAt(5);
                                a.serial_number = parts.ElementAt(6);
                                break;
                        }
                        //add source file value from what file it came from
                        a.source = x.name;
                        //this is because i did a deep copy from asset a to b. L because special case that has children SN 
                        if (a.serial_number != "" && csvType != 'B' && csvType != 'L')
                        {
                            imported_devices.Add(a);
                        }
                    }
                }
            }
        }



        //This is going to handle adding a child asset to the device list.
        public void addChild(asset a, string childSN, fileNaming x)
        {
            a.children.Add(childSN);
            asset b = new asset(a);
            b.serial_number = childSN;
            b.children.Clear();
            b.master = a.serial_number;
            b.source = x.name;
            imported_devices.Add(b);
        }


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
                        open_csv_file(x, 'T', "InventorySerialNumber");
                    else if (x.name.Contains("aps_wireless"))
                        open_csv_file(x, 'A', "AP Name", "Disassociated AP(s)");
                    else if (x.name.Contains("Wireless_Controllers"))
                        open_csv_file(x, 'W', "Controller Name");
                    else if (x.name.Contains("Brocade switch"))
                        open_csv_file(x, 'B', "Product Status");
                    else if (x.name.Contains("Device List- total-assest"))
                        open_csv_file(x, 'D', "Type");
                    else if (x.name.Contains("UPS") && x.name.Contains("Device List"))
                        open_csv_file(x, 'U', "Model");
                    else if (x.name.Contains("LMS switch and Router report"))
                        open_csv_file(x, 'L', "Device Name");
                    else if (x.name.Contains("device type - UPS"))
                        open_csv_file(x, 'S');
                    else if (x.name.Contains("Detailed_Router_Report_-_Yearly_Inventory") || x.name.Contains("Detailed_Switch_Report_-_Yearly_Inventory"))
                        open_csv_file(x, 'R', "Product Series");
                    else
                        throw new NotSupportedException();
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
                        //string masterSN = "";
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
                                        //if (masterSN == "")
                                        //    masterSN = line;
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
            foreach (asset a in exportList)
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
                ws.Cells[rows, columns++] = a.master;
                ws.Cells[rows, columns] = "";
                foreach (string child in a.children)
                {
                    var cellValue = (string)(ws.Cells[rows, columns] as Excel.Range).Value;
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

            //found_devices = fb_assets;

            //loops through each asset in list
            foreach (asset a in imported_devices)
            {
                //grab the asset from fiscal book that has matching serials from imported device data

                asset existingAsset = fb_assets.FirstOrDefault(x => x.serial_number.Contains(a.serial_number));
                if (existingAsset != null)
                {
                    asset found = new asset(existingAsset);
                    updateFoundAsset(found, a);
                    found_devices.Add(found);

                }
                else
                {
                    //doesnt exist
                    asset missing = new asset(a);
                    missing_devices.Add(missing);
                }
            }
        }

        public void updateFoundAsset(asset a, asset b)
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
            a.children = b.children;
            a.master = b.master;
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
            foreach (fileNaming file in files)
            {
                open_file(file);
            }

            //then compare the data
            compare();
            //get the current date
            string date = DateTime.Now.ToString("yyyyMMdd");

            //add report to finished files list
            finished_files.Add(new fileNaming(date + " Inventory Report "));
            finished_files.Add(new fileNaming(date + " Missing Inventory Report "));
        }


        //this function makes the necessay calls to import the data only from the inventory review file
        //calling functions: run button on form ->run_button_Click
        public void import_review_data()
        {
            open_review_book();



            //then validate the data
            validate();
            //get the current date
            string date = DateTime.Now.ToString("yyyyMMdd");

            //add report to finished files list
            finished_files.Add(new fileNaming(date + " Inventory Report "));
        }

        private void open_review_book()
        {
            string sheetName = "Advantage_FATS";
            OleDbConnection con = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + _fiscal_book_address.path + ";Extended Properties=\"Excel 12.0 Xml;HDR=NO;IMEX=1;\"");
            con.Open();

            try
            {
                //Create Dataset and fill with imformation from the Excel Spreadsheet for easier reference
                DataSet myDataSet = new DataSet();
                OleDbDataAdapter myCommand = new OleDbDataAdapter(" SELECT * from [" + sheetName + "$]", con);
                bool ignore = true;
                myCommand.Fill(myDataSet);
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
                    var result = cells.Where(n => n.ToString() == "ASSET_NUMBER").ToList();
                    if (result.Count == 0 && ignore)
                        continue;
                    else if (ignore)
                        ignore = false;
                    else
                    {
                        asset b = new asset(cells[0],
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

        private void validate()
        {
            locationValidate_devices = new BindingList<asset>(
                fb_assets.Where(a => !a.location.Equals(a.physical_location))
                .ToList());
            serialValidate_devices = new BindingList<asset>(
                fb_assets.Where(a => !(a.serial_number == "" && (a.fats_serial_number == "#N/A" || a.fats_serial_number == "")))
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

            locationValidate_devices = new BindingList<asset>(locationValidate_devices.Except(locationRoomValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());
            locationValidate_devices = new BindingList<asset>(locationValidate_devices.Except(locationSerialValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());

            serialValidate_devices = new BindingList<asset>(serialValidate_devices.Except(serialRoomValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());

            //ALL serialValidates are in locationSerial due to all records being included
            //serialValidate_devices = new BindingList<asset>(serialValidate_devices.Except(locationSerialValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());

            roomValidate_devices = new BindingList<asset>(roomValidate_devices.Except(serialRoomValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());
            //roomValidate_devices = new BindingList<asset>(roomValidate_devices.Except(locationRoomValidate_devices, new ISSIAS_Library.assetEqualityComparer()).ToList());

        }

        //this function handles writing the compared data to excel file.
        //calling functions: save button on form.
        public void write_validate_to_excel(string x)
        {
            ICollection<KeyValuePair<string, BindingList<asset>>> worksheets = new Dictionary<string, BindingList<asset>>();
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations", locationValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Serials", serialValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Rooms", roomValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations & Rooms", locationRoomValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations & Serials", locationSerialValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Serials & Rooms", serialRoomValidate_devices));
            worksheets.Add(new KeyValuePair<string, BindingList<asset>>("Locations & Rooms & Serials", locationRoomSerialValidate_devices));

            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = false;
            Excel.Workbook wb = xlApp.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);


            foreach (KeyValuePair<string, BindingList<asset>> worksheet in worksheets)
            {

                //Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[worksheet.Key];
                Excel.Worksheet ws;
                if (wb.Sheets.Count == 0)
                    ws = wb.Sheets.Add();
                else
                    ws = wb.Sheets.Add(After: wb.Sheets[wb.Sheets.Count]);
                ws.Name = worksheet.Key;
                if (ws == null)
                {
                    Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
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
                foreach (asset a in worksheet.Value)
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

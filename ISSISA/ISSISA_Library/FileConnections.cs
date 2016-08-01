using LinqToExcel;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using Office = Microsoft.Office.Core;


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
        //public void open_fiscal_book_with_linqtoexcel()
        {
            string year = fiscal_book_address.Substring(3, 4);
            string sheetName = "ISS Assets Inventory " + year;
            //This is used to open the file
            var excelFile = new ExcelQueryFactory(_fiscal_book_address.path);
            int assetLocation = 0;

            //selects all the rows in the asset sheet starting at the top
            var AssetSheetRow = from a in excelFile.Worksheet(sheetName) select a;

            //this loop is used to skip down to the starting point of our assets. 
            foreach (var a in AssetSheetRow)
            {
                string source = Convert.ToString(a["F1"].Value);
                if (source != null && source.Equals("Asset #"))
                {
                    assetLocation++;
                    break;
                }
                assetLocation++;
            }
            //selects all the rows in the asset sheet starting at the location of the asset row
            AssetSheetRow = (from a in excelFile.Worksheet(sheetName) select a).Skip(assetLocation);

            //this loop selects each row in the list and grabs data from it.
            foreach (var a in AssetSheetRow)
            {
                //var a is a list of cells that can be accessed by referenced by the column's header
                //columns that do not have a header begin with F
                asset b = new asset(
                    a["F1"].Value,
                    a["F2"].Value,
                    a["F3"].Value,
                    a["F4"].Value,
                    a["# Assets"].Value,
                    a["Cost"].Value,
                    a["F7"].Value,
                    a["F8"].Value,
                    a["F9"].Value,
                    a["F10"].Value,
                    a["F11"].Value,
                    a["F12"].Value,
                    a["F13"].Value,
                    a["F14"].Value,
                    a["F15"].Value
                    );
                fb_assets.Add(b);
            }
        }


        //stupid and bloated, uses too much cpu and time
        public void open_fiscal_book_without_linqtoexcel()
        //public void open_fiscal_book()
        {
            string year = fiscal_book_address.Substring(3, 4);
            string sheetName = "ISS Assets Inventory " + year;
            //This is used to open the file
            if (xlApp == null)
            {
                Console.WriteLine("EXCEL could not be started. Check that your office installation and project references are correct.");
                return;
            }
            xlApp.Visible = false;
            Excel.Workbook wb = xlApp.Workbooks.Open(_fiscal_book_address.path);
            Excel.Worksheet ws = (Excel.Worksheet)wb.Worksheets[1];

            if (ws == null)
            {
                Console.WriteLine("Worksheet could not be created. Check that your office installation and project references are correct.");
            }
            int rows = 1, columns = 1;
            bool skip = true;
            Excel.Range currentFind = null;

            //selects all the rows in the asset sheet starting at the top
            currentFind = ws.Cells.Find(What: "Asset #", LookIn: Excel.XlFindLookIn.xlValues, LookAt: Excel.XlLookAt.xlPart);
            currentFind = ws.Range[currentFind.Cells[1, 1], ws.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing)];
            foreach (Excel.Range a in currentFind.Rows)
            {
                if (skip) { skip = false; continue; }
                asset b = new asset(
                    a.Cells[1, 1].Text,
                    a.Cells[1, 2].Text,
                    a.Cells[1, 3].Text,
                    a.Cells[1, 4].Text,
                    a.Cells[1, 5].Text,
                    a.Cells[1, 6].Text,
                    a.Cells[1, 7].Text,
                    a.Cells[1, 8].Text,
                    a.Cells[1, 9].Text,
                    a.Cells[1, 10].Text,
                    a.Cells[1, 11].Text,
                    a.Cells[1, 12].Text,
                    a.Cells[1, 13].Text,
                    a.Cells[1, 14].Text,
                    a.Cells[1, 15].Text
                    );

                fb_assets.Add(b);
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

        //this function handles the importing of data from tropos type csv files.
        //calling funcitons: open_file
        public void open_tropos(fileNaming x)
        {
            var csvFile = new ExcelQueryFactory(x.path);
            var AssetSheetRow = from a in csvFile.Worksheet(x.name) select a;

            //Tropos Column Names: InventorySerialNumber, InventoryIp, InventoryStatus, InventoryLocation, InventoryReason, Id
            //this loop selects each row in the list and grabs data from it.
            foreach (var col in AssetSheetRow)
            {
                string serial = col["InventorySerialNumber"];
                string status = col["InventoryStatus"];
                string loc = col["InventoryLocation"];
                asset myAsset = new asset();
                myAsset.serial_number = serial;
                myAsset.status = status;
                myAsset.physical_location = loc;
                imported_devices.Add(myAsset);
            }
        }

        //this function handles the importing of data from aps_wireless type csv files.
        //calling funcitons: open_file
        public void open_aps_wireless(fileNaming x)
        {
            var csvFile = new ExcelQueryFactory(x.path);
            int start = 0;
            var AssetSheetRow = from a in csvFile.Worksheet(x.name) select a;

            //this loop is used to skip down to the starting point of our assets. 
            foreach (var a in AssetSheetRow)
            {
                string source = Convert.ToString(a["aps wireless with serial number"].Value);
                if (source != null && source.Equals("AP Name"))
                {
                    start++;
                    break;
                }
                start++;
            }

            //selects all the rows in the asset sheet starting at the location of the asset row
            var allRows = csvFile.WorksheetNoHeader().Skip(++start);
            //this loop selects each row in the list and grabs data from it.
            foreach (var row in allRows)
            {
                //this list holds the values of a single row
                List<string> values = new List<string>();
                //this loop selects each column in the row and adds it to a list.
                foreach (var col in row)
                {
                    string source = Convert.ToString(col.Value);
                    //if a value is equal to this string, all necessary data has already been captured and
                    //function can terminate
                    if (source != null && source.Equals("Disassociated AP(s)"))
                    {
                        //could possibly use return;
                        break;
                    }
                    values.Add(Convert.ToString(col.Value));
                }
                //if no data was selected terminate the loop and function ends.
                if (values.Count == 0)
                {
                    break;
                }
                string div_name = values[0];
                string mac = values[1];
                string serial = values[3];
                string model = values[4];
                string loc = values[5];
                string con_name = values[6];
                values.Clear();

                asset myAsset = new asset();
                myAsset.device_name = div_name;
                myAsset.mac_address = mac;
                myAsset.serial_number = serial;
                myAsset.model = model;
                myAsset.physical_location = loc;
                myAsset.controller_name = con_name;
                imported_devices.Add(myAsset);
            }
        }

        //this function handles the importing of data from wireless_controllers type csv files.
        //calling funcitons: open_file
        public void open_wireless_controllers(fileNaming x)
        {
            var csvFile = new ExcelQueryFactory(x.path);
            int serialLocation = 0;
            var AssetSheetRow = from a in csvFile.Worksheet(x.name) select a;

            //this loop is used to skip down to the starting point of our assets. 
            foreach (var a in AssetSheetRow)
            {
                string source = Convert.ToString(a["Wireless Controllers"].Value);
                if (source != null && source.Equals("Controller Name"))
                {
                    serialLocation++;
                    break;
                }
                serialLocation++;
            }

            //selects all the rows in the asset sheet starting at the location of the asset row
            AssetSheetRow = (from a in csvFile.Worksheet(x.name) select a).Skip(serialLocation);

            //Wireless names: Wireless Controllers, F2-6
            //this loop selects each row in the list and grabs data from it.
            foreach (var col in AssetSheetRow)
            {
                string con_name = col["Wireless Controllers"];
                string loc = col["F3"];
                string serial = col["F5"];
                string model = col["F6"];
                asset myAsset = new asset();
                myAsset.serial_number = serial;
                myAsset.model = model;
                myAsset.controller_name = con_name;
                myAsset.physical_location = loc;
                imported_devices.Add(myAsset);
            }
        }

        //this function handles the type of file to be opened.
        //this is based both of file extension as well if the file contains the specified name
        //calling functions: import_data
        public void open_file(fileNaming x)
        {
            switch (x.type)
            {
                case ".xlsx":
                    break;
                case ".csv":
                    if (x.name.Contains("Tropos"))
                    {
                        open_tropos(x);
                    }
                    else if (x.name.Contains("aps_wireless"))
                    {
                        open_aps_wireless(x);

                    }
                    else if (x.name.Contains("Wireless_Controllers"))
                    {

                        open_wireless_controllers(x);
                    }
                    break;
                case ".txt":
                    open_text_dump(x);
                    break;
                default:
                    throw new NotSupportedException();
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
                            imported_devices.Add(a);
                        }
                    }
                }
            }
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
            ws.Cells[rows, columns++] = "Cost";
            ws.Cells[rows, columns++] = "Last Inv";
            ws.Cells[rows, columns++] = "Serial # ";
            ws.Cells[rows, columns++] = "FATS Owner";
            ws.Cells[rows, columns++] = "Notes";
            ws.Cells[rows, columns++] = "Status";
            ws.Cells[rows, columns++] = "Device Name";
            ws.Cells[rows, columns++] = "Mac Address";
            ws.Cells[rows++, columns] = "Controller Name";

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
                ws.Cells[rows, columns++] = a.cost;
                ws.Cells[rows, columns++] = a.last_inv.ToString();
                ws.Cells[rows, columns++] = a.serial_number;
                ws.Cells[rows, columns++] = a.fats_owner;
                ws.Cells[rows, columns++] = a.notes;
                ws.Cells[rows, columns++] = a.status;
                ws.Cells[rows, columns++] = a.device_name;
                ws.Cells[rows, columns++] = a.mac_address;
                ws.Cells[rows++, columns] = a.controller_name;
                columns = 1;
            }

            //save the file using the param as the path and name
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

                    //update some of the found device fields by combining data
                    existingAsset.description = existingAsset.description + a.description;
                    if (!String.IsNullOrEmpty(a.model))
                        existingAsset.model = existingAsset.model + "(" + a.model + ")";
                    if (!String.IsNullOrEmpty(a.physical_location))
                        existingAsset.physical_location = existingAsset.location + "(" + a.physical_location + ")";
                    existingAsset.serial_number = a.serial_number;
                    existingAsset.status = a.status;
                    existingAsset.device_name = a.device_name;
                    existingAsset.mac_address = a.mac_address;
                    existingAsset.controller_name = a.controller_name;
                    //this field is used to show that a match has been found between data sets.
                    existingAsset.found = true;
                }
                else
                {
                    //doesnt exist
                }
            }
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
            finished_files.Add(new fileNaming("Inventory Report " + date));
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
            xlApp.Workbooks.Open(x);
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

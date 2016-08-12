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
            char[] delimiters = new char[] { ',' };

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
            else
                throw new NotSupportedException();

            using (StreamReader reader = new StreamReader(x.path))
            {
                while (true)
                {
                    string line = reader.ReadLine();
                    if (line == null)
                        break;
                    List<string> parts = line.Split(delimiters, StringSplitOptions.None).ToList();
                    if (hasSkip && ignore && parts.Where(n => n.ToString() == skipUntil).ToList().Count() == 0)
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
                                a.status = parts.ElementAt(2);
                                a.physical_location = parts.ElementAt(3);
                                break;
                            case 'W':
                                a.controller_name = parts.ElementAt(0);
                                a.physical_location = parts.ElementAt(2);
                                a.serial_number = parts.ElementAt(4);
                                a.model = parts.ElementAt(5);
                                break;
                            case 'A':
                                a.device_name = parts.ElementAt(0);
                                a.mac_address = parts.ElementAt(1);
                                a.serial_number = parts.ElementAt(3);
                                a.model = parts.ElementAt(4);
                                a.physical_location = parts.ElementAt(5);
                                a.controller_name = parts.ElementAt(6);
                                break;
                        }
                        imported_devices.Add(a);
                    }
                }
                reader.Close();

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
                        open_csv_file(x, "InventorySerialNumber");
                    else if (x.name.Contains("aps_wireless"))
                        open_csv_file(x, "AP Name", "Disassociated AP(s)");
                    else if (x.name.Contains("Wireless_Controllers"))
                        open_csv_file(x, "Controller Name");
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
                sr.Close();
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

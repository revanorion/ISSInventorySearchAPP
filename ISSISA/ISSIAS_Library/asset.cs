using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace ISSIAS_Library
{
    public class asset
    {
        //Properties found in Fiscal Book
        public string asset_number { get; set; }
        public DateTime missing { get; set; }
        public double cost { get; set; }
        public DateTime last_inv { get; set; }
        public Image AssetBarcode { get; set; }
        public string serial_number { get; set; }
        public string fats_serial_number { get; set; }
        public string description { get; set; }
        public string iss_division { get; set; }
        public string model { get; set; }
        public string asset_type { get; set; }
        public string location { get; set; }
        public string physical_location { get; set; }
        public string room_per_advantage { get; set; }
        public string room_per_fats { get; set; }
        public string fats_owner { get; set; }
        public string notes { get; set; }
        //Typical properties found in imported devices
        public string status { get; set; }
        public string device_name { get; set; }
        public string mac_address { get; set; }
        public string ip_address { get; set; }
        public string hostname { get; set; }
        public string firmware { get; set; }
        public string controller_name { get; set; }
        public string source { get; set; }
        public string contact { get; set; }
        public string last_scanned { get; set; }
        public string room_number { get; set; }
        public bool found { get; set; }
        public List<string> children { get; set; }
        public string master { get; set; }

        [DisplayName("Children")]
        public string childrenDisplay => children != null && children.Any() ? string.Join(";", children) : "";

        //Default constructor that sets up all the property fields.
        public asset()
        {
            asset_number = "";
            missing = DateTime.Now;
            cost = 0.0;
            last_inv = DateTime.Now;
            serial_number = "";
            fats_serial_number = "";
            firmware = "";
            hostname = "";
            description = "";
            iss_division = "";
            model = "";
            asset_type = "";
            location = "";
            physical_location = "";
            room_per_advantage = "";
            room_per_fats = "";
            fats_owner = "";
            notes = "";
            status = "";
            device_name = "";
            mac_address = "";
            ip_address = "";
            controller_name = "";
            source = "";
            contact = "";
            last_scanned = "";
            room_number = "";
            children = new List<string>();
            master = "";
            found = false;

        }

        //This constructor is used when getting data from an excel .xlsx fiscal book
        public asset(object asset_number, object missing, object iss_division, object description,
            object model, object asset_type, object location, object physical_location,
            object room_per_advantage, object room_per_fats, object cost, object last_inv,
            object serial_number, object serial_number_per_fats, object fats_owner, object notes)
        {
            //Make sure the object params are not null then convert them into specified data type
            DateTime outDate;
            if (asset_number != Convert.DBNull)
                this.asset_number = Convert.ToString(asset_number);
            if (missing != Convert.DBNull && DateTime.TryParse(missing.ToString(), out outDate))
                this.missing = Convert.ToDateTime(missing);
            if (iss_division != Convert.DBNull)
                this.iss_division = Convert.ToString(iss_division);
            if (description != Convert.DBNull)
                this.description = Convert.ToString(description);
            if (model != Convert.DBNull)
                this.model = Convert.ToString(model);
            if (asset_type != Convert.DBNull)
                this.asset_type = Convert.ToString(asset_type);
            if (location != Convert.DBNull)
                this.location = Convert.ToString(location);
            if (physical_location != Convert.DBNull)
                this.physical_location = Convert.ToString(physical_location);
            if (room_per_advantage != Convert.DBNull)
                this.room_per_advantage = Convert.ToString(room_per_advantage);
            if (room_per_fats != Convert.DBNull)
                this.room_per_fats = Convert.ToString(room_per_fats);
            //Remove all characters for cost that would stop the process to convert to number
            if (cost != Convert.DBNull)
            {
                var myCost = cost.ToString();
                myCost = Regex.Replace(myCost, @"[^\d+\.\d*]", "");
                if (myCost != "")
                    this.cost = Convert.ToDouble(myCost);
            }
            if (last_inv != Convert.DBNull && Convert.ToString(last_inv) != "#N/A")
                this.last_inv = Convert.ToDateTime(last_inv);
            if (serial_number != Convert.DBNull)
                this.serial_number = Convert.ToString(serial_number);
            fats_serial_number = serial_number_per_fats != Convert.DBNull ? Convert.ToString(serial_number_per_fats) : "";
            if (fats_owner != Convert.DBNull)
                this.fats_owner = Convert.ToString(fats_owner);
            if (notes != Convert.DBNull)
                this.notes = Convert.ToString(notes);
            found = false;
            children = new List<string>();

            if (!string.IsNullOrEmpty(this.asset_number) && this.asset_number.Length > 3)
            {
                AddBarcode(this);
            }
        }
        //This constructor is used when getting data from an excel .xlsx review book
        public asset(object asset_number, object description, object model, object location,
            object serial_number, object room_per_advantage, object asset_type, object fats_serial_number,
            object fats_owner, object physical_location, object room_per_fats, object last_inv,
            object hostname, object ip_address) : this()
        {

            //Make sure the object params are not null then convert them into specified data type
            if (asset_number != Convert.DBNull)
                this.asset_number = Convert.ToString(asset_number);
            if (description != Convert.DBNull)
                this.description = Convert.ToString(description);
            if (model != Convert.DBNull)
                this.model = Convert.ToString(model);
            if (asset_type != Convert.DBNull)
                this.asset_type = Convert.ToString(asset_type);
            if (location != Convert.DBNull)
                this.location = Convert.ToString(location).Trim(' ');
            if (physical_location != Convert.DBNull)
                this.physical_location = Convert.ToString(physical_location).Trim(' ');
            if (room_per_advantage != Convert.DBNull)
                this.room_per_advantage = Convert.ToString(room_per_advantage).Trim(' ');
            if (room_per_fats != Convert.DBNull)
                this.room_per_fats = Convert.ToString(room_per_fats).Trim(' ');
            //Remove all characters for cost that would stop the process to convert to number

            DateTime dateValue;
            DateTime.TryParse(Convert.ToString(last_inv), out dateValue);

            this.last_inv = dateValue;
            if (serial_number != Convert.DBNull)
                this.serial_number = Convert.ToString(serial_number).Trim(' ');
            if (fats_serial_number != Convert.DBNull)
                this.fats_serial_number = Convert.ToString(fats_serial_number).Trim(' ');
            if (fats_owner != Convert.DBNull)
                this.fats_owner = Convert.ToString(fats_owner);
            if (hostname != Convert.DBNull)
                this.hostname = Convert.ToString(hostname);
            if (ip_address != Convert.DBNull)
                this.ip_address = Convert.ToString(ip_address);
            found = false;
            children = new List<string>();

            if (!string.IsNullOrEmpty(this.asset_number) && this.asset_number.Length > 3)
            {
                AddBarcode(this);
            }
        }
        public asset(asset a)
        {
            asset_number = a.asset_number;
            fats_serial_number = a.fats_serial_number;
            missing = a.missing;
            cost = a.cost;
            last_inv = a.last_inv;
            serial_number = a.serial_number;
            description = a.description;
            iss_division = a.iss_division;
            model = a.model;
            asset_type = a.asset_type;
            location = a.location;
            physical_location = a.physical_location;
            room_per_advantage = a.room_per_advantage;
            room_per_fats = a.room_per_fats;
            fats_owner = a.fats_owner;
            notes = a.notes;
            status = a.status;
            device_name = a.device_name;
            mac_address = a.mac_address;
            ip_address = a.ip_address;
            controller_name = a.controller_name;
            source = a.source;
            contact = a.contact;
            last_scanned = a.last_scanned;
            room_number = a.room_number;
            found = a.found;
            children = new List<string>();
            foreach (var c in a.children)
                children.Add(c);
            master = a.master;

            if (!string.IsNullOrEmpty(asset_number) && asset_number.Length > 3)
            {
                AddBarcode(this);
            }
        }

        private static void AddBarcode(asset asset)
        {
            asset.AssetBarcode = GenCode128.Code128Rendering.MakeBarcodeImage(asset.asset_number, 1, true);
        }
    }
}

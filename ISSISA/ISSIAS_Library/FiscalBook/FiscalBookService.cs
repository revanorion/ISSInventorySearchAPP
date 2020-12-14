using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using ISSIAS_Library.Excel;
using OfficeOpenXml.Export.ToDataTable;

namespace ISSIAS_Library.FiscalBook
{

public class FiscalBookService
{

	public FiscalBookService()
	{

	}

	//fb example: FY 2016 20160114
	//Sheet exists that must be called ISS Assets Inventory + year
	public static IEnumerable<asset> OpenBook(string book)
	{
		const string startLiteral = "Asset #";
		const string endLiteral = "FATS Owner";

		var name = Path.GetFileName(book);
		var year = name.Substring(name.IndexOf("FY ", StringComparison.Ordinal) + 3, 4);
		var sheetName = "Inventory " + year;


		var options = ToDataTableOptions.Create(tableOptions =>
		{
			tableOptions.Mappings.Add(0, "Asset");
			tableOptions.Mappings.Add(1, "MissingDate", typeof(DateTime));
			tableOptions.Mappings.Add(2, "Division");
			tableOptions.Mappings.Add(3, "Description");
			tableOptions.Mappings.Add(4, "Model");
			tableOptions.Mappings.Add(5, "AssetType");
			tableOptions.Mappings.Add(6, "Location");
			tableOptions.Mappings.Add(7, "PhysicalLocation");
			tableOptions.Mappings.Add(8, "AdvantageRoom");
			tableOptions.Mappings.Add(9, "FatsRoom");
			tableOptions.Mappings.Add(10, "Cost");
			tableOptions.Mappings.Add(11, "LastInv");
			tableOptions.Mappings.Add(12, "Serial");
			tableOptions.Mappings.Add(13, "SerialFats");
			tableOptions.Mappings.Add(14, "FatsOwner");
		});

		var dataTable = ExcelService.OpenBook(book, sheetName, startLiteral, endLiteral, options);

		var assets = dataTable
			.AsEnumerable()
			.Select(ConvertTo)
			.ToList();


		assets.Where(t => t.asset_number.Length > 3).AsParallel().ForAll(AddBarcode);
		assets.AsParallel().ForAll(t => t.source = Path.GetFileName(book));

		return assets;
	}

	private static asset ConvertTo(DataRow row)
	{


			if (!DateTime.TryParse(row["LastInv"].ToString(), out var date))
				date = DateTime.Now;


			return new asset
			{
				asset_number = row["Asset"].ToString(),
				missing = Convert.ToDateTime(row["MissingDate"]),
				iss_division = row["Division"].ToString(),
				description = row["Description"].ToString(),
				model = row["Model"].ToString(),
				asset_type = row["AssetType"].ToString(),
				location = row["Location"].ToString(),
				physical_location = row["PhysicalLocation"].ToString(),
				room_per_advantage = row["AdvantageRoom"].ToString(),
				room_per_fats = row["FatsRoom"].ToString(),
				cost = Convert.ToDouble(row["Cost"]),
				last_inv =date,
				serial_number = row["Serial"].ToString(),
				fats_serial_number = row["SerialFats"].ToString(),
				fats_owner = row["FatsOwner"].ToString()
			};
	}


	private static void AddBarcode(asset asset)
	{
		asset.AssetBarcode = GenCode128.Code128Rendering.MakeBarcodeImage(asset.asset_number, 1, true);
	}


	public void SaveBook(string path, IEnumerable<asset> assets)
	{
		if (File.Exists(path)) File.Delete(path);


		var name = Path.GetFileName(path);
		var year = name.Substring(name.IndexOf("FY ", StringComparison.Ordinal) + 3, 4);
		var sheetName = "ISS Assets Inventory " + year;


		var firstAsset = assets.First();
		var type = typeof(asset);
		var memberInfos = new MemberInfo[]
		{
				type.GetProperty(nameof(firstAsset.asset_number)),
				type.GetProperty(nameof(firstAsset.AssetBarcode)),
				type.GetProperty(nameof(firstAsset.missing)),
				type.GetProperty(nameof(firstAsset.iss_division)),
				type.GetProperty(nameof(firstAsset.description)),
				type.GetProperty(nameof(firstAsset.model)),
				type.GetProperty(nameof(firstAsset.asset_type)),
				type.GetProperty(nameof(firstAsset.location)),
				type.GetProperty(nameof(firstAsset.physical_location)),
				type.GetProperty(nameof(firstAsset.room_per_advantage)),
				type.GetProperty(nameof(firstAsset.room_per_fats)),
				type.GetProperty(nameof(firstAsset.room_number)),
				type.GetProperty(nameof(firstAsset.cost)),
				type.GetProperty(nameof(firstAsset.last_inv)),
				type.GetProperty(nameof(firstAsset.serial_number)),
				type.GetProperty(nameof(firstAsset.fats_serial_number)),
				type.GetProperty(nameof(firstAsset.master)),
				type.GetProperty(nameof(firstAsset.childrenDisplay)),
				type.GetProperty(nameof(firstAsset.fats_owner)),
				type.GetProperty(nameof(firstAsset.notes)),
				type.GetProperty(nameof(firstAsset.status)),
				type.GetProperty(nameof(firstAsset.device_name)),
				type.GetProperty(nameof(firstAsset.mac_address)),
				type.GetProperty(nameof(firstAsset.ip_address)),
				type.GetProperty(nameof(firstAsset.hostname)),
				type.GetProperty(nameof(firstAsset.controller_name)),
				type.GetProperty(nameof(firstAsset.firmware)),
				type.GetProperty(nameof(firstAsset.contact)),
				type.GetProperty(nameof(firstAsset.last_scanned)),
				type.GetProperty(nameof(firstAsset.source)),
		};


		ExcelService.SaveBook(path, sheetName, assets, memberInfos);
	}
}
}
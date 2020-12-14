using System.Collections.Generic;
using System.IO;
using System.Reflection;
using ISSIAS_Library.Excel;

namespace ISSIAS_Library.InventoryReview
{
    public class InventoryReviewService
    {
        public void SaveBook(string path,
            IEnumerable<asset> locationValidate_devices,
            IEnumerable<asset> serialValidate_devices,
            IEnumerable<asset> roomValidate_devices,
            IEnumerable<asset> locationRoomValidate_devices,
            IEnumerable<asset> locationSerialValidate_devices,
            IEnumerable<asset> serialRoomValidate_devices,
            IEnumerable<asset> locationRoomSerialValidate_devices
        )
        {
            if (File.Exists(path)) File.Delete(path);

            var firstAsset = new asset();
            var type = typeof(asset);
            var memberInfos = new MemberInfo[]
            {
                type.GetProperty(nameof(firstAsset.asset_number)),
                type.GetProperty(nameof(firstAsset.AssetBarcode)),
                type.GetProperty(nameof(firstAsset.description)),
                type.GetProperty(nameof(firstAsset.model)),
                type.GetProperty(nameof(firstAsset.asset_type)),
                type.GetProperty(nameof(firstAsset.location)),
                type.GetProperty(nameof(firstAsset.physical_location)),
                type.GetProperty(nameof(firstAsset.room_per_advantage)),
                type.GetProperty(nameof(firstAsset.room_per_fats)),
                type.GetProperty(nameof(firstAsset.serial_number)),
                type.GetProperty(nameof(firstAsset.fats_serial_number)),
                type.GetProperty(nameof(firstAsset.fats_owner)),
                type.GetProperty(nameof(firstAsset.ip_address)),
                type.GetProperty(nameof(firstAsset.hostname)),
            };


            ExcelService.SaveBook(path, "Locations", locationValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Serials", serialValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Rooms", roomValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Locations & Rooms", locationRoomValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Locations & Serials", locationSerialValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Serials & Rooms", serialRoomValidate_devices, memberInfos);
            ExcelService.SaveBook(path, "Locations & Rooms & Serials", locationRoomSerialValidate_devices, memberInfos);
        }
    }
}
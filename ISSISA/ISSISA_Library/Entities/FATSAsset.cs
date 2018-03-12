using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace ISSIAS_Library.Entities
{
    [Table("ACTIVE_ISS_ASSETS")]
   public class FatsAsset
    {
        [Column("CATEGORY")]
        public string Category { get; set; }
        [Column("ELEMENT")]
        public string Element { get; set; }
        [Key]
        [Column("ASSET_NUMBER")]
        public string AssetNumber { get; set; }
        [Column("DESCRIPTION")] 
        public string Description { get; set; }
        [Column("MANUFACTURER")]
        public string Manufacturer { get; set; }
        [Column("MODEL")]
        public string Model { get; set; }
        [Column("SERIAL_NUMBER")]
        public string SerialNumber { get; set; }
        [Column("PURCHASE_DOC_NO")]
        public string PurchaseDocNo { get; set; }
        [Column("COST")]
        public string Cost { get; set; }
        [Column("TYPE")]
        public string Type { get; set; }
        [Column("RECEIVED_DATE")]
        public string ReceivedDate { get; set; }
        [Column("LAST_INVENTORY_DATE")]
        public string LastInventoryDate { get; set; }
        [Column("LOCATION_DESC")]
        public string LocationDesc { get; set; }
        [Column("LOCATION_CODE")]
        public string LocationCode { get; set; }
        [Column("ROOM")]
        public string Room { get; set; }
        [Column("OWNER")]
        public string Owner { get; set; }
        [Column("CONDITION")]
        public string Condition { get; set; }
        [Column("WARRANTY_EXPIRE_DATE")]
        public string WarrantyExpireDate { get; set; }
    }
}

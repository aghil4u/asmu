using System;
using System.Linq;

namespace asmu
{
    class ServerEquipment
    {
        public int id { get; set; }
        public string EquipmentNumber { get; set; }
        public string AssetNumber { get; set; }
        public string AcquisitionDate { get; set; }
        public bool PendingUpdate { get; set; }
        public double AcquisitionValue { get; set; }
        public string BookValue { get; set; }
        public string AssetDescription { get; set; }
        public string EquipmentDescription { get; set; }
        public string OperationId { get; set; }
        public string SubType { get; set; }
        public string SubTypeDescription { get; set; }
        public string Weight { get; set; }
        public string WeightUnit { get; set; }
        public string Dimensions { get; set; }
        public string Tag { get; set; }
        public string Type { get; set; }
        public string Connection { get; set; }
        public string Length { get; set; }
        public string ModelNumber { get; set; }
        public string SerialNumber { get; set; }
        public string AssetLocation { get; set; }
        public string AssetLocationText { get; set; }
        public string EquipmentLocation { get; set; }

    }
}

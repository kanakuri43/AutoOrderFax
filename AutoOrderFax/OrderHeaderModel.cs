using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace AutoOrderFax
{
    public class OrderHeaderModel
    {

        public string OrderNo { get; set; }
        public string OrderDate { get; set; }
        public string SupplierName { get; set; } 
        public string DeliveryTypeName { get; set; }
        public string CustomerAddress { get; set; } 
        public string CustomerName { get; set; } 
        public string CustomerTel { get; set; }
        public string OperatorName { get; set; }
        public string PrivateNotes { get; set; }
        public string PublicNotes { get; set; }
        public string SelfZipCode { get; set; }
        public string SelfAddress { get; set; }
        public string SelfCompanyName { get; set; }
        public string SelfDepartmentName { get; set; }
        public string SelfTel { get; set; }
        public string SelfFax{ get; set; }
        public string ShippingDate { get; set; }
        public string OrderNoTimeStamp { get; set; }
        public string FixedNotes { get; set; }

        public List<OrderDetailModel> OrderDetails { get; set; } = new List<OrderDetailModel>();
    }
}

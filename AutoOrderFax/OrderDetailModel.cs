using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutoOrderFax
{
    public class OrderDetailModel
    {

        public int LineNo { get; set; }
        public string TeacherName { get; set; }
        public string IndividualName { get; set; }
        public string SchoolYear { get; set; }
        public string SchoolClass { get; set; }
        public string ItemName { get; set; }
        public string ItemCode { get; set; }
        public string SpecName { get; set; }
        public float Qty { get; set; }
        public float ReserveQty { get; set; }
        public float TeacherQty { get; set; }
        public string OperatorName { get; set; }
        public string SupplierFax { get; set; }
        public float UnitPrice { get; set; }
        public float Price { get; set; }
        public int[] ClassDivide { get; set; }
        public string LinePrivateNotes { get; set; }
        public string LinePublicNotes { get; set; }

        public OrderDetailModel()
        {
            ClassDivide = new int[8];
        }
    }

}

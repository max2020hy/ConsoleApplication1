using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ConsoleApplication1
{
  public  class Model
    {

        //序号	项目名称	分区		楼栋名称			单元	房号	客户名称
        public string 序号 { get; set; }
        public string 项目名称 { get; set; }
        public string 楼栋名称 { get; set; }
        public string 房号 { get; set; }
        public string 客户名称 { get; set; }
        public  DateTime? 收款日期 { get; set; }
        //票据类型	票据编号	款项类型	款项名称

        public string 票据类型 { get; set; }
        public string 票据编号 { get; set; }
        public string 款项类型 { get; set; }
        public string 款项名称 { get; set; }
        public double 金额 { get; set; }
        public string 支付方式 { get; set; }
        public string 银付方式 { get; set; }
        public string 摘要 { get; set; }
    }
}

using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;

namespace Model
{
    public class 宗地属性
    {
        public String 宗地代码 { get; set; }
        public String 土地权利人 { get; set; }
        public String 证件类型 { get; set; }
        public String 证件编号 { get; set; }
        public String 通讯地址 { get; set; }
        public String 土地权属性质 { get; set; }
        public String 使用权类型 { get; set; }
        public String 土地坐落 { get; set; }
        public String 预编宗地代码 { get; set; }
        public String 图幅号 { get; set; }
        public String 东至 { get; set; }
        public String 南至 { get; set; }
        public String 西至 { get; set; }
        public String 北至 { get; set; }
        public String 批准用途 { get; set; }
        public String 实际用途 { get; set; }
        public String 新证书号 { get; set; }

        private String yzmj = string.Empty;
        public String 原证面积
        {
            get
            {
                if (yzmj == "0") return string.Empty;

                return yzmj;
            }
            set
            {
                yzmj = value;
            }
        }
        public String 其他面积 { get; set; }
        public String 晒坝面积 { get; set; }
        public String 批准面积 { get; set; }
        public String 宗地面积 { get; set; }
        public String 原证书号 { get; set; }

        public String 核实面积文本
        {
            get
            {
                return String.Format(@"经审查，该宗地申请登记的权属资料齐全且合法有效；地籍调查结果正确，四至清楚无争议，面积准确；权属性质为宅基地使用权，建议对该权利人，进行确权登记，登记面积为  {0}  平方米，其余   {1}   平方米为本集体经济组织其他建设用地。", this.批准面积, this.总面积 - ToInt32(this.批准面积));
            }
        }

        public Decimal 总面积
        {
            get
            {
                return ToInt32(this.宗地面积) + ToInt32(this.其他面积) + ToInt32(this.晒坝面积);
            }
        }

        public Decimal ToInt32(String strNumber)
        {
            Decimal result = 0;

            Decimal.TryParse(strNumber, out result);

            return result;
        }
    }
}

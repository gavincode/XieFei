using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Model
{
    public class 承包方调查表
    {
        public String 承包方代表编码 { get; set; }
        public String 承包方名称 { get; set; }
        public String 承包方类型 { get; set; }
        public String 成员个数 { get; set; }
        public String 成员姓名 { get; set; }
        public String 成员性别 { get; set; }
        public String 成员证件类型 { get; set; }
        public String 证件号码 { get; set; }
        public String 家庭关系 { get; set; }
        public String 是否共有人 { get; set; }
        public String 有无自由地 { get; set; }
        public String 备注 { get; set; }
        public String 承包方地址 { get; set; }
        public String 邮政编码 { get; set; }
        public String 联系电话 { get; set; }
        public String 调查员 { get; set; }
        public String 调查日期 { get; set; }
        public String 调查记事 { get; set; }
        public String 审核人 { get; set; }
        public String 审核日期 { get; set; }
        public String 审核意见 { get; set; }
        public Object 有承包地 { get { return ToNumber(是否共有人 == "1", 1); } }
        public Object 无承包地 { get { return ToNumber(是否共有人 != "1", 1); } }

        public List<承包方调查表> 所属家属列表 = new List<承包方调查表>();

        public String 非在册户口
        {
            get
            {
                return 所属家属列表.All(p => p.备注 != "在册") ? "是" : "否";
            }
        }
        public static Object ToNumber(Boolean boolValue, Object dataValue)
        {
            if (boolValue)
                return dataValue;
            else
                return null;
        }
    }
}

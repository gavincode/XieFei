using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Model
{
    public class 地块属性表
    {
        public String 地块预编码 { get; set; }
        public String 发包方编码 { get; set; }
        public String 行政区名称 { get; set; }
        public String 承包方代表编码 { get; set; }
        public String 承包方代表姓名 { get; set; }
        public String 地块名称 { get; set; }
        public String 地块编码 { get; set; }
        public String 北至 { get; set; }
        public String 东至 { get; set; }
        public String 南至 { get; set; }
        public String 西至 { get; set; }
        public String 土地利用类型 { get; set; }
        public String 地块类别 { get; set; }
        public String 主要树种 { get; set; }
        public String 经营类型 { get; set; }
        public String 森林类别 { get; set; }
        public String 林种 { get; set; }
        public String 造林年度 { get; set; }
        public String 林地使用期 { get; set; }
        public String 合同终止日期 { get; set; }
        //public String 备注
        //{
        //    get
        //    {
        //        if (地块类别 == null) return String.Empty;

        //        if (地块类别.ToLower() == "c")
        //            return "J";
        //        else if (地块类别.ToLower() == "z")
        //            return "Z";
        //        else if (地块类别.ToLower() == "q")
        //            return "Q";

        //        return String.Empty;
        //    }
        //}
        public Decimal 面积 { get; set; }

        public Object 水田
        {
            get
            {
                if (土地利用类型 == "1")
                    return 面积;
                else return String.Empty;
            }
        }
        public Object 旱地
        {
            get
            {
                if (土地利用类型 == "2")
                    return 面积;
                else return String.Empty;
            }
        }
        public String 个体编码
        {
            get
            {
                return this.发包方编码 + this.承包方代表编码 + "J";
            }
        }
    }
}

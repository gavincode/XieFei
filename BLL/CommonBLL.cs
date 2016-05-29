using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using DAL;

namespace BLL
{
    public class CommonBLL
    {
        #region First

        /// <summary>
        /// 构造承包方户主数据
        /// </summary>
        /// <returns></returns>
        public static Dictionary<String, 承包方调查表> 承包方调查表()
        {
            var dataList = CommonDAL.Read承包方调查表();

            List<承包方调查表> list = new List<承包方调查表>();
            foreach (var group in dataList.GroupBy(p => p.承包方代表编码))
            {
                var 户主 = group.FirstOrDefault(p => (p.家庭关系 ?? "").Trim() == "户主");

                if (group.Key == null || 户主 == null) continue;

                foreach (var item in group)
                {
                    //if (item.家庭关系 == "户主") continue;

                    户主.所属家属列表.Add(item);
                }

                list.Add(户主);
            }


            return list.ToDictionary(p => p.承包方代表编码);
        }

        /// <summary>
        /// 组织二轮承包信息汇总表
        /// </summary>
        /// <returns></returns>
        public static Dictionary<String, List<二轮承包信息汇总表>> 二轮承包信息汇总表()
        {
            Dictionary<String, List<二轮承包信息汇总表>> dict = new Dictionary<String, List<二轮承包信息汇总表>>();

            var dataList = CommonDAL.Read二轮承包信息汇总表();

            foreach (var item in dataList.GroupBy(p => p.承包方代表编码))
            {
                //排除空数据
                if (String.IsNullOrEmpty(item.Key)) continue;

                if (!dict.ContainsKey(item.Key))
                    dict[item.Key] = new List<二轮承包信息汇总表>();

                dict[item.Key].AddRange(item);
            }

            return dict;
        }

        /// <summary>
        /// 组织地块属性表
        /// </summary>
        /// <returns></returns>
        public static Dictionary<String, List<地块属性表>> 地块属性表()
        {
            Dictionary<String, List<地块属性表>> dict = new Dictionary<String, List<地块属性表>>();

            var dataList = CommonDAL.Read地块属性表();

            foreach (var item in dataList.GroupBy(p => p.承包方代表编码))
            {
                //排除空数据
                if (String.IsNullOrEmpty(item.Key)) continue;

                if (!dict.ContainsKey(item.Key))
                    dict[item.Key] = new List<地块属性表>();

                dict[item.Key].AddRange(item);
            }

            return dict;
        }

        /// <summary>
        /// 组织二轮承包表
        /// </summary>
        /// <returns></returns>
        public static Dictionary<String, List<二轮承包表>> 二轮承包表()
        {
            Dictionary<String, List<二轮承包表>> dict = new Dictionary<String, List<二轮承包表>>();

            var dataList = CommonDAL.Read二轮承包表();

            foreach (var item in dataList.GroupBy(p => p.承包方代表编码))
            {
                //排除空数据
                if (String.IsNullOrEmpty(item.Key)) continue;

                if (!dict.ContainsKey(item.Key))
                    dict[item.Key] = new List<二轮承包表>();

                dict[item.Key].AddRange(item);
            }

            return dict;
        }

        #endregion

        #region Second

        /// <summary>
        /// 宗地属性
        /// </summary>
        /// <returns></returns>
        public static List<宗地属性> 宗地属性()
        {
            var dataList = CommonDAL.Read宗地属性();

            return dataList;
        }

        #endregion

        /// <summary>
        /// 设置db文件 .mdb
        /// </summary>
        /// <param name="dbFile">数据库文件</param>
        public static void SetDbFile(String dbFile)
        {
            BaseDAL.DbFile = dbFile;
        }
    }
}

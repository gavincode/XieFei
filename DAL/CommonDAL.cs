using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Model;
using Dapper;

namespace DAL
{
    public class CommonDAL : BaseDAL
    {
        /// <summary>
        /// Read承包方调查表
        /// </summary>
        /// <returns></returns>
        public static List<承包方调查表> Read承包方调查表()
        {
            return getConn().Query<承包方调查表>("select * from 承包方调查表;", null).ToList();
        }
        /// <summary>
        /// Read承包方调查表
        /// </summary>
        /// <returns></returns>
        public static List<地块属性表> Read地块属性表()
        {
            return getConn().Query<地块属性表>("select * from 地块属性表;", null).ToList();
        }
        /// <summary>
        /// Read承包方调查表
        /// </summary>
        /// <returns></returns>
        public static List<二轮承包信息汇总表> Read二轮承包信息汇总表()
        {
            return getConn().Query<二轮承包信息汇总表>("select * from 二轮承包信息汇总表;", null).ToList();
        }
        /// <summary>
        /// Read二轮承包表
        /// </summary>
        /// <returns></returns>
        public static List<二轮承包表> Read二轮承包表()
        {
            return getConn().Query<二轮承包表>("select * from 二轮承包信息;", null).ToList();
        }

        #region Second

        /// <summary>
        /// Read宗地属性
        /// </summary>
        /// <returns></returns>
        public static List<宗地属性> Read宗地属性()
        {
            return getConn().Query<宗地属性>("select * from 宗地属性;", null).ToList();
        }

        #endregion
    }
}

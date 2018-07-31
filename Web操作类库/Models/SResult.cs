using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models
{
    public class SResult
    {
        public bool IsSuccessful { get; set; }
        /// <summary>错误信息</summary>
        public string ErrMsg { get; set; }

        public object Data { get; set; }


        /// <summary>
        /// 返回服务知行成功的一个实例
        /// </summary>
        public static SResult Success
        {
            get
            {
                return new SResult
                {
                    IsSuccessful = true
                };
            }
        }

        public static SResult GetSuccess<T>(T obj) where T : new()
        {
            SResult x = new SResult();
            x.IsSuccessful = true;
            x.Data = obj;
            return x;
        }
        /// <summary>
        /// 返回服务执行失败的一个实例
        /// </summary>
        /// <param name="errorMsg">错误描述</param>
        public static SResult Error(string errorMsg)
        {
            return new SResult
            {
                IsSuccessful = false,
                ErrMsg = errorMsg
            };
        }
    }
}

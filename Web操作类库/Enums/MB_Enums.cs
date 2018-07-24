using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Models.Enums
{
    /// <summary>
    /// 模板类型
    /// </summary>
    public enum MB_Enums
    {
        周报 =1,
        月报 =2
    }

    public enum RW_ZT
    {
        起草阶段=0,
        数据筹备阶段=1,
        参数填写阶段=2,
        文档生成中=3,
        完成可下载=4
    }
    public enum DATA_ZT
    {
        未上传=0,
        已上传=1,
        确认忽略=2
    }
    public enum DATA_LX
    {
        成交记录=1,
        新增预售=2,
        土地成交=3,
        认购数据=4
    }

    public enum CS_LX
    {
        文字型=1,
        筛选型=2,
        文件型=3
    }

    public enum YH_LX
    {
        默认账号=1,
        普通账号=2
    }
    /// <summary>
    /// 模板细分类型
    /// </summary>
    public enum MB_XFLX
    {
        主模板=1,
        竞品模板=2
    }
}

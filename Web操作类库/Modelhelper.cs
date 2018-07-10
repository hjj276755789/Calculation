using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;

namespace Calculation.Models
{
    public class Modelhelper
    {
        public static List<T> 类列表赋值<T>(T model, DataTable dt)
        {
            List<T> list = new List<T>();
            Type type = model.GetType();
            for (int i = 0; i < dt.Rows.Count; i++)
            {
                T temp = (T)type.Assembly.CreateInstance(type.FullName);
                foreach (var item in type.GetProperties())
                {
                    try
                    {
                        object sss = dt.Rows[i][item.Name];
                        if (sss != System.DBNull.Value)
                            item.SetValue(temp,sss, null);
                        
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e.Message);
                        continue;
                    }

                }
                list.Add(temp);
            }
            return list;
        }
        public static T 类对象赋值<T>(T model, DataTable dt)
        {
            Type type = model.GetType();
            T temp = (T)type.Assembly.CreateInstance(type.FullName);
            foreach (var item in type.GetProperties())
            {
                try
                {
                    object sss = dt.Rows[0][item.Name];
                    if (sss != System.DBNull.Value)
                        item.SetValue(temp, sss, null);
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.Message);
                    continue;
                }

            }
            return temp;
        }
        public static DataTable GetTableSchema<T>(List<T> modelList)
        {
            if (modelList == null || modelList.Count == 0)
            {
                return null;
            }
            DataTable dt = CreateData(modelList[0]);

            foreach (T model in modelList)
            {
                DataRow dataRow = dt.NewRow();
                foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
                {
                    dataRow[propertyInfo.Name] = propertyInfo.GetValue(model, null);
                }
                dt.Rows.Add(dataRow);
            }
            return dt;

        }

        private static DataTable CreateData<T>(T model)
        {
            DataTable dataTable = new DataTable(typeof(T).Name);
            foreach (PropertyInfo propertyInfo in typeof(T).GetProperties())
            {
                dataTable.Columns.Add(new DataColumn(propertyInfo.Name, propertyInfo.PropertyType));
            }
            return dataTable;
        }
        public static List<Dictionary<string, object>> ToJson(DataTable dt)
        {
            List<Dictionary<string, object>> list = new List<Dictionary<string, object>>();
            foreach (DataRow dr in dt.Rows)
            {
                Dictionary<string, object> result = new Dictionary<string, object>();
                foreach (DataColumn dc in dt.Columns)
                {
                    result.Add(dc.ColumnName, dr[dc]);
                }
                list.Add(result);
            }
            return list;
        }
        public static object GetColumns(DataTable dt)
        {
            string str = "[[";
            foreach (DataColumn dc in dt.Columns)
            {
                str+="{ field:'"+dc.ColumnName+"', title:'"+dc.ColumnName+"', align :'left' , width : "+12 * dc.ColumnName.Length +" },";
            }
            
            return str.Substring(0,str.Length-1)+"]]";
        }
        public static object GetCheckColumns(DataTable dt)
        {
            string str = "[[";
            str += "{ field:'ck', checkbox:true },";
            foreach (DataColumn dc in dt.Columns)
            {
                str += "{ field:'" + dc.ColumnName + "', title:'" + dc.ColumnName + "', align :'left' , width : " + 12 * dc.ColumnName.Length + " },";
            }

            return str.Substring(0, str.Length - 1) + "]]";
        }
     
    }
  
}

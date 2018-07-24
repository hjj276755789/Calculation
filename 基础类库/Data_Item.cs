using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace Calculation.Base
{
    public abstract class Data_Item<T>
       where T : Data_Item<T>, new()
    {
        public static IEnumerable<T> BuildFrom(DataTable table)
        {
            var rows = table.AsEnumerable().ToArray();
            return BuildFrom(rows);
        }
        public static IEnumerable<T> BuildFrom(DataRow[] rows)
        {
            return rows.Select(item => BuildFrom(item));
        }
        public static T BuildFrom(DataRow row)
        {
            if (row == null)
            {
                return null;
            }
            else
            {
                var t = new T();
                t.ReadDataRow(row);
                return t;
            }
        }
        /// <summary>
        /// 执行该方法将自动绑定DataRow数据，可以通过重载该方法来实现自定义的数据绑定
        /// </summary>
        /// <param name="row">DataRow</param>
        public virtual void ReadDataRow(DataRow row)
        {
            ReadDataRowLazy((T)this, row);
        }



        #region 私有方法
        private static void ReadDataRowLazy(T obj, DataRow row)
        {
            foreach (var property in typeof(T).GetProperties())
            {
                if (row.Table.Columns.Contains(property.Name))
                {
                    object value = row[property.Name];
                    if (value is DBNull)
                    {
                        continue;
                    }
                    else
                    {
                        property.SetValue(obj, GetObject(property, value.ToString()));
                    }
                }
            }

        }
        private static object GetObject(PropertyInfo property, string value)
        {
            if (property.PropertyType.IsEnum)
            {
                return int.Parse(value);
            }

            string propertyTypeName = IsNullable(property.PropertyType) ?
                                      property.PropertyType.GenericTypeArguments[0].Name :
                                      property.PropertyType.Name;


            return GetObject(propertyTypeName, value);
        }
        private static object GetObject(string propertyTypeName, string value)
        {
            switch (propertyTypeName.ToLower())
            {
                case "int16":
                    return Convert.ToInt16(value);
                case "int32":
                    return Convert.ToInt32(value);
                case "int64":
                    return Convert.ToInt64(value);
                case "string":
                    return Convert.ToString(value);
                case "datetime":
                    return Convert.ToDateTime(value);
                case "boolean":
                    {
                        if (value == "0")
                            return false;
                        else if (value == "1")
                            return true;
                        else
                            return Convert.ToBoolean(value);
                    }
                case "char":
                    return Convert.ToChar(value);
                case "double":
                    return Convert.ToDouble(value);
                default:
                    return value;
            }

        }
        private static bool IsNullable(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>);
        }
        #endregion
    }

}

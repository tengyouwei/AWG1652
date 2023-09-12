using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Reflection;

namespace CommonLibrary.EnumHelper
{
    /// <summary>
    /// 枚举帮助类
    /// 用于获取枚举的描述信息
    /// </summary>
    public class EnumHelper
    {
        /// <summary>
        /// 获取枚举描述信息
        /// </summary>
        /// <typeparam name="T">枚举类型</typeparam>
        /// <param name="value">枚举值</param>
        /// <returns></returns>
        public static string GetEnumDescription<T>(int value)
        {
            Type enumType = typeof(T);
            DescriptionAttribute attr = null;

            //获取枚举常数名称
            string name = System.Enum.GetName(enumType, value);
            if (!string.IsNullOrEmpty(name))
            {
                FieldInfo fieldInfo = enumType.GetField(name);
                if (fieldInfo != null)
                {
                    //获取描述属性
                    attr = Attribute.GetCustomAttribute(fieldInfo, typeof(DescriptionAttribute), false) as DescriptionAttribute;
                }
            }

            if (attr != null && !string.IsNullOrEmpty(attr.Description))
            {
                return attr.Description;
            }
            else
            {
                return String.Empty;
            }
        }



        public static string GetEnumDescription(System.Enum e)
        {
            if (e == null)
            {
                return string.Empty;
            }

            Type enumType = e.GetType();
            DescriptionAttribute attr = null;

            //获取枚举字段
            FieldInfo fieldInfo = enumType.GetField(e.ToString());
            if (fieldInfo != null)
            {
                //获取描述属性
                attr = Attribute.GetCustomAttribute(fieldInfo, typeof(DescriptionAttribute), false) as DescriptionAttribute;
            }

            //返回结果
            if (attr != null && !string.IsNullOrEmpty(attr.Description))
            {
                return attr.Description;
            }
            else
            {
                return String.Empty;
            }
        }

       


        public static List<EnumKeyValue> EnumToList<T>(bool isHasAll, params string[] filterItem)
        {
            List<EnumKeyValue> list = new List<EnumKeyValue>();

            //如果包含全部，则添加该项
            if (isHasAll)
            {
                list.Add(new EnumKeyValue() { Key = 0, Name = "全部" });
            }

            foreach (int item in System.Enum.GetValues(typeof(T)))
            {
                string name = System.Enum.GetName(typeof(T), item);
                //跳过过滤项
                if (Array.IndexOf<string>(filterItem, name) != -1)
                {
                    continue;
                }
                //添加
                EnumKeyValue model = new EnumKeyValue();
                model.Key = item;
                model.Name = name;
                list.Add(model);
            }

            return list;
        }




    }

    public class EnumKeyValue
    {
        public int Key { get; set; }

        public string Name { get; set; }
    }
}

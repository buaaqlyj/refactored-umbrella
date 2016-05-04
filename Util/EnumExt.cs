using System;
using System.ComponentModel;
using System.Reflection;

namespace Util
{
    public static class EnumExt
    {
        /// <summary>
        /// http://stackoverflow.com/a/4367868
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="description"></param>
        /// <returns></returns>
        public static T GetEnumValueFromDescription<T>(string description)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();
            foreach (var field in type.GetFields())
            {
                var attribute = Attribute.GetCustomAttribute(field,
                    typeof(DescriptionAttribute)) as DescriptionAttribute;
                if (attribute != null)
                {
                    if (attribute.Description == description)
                        return (T)field.GetValue(null);
                }
                else
                {
                    if (field.Name == description)
                        return (T)field.GetValue(null);
                }
            }
            throw new ArgumentException("Not found.", "description");
        }

        public static string GetDescriptionFromEnumValue<T>(T enumValue)
        {
            var type = typeof(T);
            if (!type.IsEnum) throw new InvalidOperationException();

            FieldInfo field = enumValue.GetType().GetField(enumValue.ToString());

            DescriptionAttribute attr = Attribute.GetCustomAttribute(field, typeof(DescriptionAttribute)) as DescriptionAttribute; ;

            return attr == null ? enumValue.ToString() : attr.Description;
        }
    }
}

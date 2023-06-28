using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace ExcelObjectMapper.Extensions
{
	internal static class ObjectExtensions
	{
		internal static void SetProperty(this object obj, string propertyName, object value)
		{
			string[] propertyPath = propertyName.Split('.');
			SetPropertyRecursive(obj, propertyPath, 0, value);
		}

		private static void SetPropertyRecursive(object obj, string[] propertyPath, int currentIndex, object value)
		{
			string currentProperty = propertyPath[currentIndex];

			var propertyInfo = obj.GetType().GetProperties()
				.FirstOrDefault(p => string.Equals(p.Name, currentProperty, StringComparison.OrdinalIgnoreCase));
			if (propertyInfo == null || !propertyInfo.CanWrite)
			{
				return;
			}

			if (currentIndex == propertyPath.Length - 1)
			{
				// Handle value types and DateTime
				if (propertyInfo.PropertyType.IsValueType || propertyInfo.PropertyType == typeof(DateTime))
				{
					object convertedValue = ConvertValueToType(value, propertyInfo.PropertyType);
					if (convertedValue == null)
					{
						return;
					}

					propertyInfo.SetValue(obj, convertedValue);
				}
				// Handle generic lists
				else if (propertyInfo.PropertyType.IsGenericType &&
								 propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
				{
					var list = propertyInfo.GetValue(obj);
					if (list == null)
					{
						list = Activator.CreateInstance(propertyInfo.PropertyType);
						propertyInfo.SetValue(obj, list);
					}

					var addMethod = propertyInfo.PropertyType.GetMethod("Add");
					addMethod.Invoke(list, new[] { value });
				}
				// Handle arrays
				else if (propertyInfo.PropertyType.IsArray)
				{
					var array = propertyInfo.GetValue(obj) as Array;
					if (array == null || array.Length == 0)
					{
						array = Array.CreateInstance(propertyInfo.PropertyType.GetElementType(), 1);
						array.SetValue(value, 0);
						propertyInfo.SetValue(obj, array);
					}
					else
					{
						array = ResizeArrayAndSetValue(array, value);
						propertyInfo.SetValue(obj, array);
					}
				}
				// Handle strings and other reference types
				else
				{
					propertyInfo.SetValue(obj, value.ToString());
				}
			}
			else
			{
				var nextObj = propertyInfo.GetValue(obj);
        if (nextObj == null)
        {
          if (propertyInfo.PropertyType == typeof(string))
          {
            nextObj = string.Empty;
          }
          else
          {
            nextObj = Activator.CreateInstance(propertyInfo.PropertyType);
          }

          propertyInfo.SetValue(obj, nextObj);
        }

        if (propertyInfo.PropertyType.IsGenericType &&
						propertyInfo.PropertyType.GetGenericTypeDefinition() == typeof(List<>))
				{
					var list = (IList)nextObj;
					if (list.Count == 0)
					{
						var itemType = propertyInfo.PropertyType.GetGenericArguments()[0];
						var newItem = Activator.CreateInstance(itemType);
						list.Add(newItem);
					}

					nextObj = list[0];
				}

				SetPropertyRecursive(nextObj, propertyPath, currentIndex + 1, value);
			}
		}

		private static Array ResizeArrayAndSetValue(Array array, object value)
		{
			int newSize = array.Length + 1;
			Type elementType = array.GetType().GetElementType();
			Array newArray = Array.CreateInstance(elementType, newSize);
			Array.Copy(array, newArray, array.Length);
			newArray.SetValue(value, newSize - 1);
			return newArray;
		}

		private static object ConvertValueToType(object value, Type targetType)
		{
			if (value == null)
			{
				return null;
			}
			if (targetType == typeof(DateTime) && DateTime.TryParse(value.ToString(), out DateTime dateTimeResult))
			{
				return dateTimeResult;
			}
			else if (targetType == typeof(double) && double.TryParse(value.ToString(), out double doubleResult))
			{
				return doubleResult;
			}
			else if (targetType == typeof(int) && int.TryParse(value.ToString(), out int intResult))
			{
				return intResult;
			}
			else if (targetType == typeof(decimal) && decimal.TryParse(value.ToString(), out decimal decimalResult))
			{
				return decimalResult;
			}
			else if (targetType == typeof(float) && float.TryParse(value.ToString(), out float floatResult))
			{
				return floatResult;
			}
			else if (targetType == typeof(bool) && bool.TryParse(value.ToString(), out bool boolResult))
			{
				return boolResult;
			}
			else
			{
				return null;
			}
		}
	}
}
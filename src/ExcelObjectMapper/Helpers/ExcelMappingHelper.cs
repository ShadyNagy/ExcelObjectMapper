using System.Collections.Generic;
using ExcelObjectMapper.Models;

namespace ExcelObjectMapper.Helpers
{
	/// <summary>
	/// A class to manage mapping between Excel column names and object property names.
	/// </summary>
	public class ExcelMappingHelper
	{
		private readonly Dictionary<string, string> _mapping = new Dictionary<string, string>();
		private readonly List<PropertyMapping> _propertyMapping = new List<PropertyMapping>();

		/// <summary>
		/// Creates a new instance of the ExcelMapping class.
		/// </summary>
		/// <returns>An instance of the ExcelMapping class.</returns>
		public static ExcelMappingHelper Create()
		{
			return new ExcelMappingHelper();
		}

		/// <summary>
		/// Adds a new mapping between a column name and a property name.
		/// </summary>
		/// <param name="propertyName">The corresponding object property name.</param>
		/// <param name="columnName">The Excel column name to map.</param>
		/// <returns>The current ExcelMapping instance, allowing for method chaining.</returns>
		public ExcelMappingHelper Add(string propertyName, string columnName)
		{
			_mapping.Add(columnName, propertyName);
			_propertyMapping.Add(new PropertyMapping(propertyName, columnName));

			return this;
		}

		/// <summary>
		/// Adds a new mapping between a column name and a property name.
		/// </summary>
		/// <param name="propertyName">The corresponding object property name.</param>
		/// <param name="columnName">The Excel column name to map.</param>
		/// <param name="staticData">The Excel column static to set.</param>
		/// <returns>The current ExcelMapping instance, allowing for method chaining.</returns>
		public ExcelMappingHelper Add(string propertyName, string columnName, object staticData)
		{
			_mapping.Add(columnName, propertyName);
			_propertyMapping.Add(new PropertyMapping(propertyName, columnName, staticData));

			return this;
		}

		/// <summary>
		/// Returns the current mapping dictionary.
		/// </summary>
		/// <returns>A dictionary containing the mapping between column names and property names.</returns>
		public Dictionary<string, string> Build()
		{
			return _mapping;
		}

		/// <summary>
		/// Returns the current property mapping list.
		/// </summary>
		/// <returns>A dictionary containing the mapping between column names and property names.</returns>
		public List<PropertyMapping> BuildPropertyMapping()
		{
			return _propertyMapping;
		}
	}
}
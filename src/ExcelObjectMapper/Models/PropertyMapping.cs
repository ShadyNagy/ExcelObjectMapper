namespace ExcelObjectMapper.Models
{

	/// <summary>
	/// Represents a mapping between a property name and an Excel column name, with an optional static value.
	/// </summary>
	public class PropertyMapping
	{
		/// <summary>
		/// Gets or sets the name of the property.
		/// </summary>
		public string PropertyName { get; set; }

		/// <summary>
		/// Gets or sets the name of the Excel column.
		/// </summary>
		public string ColumnName { get; set; }

		/// <summary>
		/// Gets or sets the static value to be assigned when the Excel column is not found.
		/// </summary>
		public object StaticValue { get; set; }

		/// <summary>
		/// Initializes a new instance of the <see cref="PropertyMapping"/> class.
		/// </summary>
		/// <param name="propertyName">The name of the property.</param>
		/// <param name="columnName">The name of the Excel column.</param>
		/// <param name="staticValue">The optional static value to be assigned when the Excel column is not found (default: null).</param>
		public PropertyMapping(string propertyName, string columnName, object staticValue = null)
		{
			PropertyName = propertyName;
			ColumnName = columnName;
			StaticValue = staticValue;
		}
	}
}
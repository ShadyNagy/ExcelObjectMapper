using System;
using System.Collections.Generic;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Xml;
using ExcelObjectMapper.Extensions;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using ExcelObjectMapper.Interfaces;
using ExcelObjectMapper.Models;

namespace ExcelObjectMapper.Readers
{
	/// <summary>
	/// A generic class for reading data from Excel sheets into a list of objects.
	/// </summary>
	/// <typeparam name="T">The type of objects to read data into.</typeparam>
	public class ExcelReader<T> : IExcelReader<T> where T : new()
	{
		private readonly ExcelPackage _package;

		/// <summary>
		/// Initializes a new instance of the ExcelReader class using the provided file path.
		/// </summary>
		/// <param name="excelFilePath">The file path of the Excel file to read.</param>
		public ExcelReader(string excelFilePath)
		{
			if (!File.Exists(excelFilePath))
			{
				throw new FileNotFoundException("File not found.", excelFilePath);
			}

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			_package = new ExcelPackage(new FileInfo(excelFilePath));
		}

		/// <summary>
		/// Initializes a new instance of the ExcelReader class using the provided byte array.
		/// </summary>
		/// <param name="excelFileBytes">The byte array of the Excel file to read.</param>
		public ExcelReader(byte[] excelFileBytes)
		{
			if (excelFileBytes == null || excelFileBytes.Length == 0)
			{
				throw new ArgumentException("Invalid excel file bytes.", nameof(excelFileBytes));
			}

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
			_package = new ExcelPackage(new MemoryStream(excelFileBytes));
		}

		/// <summary>
		/// Initializes a new instance of the ExcelReader class using the provided IFormFile instance.
		/// </summary>
		/// <param name="excelFile">The IFormFile instance representing the Excel file to read.</param>
		public ExcelReader(IFormFile excelFile)
		{
			if (excelFile == null || excelFile.Length == 0)
			{
				throw new ArgumentException("Invalid excel file.", nameof(excelFile));
			}

			ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

			using (var stream = new MemoryStream())
			{
				excelFile.CopyTo(stream);
				_package = new ExcelPackage(stream);
			}
		}

		/// <summary>
		/// Reads data from the first sheet into a list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <returns>A list of objects with data read from the first sheet.</returns>
		public List<T> ReadSheet(Dictionary<string, string> mapping)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheet(firstWorksheet.Name, mapping);
		}

    /// <summary>
    /// Reads data from the first sheet into a list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the first sheet.</returns>
    public List<T> ReadSheet(Dictionary<string, string> mapping, IReadOnlyList<string> requiredProperties)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheet(firstWorksheet.Name, mapping, requiredProperties);
		}

		/// <summary>
		/// Reads data from the first sheet into a list of objects using property mapping.
		/// </summary>
		/// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
		/// <returns>A list of objects with data read from the first sheet.</returns>
		public List<T> ReadSheet(IReadOnlyList<PropertyMapping> mapping)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheet(firstWorksheet.Name, mapping);
		}

    /// <summary>
    /// Reads data from the first sheet into a list of objects using property mapping.
    /// </summary>
    /// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the first sheet.</returns>
    public List<T> ReadSheet(IReadOnlyList<PropertyMapping> mapping, IReadOnlyList<string> requiredProperties)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheet(firstWorksheet.Name, mapping, requiredProperties);
		}

    /// <summary>
    /// Reads the first worksheet of the Excel file and returns a list of objects of type T where each object represents a row in the worksheet.
    /// The objects are created based on the provided mapping and filter, and only rows that contain non-null and non-empty values for all required properties are included in the list.
    /// </summary>
    /// <param name="mapping">A dictionary where the keys are property names of type T and the values are corresponding column names in the worksheet. The method uses this mapping to create the objects of type T.</param>
    /// <param name="filter">A function that takes an object of type T and returns a boolean. The method uses this function to filter the rows in the worksheet. Only rows for which the function returns true are included in the list.</param>
    /// <returns>A list of objects of type T where each object represents a row in the worksheet. The objects are created based on the provided mapping and filter, and only rows that contain non-null and non-empty values for all required properties are included in the list.</returns>
    /// <exception cref="InvalidOperationException">Thrown when no worksheets are found in the Excel file.</exception>
    public List<T> ReadSheetFiltered(Dictionary<string, string> mapping, Func<T, bool> filter)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetFiltered(firstWorksheet.Name, mapping, filter);
		}

    /// <summary>
    /// Reads the first worksheet of the Excel file and returns a list of objects of type T where each object represents a row in the worksheet.
    /// The objects are created based on the provided mapping and filter, and only rows that contain non-null and non-empty values for all required properties are included in the list.
    /// </summary>
    /// <param name="mapping">A dictionary where the keys are property names of type T and the values are corresponding column names in the worksheet. The method uses this mapping to create the objects of type T.</param>
    /// <param name="filter">A function that takes an object of type T and returns a boolean. The method uses this function to filter the rows in the worksheet. Only rows for which the function returns true are included in the list.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row in the worksheet contains null or empty values for any of these properties, the row will not be included in the list.</param>
    /// <returns>A list of objects of type T where each object represents a row in the worksheet. The objects are created based on the provided mapping and filter, and only rows that contain non-null and non-empty values for all required properties are included in the list.</returns>
    /// <exception cref="InvalidOperationException">Thrown when no worksheets are found in the Excel file.</exception>
    public List<T> ReadSheetFiltered(Dictionary<string, string> mapping, Func<T, bool> filter, IReadOnlyList<string> requiredProperties)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetFiltered(firstWorksheet.Name, mapping, filter, requiredProperties);
		}

		/// <summary>
		/// Reads data from the specified sheet into a sorted list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A sorted list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheetSorted(Dictionary<string, string> mapping, Comparison<T> comparison)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetSorted(firstWorksheet.Name, mapping, comparison);
		}

    /// <summary>
    /// Reads data from the specified sheet into a sorted list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A sorted list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheetSorted(Dictionary<string, string> mapping, Comparison<T> comparison, IReadOnlyList<string> requiredProperties)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetSorted(firstWorksheet.Name, mapping, comparison, requiredProperties);
		}

		/// <summary>
		/// Reads data from the specified sheet into a filtered and sorted list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheetFilteredAndSorted(Dictionary<string, string> mapping, Func<T, bool> filter,
			Comparison<T> comparison)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetFilteredAndSorted(firstWorksheet.Name, mapping, filter, comparison);
		}

    /// <summary>
    /// Reads data from the specified sheet into a filtered and sorted list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheetFilteredAndSorted(Dictionary<string, string> mapping, Func<T, bool> filter,
			Comparison<T> comparison, IReadOnlyList<string> requiredProperties)
		{
			var firstWorksheet = _package.Workbook.Worksheets.FirstOrDefault();
			if (firstWorksheet == null)
			{
				throw new InvalidOperationException("No worksheets found in the Excel file.");
			}

			return ReadSheetFilteredAndSorted(firstWorksheet.Name, mapping, filter, comparison, requiredProperties);
		}

		/// <summary>
		/// Reads data from the specified sheet into a list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <returns>A list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheet(string sheetName, Dictionary<string, string> mapping)
		{
			var result = new List<T>();
			var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
				string.Equals(sheetName, ws.Name, StringComparison.CurrentCultureIgnoreCase));

			if (worksheet == null || worksheet.Dimension == null)
			{
				return result;
			}

			var rowCount = worksheet.Dimension.End.Row;
			var columnCount = worksheet.Dimension.End.Column > 100000 ? 100000 : worksheet.Dimension.End.Column;

			for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
			{
				var toAdd = new T();

				for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
				{
					var columnName = GetColumnNameByIndex(worksheet, columnIndex);
					var mapped = mapping.FirstOrDefault(x => string.Equals(x.Value.RemoveSpecialCharacters(),
						columnName.RemoveSpecialCharacters(), StringComparison.CurrentCultureIgnoreCase));
					if (string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(mapped.Key))
					{
						continue;
					}

					var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;

					toAdd.SetProperty(mapped.Key, cellValue);
				}

				result.Add(toAdd);
			}

			return result;
		}

    /// <summary>
    /// Reads data from the specified worksheet into a list of dynamic objects (ExpandoObject).
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <returns>A list of dynamic objects with data read from the specified sheet. If the sheet doesn't exist or has no data, an empty list is returned.</returns>
    public dynamic ReadSheetToDynamic(string sheetName, Dictionary<string, string> mapping)
    {
      var result = new List<dynamic>();
      var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
        string.Equals(sheetName, ws.Name, StringComparison.CurrentCultureIgnoreCase));

      if (worksheet == null || worksheet.Dimension == null)
      {
        return result;
      }

      var rowCount = worksheet.Dimension.End.Row;
      var columnCount = worksheet.Dimension.End.Column > 100000 ? 100000 : worksheet.Dimension.End.Column;

      for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
      {
        var toAdd = new ExpandoObject() as IDictionary<string, Object>;

        for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
        {
          var columnName = GetColumnNameByIndex(worksheet, columnIndex);
          var mapped = mapping.FirstOrDefault(x => string.Equals(x.Value.RemoveSpecialCharacters(),
            columnName.RemoveSpecialCharacters(), StringComparison.CurrentCultureIgnoreCase));
          if (string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(mapped.Key))
          {
            continue;
          }

          var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;

          // Using the IDictionary interface to add new properties to the ExpandoObject
          toAdd[mapped.Key] = cellValue;
        }

        result.Add(toAdd);
      }

      return result;
    }

    /// <summary>
    /// Reads data from the specified sheet into a list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheet(string sheetName, Dictionary<string, string> mapping, IReadOnlyList<string> requiredProperties)
    {
      var result = new List<T>();
      var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
          string.Equals(sheetName, ws.Name, StringComparison.CurrentCultureIgnoreCase));

      if (worksheet == null || worksheet.Dimension == null)
      {
        return result;
      }

      var rowCount = worksheet.Dimension.End.Row;
      var columnCount = worksheet.Dimension.End.Column > 100000 ? 100000 : worksheet.Dimension.End.Column;

      for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
      {
        var toAdd = new T();
        bool shouldAdd = true;

        for (var columnIndex = 1; columnIndex <= columnCount; columnIndex++)
        {
          var columnName = GetColumnNameByIndex(worksheet, columnIndex);
          var mapped = mapping.FirstOrDefault(x => string.Equals(x.Value.RemoveSpecialCharacters(),
              columnName.RemoveSpecialCharacters(), StringComparison.CurrentCultureIgnoreCase));
          if (string.IsNullOrEmpty(columnName) || string.IsNullOrEmpty(mapped.Key))
          {
            continue;
          }

          var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;
          toAdd.SetProperty(mapped.Key, cellValue);

          // Check if the property is in the required properties list and if it's null or empty.
          if (requiredProperties.Contains(mapped.Key))
          {
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
            {
              shouldAdd = false;
              break;
            }
          }
        }

        // Only add to the result if all required properties are not null or empty.
        if (shouldAdd)
        {
          result.Add(toAdd);
        }
      }

      return result;
    }

    /// <summary>
    /// Reads data from the specified sheet into a list of objects using property mapping.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
    /// <returns>A list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheet(string sheetName, IReadOnlyList<PropertyMapping> mapping)
		{
			var result = new List<T>();

			var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
				string.Equals(sheetName, ws.Name, StringComparison.CurrentCultureIgnoreCase));

			if (worksheet == null || worksheet.Dimension == null)
			{
				return result;
			}

			var rowCount = worksheet.Dimension.End.Row;
			var columnCount = worksheet.Dimension.End.Column > 100000 ? 100000 : worksheet.Dimension.End.Column;

			for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
			{
				var toAdd = new T();

				foreach (var entry in mapping)
				{
					var columnName = entry.ColumnName.RemoveTabAndEnter();
					var columnIndex = GetColumnIndexByName(worksheet, columnName);

					if (columnIndex == -1)
					{
						if (entry.StaticValue != null)
						{
							toAdd.SetProperty(entry.PropertyName, entry.StaticValue);
						}

						continue;
					}

					var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;

					toAdd.SetProperty(entry.PropertyName, cellValue);
				}

				result.Add(toAdd);
			}

			return result;
		}

    /// <summary>
    /// Reads data from the specified sheet into a list of objects using property mapping.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
		/// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheet(string sheetName, IReadOnlyList<PropertyMapping> mapping, IReadOnlyList<string> requiredProperties)
    {
      var result = new List<T>();

      var worksheet = _package.Workbook.Worksheets.FirstOrDefault(ws =>
          string.Equals(sheetName, ws.Name, StringComparison.CurrentCultureIgnoreCase));

      if (worksheet == null || worksheet.Dimension == null)
      {
        return result;
      }

      var rowCount = worksheet.Dimension.End.Row;
      var columnCount = worksheet.Dimension.End.Column > 100000 ? 100000 : worksheet.Dimension.End.Column;

      for (var rowIndex = 2; rowIndex <= rowCount; rowIndex++)
      {
        var toAdd = new T();
        bool shouldAdd = true;

        foreach (var entry in mapping)
        {
          var columnName = entry.ColumnName.RemoveTabAndEnter();
          var columnIndex = GetColumnIndexByName(worksheet, columnName);

          if (columnIndex == -1)
          {
            if (entry.StaticValue != null)
            {
              toAdd.SetProperty(entry.PropertyName, entry.StaticValue);
            }

            continue;
          }

          var cellValue = worksheet.Cells[rowIndex, columnIndex].Value;
          toAdd.SetProperty(entry.PropertyName, cellValue);

          // Check if the property is in the required properties list and if it's null or empty.
          if (requiredProperties.Contains(entry.PropertyName))
          {
            if (cellValue == null || string.IsNullOrWhiteSpace(cellValue.ToString()))
            {
              shouldAdd = false;
              break;
            }
          }
        }

        // Only add to the result if all required properties are not null or empty.
        if (shouldAdd)
        {
          result.Add(toAdd);
        }
      }

      return result;
    }

    /// <summary>
    /// Retrieves metadata from the Excel file.
    /// </summary>
    /// <returns>A read-only dictionary containing metadata key-value pairs.</returns>
    public IReadOnlyDictionary<string, string> GetSheetMetadata()
		{
			var metadata = new Dictionary<string, string>();
			var customPropertiesXml = _package.Workbook.Properties.CustomPropertiesXml;

			if (customPropertiesXml?.DocumentElement == null)
			{
				return metadata;
			}

			var xmlNamespaceManager = new XmlNamespaceManager(customPropertiesXml.NameTable);
			xmlNamespaceManager.AddNamespace("vt", "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes");

			var properties = customPropertiesXml.DocumentElement.SelectNodes("property");
			if (properties == null)
			{
				return metadata;
			}

			foreach (XmlNode prop in properties)
			{
				if (prop.Attributes?["name"] == null)
				{
					continue;
				}

				var name = prop.Attributes["name"].Value;
				var value = prop.SelectSingleNode("vt:lpwstr", xmlNamespaceManager)?.InnerText;
				if (value != null)
				{
					metadata[name] = value;
				}
			}

			return metadata;
		}

		/// <summary>
		/// Reads data from the specified sheet into a filtered list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <returns>A filtered list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheetFiltered(string sheetName, Dictionary<string, string> mapping, Func<T, bool> filter)
		{
			var data = ReadSheet(sheetName, mapping);
			return data.Where(filter).ToList();
		}

    /// <summary>
    /// Reads data from the specified sheet into a filtered list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheetFiltered(string sheetName, Dictionary<string, string> mapping, Func<T, bool> filter, IReadOnlyList<string> requiredProperties)
		{
			var data = ReadSheet(sheetName, mapping, requiredProperties);
			return data.Where(filter).ToList();
		}

		/// <summary>
		/// Reads data from the specified sheet into a sorted list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A sorted list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheetSorted(string sheetName, Dictionary<string, string> mapping, Comparison<T> comparison)
		{
			var data = ReadSheet(sheetName, mapping);
			data.Sort(comparison);
			return data;
		}

    /// <summary>
    /// Reads data from the specified sheet into a sorted list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A sorted list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheetSorted(string sheetName, Dictionary<string, string> mapping, Comparison<T> comparison, IReadOnlyList<string> requiredProperties)
		{
			var data = ReadSheet(sheetName, mapping, requiredProperties);
			data.Sort(comparison);
			return data;
		}

		/// <summary>
		/// Reads data from the specified sheet into a filtered and sorted list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
		public List<T> ReadSheetFilteredAndSorted(string sheetName, Dictionary<string, string> mapping,
			Func<T, bool> filter, Comparison<T> comparison)
		{
			var data = ReadSheet(sheetName, mapping);
			var filteredData = data.Where(filter).ToList();
			filteredData.Sort(comparison);
			return filteredData;
		}

    /// <summary>
    /// Reads data from the specified sheet into a filtered and sorted list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
    public List<T> ReadSheetFilteredAndSorted(string sheetName, Dictionary<string, string> mapping,
			Func<T, bool> filter, Comparison<T> comparison, IReadOnlyList<string> requiredProperties)
		{
			var data = ReadSheet(sheetName, mapping, requiredProperties);
			var filteredData = data.Where(filter).ToList();
			filteredData.Sort(comparison);
			return filteredData;
		}

		private static string GetColumnNameByIndex(ExcelWorksheet sheet, int columnIndex)
		{
			var name = sheet.Cells[1, columnIndex].FirstOrDefault()?.Value;
			return name == null ? string.Empty : name.ToString();
		}

		private static int GetColumnIndexByName(ExcelWorksheet sheet, string columnName)
		{
			var columnIndex = -1;

			for (int i = 1; i <= sheet.Dimension.End.Column; i++)
			{
				if (string.Equals(sheet.Cells[1, i].Value?.ToString(), columnName, StringComparison.OrdinalIgnoreCase))
				{
					columnIndex = i;
					break;
				}
			}

			return columnIndex;
		}
	}
}
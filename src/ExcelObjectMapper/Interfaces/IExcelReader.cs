﻿using System;
using System.Collections.Generic;
using ExcelObjectMapper.Models;

namespace ExcelObjectMapper.Interfaces
{

	/// <summary>
	/// Interface for reading data from Excel sheets into a list of objects.
	/// </summary>
	/// <typeparam name="T">The type of objects to read data into.</typeparam>
	public interface IExcelReader
	{
		/// <summary>
		/// Reads data from the specified sheet into a list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <returns>A list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheet(string sheetName, Dictionary<string, string> mapping);

    /// <summary>
    /// Reads data from the specified sheet into a list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheet(string sheetName, Dictionary<string, string> mapping, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the first sheet into a list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <returns>A list of objects with data read from the first sheet.</returns>
		List<dynamic> ReadSheet(Dictionary<string, string> mapping);

    /// <summary>
    /// Reads data from the first sheet into a list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the first sheet.</returns>
    List<dynamic> ReadSheet(Dictionary<string, string> mapping, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a list of objects using property mapping.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
		/// <returns>A list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheet(string sheetName, IReadOnlyList<PropertyMapping> mapping);

    /// <summary>
    /// Reads data from the specified sheet into a list of objects using property mapping.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheet(string sheetName, IReadOnlyList<PropertyMapping> mapping, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the first sheet into a list of objects using property mapping.
		/// </summary>
		/// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
		/// <returns>A list of objects with data read from the first sheet.</returns>
		List<dynamic> ReadSheet(IReadOnlyList<PropertyMapping> mapping);

    /// <summary>
    /// Reads data from the first sheet into a list of objects using property mapping.
    /// </summary>
    /// <param name="mapping">A list of property mappings that map object property names to Excel column names and static values.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A list of objects with data read from the first sheet.</returns>
    List<dynamic> ReadSheet(IReadOnlyList<PropertyMapping> mapping, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Retrieves metadata from the Excel file.
		/// </summary>
		/// <returns>A read-only dictionary containing metadata key-value pairs.</returns>
		IReadOnlyDictionary<string, string> GetSheetMetadata();

		/// <summary>
		/// Reads data from the specified sheet into a filtered list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <returns>A filtered list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetFiltered(string sheetName, Dictionary<string, string> mapping, Func<dynamic, bool> filter);

    /// <summary>
    /// Reads data from the specified sheet into a filtered list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetFiltered(string sheetName, Dictionary<string, string> mapping, Func<dynamic, bool> filter, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a filtered list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <returns>A filtered list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetFiltered(Dictionary<string, string> mapping, Func<dynamic, bool> filter);

    /// <summary>
    /// Reads data from the specified sheet into a filtered list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetFiltered(Dictionary<string, string> mapping, Func<dynamic, bool> filter, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a sorted list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A sorted list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetSorted(string sheetName, Dictionary<string, string> mapping, Comparison<dynamic> comparison);

    /// <summary>
    /// Reads data from the specified sheet into a sorted list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A sorted list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetSorted(string sheetName, Dictionary<string, string> mapping, Comparison<dynamic> comparison, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a sorted list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A sorted list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetSorted(Dictionary<string, string> mapping, Comparison<dynamic> comparison);

    /// <summary>
    /// Reads data from the specified sheet into a sorted list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A sorted list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetSorted(Dictionary<string, string> mapping, Comparison<dynamic> comparison, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a filtered and sorted list of objects.
		/// </summary>
		/// <param name="sheetName">The name of the sheet to read data from.</param>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetFilteredAndSorted(string sheetName, Dictionary<string, string> mapping, Func<dynamic, bool> filter,
			Comparison<dynamic> comparison);

    /// <summary>
    /// Reads data from the specified sheet into a filtered and sorted list of objects.
    /// </summary>
    /// <param name="sheetName">The name of the sheet to read data from.</param>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetFilteredAndSorted(string sheetName, Dictionary<string, string> mapping, Func<dynamic, bool> filter,
			Comparison<dynamic> comparison, IReadOnlyList<string> requiredProperties);

		/// <summary>
		/// Reads data from the specified sheet into a filtered and sorted list of objects.
		/// </summary>
		/// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
		/// <param name="filter">A function to filter the rows based on their data.</param>
		/// <param name="comparison">A comparison function to sort the rows based on their data.</param>
		/// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
		List<dynamic> ReadSheetFilteredAndSorted(Dictionary<string, string> mapping, Func<dynamic, bool> filter,
			Comparison<dynamic> comparison);

    /// <summary>
    /// Reads data from the specified sheet into a filtered and sorted list of objects.
    /// </summary>
    /// <param name="mapping">A dictionary that maps object property names to Excel column names.</param>
    /// <param name="filter">A function to filter the rows based on their data.</param>
    /// <param name="comparison">A comparison function to sort the rows based on their data.</param>
    /// <param name="requiredProperties">A list of property names that are required to have non-null and non-empty values. If a row contains null or empty values for any of these properties, the row will not be included in the result list.</param>
    /// <returns>A filtered and sorted list of objects with data read from the specified sheet.</returns>
    List<dynamic> ReadSheetFilteredAndSorted(Dictionary<string, string> mapping, Func<dynamic, bool> filter,
			Comparison<dynamic> comparison, IReadOnlyList<string> requiredProperties);
	}
}
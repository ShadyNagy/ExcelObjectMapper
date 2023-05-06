[![publish to nuget](https://github.com/ShadyNagy/ExcelObjectMapper/actions/workflows/nugt-publish.yml/badge.svg)](https://github.com/ShadyNagy/ExcelObjectMapper/actions/workflows/nugt-publish.yml)
[![ExcelObjectMapper on NuGet](https://img.shields.io/nuget/v/ExcelObjectMapper?label=ExcelObjectMapper)](https://www.nuget.org/packages/ExcelObjectMapper/)
[![NuGet](https://img.shields.io/nuget/dt/ExcelObjectMapper)](https://www.nuget.org/packages/ExcelObjectMapper)
[![License](https://img.shields.io/badge/License-MIT-blue.svg)](https://github.com/ShadyNagy/ExcelObjectMapper/blob/main/LICENSE)
[![paypal](https://img.shields.io/badge/PayPal-tip%20me-green.svg?logo=paypal)](https://www.paypal.me/shadynagy)

# ExcelObjectMapper

A simple and efficient .NET library for mapping Excel files to C# objects. This library makes it easy to read Excel files and convert rows to custom C# objects, with minimal code.

## Features

- Read Excel files (.xlsx) with ease
- Map Excel rows to C# objects
- Auto-detect column names and map them to C# object properties
- Easily specify custom column-to-property mappings
- Supports .NET Standard 2.0 and higher

## Getting Started

### Installation

Install the ExcelObjectMapper package from NuGet:

```
Install-Package ExcelObjectMapper
```

### Usage

1. Create a class that represents the structure of your Excel file:

```csharp
public class Employee
{
    public int EmployeeId { get; set; }
    public string FirstName { get; set; }
    public string LastName { get; set; }
    public DateTime DateOfBirth { get; set; }
}
```
2. Read the Excel file and map the rows to the `Employee` class:

```csharp
using ExcelObjectMapper.Readers;

// ...

var filePath = "path/to/your/excel-file.xlsx";
var mapping = new Dictionary<string, string>
{
    { "EmployeeId", "Excel ColumnName1" },
    { "FirstName", "Excel ColumnName2" },
    { "LastName", "Excel ColumnName3" },
    { "DateOfBirth", "Excel ColumnName4" }
};

var excelReader = new ExcelReader<Employee>(filePath);
var data = excelReader.ReadSheet(mapping);

foreach (var employee in employees)
{
    Console.WriteLine($"{employee.EmployeeId}: {employee.FirstName} {employee.LastName}");
}
```

### Examples

#### Reading Excel file using `byte[]`

You can read an Excel file directly from a byte array:

```csharp
using System.IO;
using ExcelObjectMapper.Readers;

byte[] fileBytes = File.ReadAllBytes("path/to/your/excel-file.xlsx");
var mapping = new Dictionary<string, string>
{
    { "EmployeeId", "Excel ColumnName1" },
    { "FirstName", "Excel ColumnName2" },
    { "LastName", "Excel ColumnName3" },
    { "DateOfBirth", "Excel ColumnName4" }
};

var excelReader = new ExcelReader<Employee>(fileBytes);
var data = excelReader.ReadSheet(mapping);

foreach (var employee in employees)
{
    Console.WriteLine($"{employee.EmployeeId}: {employee.FirstName} {employee.LastName}");
}
```

#### Reading Excel file using `IFormFile` in ASP.NET Core

When working with file uploads in ASP.NET Core, you can use the `IFormFile` interface to read an Excel file:

```csharp
using ExcelObjectMapper;
using Microsoft.AspNetCore.Http;

public async Task<IActionResult> Upload(IFormFile file)
{
    var mapping = ExcelMappingHelper.Create()
    .Add("EmployeeId", "Excel ColumnName1")
    .Add("FirstName", "Excel ColumnName2")
    .Add("LastName", "Excel ColumnName3")
    .Add("DateOfBirth", "Excel ColumnName4")
    .Build();
    
    var excelReader = new ExcelReader<Employee>(file);
    var data = excelReader.ReadSheet(mapping);

    foreach (var employee in employees)
    {
        Console.WriteLine($"{employee.EmployeeId}: {employee.FirstName} {employee.LastName}");
    }

    return Ok("File processed successfully.");
}
```

These examples demonstrate how you can read Excel files using different input sources with the ExcelObjectMapper library.

For more information, please refer to the [documentation](https://github.com/ShadyNagy/ExcelObjectMapper/blob/main/README.md).

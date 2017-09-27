# Genexcel
Wrapper for easily generating excel files using OpenXml

[![Build status](https://ci.appveyor.com/api/projects/status/jrh1n1jmk8glk1xb?svg=true)](https://ci.appveyor.com/project/guimabdo/genexcel) [![Latest version](https://img.shields.io/nuget/v/Genexcel.svg)](https://www.nuget.org/packages?q=genexcel)

## Install
PM> Install-Package Genexcel

## Usage
```csharp
var excel = new Document();

//Get the first sheet 
var sheet1 = excel.GetSheets().First();
//Change sheet name 
sheet1.Name = "My first sheet";

//Write some data 
sheet1.Add(new Cell(1, 1, "Test 1"));

//Create another sheet 
var sheet2 = excel.AddSheet("My second sheet");

//Write some data 
sheet2.Add(new Cell(1, 1, "Test 2"));

//Save to file, or stream... 
excel.Save("myFile.xlsx");
```
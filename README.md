# LINQ to Excel Examples

This project contains examples for using LINQ to process Excel files.
Currently, I'm implementing the examples (which are done in Microsoft Excel 2007) in [https://post.smzdm.com/p/alpokzgo/](https://post.smzdm.com/p/alpokzgo/) via LINQ to Excel in C#.

## Setup
To setup, simply clone the project and open it with Visual studio 2019, and run the C# console application.
Or, you can just create a new C# project in Visual studio targeting .NET Core 3.0/.NET Standard 2.1 or above (this is required since this project use C# 8.0 features), install NuGet packages: System.ValueTuple, LinqToExcel, Remotion.Linq, MoreLINQ and ComparerExtensions, and then copy the code to your C# file to have a try.

#### Access Database Engine
In order to use LinqToExcel, you need to install the [Microsoft Access Database Engine Redistributable](https://www.microsoft.com/en-us/download/confirmation.aspx?id=54920). If it's not installed, you'll get the following exception:

	The 'Microsoft.ACE.OLEDB.12.0' provider is not registered on the local machine.'

* Both a 32-bit and 64-bit version are available, select the one that matches your project settings.
* You can only have one of them installed at a time.

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LinqToExcel;
using LinqToExcel.Query;
using MoreLinq.Extensions;

// Notice: SheetRowRange(2, 4) means rows [2, 4) in the Excel sheet, and the index start from 1 just as Microsoft Excel,
// so row 2 is actual the first data row and the first item in the `ExcelQueryable<Row>` with index 0.
using SheetRowRange = System.ValueTuple<int, int>;
using SheetColumnRange = System.ValueTuple<char, char>;
using CellPosition = System.ValueTuple<char, int>;
using System.Diagnostics.Contracts;
using ComparerExtensions;

namespace ExcelLinqApp {

  sealed class Program {

    static void Main(string[] args) {

      string currentDirectory = System.IO.Directory.GetCurrentDirectory();
      Program.ExcelFilePath = Path.Combine(currentDirectory, "../../Resources", "TestInput.xlsx");
      Console.WriteLine($"Input xlsx file path: {Program.ExcelFilePath}.");
      var separatorLine = String.Concat(Enumerable.Repeat('-', 60));
      Console.WriteLine(separatorLine);

      Console.WriteLine("Example 1: Count distinct elements");
      {
        var sheetName = "Sheet1";
        var columnName = "UserName";
        var rowRange = (2, 15);
        var distinctElementsCount = RunExample1CountDistinctElements(sheetName, columnName, rowRange);
        Console.WriteLine($"Find {distinctElementsCount} elements in column {columnName}, row { rowRange.Item1 } to { rowRange.Item2 } from {sheetName}.");
      }
      Console.WriteLine(separatorLine);

      Console.WriteLine("Example 3: VLOOKUP without custom function");
      {
        var findValue = RunExample3VlookupV1();
        Console.WriteLine($"Find Value: {findValue} in Sheet3.");
      }
      Console.WriteLine(separatorLine);

      Console.WriteLine("Example 3: VLOOKUP with custom function");
      {
        string lookUpValue = ProcessExcelData("Sheet3", sheet => sheet.CellValueAsString(8, "Id"));
        var sheetRowRange = (2, 7);
        string[] columns = { "Id", "Length" };
        var selectColumnIndex = 1;
        bool ValueComparator(string str1, string str2) => str1 == str2; // local function (since C# 7)

        var findValue = RunExample3VlookupV2(lookUpValue, sheetRowRange, columns, selectColumnIndex, ValueComparator);
        Console.WriteLine($"Find Value: {findValue} in Sheet3.");
      }
      Console.WriteLine(separatorLine);

      Console.WriteLine("Example 6: Continuoues rank");
      {
        RunExample6ContinuousRank((row, rank) => Console.WriteLine($"Name: {row["UserName"].Value} , Rank: {rank}, Length: {row["Length"].Value}."));
      }
      Console.WriteLine(separatorLine);

      Console.WriteLine("Example 6: Multi rows rank");
      {
        RunExample6MultiFieldsRank((row, rank) => Console.WriteLine($"Name: {row["UserName"].Value} , Rank: {rank}."));
      }
      Console.WriteLine(separatorLine);

      Console.ReadKey();
    }

    private static String ExcelFilePath;

    private static T ProcessExcelData<T>(string sheetName, Func<ExcelQueryable<Row>, T> processFunc) {
      using (var execelfile = new ExcelQueryFactory(ExcelFilePath)) {
        var tsheet = execelfile.Worksheet(sheetName);
        return processFunc(tsheet);
      }
    }

    private static int RunExample1CountDistinctElements(string sheetName, string columnName, ValueTuple<int, int> rowRange) {
      return ProcessExcelData<int>(sheetName, (sheet) => {
        return sheet.Slice(rowRange)
                    .DistinctBy(e => e[columnName].Value.ToString()) // or use Select( e => ... ).Distinct()
                    .Count();
      });
    }

    private static String RunExample3VlookupV1() {
      var columnNames = new string[2] { "Id", "Length" };
      return ProcessExcelData("Sheet3", (sheet) => {
        // FIXME: `ElementAt` not supported
        var content = (sheet.Slice(6, 6).First())["Id"].Value.ToString(); // get Cell at row (6+2) and column `Id`
        var selectValues = sheet.Slice(0, 4) // get Row[1:1+5]
                           .FirstOrDefault(row => columnNames.Any(c => row[c].Value.ToString() == content));

        if (selectValues != null) { return selectValues[columnNames[1]].Value.ToString(); }
        else { return null; }

        // Note: Error: sub query not supported
        // .Select(e => new string[2] { e["Id"].Value.ToString(), e["Length"].Value.ToString() })
      });
    }

    private static String RunExample3VlookupV2(string lookUpValue,
                                               SheetRowRange sheetRowRange,
                                               string[] columns,
                                               int selectColumnIndex,
                                               Func<string, string, bool> valueComparator) {
      return ProcessExcelData("Sheet3", (sheet) =>
        sheet.Vlookup(lookUpValue, sheetRowRange, columns, LinqExtensions.cellValueToStringFunc, selectColumnIndex, valueComparator)
      );
    }

    private static int RunExample6ContinuousRank(Action<Row, int> rowRankHandler) {
      return ProcessExcelData<int>("Sheet6", (sheet) => {

        var sheetRows = sheet.ToArray(); // Hack for the issue that group is not supported, convert to objects to use LINQ to objects
        var groups = sheetRows.GroupBy(e => (double)(e["Length"].Value));    // group items by length
        var groupsRank = groups.RankBy(group => group.Key);                  // rank groups
        // var groupWithRanks = groups.Zip(groupsRank, (group, rank) => Tuple.Create(group, rank));   // zip groups with their ranks

        var groupsRankEnumerator = groupsRank.GetEnumerator();
        foreach (IGrouping<double, Row> group in groups) {
          groupsRankEnumerator.MoveNext();
          var groupRank = groupsRankEnumerator.Current;
          foreach (Row row in group) {
            rowRankHandler(row, groupRank);
          }
        }

        return 0;
      });
    }

    private static int RunExample6MultiFieldsRank(Action<Row, int> rowRankHandler) {
      return ProcessExcelData("Sheet6", (sheet) => {
        var comparer = KeyComparer<Row>.OrderBy(e => e["Type"].Value.ToString())
                                       .ThenBy(e => (double)e["Length"].Value)
                                       .ThenBy(e => (double)e["HP"].Value);
        var rowsRank = sheet.Rank(comparer);

        var rowsRankEnumerator = rowsRank.GetEnumerator();
        foreach (var row in sheet.ToArray()) {
          rowsRankEnumerator.MoveNext();
          var rowRank = rowsRankEnumerator.Current;
          rowRankHandler(row, rowRank);
        }

        return 0;
      });
    }

  }

  public static class LinqExtensions {

    public static readonly Func<Cell, string> cellValueToStringFunc = cell => cell.Value.ToString();

    public static int RowNumberToQueryIndex(int rowNumber) { return rowNumber - 2; }

    public static IEnumerable<T> Slice<T>(this IEnumerable<T> sequence, SheetRowRange sheetRowRange) {
      Contract.Requires(sheetRowRange.Item1 <= sheetRowRange.Item2);

      return sequence.Slice(sheetRowRange.Item1 - 2, sheetRowRange.Item2 - sheetRowRange.Item1);
    }

    public static T CellValue<T>(this ExcelQueryable<Row> sheet, int rowNumber, string columnName, Func<Cell, T> cellToValue) {
      Contract.Requires(rowNumber > 1); // Excel row number start from 1 and row 1 is for colunm names
      Contract.Requires(columnName != null && columnName.Length == 0);
      Contract.Requires(cellToValue != null);

      var row = sheet.Slice(RowNumberToQueryIndex(rowNumber), 1).Single();
      return cellToValue(row[columnName]);
    }

    public static string CellValueAsString(this ExcelQueryable<Row> sheet, int rowNumber, string columnName) {
      return CellValue(sheet, rowNumber, columnName, cellValueToStringFunc);
    }

    // VLOOKUP(lookup_value,table_array,col_index_num,range_lookup)
    public static T Vlookup<T>(this ExcelQueryable<Row> rows,
                               T lookUpValue,
                               SheetRowRange sheetRowRange,
                               string[] columns,
                               Func<Cell, T> cellValueSelector,
                               int selectColumnIndex,
                               Func<T, T, bool> valueComparator) {

      Contract.Requires(columns != null && columns.Length > 0, $"{nameof(columns)} can not be null.");
      Contract.Requires(selectColumnIndex >= 0 && selectColumnIndex < columns.Length,
                        $"{nameof(selectColumnIndex)} : {selectColumnIndex} should be in the indexes of ${nameof(columns)}. ");

      // Note: `Requires<ArgumentException>` requires Code Contracts to do binary rewrite
      //Contract.Requires<ArgumentException>(columns != null && columns.Length > 0,
      //                                     $"{nameof(columns)} can not be null.");
      //Contract.Requires<ArgumentException>(selectColumnIndex >= 0 && selectColumnIndex < columns.Length,
      //                                     $"{nameof(selectColumnIndex)} : {selectColumnIndex} should be in the indexes of ${nameof(columns)}. ");

      Row selectRowOrNull = rows.Slice(sheetRowRange)
                                .FirstOrDefault(row => columns.Any(col => valueComparator(cellValueSelector(row[col]), lookUpValue)));

      if (selectRowOrNull != null) { return cellValueSelector(selectRowOrNull[columns[selectColumnIndex]]); }
      else { return default; }
    }
  }

}

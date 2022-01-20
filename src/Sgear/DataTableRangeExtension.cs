using System;
using System.Data;
using System.Collections.Generic;

using Microsoft.Extensions.Logging;
using SpreadsheetGear;
using SpreadsheetGear.Advanced.Cells;


namespace Tlabs.CalcNgn.Sgear {
  using CellType= SpreadsheetGear.Advanced.Cells.ValueType;

  internal static class DataTableRangeExtension {
    private static readonly ILogger<Intern.ICalcNgnModelParser> log= App.Logger<Intern.ICalcNgnModelParser>();

    private static IFormatProvider FRMP= System.Globalization.CultureInfo.InvariantCulture;
    private static List<string> ERRmap;
    static DataTableRangeExtension() {
      ERRmap= new List<string>(new string[32]);
      ERRmap[(int)ValueError.None]= null;
      ERRmap[(int)ValueError.Div0]= "#DIV/0!";
      ERRmap[(int)ValueError.NA]= "#N/A";
      ERRmap[(int)ValueError.Name]= "#NMAE?";
      ERRmap[(int)ValueError.Null]= "#NULL!";
      ERRmap[(int)ValueError.Num]= "#NUM!";
      ERRmap[(int)ValueError.Ref]= "#REF!";
      ERRmap[(int)ValueError.Value]= "#VALUE!";
    }

    public static void CopyFromDataTable(this IRange rng, DataTable data, bool insertRows, bool hasHeader) {
      if (null == data) throw new ArgumentNullException(nameof(data));

      int nCols= Math.Min(data.Columns.Count, rng.ColumnCount);
      int nRows= data.Rows.Count;
      int startCol= rng.Column;
      int startRow= rng.Row;
      int hdrRow= startRow -1;
      IValues cells= (IValues)rng.Worksheet;
      string[] hdr= null;

      if (hasHeader) {
        if (hdrRow < 0) throw new ArgumentException("No header above range.");
        hdr= new string[nCols];
        for (int l= 0; l < nCols; ++l)
          if (null == (hdr[l]= cells[hdrRow, startCol+l]?.Text)) {
            nCols= l;
            break;//throw new CalcNgnCellException(hdrRow, startCol+l, "No header key");
          }
      }
      int endCol= startCol + nCols;
      int endRow= startRow + nRows;

      if (insertRows) {
        if (rng.RowCount != 2) throw new ArgumentException("Dynamic range must have exactly two rows.");
        if (rng.RowCount < nRows) {  //extend range
          var insRng= rng[1, 0, nRows-2, 0].EntireRow;
          log.LogDebug("Inserting range: {rng}", insRng.RelativeAddr());
          insRng.Insert();
        }
        if (nRows > 1) {
          var fillRng= rng[0, 0, nRows-1, 0].EntireRow;
          log.LogDebug("filling down range: {rng}", fillRng.RelativeAddr());
          fillRng.FillDown();
        }
      }

      //import table
      for (int r= startRow; r < endRow; ++r) {
        var row= data.Rows[r-startRow];
        for (int c= startCol; c < endCol; ++c) {
          var ci= c-startCol;
          var val= (hasHeader ? row[hdr[ci]] : row[ci]) as IConvertible;
          var tc= val?.GetTypeCode();

          if (tc >= TypeCode.SByte && tc <= TypeCode.Decimal)
            cells.SetNumber(r, c, val.ToDouble(FRMP));
          else if (tc == TypeCode.Boolean)
            cells.SetLogical(r, c, (bool)val);
          else if (tc == TypeCode.DateTime)
            cells.SetNumber(r, c, rng.Worksheet.Workbook.DateTimeToNumber((DateTime)val));
          else if (null != val)
            cells.SetText(r, c, val.ToString());
          else
            cells.Clear(r, c);
        }
      }
    }


    public static DataTable GetDataTable(this IRange rng, bool hasHeader) {
      int nCols= rng.ColumnCount;
      int nRows= rng.RowCount;
      int startCol= rng.Column;
      int startRow= rng.Row;
      int endCol= startCol + nCols;
      int endRow= startRow + nRows;
      int hdrRow= 0;
      IValues cells= (IValues)rng.Worksheet;
      var table= new DataTable();

      if (hasHeader) {
        hdrRow= startRow++;
        --nRows;
      }

      var tabCols= table.Columns;
      for (int l = 0; l < nCols; ++l)
        tabCols.Add(new DataColumn(  hasHeader
                                   ? cells[hdrRow, startCol+l].Text
                                   : $"_c{l}"));

      //export table
      for (int r = startRow; r < endRow; ++r) {
        DataRow tabRow= table.NewRow();
        table.Rows.Add(tabRow);
        for (int c= startCol, ci= 0; c < endCol; ++c, ++ci) {
          var cell= cells[r, c];
          object val= null;
          if (null != cell) switch (cell.Type) {
            case CellType.Text: 
              val= cell.Text;
            break;

            case CellType.Number:
              val= cell.Number;
            break;

            case CellType.Logical:
              val= cell.Logical;
            break;

            default:
              val= ERRmap[(int)cell.Error];
            break;
          }
          tabRow[ci]= val;
        }
      }
      return table;
    }

    public static string AbsolutAddr(this IRange rng) {
      if (null == rng) return "NULL";
      return rng.WorkbookSet.GetAddress(rng.Row, rng.Column, rng.Row + rng.RowCount-1, rng.Column + rng.ColumnCount-1,
                                        0, 0, //bases
                                        true, true, true, true, //absolute
                                        forceRange: false,
                                        referenceStyle: ReferenceStyle.A1);
    }
    public static string RelativeAddr(this IRange rng) {
      if (null == rng) return "NULL";
      return rng.WorkbookSet.GetAddress(rng.Row, rng.Column, rng.Row + rng.RowCount-1, rng.Column + rng.ColumnCount-1,
                                        0, 0, //bases
                                        false, false, false, false, //relative
                                        forceRange: false,
                                        referenceStyle: ReferenceStyle.A1);
    }

    public static string SheetAddress(this IRange rng) {
      if (null == rng) return "NULL";
      return $"{rng.Worksheet.Name}@{rng.RelativeAddr()}";
    }
  }
}
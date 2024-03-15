using System;
using System.Linq;
using System.Collections.Generic;

using Microsoft.Extensions.Logging;
using SpreadsheetGear;


namespace Tlabs.CalcNgn.Sgear {

  internal static class TableHeaderRangeExtension {
    static readonly ILogger log= App.Logger<ITableData>();
    public static IRange FindTableHeader(this IRange searchArea, IEnumerable<string> header) {
      var hdrCnt= header.Count();
      if (0 == hdrCnt) throw new ArgumentException(nameof(header));
      var hdrCells= new List<IRange>();
      var cell= searchArea[0, 0];    //search start
      do {
        hdrCells.Clear();
        foreach(var hdr in header) {
          var fndCell= searchArea.Find(hdr, cell, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);
          if (null == fndCell) break; //header not found
          if (   hdrCells.Count > 0
              && (fndCell.Row != hdrCells[^1].Row || 1 != (fndCell.Column - hdrCells[^1].Column))) break;   //bad header position
          hdrCells.Add(cell= fndCell);
        }
      } while (hdrCells.Count > 0 && header.Count() != hdrCells.Count);
      if (header.Count() != hdrCells.Count) throw EX.New<CalcNgnModelException>("No table found in range: {rng}", searchArea.RelativeAddr());

      var hdrRng= searchArea.Worksheet.Cells[hdrCells[0].Row, hdrCells[0].Column, hdrCells[^1].Row, hdrCells[^1].Column];
      log.LogDebug("Table header range found: {addr}", hdrRng.RelativeAddr());
      return hdrRng;
    }

    public static IRange DataRange(this IRange hdrRng) {
      var startCell= hdrRng[1, 0];  //row below of header
      if (null == startCell.Value)
        startCell= hdrRng.EndDown;  //first non empty row below header
      var endCell= startCell.EndDown;

      var dataRng= hdrRng.Worksheet.Cells[startCell.Row, hdrRng.Column, endCell.Row, hdrRng.Column + hdrRng.ColumnCount-1];
      log.LogDebug("Data range identified: {addr}", dataRng.RelativeAddr());
      return dataRng;
    }
  }

}
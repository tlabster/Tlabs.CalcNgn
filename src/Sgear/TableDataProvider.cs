﻿using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.IO;

using SpreadsheetGear;

namespace Tlabs.CalcNgn.Sgear {

  /// <summary>I<see cref="ITableData"/> implementation for spreadsheet data.</summary>
  /// <remarks>Uitlity to extract tabular (either 2D-array or a <see cref="IRowData"/> from a spreadsheet file given as a <see cref="Stream"/>.
  /// </remarks>
  public class TableDataProvider : CalcNgnModelDef.AbstractModel, ITableData {
    /// <summary>Ctor from optional <paramref name="culture"/> and <paramref name="licKey"/>.</summary>
    public TableDataProvider(CultureInfo? culture= null, string? licKey= null) : base(culture, licKey) {  }

    ///<inheritdoc/>
    public object[,] ReadDataWithHeader(Stream dataStream, IEnumerable<string> header) {
      IWorkbook? wbk= null;
      try {
        wbk= wbSet.Workbooks.OpenFromStream(dataStream);
        if (wbk.Worksheets.Count != 1) throw EX.New<CalcNgnModelException>("Invalid worksheet count: {cnt}", wbk.Worksheets.Count);
        var wks= wbk.Worksheets[0];
        var dataRng= wks.UsedRange.FindTableHeader(header)
                                  .DataRange();
        var tableData= (object[,])dataRng.Value;
        if (tableData.GetLength(1) != header.Count()) throw EX.New<CalcNgnModelException>("Invalid table data shape: [{rows}, {cols}]", tableData.GetLength(0), tableData.GetLength(1));
        return tableData;
      }
      finally {
        wbk?.Close();
      }
    }

    ///<inheritdoc/>
    public IRowData ReadRowDataWithHeader(Stream dataStream, IEnumerable<string> header) => new RowData(ReadDataWithHeader(dataStream, header), header);

    class RowData : IRowData {
      readonly object[,] data;
      readonly IEnumerable<string> header;
      IReadOnlyDictionary<string, int>? hdrIdx;
      public RowData(object[,] data, IEnumerable<string> header) {
        this.data= data;
        this.header= header;
      }

      public IReadOnlyList<object> this[int index] => new Misc.Array2DRowSlice<object>(data, index);

      public int Count => data.GetLength(0);

      public IReadOnlyDictionary<string, int> HeaderIndex {
        get {
          var idx= 0;
          return hdrIdx??= header.ToDictionary(hdr => hdr, _=> idx++);  //defer index creation on demand
        }
      }
    }

  }
}
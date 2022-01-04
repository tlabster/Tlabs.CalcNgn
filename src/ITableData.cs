using System.Collections.Generic;
using System.IO;

namespace Tlabs.CalcNgn {

  /// <summary>Representation of table based data.</summary>
  public interface ITableData {
    /// <summary>Read table data with <paramref name="header"/> from <paramref name="dataStream"/>.</summary>
    object[,] ReadDataWithHeader(Stream dataStream, IEnumerable<string> header);
    /// <summary>Read table row data with <paramref name="header"/> from <paramref name="dataStream"/>.</summary>
    IRowData ReadRowDataWithHeader(Stream dataStream, IEnumerable<string> header);
  }

   ///<summary>Row data (list).</summary>
  public interface IRowData {
    ///<summary>Count of rows.</summary>
    int Count { get; }
    ///<summary>Header index.</summary>
    IReadOnlyDictionary<string, int> HeaderIndex { get; }
    ///<summary>Row with <paramref name="index"/>.</summary>
    IReadOnlyList<object> this[int index] { get; }
  }
}
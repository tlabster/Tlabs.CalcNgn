using System.Collections.Generic;
using System.IO;

namespace Tlabs.CalcNgn {

  /// <summary>Representation of table based data.</summary>
  public interface ITableData {
    /// <summary>Read data as 2D-array identified with list of column <paramref name="header"/>(s) from a spreadsheet (file) <paramref name="dataStream"/>.</summary>
    /// <remarks>The data area to be returned from the spreadsheet <paramref name="dataStream"/>
    /// is identified with adjacent columns heaving header cells that exactly match the <paramref name="header"/> list.
    /// <para>The retuned 2D data array start with the first non empty cell below the first column specified with the <paramref name="header"/> list.</para>
    /// <para>The data ends with the first empty cell below the first data value of the first column (specified with the <paramref name="header"/> list).</para>
    /// </remarks>
    object[,] ReadDataWithHeader(Stream dataStream, IEnumerable<string> header);

    /// <summary>Returns <see cref="IRowData"/> identified with list of column <paramref name="header"/>(s) from a spreadsheet (file) <paramref name="dataStream"/>.</summary>
    /// <remarks>The <see cref="IRowData"/> to be returned from the spreadsheet <paramref name="dataStream"/>
    /// is identified with adjacent columns heaving header cells that exactly match the <paramref name="header"/> list.
    /// <para>The retuned <see cref="IRowData"/> start with the first non empty cell below the first column specified with the <paramref name="header"/> list.</para>
    /// <para>The <see cref="IRowData"/> ends with the first empty cell below the first data value of the first column (specified with the <paramref name="header"/> list).</para>
    /// </remarks>
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
using System;
using System.Data;
using System.Collections;
using System.Collections.Generic;

namespace Tlabs.CalcNgn.Util {

  /// <summary>Extension class to convert a list of objects into <see cref="DataTable"/>.</summary>
  public static class ListAsDataTableExtension {
    /// <summary>Convert a List of dynamic JsonObject(s) into a <see cref="DataTable"/>.</summary>
    /// <remarks>
    /// The first List entry is used to define the schema (columns) of the table.
    /// Subsequent entries MUST NOT introduce new properties/columns...
    /// </remarks>
    /// <param name="lst"></param>
    /// <returns><see cref="DataTable"/></returns>
    public static DataTable AsDataTable(this ICollection lst) {
      var table= new DataTable();
      var tabCols= table.Columns;
      var tabRows= table.Rows;

      foreach (object jobj in lst) {
        DataRow tabRow= table.NewRow();

        if (jobj is IDictionary<string, object> dict)
          rowFromDictionary(tabRow, dict, tabCols);
        else
          rowFromList(tabRow, jobj, tabCols);

        tabRows.Add(tabRow);
        tabCols= null;
      }
      return table;
    }

    private static void rowFromDictionary(DataRow tabRow, IDictionary<string, object> dict, DataColumnCollection colDef) {
      if (null == dict) return;

      if (null != colDef) {
        /* Setup table schema (columns) from first list entry
         * and also add first data row:
         */
        var row= new object[dict.Count];
        int l= 0;
        foreach (var prop in dict) {
          colDef.Add(new DataColumn(prop.Key, null != prop.Value ? prop.Value.GetType() : typeof(object)));
          row[l++]= prop.Value;
        }
        tabRow.ItemArray= row;
      }
      else foreach (var prop in dict)
        tabRow[prop.Key]= prop.Value ?? DBNull.Value;
    }

    private static void rowFromList(DataRow tabRow, object obj, DataColumnCollection colDef) {
      var lst= obj as IList<object>;

      if (null == lst) return;

      if (null != colDef) {
        /* Setup table schema (columns) from first list entry
         * and also add first data row:
         */
        var row= new object[lst.Count];
        for (int l= 0; l < lst.Count; ++l) {
          var v= row[l]= lst[l];
          colDef.Add(new DataColumn(string.Format("C{0}", l), null != v ? v.GetType() : typeof(object)));
        }
        tabRow.ItemArray= row;
      }
      else for (int l= 0; l < lst.Count; ++l)
        tabRow[l]= lst[l];
    }

  }

}

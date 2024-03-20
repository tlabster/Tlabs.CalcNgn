using System;
using System.Data;
using System.Globalization;

using Xunit;

using Tlabs.CalcNgn.Sgear;
namespace Tlabs.CalcNgn.Tests {

  public class DataTableTest {

    [Fact]
    public void TypeConversionTest() {
      var tb= new DataTable();
      var tbCols= tb.Columns;
      tbCols.Add("null");
      tbCols.Add("string", typeof(string));
      tbCols.Add("object", typeof(object));
      tbCols.Add("double", typeof(double));

      var row= tb.NewRow();
      tb.Rows.Add(row);
      Assert.IsAssignableFrom<DBNull>(row["null"]);
      row["string"]= "xyz";
      Assert.Equal("xyz", row["string"]);
      row["string"]= ((IConvertible)1.2).ToString(CultureInfo.InvariantCulture);
      Assert.NotEqual("1,2", row["string"]);
      Assert.IsAssignableFrom<DBNull>(row["object"]);
      row["object"]= 123;
      Assert.Equal(123, row["object"]);
      row["double"]= "1.2";
      Assert.Equal(1.2, (row["double"]));    //uses current culture!!!
    }
  }
}
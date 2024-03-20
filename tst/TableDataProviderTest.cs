using System.Linq;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

using Xunit;

using Tlabs.CalcNgn.Sgear;
namespace Tlabs.CalcNgn.Tests {

  public class TableDataProviderTest {
    private Calculator cngn= new Calculator(new Sgear.CalcNgnModelParser(null, CalcNgnTest.CALCNGN_LIC));

    [Fact]
    public void BasicTest() {
      var calcMod= cngn.LoadModel(CalcNgnTest.createCalcSheet());
      calcMod.Data= new Dictionary<string, object> {
        ["data"]= CalcNgnTest.TABvals
      };

      var modelPath= Path.Combine(App.ContentRoot, "calcngnTableData01.xls");
      calcMod.SaveCopy(modelPath);
      MemoryStream strm= new MemoryStream(20 * 1024);
      calcMod.Definition.WriteStream(strm);
      strm.Position= 0;
      calcMod.Dispose();


      ITableData tabData= new TableDataProvider(App.DfltFormat, CalcNgnTest.CALCNGN_LIC);
      var rowData= tabData.ReadRowDataWithHeader(strm, new string[] {"prop01", "prop02", "frml", "prop03"});
      Assert.NotEqual(0, rowData.Count);
      for (var rno= 0; rno < rowData.Count; ++rno) {
        var row= rowData[rno];
        var tabRow= CalcNgnTest.TABvals.Rows[rno].ItemArray;
        for (var l = 0; l < tabRow.Length; ++l) {
          var colName= ((DataColumn)CalcNgnTest.TABvals.Columns[l]).ColumnName;
          var idx= rowData.HeaderIndex[colName];
          Assert.Equal(l, idx);
          if ("dummy" != tabRow[l].ToString()) Assert.True(tabRow[l].Equals(row[idx]), $"@R{rno}C{l} {tabRow[l]} != {row[idx]}");
        }
      }
    }
  }
}
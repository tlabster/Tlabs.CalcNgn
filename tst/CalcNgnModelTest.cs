﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using SpreadsheetGear;

using Tlabs.Misc;
using Tlabs.CalcNgn.Util;

namespace Tlabs.CalcNgn.Tests {
  using System.Linq;

  using Xunit;

  public class CalcNgnTest {
    // private SpreadsheetGear Licence:
    public static string CALCNGN_LIC= Environment.GetEnvironmentVariable("SPREADSHEETGEAR-LICENSE");
    private Calculator cngn= new Calculator(new Sgear.CalcNgnModelParser(null, CALCNGN_LIC));

    internal static DataTable TABvals= new List<object> {
        new Dictionary<string, object> {["prop01"]= 1.1, ["prop02"]= "Aaaa", ["frml"]= "dummy", ["prop03"]= 10.0},
        new Dictionary<string, object> {["prop01"]= 2.2, ["prop02"]= "Bbbb", ["frml"]= "dummy", ["prop03"]= 20.0},
        new Dictionary<string, object> {["prop01"]= 3.3, ["prop02"]= "Cccc", ["frml"]= "dummy", ["prop03"]= 30.0},
        new Dictionary<string, object> {["prop01"]= 4.4, ["prop02"]= "Dddd", ["frml"]= "dummy", ["prop03"]= 40.0}
      }.AsDataTable();


    [Fact]
    public void AsDataTableTest() {
      var rowList= new List<IList<object>>();
      for (var rno= 0; rno < TABvals.Rows.Count; ++rno) rowList.Add(TABvals.Rows[rno].ItemArray.ToList());
      var tab= rowList.AsDataTable();
      Assert.Equal(TABvals.Rows.Count, tab.Rows.Count);
      Assert.Equal(TABvals.Columns.Count, tab.Columns.Count);
      for (var rno = 0; rno < TABvals.Rows.Count; ++rno)
        for (var cno = 0; cno < TABvals.Columns.Count; ++cno)
          Assert.Equal(TABvals.Rows[rno][cno], tab.Rows[rno][cno]);
    }

    [Fact]
    public void RangeImpTest() {
      var calcMod= cngn.LoadModel(createCalcSheet());

      var namedVals= new Dictionary<string, object> {
        ["A5"]= 555.0,
        ["A6"]= 666.0
      };
      calcMod.Definition.ImportNamedValues(namedVals);

      MemoryStream strm= new MemoryStream(20 * 1024);
      calcMod.Definition.WriteStream(strm);
      strm.Position= 0;

      var wbk= Factory.GetWorkbookSet().Workbooks.OpenFromStream(strm);
      var wks= wbk.Worksheets[0];
      Assert.NotEqual(0, wbk.Names.Count);
      var name= wbk.Names["WEB_A5"];
      var frml= name.RefersTo;
      var val0= ((IConvertible)namedVals["A5"]).ToDouble(System.Globalization.CultureInfo.InvariantCulture);
      var val1= wks.EvaluateRange(name.Name).Value;
      Assert.Equal(val0, val1);
      Assert.Equal(namedVals["A6"], wks.EvaluateRange("WEB_A6").Value);
    }


    [Fact]
    public void TableImpTest() {
      var calcMod= cngn.LoadModel(createCalcSheet(withColumnSpec: false));
      // calcMod.SaveCopy(Path.Combine(App.ContentRoot, "calcngn0.xls"));

      calcMod.Data= new Dictionary<string, object> {
        ["data"]= TABvals
      };
      // calcMod.SaveCopy("D:\\calcngn.xls");

      MemoryStream strm= new MemoryStream(20 * 1024);
      calcMod.Definition.WriteStream(strm);
      strm.Position= 0;

      var wbk= Factory.GetWorkbookSet().Workbooks.OpenFromStream(strm);
      var wks= wbk.Worksheets[0];

      Assert.Equal(1.1, wks.Cells["D10"].Value);
      Assert.Equal(2.2, wks.Cells["D11"].Value);
      Assert.Equal(3.3, wks.Cells["D12"].Value);
      Assert.Equal(4.4, wks.Cells["F1"].Value);
      Assert.Equal(11.1, wks.Cells["F10"].Value);
      Assert.Equal(366.3, wks.Cells["B9"].Value);
    }


    [Fact]
    public void DataExportTest() {
      var calcMod= cngn.LoadModel(createCalcSheet());
      calcMod.SaveCopy(Path.Combine(App.ContentRoot, "calcngn0.xls"));
      calcMod.Data= new Dictionary<string, object> {
        ["data"]= TABvals
      };
      calcMod.SaveCopy(Path.Combine(App.ContentRoot, "calcngn.xls"));

      var exportData= calcMod.Data;
      object vo;
      string prop;
      Assert.True(exportData.TryResolveValue("data.export.A5", out vo, out prop));
      Assert.Equal("A5", prop);
      Assert.Equal("A5!!!", vo.ToString());
      Assert.True(exportData.TryResolveValue("data.export.A6", out vo, out prop));
      Assert.Equal("A6", prop);
      Assert.Equal("A6!!!", vo.ToString());

      Assert.True(exportData.TryResolveValue("data.export.table", out vo, out prop));
      Assert.Equal("table", prop);
      Assert.IsType<DataTable>(vo);
      var tab= (DataTable)vo;
      Assert.Equal(4, tab.Rows.Count);
      var row= tab.Rows[3];
      Assert.Equal(double.Parse(row[2].ToString()), double.Parse(row[0].ToString()) + double.Parse(row[3].ToString()));
    }

    internal static Stream createCalcSheet(bool withColumnSpec= true) {
      MemoryStream strm= new MemoryStream(20 * 1024);
      IWorkbook wbk= Factory.GetWorkbook();
      IWorksheet wks= wbk.Worksheets[0];
      var wksName= wks.Name;

      //wbk.WorkbookSet.Calculate
      wbk.Names.Add("WEB_A5", /*wksName + "!" +*/ "=$A$5");
      var name= wbk.Names["WEB_A5"];
      Assert.NotNull(name);
      Assert.NotNull(wks.EvaluateRange(name.Name));
      wbk.Names.Add("WEB_A6", "=" + wksName + "!$A$6");


      wks.Cells["D5"].Formula= "=WEB_A5";
      wks.Cells["D6"].Formula= "=WEB_A6";

      wks.Cells["A5"].Formula= "A5!!!";
      wks.Cells["A5"].AddComment("@Data_Export(CELL, data.export.A5)");
      wks.Cells["A6"].Formula= "A6!!!";
      wks.Cells["A6"].AddComment("@Data_Export(CELL, data.export.A6)");

      wks.Cells["D10"].AddComment($"@Data_Import(BEGIN, data{(withColumnSpec ? ", 4" : "")}) @Data_Export(TABLE, data.export.table)");
      wbk.Names.Add("EXPORT_RNG", "=" + wksName + "!" + "$D$10:$G$11");
      Assert.NotNull(wks.Cells["D10"].Comment);
      wks.Cells["D9"].Formula= "prop01";
      wks.Cells["E9"].Formula= "prop02";
      wks.Cells["F9"].Formula= "frml";
      wks.Cells["G9"].Formula= "prop03";

      wks.Cells["B10"].Formula= "=D10 * F10";
      wks.Cells["F10"].Formula= "=D10 + G10";
      wks.Cells["B9"].Formula= "=SUM(B10:B11)";
      wks.Cells["F1"].Formula= "=D11";
      // wks.Cells["H15"].Formula= "X";    //extend used range

      wks.SaveToStream(strm, FileFormat.Excel8);
      strm.Position= 0;
      return strm;
    }

    [Fact]
    public void LoadModelTest() {
      var calcMod= cngn.LoadModel(createCalcSheet());
      var modelPath= Path.Combine(App.ContentRoot, "calcngn0.xls");
      calcMod.SaveCopy(modelPath);

      using (var strm= new FileStream(modelPath, FileMode.Open)) {
        var memstrm= new MemoryStream();
        strm.CopyTo(memstrm);
        calcMod= cngn.LoadModel(strm);

        Assert.True(calcMod.Definition.Imports.Count > 0);
        Assert.NotEmpty(calcMod.Definition.Imports);
        // Assert.True(calcMod.Definition.Exports.Count > 0);
        Assert.NotEmpty(calcMod.Definition.Exports);

      }
    }

    [Fact]
    public void ExportNamesTest() {
      var calcMod= cngn.LoadModel(createCalcSheet());
      // calcMod.SaveCopy(Path.Combine(App.ContentRoot, "calcngn0.xls"));
      calcMod.Data= new Dictionary<string, object> {
        ["data"]= TABvals
      };
      // calcMod.SaveCopy(Path.Combine(App.ContentRoot, "calcngn.xls"));

      var exportData= calcMod.Data;
      string prop;
      object o;
      IDictionary<string, object> dict;

      Assert.True(exportData.TryResolveValue("data", out o, out prop));
      dict= o as IDictionary<string, object>;
      Assert.NotNull(dict);
      Assert.True(dict.ContainsKey("export"));

      Assert.True(dict.TryResolveValue("export", out o, out prop));
      dict= o as IDictionary<string, object>;
      Assert.True(dict.ContainsKey("table"));

      Assert.True(exportData.TryResolveValue("data.export", out o, out prop));
      Assert.True(exportData.TryResolveValue("data.export.table", out o, out prop));

    }
  }
  public static class CalcNgnModelExt {
    public static void SaveCopy(this Calculator.Model cmod, string path) {
      using (var strm= File.Create(path)) {
        cmod.Definition.WriteStream(strm);
      }
    }
  }
}
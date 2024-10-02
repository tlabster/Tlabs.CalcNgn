using System;
using System.IO;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using SpreadsheetGear;

using Tlabs.IO;

namespace Tlabs.CalcNgn.Tests {
  using System.Linq;

  using Xunit;

  public class SaveModelTest {
    // private SpreadsheetGear Licence:
    public static string CALCNGN_LIC= Environment.GetEnvironmentVariable("SPREADSHEETGEAR-LICENSE");
    private Calculator cngn= new Calculator(new Sgear.CalcNgnModelParser(null, CALCNGN_LIC));


    [Fact(Skip= "Base64Stream.CreateForEncoding() not ready")]
    public void LoadSaveFromBase64Test() {
      MemoryStream srcStrm= new MemoryStream(50 * 1024);
      createCalcSheet(srcStrm, true);
      srcStrm.Flush();
      // var calcSheetBin= srcStrm.ToArray();
      srcStrm.SetLength(0);
      srcStrm.Position= 0;

      var b64Strm= Base64Stream.CreateForEncoding(srcStrm);
      createCalcSheet(b64Strm, true);
      b64Strm.Flush();
      srcStrm.Position= 0;
      b64Strm= Base64Stream.CreateForDecoding(srcStrm);
      var calcMod2= cngn.LoadModel(b64Strm);
      return;
    }


    internal static void createCalcSheet(Stream strm, bool withColumnSpec= true) {
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
    }

  }
}
using System;
using Xunit;
using SpreadsheetGear;

using Tlabs.CalcNgn.Sgear;
namespace Tlabs.CalcNgn.Tests {

  public class RangeTest {

    [Fact]
    public void RowInsertion() {
      // Create a workbook.
      IWorkbook wbk= createWbk();
      IWorksheet wks= wbk.Worksheets[0];
      IRange cells= wks.Cells;
      
      var rowCnt= cells.RowCount;
      rowCnt= cells.RowCount;

      // Set some text in A1:B2.
      cells["A1"].Formula= "A1";
      cells["A2"].Formula= "A2";
      cells["B1"].Formula= "B1";
      cells["B2"].Formula= "B2";
      cells["C1"].Formula= "C1";
      wbk.SaveAs("d:\\tst0.xls", FileFormat.Excel8);
      rowCnt= wks.UsedRange.RowCount;

      // Insert two rows before row 2.
      // cells["A1:B2"][1, 0, 2, 0].EntireRow.Insert();
      cells[1, 0, 2, 0].EntireRow.Insert();
      cells["A1:B2"][1, 0, 2, 0].Formula= "new row";
      cells[0, 0, 2, 1].FillDown();
      wbk.SaveAs("d:\\tst.xls", FileFormat.Excel8);

      // Verify new locations - A1 has not moved.
      Assert.Equal("A1", cells["A1"].Formula);
      // A2 has been moved down two rows.
      Assert.Equal("A2", cells["A4"].Formula);
      // A1 has been copied down.
      Assert.Equal("A1", cells["A2"].Formula);
      Assert.Equal("A1", cells["A3"].Formula);
    }

    [Fact]
    public void EndDownTest() {
      // Create a workbook.
      IWorkbook wbk= createWbk();
      IWorksheet wks= wbk.Worksheets[0];
      IRange cells= wks.Cells;

      // Set header in C4:E4.
      cells["C4"].Formula= "H_C4";
      cells["D4"].Formula= "H_D4";
      cells["E4"].Formula= "H_E4";
      // Set some text in C5:E7.
      cells["C5"].Formula= "C5";
      cells["D5"].Formula= "D5";
      cells["E5"].Formula= "E5";
      cells["C6"].Formula= "C6";
      cells["D6"].Formula= "D6";
      cells["E6"].Formula= "E6";
      cells["C7"].Formula= "C7";
      cells["D7"].Formula= "D7";
      cells["E7"].Formula= "E7";

      cells["C10"].Formula= "C10";
      cells["C11"].Formula= "C11";

      wbk.SaveAs("d:\\tst0.xls", FileFormat.Excel8);


      var rng= cells["C5:E7"];
      var drng= cells["C4:E4"].DataRange();
      Assert.True(rng.RowCount == drng.RowCount && rng.ColumnCount == drng.ColumnCount && rng.Row == drng.Row);

      rng= cells["C10:C11"];
      drng= cells["C7:E7"].DataRange();
      Assert.True(rng.RowCount == drng.RowCount && 3 == drng.ColumnCount && rng.Row == drng.Row);
    }

    IWorkbook createWbk() {
      // private SpreadsheetGear Licence:
      const string CALCNGN_LIC= "SpreadsheetGear.License, Type=Standard, Hash=aiwZIy8On2qfzyoNK64Eqmk, Product=NST, NewVersionsUntil=2019-06-04, Company=Tomorrow Labs GmbH, Email=p.oltmanns@tomorrowlabs.io, Signature=I7b/hDUp/VgSyKK0qo2P+FIfyDbFx/qyn0/D1VifbJkA#xatGw3JAizFv76MwMv96/QiZkZodf6TYeo2056WK8OkA#J";
      try {
        Factory.SetSignedLicense(CALCNGN_LIC);
      }
      catch (InvalidOperationException) { }

      // Create a workbook.
      return Factory.GetWorkbook();
    }
  }
}

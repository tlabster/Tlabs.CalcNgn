using System;
using Xunit;
using SpreadsheetGear;

namespace Tlabs.CalcNgn.Tests {

  public class RangeTest {

    [Fact]
    public void RowInsertion() {
      // Create a workbook.
      try {
        Factory.SetSignedLicense("SpreadsheetGear.License, Type=Standard, Hash=aiwZIy8On2qfzyoNK64Eqmk, Product=NST, NewVersionsUntil=2019-06-04, Company=Tomorrow Labs GmbH, Email=p.oltmanns@tomorrowlabs.io, Signature=I7b/hDUp/VgSyKK0qo2P+FIfyDbFx/qyn0/D1VifbJkA#xatGw3JAizFv76MwMv96/QiZkZodf6TYeo2056WK8OkA#J");
      }
      catch (InvalidOperationException) { }

      IWorkbook wbk= Factory.GetWorkbook();
      IWorksheet wks= wbk.Worksheets[0];
      IRange cells= wks.Cells;
      
      var rowCnt= cells.RowCount;
      rowCnt= cells.RowCount;

      // Set some text in A1:B2.
      cells["A1"].Formula = "A1";
      cells["A2"].Formula = "A2";
      cells["B1"].Formula = "B1";
      cells["B2"].Formula = "B2";
      cells["C1"].Formula = "C1";
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
  }
}

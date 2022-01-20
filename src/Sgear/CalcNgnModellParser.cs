using System;
using System.Collections;
using System.Collections.Generic;

using SpreadsheetGear;

using Tlabs.CalcNgn.Intern;

namespace Tlabs.CalcNgn.Sgear {
  using CommentParser = Tlabs.CalcNgn.Parser.CmdParser;
  using DataTable = System.Data.DataTable;
  using System.Globalization;
  using Tlabs.CalcNgn.Util;

  /// <summary>Impl. of <see cref="ICalcNgnModelParser"/> using comment directives specified with <see cref="TemplateDirectives"/>.</summary>
  public sealed class CalcNgnModelParser : CalcNgnModelDef.AbstractModelParser {

    internal enum TemplateDirectives {
      /// <summary>@DATA_IMPORT</summary>
      DATA_IMPORT,
#if Shape_support
      /// <summary>@DynamicAutoShape</summary>
      DynamicAutoShape,
#endif
      /// <summary>@DATA_EXPORT</summary>
      DATA_EXPORT
    }
    /// <summary>DATA_IMPORT parameters</summary>
    internal enum DataImportParamTokens : int {
      /// <summary>Begin of fixed size import range</summary>
      BEGIN = 550,
      /// <summary>End of fixed size import range</summary>
      END,
      /// <summary>Single import value</summary>
      CELL

    }
#if Shape_support
    internal enum ShapeParamTokens : int {
      TEXT = 560,
      BackColor
    }
#endif
    /// <summary>DATA_IMPORT parameters</summary>
    internal enum DataExportParamTokens : int {
      /// <summary>Export single cell value</summary>
      CELL = 570,
      /// <summary>Export range of values</summary>
      RANGE,
      /// <summary>Export range of values with header row</summary>
      TABLE
    };

    internal static readonly IReadOnlyDictionary<string, Type> EXPORT_TYPE= new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase) {
      ["DATETIME"]= typeof(DateTime?)
    };

    private static readonly string IMPORT_CMD= Enum.GetName(typeof(TemplateDirectives), TemplateDirectives.DATA_IMPORT);
    private static readonly string IMPORT_PATTERN= Parser.CmdTokenizer.GetCommand(TemplateDirectives.DATA_IMPORT);
#if Shape_support
    private static readonly string SHAPE_CMD= Enum.GetName(typeof(TemplateDirectives), TemplateDirectives.DynamicAutoShape);
    private static readonly string SHAPE_PATTERN= Parser.CmdTokenizer.GetCommand(TemplateDirectives.DynamicAutoShape);
#endif
    private static readonly string EXPORT_CMD= Enum.GetName(typeof(TemplateDirectives), TemplateDirectives.DATA_EXPORT);
    private static readonly string EXPORT_PATTERN= Parser.CmdTokenizer.GetCommand(TemplateDirectives.DATA_EXPORT);

    /// <summary>Ctor from <paramref name="culture"/>.</summary>
    public CalcNgnModelParser(CultureInfo culture= null, string licKey= null) : base(culture, licKey) { }

    /// <inheritdoc/>
    protected override void ParseWorkbook(IWorkbook wbk, out IDictionary<string, IModelImport> imp, out IDictionary<string, IModelExport> exp) {
      var parsedWbk= new ParsedWbk(wbk);
      imp= parsedWbk.impDefs;
      exp= parsedWbk.expDefs;
    }

    private class ParsedWbk {
      public IDictionary<string, IModelImport> impDefs= new Dictionary<string, IModelImport>();
      public IDictionary<string, IModelExport> expDefs= new Dictionary<string, IModelExport>();
      private IWorksheet currSheet;

      public ParsedWbk(IWorkbook wbk) {
        CommentParser commentParser= new CommentParser(typeof(TemplateDirectives));
        commentParser.registerEvent(new Parser.CmdTokenizer(typeof(TemplateDirectives), IMPORT_PATTERN), ParseImportCmd, true);
#if Shape_support
        commentParser.registerEvent(new Parser.CmdTokenizer(typeof(TemplateDirectives), SHAPE_PATTERN), OnShapeCmd, true);
#endif
        commentParser.registerEvent(new Parser.CmdTokenizer(typeof(TemplateDirectives), EXPORT_PATTERN), ParseExportCmd, true);

        parse(wbk, commentParser);
      }

      private void parse(IWorkbook wbook, CommentParser commentParser) {
        foreach (object o1 in wbook.Worksheets) {
          currSheet= (IWorksheet)o1;
          var cells= currSheet.UsedRange;
          for (int rowNo = 0; rowNo < cells.RowCount; ++rowNo) {
            for (int colNo = 0; colNo < cells.ColumnCount; ++colNo) {
              var cell= cells[rowNo, colNo];
              var com= cell.Comment;
              if (null != com)
                commentParser.parse(com.ToString(), cell.Row, cell.Column);
            }
          }
        }
      }

      private void ParseImportCmd(object src, String cmd, IDictionary<String, IList<string>> parserResult, int row, int column) {
        var cell= currSheet.Cells[row, column];
        var impParam= parserResult[cmd];
        var minArgCnt= 2;
        var collCnt= 0;
        if (impParam.Count < minArgCnt) throw new CalcNgnException($"Wrong number of arguments for {cmd} in {cell.SheetAddress()}");
        DataImportParamTokens impType= (DataImportParamTokens)Enum.Parse(typeof(DataImportParamTokens), impParam[0], true);
        if (impType == DataImportParamTokens.BEGIN && impParam.Count != minArgCnt + 1) throw new CalcNgnException($"Wrong number of arguments for {cmd} in {cell.SheetAddress()}");
        string impID= impParam[1].Trim();

        if (impType == DataImportParamTokens.END) {
          string matchingCmd= Parser.CmdTokenizer.GetCommand(TemplateDirectives.DATA_IMPORT, DataImportParamTokens.BEGIN);
          IModelImport modImp;
          matchingCmd= matchingCmd.Substring(0, matchingCmd.Length-1) + ", \"" + impID + "\")";
          if (!this.impDefs.TryGetValue(impID, out modImp)) throw new CalcNgnException($"Data import end in {cell.SheetAddress()} not matching {matchingCmd}");
          var idef= (ImportDef)modImp;
          if (idef.startCell.Row >= cell.Row) throw new CalcNgnException($"Data import end in {cell.SheetAddress()} has to be below of {matchingCmd}");
          if (idef.startCell.Column != cell.Column) throw new CalcNgnException($"Data import end in {cell.SheetAddress()} has to be in same column as {matchingCmd}");
          idef.rowCnt= cell.Row - idef.startCell.Row;
          return;
        }

        if (impType == DataImportParamTokens.BEGIN) {
          foreach (var modImp in this.impDefs.Values) {
            var impDef= (ImportDef)modImp;
            if (   impDef.Type == impType
                && cell.Worksheet.Index == impDef.startCell.Worksheet.Index) throw new CalcNgnException($"Duplicate {cmd} BEGIN in {cell.SheetAddress()}, only one import BEGIN per sheet supported.");
          }
          if (!Int32.TryParse(impParam[2].Trim(), out collCnt)) throw new CalcNgnException($"Invalid column count parameter for {cmd} in {cell.SheetAddress()}.");
        }

        this.impDefs[impID]= new ImportDef(impID, cell, impType, collCnt);
      }

      public void ParseExportCmd(object src, String cmd, IDictionary<String, IList<string>> parserResult, int row, int col) {
        var cell= currSheet.Cells[row, col];
        var param= parserResult[cmd];
        if (param.Count < 2 || param.Count > 3) throw new CalcNgnException(string.Format("Wrong arguments number for {0} in cell[{1:D}, {2:D}]", cmd, cell.Row, cell.Column));
        DataExportParamTokens expType= (DataExportParamTokens)Enum.Parse(typeof(DataExportParamTokens), param[0], true);

        if (String.IsNullOrWhiteSpace(param[1])) throw new CalcNgnException(string.Format("Bad data key {0} in cell[{1:D}, {2:D}]", cmd, cell.Row, cell.Column));
        var dataKey= param[1].Trim();
        
        Type targetType= null;
        if (param.Count > 2)
          EXPORT_TYPE.TryGetValue(param[2].Trim(), out targetType);
        
        IName rngName= lookupRngName(cell);

        expDefs.Add(dataKey, 
                    new ExportDef {
                      namedRng= rngName,
                      hasHeader= (DataExportParamTokens.TABLE == expType),
                      type= targetType
        });
      }

      private static IName findInNames(IRange cell, INames names) {
        for (int l = 0; l < names.Count; ++l) {
          var nam= names[l];
          var rng= nam.RefersToRange;
          if (null != rng && null != cell.Intersect(rng))
            return nam;
        }
        return null;
      }

      private static IName createRngName(IRange cell) {
        var cellName= "__.cell." + cell.RelativeAddr();
        return cell.Worksheet.Names.Add(cellName, "=" + cell.AbsolutAddr(), ReferenceStyle.A1);
      }

      private static IName lookupRngName(IRange cell) {
        return    findInNames(cell, cell.Worksheet.Names)
               ?? findInNames(cell, cell.Worksheet.Workbook.Names)
               ?? createRngName(cell);
      }

#if Shape_support
      private void OnShapeCmd(object src, String cmd, IDictionary<String, IList<string>> parserResult, int row, int column) {
        var cell= currRange.Worksheet.Cells[row, column];
        var shapeParam= parserResult[cmd];
        if (shapeParam.Count < 2) throw new CalcNgnException(string.Format("Wrong arguments number for {0} in cell[{1:D}, {2:D}]", cmd, cell.Row, cell.Column));
        string shapeName= shapeParam[0];
        var shape= cell.Worksheet.Shapes[shapeName];
        if (null == shape) throw new CalcNgnException(string.Format("{0} with name '{3}' not found (cell[{1:D}, {2:D}])", cmd, cell.Row, cell.Column, shapeName));
        var shapeTokens= new ShapeParamTokens[shapeParam.Count-1];
        for (int l= 1; l < shapeParam.Count; ++l)
          shapeTokens[l-1]= (ShapeParamTokens)Enum.Parse(typeof(ShapeParamTokens), shapeParam[l], true);
        dynShapes.Add(new DynamicShape(cell, shape, shapeTokens));
      }
#endif


    }//class ParsedWbk

    private class ImportDef : IModelImport {
      public string id;
      public IRange startCell;
      public DataImportParamTokens Type;
      public int colCnt;
      public int rowCnt;

      public ImportDef(string id, IRange cell, DataImportParamTokens type, int collCnt) { this.id= id; this.startCell= cell; this.Type= type; this.colCnt= collCnt; }

      public void Input(object data) {
        if (this.Type == DataImportParamTokens.CELL) {
          startCell.Value= data;
          return;
        }

        if (rowCnt < 0)
          startCell[0, 0, -rowCnt-1, Math.Max(colCnt-1, 0)].ClearContents();   //clear previous contents

        var dataTab= data as DataTable;
        if (null == dataTab) {
          var lst= data as ICollection;
          if (null == lst) {
            if (0 == rowCnt || 0 == colCnt) {
              startCell.Value= ValueError.Value;
              return;
            }
            lst= new object[0];
          }
          dataTab= lst.AsDataTable();
        }
        
        int rngRows= rowCnt;
        // int rngCols= 0 == colCnt ? (colCnt= dataTab.Columns.Count) : colCnt;
        bool insRows= false;

        /* Note: If rngRows == 0 this is resulting from a @DATA_IMPORT(BEGIN,"<id>") with no @DATA_IMPORT(END) and denotes
         *       a dynamically sized data input range (that initialy must consist of exactly two rows [SpreadsheetGear requirement] !!!).
         *       In this case this 'template' range will be filled with rng.CopyFromDataTable(dataTab, CopyTableFlags.InsertCells),
         *       while any references (from formulas) are getting adjusted accordingly.
         *       
         *       If rngRows < 0 that is the result from a previous import into a dynamic template range. In this case the range
         *       must be shrunk to exaactly two rows again for being prepared for rng.CopyFromDataTable(dataTab, CopyTableFlags.InsertCells)...
         *       
         * Caution: If the DataTable being imported does not contain at least two rows, the import target range gets shrunk to
         *          less than two rows, making it impossible to further import into this range!!!
         *       
         */
        if (rngRows <= 0) {
          int delRows= -rngRows;
          rngRows= 2;
          insRows= true;
          if (delRows > 2) {
            startCell[2, 0, delRows-1, 0].EntireRow.Delete();
          }
        }

        IRange rng= startCell[0, 0, rngRows-1, Math.Max(colCnt-1, 0)];
        rng.ClearContents();

        if (rowCnt > 0 && dataTab.Rows.Count > rng.RowCount) {
          /* If DataTable row count > fix target range, shrink a copy of the table to fit
           * the size of the target range:
           */
          dataTab= dataTab.Copy();
          while (dataTab.Rows.Count > rng.RowCount)
            dataTab.Rows.RemoveAt(dataTab.Rows.Count-1);
        }

        rng.CopyFromDataTable(dataTab, insRows, hasHeader: true);
        if (rowCnt < 1)
          rowCnt= -dataTab.Rows.Count;      //mark ImportDef to clear shrink dynamic target range
      }
    }

    private class ExportDef : IModelExport {
      public IName namedRng;
      public bool hasHeader;
      public Type type;
      public object Value {
        get {
          IRange rng= namedRng.RefersToRange;
          object v= null;
          double? num;
          if (null == rng) return v;
          
          if (1 == rng.CellCount || rng.MergeCells) {
            v= rng.Value;
            if (   type == typeof(DateTime?)
                && null != (num= v as Double?))
              v= rng.Worksheet.Workbook.NumberToDateTime(num.Value);
            return v;
          }

          if (hasHeader)
            rng= rng[-1, 0, rng.RowCount-1, rng.ColumnCount-1]; //extend range to header
          return v= rng.GetDataTable(hasHeader);
        }
      }
    }//class ExportDef

#if Shape_support
    /* This makes a shape dynamically changing based on the properties of a 'anchor' cell/range.
     * NOTE: SpreadsheetGear does not display the TextFrame of a shape (excpet for a textbox shape).
     * For that reason we adjust the shape by adding a transparent textbox with same size on top...
     */
    internal class DynamicShape {
      private IRange anchorCell;
      private IShape shape;
      private ICharacters txtCaption;
      private ShapeParamTokens[] shapeToks;

      public DynamicShape(IRange anchorCell, IShape shape, ShapeParamTokens[] toks) {
        if (null == (this.anchorCell= anchorCell)) throw new ArgumentNullException("anchorCell");
        if (null == (this.shape= shape)) throw new ArgumentNullException("shape");
        this.shapeToks= toks;


        var wks= anchorCell.Worksheet;
        var tboxName= shape.Name+"_text";
        var tbox= wks.Shapes[tboxName];
        if (null == tbox) {
          /* Create txtCaption shape:
           */
          shape.Placement= Placement.Move;
          tbox= wks.Shapes.AddTextBox(shape.Left, shape.Top, shape.Width, shape.Height);
          tbox.Name= tboxName;
          tbox.Placement= shape.Placement;
          tbox.Fill.Visible= false;
          tbox.Line.Visible= false;
          tbox.BringToFront();
          txtCaption= CopyTextFrame(shape.TextFrame, tbox.TextFrame).Characters;
        }
        txtCaption.Text= shape.TextFrame.Characters.Text;
        shape.TextFrame.Characters.Delete();
      }

      private ITextFrame CopyTextFrame(ITextFrame src, ITextFrame dst) {
        dst.VerticalAlignment= src.VerticalAlignment;
        dst.HorizontalAlignment= src.HorizontalAlignment;
        dst.Orientation= src.Orientation;

        dst.MarginLeft= src.MarginLeft;
        dst.MarginRight= src.MarginRight;
        dst.MarginTop= src.MarginTop;
        dst.MarginBottom= src.MarginBottom;

        CopyFont(src.Characters.Font, dst.Characters.Font);
        return dst;
      }

      private IFont CopyFont(IFont src, IFont dst) {
        dst.Name= src.Name;
        dst.TintAndShade= src.TintAndShade;
        dst.Color= src.Color;
        dst.Size= src.Size;
        dst.Bold= src.Bold;
        dst.Italic= src.Italic;
        dst.Underline= src.Underline;
        return dst;
      }

      public void UpdateRendition() {
        /* for each shape's property group specified in the @DynamicAutoShape(...) directive, update
         * the shape's properties:
         */
        foreach (var tok in shapeToks) switch (tok) {
            case ShapeParamTokens.TEXT:
              txtCaption.Text= anchorCell.Text;
              break;

            case ShapeParamTokens.BackColor:
              shape.Fill.ForeColor.RGB= anchorCell.Interior.Color;
              break;
          }
      }
    }
#endif

  }

}

using System;
using System.Collections.Generic;
using System.IO;
using System.Globalization;

using Microsoft.Extensions.Logging;
using SpreadsheetGear;
using SpreadsheetGear.Shapes;

using Tlabs.Misc;

namespace Tlabs.CalcNgn {
  using DataDictionary= IDictionary<string, object>;

  /// <summary>Calculation Engine</summary>
  public sealed class Calculator {
    private static readonly ILogger log= App.Logger<Calculator>();
    private Intern.ICalcNgnModelParser modelParser;


    /// <summary>Ctor from <paramref name="modelParser"/>.</summary>
    public Calculator(Intern.ICalcNgnModelParser modelParser) { this.modelParser= modelParser; }

    /// <summary>Load calculation model from <paramref name="modelStream"/>.</summary>
    public Model LoadModel(Stream modelStream) {
      return new Model(modelParser.ParseModelStream(modelStream));
    }

    /// <summary>Calculation model (spreadsheet abstraction used as computation/business logic).</summary>
    public class Model {
      private static readonly ILogger log= App.Logger<Calculator>();
      private Intern.ICalcNgnModelDef modelDef;

      /// <summary>Ctor from <paramref name="modelDef"/>.</summary>
      public Model(Intern.ICalcNgnModelDef modelDef) {
        this.modelDef= modelDef;
      }

      /// <summary>Model definition info.</summary>
      public Intern.ICalcNgnModelDef Definition => modelDef;

      /// <summary>Model data.</summary>
      public DataDictionary Data {
        get => ExportInto(new Dictionary<string, object>());
        set {
          string k;
          object val;
          /* ** For each import cell:
             Try to resolve the import cell key with the given data (in value).
             If the resolved value is null or key could not be resolved, val == null. Null as Input value causes the cell to be set as 'empty'...
           */
          foreach (var imp in modelDef.Imports) {
            value.TryResolveValue(imp.Key, out val, out k);
            imp.Value.Input(val);
          }
        }
      }

      /// <summary>Exports model (data) provided into <paramref name="data"/>.</summary>
      public DataDictionary ExportInto(DataDictionary data) {
        string k;
        object xVal= null;
        foreach (var exp in modelDef.Exports) try {
          xVal= exp.Value.Value;  //poke value from cell
          /*** Place cell value into into target data (dictionary):
           *   NOTE:
           *   The actual DataDictionary implementation could possibly throw from setter!
           *   (e.g. if the setter tries to cast the object value to a target type...)
           */
          data.SetResolvedValue(exp.Key, xVal, out k);
        }
        catch (Exception e) {
          log.LogWarning(0, e, "Problem exporting [{key}]= {val} ({msg})", exp.Key, xVal, e.Message);
        }
        return data;
      }

      /// <summary>Compute <paramref name="data"/> with calculation model.</summary>
      public void Compute(DataDictionary data) {
        Data= data;
        ExportInto(data);
      }

    }//clas Model
  } //class CalcNgn

}
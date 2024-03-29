﻿using System;
using System.Collections.Generic;
using System.IO;

using Microsoft.Extensions.Logging;

using Tlabs.Misc;

namespace Tlabs.CalcNgn {
  using DataDictionary= IDictionary<string, object?>;

  /// <summary>Calculation Engine</summary>
  public class Calculator {
    readonly Intern.ICalcNgnModelParser modelParser;

    /// <summary>Ctor from <paramref name="modelParser"/>.</summary>
    public Calculator(Intern.ICalcNgnModelParser modelParser) { this.modelParser= modelParser; }

    /// <summary>Load calculation model from <paramref name="modelStream"/>.</summary>
    public Model LoadModel(Stream modelStream) {
      return new Model(modelParser.ParseModelStream(modelStream));
    }

    /// <summary>Calculation model (spreadsheet abstraction used as computation/business logic).</summary>
    public sealed class Model : IDisposable {
      private static readonly ILogger log= App.Logger<Calculator>();
      private Intern.ICalcNgnModelDef modelDef;

      /// <summary>Ctor from <paramref name="modelDef"/>.</summary>
      internal Model(Intern.ICalcNgnModelDef modelDef) {
        this.modelDef= modelDef;
      }

      /// <summary>Model definition info.</summary>
      public Intern.ICalcNgnModelDef Definition => modelDef;

      /// <summary>Model data.</summary>
#pragma warning disable CA2227
      public DataDictionary Data {
        get => ExportInto(new Dictionary<string, object?>());
        set {
          /* ** For each import cell:
             Try to resolve the import cell key with the given data (in value).
             If the resolved value is null or key could not be resolved, val == null. Null as Input value causes the cell to be set as 'empty'...
           */
          foreach (var imp in modelDef.Imports) {
            if (value.TryResolveValue(imp.Key, out var val, out var _))
              imp.Value.Input(val);
          }
        }
      }
#pragma warning restore CA2227

      /// <summary>Exports model (data) into provided <paramref name="data"/>.</summary>
      public DataDictionary ExportInto(DataDictionary data) {
        object? xVal= null;
        foreach (var exp in modelDef.Exports) try {
          xVal= exp.Value.Value;  //poke value from cell
          /*** Place cell value into into target data (dictionary):
           *   NOTE:
           *   The actual DataDictionary implementation could possibly throw from setter!
           *   (e.g. if the setter tries to cast the object value to a target type...)
           */
          data.SetResolvedValue(exp.Key, xVal, out var _);
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

      ///<inheritdoc/>
      public void Dispose() {
        modelDef.Dispose();
        GC.SuppressFinalize(this);
      }
    }//clas Model
  } //class CalcNgn

}
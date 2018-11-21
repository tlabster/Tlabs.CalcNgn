using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Globalization;

using Microsoft.Extensions.Logging;
using SpreadsheetGear;
using SpreadsheetGear.Shapes;

using Tlabs.Misc;
using Tlabs.CalcNgn.Intern;

namespace Tlabs.CalcNgn.Sgear {
  using DataDictionary= IDictionary<string, object>;

  /// <summary>Impl. of a <see cref="Intern.ICalcNgnModelDef"/>.</summary>
  public class CalcNgnModelDef : Intern.ICalcNgnModelDef {
    private IWorkbook wbk;
    private string namedValPrefix= "WEB_";
    private IReadOnlyDictionary<string, IModelExport> exports;
    private CalcNgnModelDef(IWorkbook wbk, IDictionary<string, IModelImport> imp, IDictionary<string, IModelExport> exp) {
      this.wbk= wbk;
      this.Imports= new ReadOnlyDictionary<string, IModelImport>(imp);
      this.exports= new ReadOnlyDictionary<string, IModelExport>(exp);
    }

    /// <inherit/>
    public string Name => wbk.Name;

    /// <inherit/>
    public IReadOnlyDictionary<string, IModelImport> Imports { get; }

    /// <inherit/>
    public IReadOnlyDictionary<string, IModelExport> Exports {
      get {
        wbk.WorkbookSet.Calculate();
        return exports;
      }
    }

    /// <inherit/>
    public void WriteStream(Stream strm) {
      wbk.WorkbookSet.Calculate();
      wbk.SaveToStream(strm, FileFormat.Excel8);
    }

    /// <inherit/>
    public string NamedValuesPrefix {
      get { return namedValPrefix; }
      set {
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Invalid named value prefix");
        namedValPrefix= value;
      }
    }

    /// <inherit/>
    public void ImportNamedValues(DataDictionary namedVals) {
      IRange rng= null;
      string rngKey= "?";
#if do_not_set_or_clear_import_ranges
      ISheet sht= wbk.Sheets[0];
      foreach (var ent in namedVals) try {
        rngKey= this.namedValPrefix + ent.Key;
        if (null != (rng= sht.EvaluateRange(rngKey)))
          rng.Value= ent.Value;
      }
#else
      for (int l = 0, n = wbk.Names.Count; l < n; ++l) try {
          var rngName= wbk.Names[l];
          if (!(rngKey= rngName.Name).StartsWith(this.namedValPrefix)
              || null == (rng= rngName.RefersToRange)) continue;
          object namedVal;
          if (namedVals.TryGetValue(rngKey.Substring(this.namedValPrefix.Length), out namedVal))
            rng.Value= namedVal;
          else
            rng.ClearContents();
      }
#endif
      catch (Exception e) {
        if (null == e) throw;
        App.Logger<CalcNgnModelDef>().LogWarning("Problem setting value in range: '{key}' ({msg})", rngKey, e.Message);
      }

    }


    /// <summary>Abstract base impl. of a <see cref="Intern.ICalcNgnModelParser"/>.</summary>
    public abstract class AbstractModelParser : Intern.ICalcNgnModelParser, IDisposable {
      /// <summary>Workbook set.</summary>
      protected IWorkbookSet wbSet;

      /// <summary>Ctor from optional <paramref name="culture"/> and <paramref name="licKey"/>.</summary>
      protected AbstractModelParser(CultureInfo culture= null, string licKey= null) {
        try { //ensure licence
          if (null != licKey)
            Factory.SetSignedLicense(licKey);
        }
        catch (Exception e) when (Misc.Safe.NoDisastrousCondition(e)) { }

        this.wbSet= Factory.GetWorkbookSet(culture ?? CultureInfo.InvariantCulture);

        /* Setup workbook set configuration:
         */
        this.wbSet.NeverRecalc= false;
        this.wbSet.CalculateBeforeSave= true;
        this.wbSet.CalculationOnDemand= false;
        this.wbSet.Calculation= Calculation.Manual;//Calculation.Automatic;
        this.wbSet.ReadObjects= true;
        this.wbSet.ReadVBA= false;
      }

      /// <inherit/>
      public ICalcNgnModelDef ParseModelStream(Stream modelStream) {
        IDictionary<string, IModelImport> imp;
        IDictionary<string, IModelExport> exp;
        var wbk= wbSet.Workbooks.OpenFromStream(modelStream);
        this.ParseWorkbook(wbk, out imp, out exp);
        return new CalcNgnModelDef(wbk, imp, exp);
      }

      /// <inherit/>
      protected abstract void ParseWorkbook(IWorkbook wbk, out IDictionary<string, IModelImport> imp, out IDictionary<string, IModelExport> exp);

      /// <inherit/>
      public void Dispose() {
        var wbs= this.wbSet;
        if (null != wbs) {
          wbs.Workbooks.Close(); //close all
          this.wbSet= wbs= null;
        }
      }
    }

  }


}
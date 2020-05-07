using System;
using System.IO;
using System.Collections.Generic;

using Microsoft.Extensions.Logging;

using Tlabs.CalcNgn;
using Tlabs.Data.Entity;
using Tlabs.Data.Serialize;

namespace Tlabs.Data.Processing.Intern {

  /// <summary><see cref="DocSchemaProcessor"/> to perform <see cref="Calculator.Model"/> specific document computation(s).</summary>
  internal class DocSchemaCalcProcessor : DocSchemaProcessor {

    private Calculator.Model calcNgnModel;

    /// <summary>Ctor from <paramref name="compSchema"/>, <paramref name="docSeri"/> and <paramref name="calcNgn"/>.</summary>
    public DocSchemaCalcProcessor(ICompiledDocSchema compSchema, IDynamicSerializer docSeri, Calculator calcNgn)
    : base(compSchema, docSeri) {
      if (null == calcNgn) throw new ArgumentNullException(nameof(calcNgn));

      if (Schema.HasCalcModel) {    // check calc model
        this.calcNgnModel= calcNgn.LoadModel(Schema.CalcModelStream);
        var impCnt= calcNgnModel.Definition.Imports.Count;
        var expCnt= calcNgnModel.Definition.Exports.Count;
        DocSchemaProcessor.log.LogDebug("{schema} has {imp} import(s) and {exp} export(s).", Schema.TypeId, impCnt, expCnt);
        if (0 == impCnt + expCnt) DocSchemaProcessor.log.LogWarning("No data import/export definition found in calcModel of {schema}.", Schema.TypeId);
      }
    }

    ///<summary>Perform any calc. model specific computation(s).</summary>
    protected override object processBodyObject(object bodyObj, Func<object, IDictionary<string, object>> setupData) {
      if (null == calcNgnModel) return bodyObj;
      lock(calcNgnModel) {
        var model=   null != setupData
                  ? setupData(bodyObj)
                  : BodyAccessor.ToDictionary(bodyObj);
#if DEBUG
        SaveModel(Path.Combine(Path.GetDirectoryName(App.MainEntryPath), "calcNgnModel", Schema.TypeName + "0.xls"));
#endif
        calcNgnModel.Compute(model);
#if DEBUG
        SaveModel(Path.Combine(Path.GetDirectoryName(App.MainEntryPath), "calcNgnModel", Schema.TypeName + ".xls"));
#endif
        return bodyObj;
      }
    }

    private void SaveModel(string path) {
      if (null == calcNgnModel) return;
      new DirectoryInfo(Path.GetDirectoryName(path)).Create();  //ensure path
      using (var strm = File.Create(path)) {
        calcNgnModel.Definition.WriteStream(strm);
      }
    }


  }

}
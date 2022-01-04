using System;
using System.Collections.Generic;
using System.IO;

namespace Tlabs.CalcNgn.Intern {

  /// <summary>Interface of a <see cref="Calculator.Model"/> definition parser.</summary>
  public interface ICalcNgnModelParser {
    /// <summary>Parse a <paramref name="modelStream"/> into a <see cref="ICalcNgnModelDef"/>.</summary>
    ICalcNgnModelDef ParseModelStream(Stream modelStream);
  }

  /// <summary><see cref="Calculator.Model"/> definition.</summary>
  public interface ICalcNgnModelDef : IDisposable {
    /// <summary>Model name (calc. workbook name).</summary>
    string Name { get; }
    /// <summary>Read-only dictionary of <see cref="IModelImport"/>.</summary>
    IReadOnlyDictionary<string, IModelImport> Imports { get; }
    /// <summary>Read-only dictionary of <see cref="IModelExport"/>.</summary>
    IReadOnlyDictionary<string, IModelExport> Exports { get; }
    /// <summary>Write model data as binary stream.</summary>
    void WriteStream(Stream strm);
    /// <summary>Import the named values from <paramref name="namedVals"/> into named cells of the calc.sheet model.</summary>
    /// <remarks>
    /// The target range names MUST be prefixed with <see cref="NamedValuesPrefix"/> to match the named values.
    /// <para>i.e. named value: 'foo' is matching the named range: 'WEB_foo' (if <see cref="NamedValuesPrefix"/> is 'WEB_').</para>
    /// <para>NOTE:<br/>
    /// The cells of prefixed named ranges that are not matching any named value are cleared!
    /// </para>
    /// </remarks>
    void ImportNamedValues(IDictionary<string, object> namedVals);
    /// <summary>Prefix of named range name(s) to be considered with <see cref="ImportNamedValues(IDictionary{string, object})"/>.</summary>
    string NamedValuesPrefix { get; set; }

    }

  /// <summary><see cref="Calculator.Model"/> import interface.</summary>
  public interface IModelImport {
    /// <summary>Input <paramref name="value"/>.</summary>
    void Input(object value);
  }

  /// <summary><see cref="Calculator.Model"/> export interface.</summary>
  public interface IModelExport {
    /// <summary>Export value.</summary>
    object Value { get; }
  }
}
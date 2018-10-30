using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Tlabs.CalcNgn {
  /// <summary>Exception thrown from a calc-engine.</summary>
  public class CalcNgnException : GeneralException {

    /// <summary>Default ctor</summary>
    public CalcNgnException() : base() { }

    /// <summary>Ctor from message</summary>
    public CalcNgnException(string message) : base(message) { }

    /// <summary>Ctor from message and inner exception.</summary>
    public CalcNgnException(string message, Exception e) : base(message, e) { }
  }

  /// <summary>Exception thrown from a calc-engine cell.</summary>
  public class CalcNgnCellException : CalcNgnException {
    /// <summary>Ctor from cell index and <paramref name="message"/>.</summary>
    public CalcNgnCellException(int row, int col, string message) : base($"{message} {CellID(row, col)}") {}

    /// <summary>Ctor from cell index and <paramref name="message"/> and <paramref name="e"/>.</summary>
    public CalcNgnCellException(int row, int col, string message, Exception e) : base($"{message} {CellID(row, col)}", e) { }

    /// <summary>Cell ID from index.</summary>
    public static string CellID(int row, int col) => $"@R{row+1}C{col+1}";
  }
}

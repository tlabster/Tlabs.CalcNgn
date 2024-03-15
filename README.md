# Tlabs.CalcNgn

### The Tlabs calculation engine module.

* Create a `Calculator` instance by specifying a `Intern.ICalcNgnModelParser` implementation.
  (`Sgear.CalcNgnModelParser` is a [SpreadsheetGear](https://www.spreadsheetgear.com) based implementation)
* A `Model` obtained from `Calculator.LoadModel()` will receive input data through setting data to `Model.Data`
  and return computed data with the `Model.Data` getter...

### .NET version dependency
*	`2.1.*` .NET 6
*	`2.2.*` .NET 8

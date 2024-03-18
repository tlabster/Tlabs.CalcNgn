# Tlabs.CalcNgn

### The Tlabs calculation engine module.
Provides an abstraction to spreadsheet like calculation engines.  
(The underlying calculation model represents the calculation model/formulas specific to the engine implementation. `Sgear.CalcNgnModelParser` is a [SpreadsheetGear](https://www.spreadsheetgear.com) based implementation. YOU WILL NEED A LICENSE to use this engine !))


* Create a `Calculator` instance by specifying a `Intern.ICalcNgnModelParser` implementation.
* A `Model` obtained from `Calculator.LoadModel()` will receive input data through setting data to `Model.Data`
  and returns computed data with the `Model.Data` getter...

### .NET version dependency
*	`2.1.*` .NET 6
*	`2.2.*` .NET 8

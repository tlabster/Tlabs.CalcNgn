using System.Collections.Generic;
using System.Globalization;

using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;

using Tlabs.Config;
using Tlabs.Data.Processing;

namespace Tlabs.CalcNgn {

  ///<summary>Configures all data repositories as services.</summary>
  public class CalcReposConfigurator : IConfigurator<IServiceCollection> {
    static readonly ILogger log= App.Logger<CalcReposConfigurator>();
    readonly CultureInfo? calcNgnCulture;
    readonly string? calcNgnLicense;


    ///<summary>Default ctor</summary>
    public CalcReposConfigurator() : this(null) { }

    ///<summary>Ctor from <paramref name="config"/>.</summary>
    public CalcReposConfigurator(IDictionary<string, string>? config) {
      config??= new Dictionary<string, string>(0);
      if (config.TryGetValue("culture", out var cul)) this.calcNgnCulture= new CultureInfo(cul);
      config.TryGetValue("licKey", out this.calcNgnLicense);
    }


    ///<inheritdoc/>
    public void AddTo(IServiceCollection services, IConfiguration cfg) {
      services.TryAddSingleton<CalcNgn.Calculator>();
      services.TryAddSingleton<CalcNgn.Intern.ICalcNgnModelParser>((svcProv) => new CalcNgn.Sgear.CalcNgnModelParser(this.calcNgnCulture, this.calcNgnLicense));
      services.AddScoped<IDocProcessorRepo, DocCalcProcessorRepo>();
      log.LogDebug("Calc. repository services added.");
    }


  }
}
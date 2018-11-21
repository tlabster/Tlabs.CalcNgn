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
    // IDictionary<string, string> config;
    CultureInfo calcNgnCulture;
    string calcNgnLicense;


    ///<summary>Default ctor</summary>
    public CalcReposConfigurator() : this(null) { }

    ///<summary>Ctor from <paramref name="config"/>.</summary>
    public CalcReposConfigurator(IDictionary<string, string> config) {
      config= config ?? new Dictionary<string, string>(0);
      string val;
      if (config.TryGetValue("culture", out val))
        this.calcNgnCulture= new CultureInfo(val);
      config.TryGetValue("licKey", out this.calcNgnLicense);
    }

    private ILogger log= App.Logger<CalcReposConfigurator>();

    ///<inherit/>
    public void AddTo(IServiceCollection services, IConfiguration cfg) {
      services.TryAddSingleton<CalcNgn.Calculator>();
      services.TryAddSingleton<CalcNgn.Intern.ICalcNgnModelParser>((svcProv) => new CalcNgn.Sgear.CalcNgnModelParser(this.calcNgnCulture, this.calcNgnLicense));
      services.AddScoped(typeof(DocCalcProcessorRepo));
      log.LogDebug("Calc. repository services added.");
    }

    
  }
}
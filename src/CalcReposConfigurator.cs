using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.DependencyInjection.Extensions;
using Microsoft.Extensions.Logging;

using Tlabs.Config;
using Tlabs.Data.Processing;

namespace Tlabs.CalcNgn {

  ///<summary>Configures all data repositories as services.</summary>
  public class CalcReposConfigurator : IConfigurator<IServiceCollection> {
    ///<summary>Default time zone</summary>
    public const string DEFAULT_TIME_ZONE= "W. Europe Standard Time";   //TODO: this is probably windows only

    private ILogger log= App.Logger<CalcReposConfigurator>();

    ///<inherit/>
    public void AddTo(IServiceCollection services, IConfiguration cfg) {
      services.TryAddSingleton<CalcNgn.Calculator>();
      services.TryAddSingleton<CalcNgn.Intern.ICalcNgnModelParser, CalcNgn.Sgear.CalcNgnModelParser>();
      services.AddScoped(typeof(DocCalcProcessorRepo));
      log.LogDebug("Calc. repository services added.");
    }

    
  }
}
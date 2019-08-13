
using Tlabs.Misc;
using Tlabs.Data.Entity;

namespace Tlabs.Data.Processing {

  ///<summary>Repository of <see cref="Intern.DocSchemaCalcProcessor"/>.</summary>
  internal class DocCalcProcessorRepo : Intern.AbstractDocProcRepo {
    private Tlabs.CalcNgn.Calculator calcNgn;
    private static readonly BasicCache<string, IDocSchemaProcessor> schemaCache= new BasicCache<string, IDocSchemaProcessor>();

    ///<summary>Ctor from services.</summary>
    public DocCalcProcessorRepo(Repo.IDocSchemaRepo schemaRepo,
                            IDocumentClassFactory docClassFactory,
                            Serialize.IDynamicSerializer docSeri,
                            Tlabs.CalcNgn.Calculator calcNgn)
    : base(schemaRepo, docClassFactory, docSeri)
    {
      this.calcNgn= calcNgn;
    }

    ///<inherit/>
    protected override IDocSchemaProcessor createProcessor<TVx, TCx>(DocumentSchema schema, CtxConverterFactory valCfac, CtxConverterFactory evaCfac)
      => new Intern.DocSchemaCalcProcessor(new Intern.CompiledDocSchema<TVx, TCx>(schema, docClassFactory, valCfac, evaCfac), docSeri, calcNgn);
  }


}
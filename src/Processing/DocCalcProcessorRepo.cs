
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
                            SchemaCtxDescriptorResolver ctxDescResolver,
                            Tlabs.CalcNgn.Calculator calcNgn)
    : base(schemaRepo, docClassFactory, docSeri, ctxDescResolver)
    {
      this.calcNgn= calcNgn;
    }

    ///<inherit/>
    protected override IDocSchemaProcessor createProcessor(Intern.ICompiledDocSchema compSchema)
      => new Intern.DocSchemaCalcProcessor(compSchema, docSeri, calcNgn);
  }


}
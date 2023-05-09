
using Tlabs.Misc;
using Tlabs.Data.Entity;

namespace Tlabs.Data.Processing {

  ///<summary>Repository of <see cref="Intern.DocSchemaCalcProcessor"/>.</summary>
  internal class DocCalcProcessorRepo : Intern.AbstractDocProcRepo {
    readonly Tlabs.CalcNgn.Calculator calcNgn;

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

    ///<inheritdoc/>
    protected override IDocSchemaProcessor createProcessor(Intern.ICompiledDocSchema compSchema)
      => new Intern.DocSchemaCalcProcessor(compSchema, docSeri, calcNgn);
  }


}
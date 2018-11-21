
using Tlabs.Misc;
using Tlabs.Data.Entity;

namespace Tlabs.Data.Processing {

  ///<summary>Repository of <see cref="DocSchemaCalcProcessor"/>.</summary>
  public class DocCalcProcessorRepo : Intern.AbstractDocProcRepo<DocSchemaCalcProcessor> {
    private Tlabs.CalcNgn.Calculator calcNgn;
    private static readonly BasicCache<string, DocSchemaCalcProcessor> schemaCache= new BasicCache<string, DocSchemaCalcProcessor>();

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
    protected override DocSchemaCalcProcessor createProcessor(DocumentSchema schema) => new DocSchemaCalcProcessor(schema, docClassFactory, docSeri, calcNgn);
  }


}
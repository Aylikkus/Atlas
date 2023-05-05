using Atlas.Data;

namespace Atlas.Interops
{
    public interface IDocGenerator
    {
        void GenerateDocs(DocAttributes attrs, string pathToTemplate);
    }
}

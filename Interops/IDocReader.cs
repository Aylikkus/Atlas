using Atlas.Data;

namespace Atlas.Interops
{
    public interface IDocReader
    {
        DocAttributes PullAttributes(string pathToFile);
    }
}

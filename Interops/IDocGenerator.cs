using Atlas.Data;
using Microsoft.Office.Interop.Excel;
using System;

namespace Atlas.Interops
{
    public interface IDocGenerator
    {
        event Action<int> DisciplineFinished;
        void GenerateDocs(DocAttributes attrs, string pathToTemplate);
    }
}

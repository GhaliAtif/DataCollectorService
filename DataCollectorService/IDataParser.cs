using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectorService
{
    public interface IDataParser
    {
        List<DataEntry1> ParseDataFile1(string filePath);
        List<DataEntry2> ParseDataFile2(string filePath);
        List<DataEntry3> ParseDataFile3(string filePath);
        List<DataEntry4> ParseDataFile4(string filePath);
        List<DataEntry5> ParseDataFile5(string filePath);
        List<DataEntry6> ParseDataFile6(string filePath);
        // Vous pouvez ajouter d'autres méthodes ou propriétés nécessaires
        DataFileType DetectFileType(string filePath);
    }
}

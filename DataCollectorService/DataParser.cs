using Microsoft.VisualBasic.FileIO;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataCollectorService
{
    public class DataParser : IDataParser
    {

        private readonly Service1 _service;
        private readonly string _errorDirectory;

        public DataParser(Service1 service, string errorDirectory)
        {
            _service = service ?? throw new ArgumentNullException(nameof(service));
            _errorDirectory = errorDirectory ?? throw new ArgumentNullException(nameof(errorDirectory));
        }


        private void MoveFileToErrorDirectory(string filePath)
        {
            try
            {
                var fileName = Path.GetFileName(filePath);
                var destinationPath = Path.Combine(_errorDirectory, fileName);

                if (File.Exists(destinationPath))
                {
                    File.Delete(destinationPath); // Remove the existing file in the error directory
                }

                File.Move(filePath, destinationPath);
                _service.LogToFile($"Moved file '{filePath}' to error directory '{_errorDirectory}'.");
            }
            catch (Exception ex)
            {
                _service.LogToFile($"Error moving file '{filePath}' to error directory: {ex.Message}");
            }
        }

        public List<DataEntry1> ParseDataFile1(string filePath)
        {
            var dataEntries = new List<DataEntry1>();
            bool hasErrors = false;

            if (!File.Exists(filePath))
            {
                _service.LogToFile($"Error: File '{filePath}' not found.");
                return dataEntries;
            }

            var lines = File.ReadAllLines(filePath);

            // Skip the first few metadata lines (assuming fixed number of lines to skip)
            int metadataLinesToSkip = 6;
            if (lines.Length < metadataLinesToSkip)
            {
                return dataEntries;
            }

            // Skip metadata lines
            var dataLines = lines.Skip(metadataLinesToSkip).ToArray();

            foreach (var line in dataLines)
            {
                if (string.IsNullOrWhiteSpace(line) || line.StartsWith(",,") || line.Contains(",,,,,,,,,,,,") || line.Contains("Somme"))
                {
                    continue; // Skip empty lines or lines with only commas
                }

                var columns = line.Split(',');

                // Ensure that the line has the expected number of columns
                if (columns.Length >= 14)
                {
                    try
                    {
                        var number = int.Parse(columns[0].Trim());
                        var sampleName = columns[1].Trim();
                        var dateTime = DateTime.ParseExact(columns[2].Trim(), "dd/MM/yyyy HH:mm", CultureInfo.InvariantCulture);
                        var sum = double.Parse(columns[3].Trim(), CultureInfo.InvariantCulture);
                        var resultType = columns[4].Trim();
                        var SiO2 = double.Parse(columns[5].Trim(), CultureInfo.InvariantCulture);
                        var Al2O3 = double.Parse(columns[6].Trim(), CultureInfo.InvariantCulture);
                        var Fe2O3 = double.Parse(columns[7].Trim(), CultureInfo.InvariantCulture);
                        var MnO = double.Parse(columns[8].Trim(), CultureInfo.InvariantCulture);
                        var MgO = double.Parse(columns[9].Trim(), CultureInfo.InvariantCulture);
                        var Na2O = double.Parse(columns[11].Trim(), CultureInfo.InvariantCulture); // Adjusted for missing column 10
                        var K2O = double.Parse(columns[12].Trim(), CultureInfo.InvariantCulture);
                        var TiO2 = double.Parse(columns[13].Trim(), CultureInfo.InvariantCulture);
                        var P2O5 = double.Parse(columns[14].Trim(), CultureInfo.InvariantCulture);

                        dataEntries.Add(new DataEntry1
                        {
                            Number = number,
                            SampleName = sampleName,
                            DateTime = dateTime,
                            Sum = sum,
                            ResultType = resultType,
                            SiO2 = SiO2,
                            Al2O3 = Al2O3,
                            Fe2O3 = Fe2O3,
                            MnO = MnO,
                            MgO = MgO,
                            Na2O = Na2O,
                            K2O = K2O,
                            TiO2 = TiO2,
                            P2O5 = P2O5
                        });
                    }
                    catch (FormatException ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Error parsing line: {line}. Error: {ex.Message}");
                    }
                }
                else
                {
                    hasErrors = true;
                    _service.LogToFile($"Skipping line with unexpected number of columns: {line}");
                }
            }

            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return dataEntries;
        }


        public List<DataEntry2> ParseDataFile2(string filePath)
        {
            var dataEntries = new List<DataEntry2>();
            bool hasErrors = false;

            if (!File.Exists(filePath))
            {
                _service.LogToFile($"Error: File '{filePath}' not found.");
                return dataEntries;
            }

            var lines = File.ReadAllLines(filePath);

            // Commencez à partir de la 21ème ligne
            for (int i = 48; i < lines.Length; i++)
            {
                var line = lines[i];

                // Ignore empty lines and lines that don't contain data
                if (string.IsNullOrWhiteSpace(line) || !line.Contains(",") ||
                    line.Contains("Analyste") || line.Contains("Date") ||
                    line.Contains("Feuille") || line.Contains("Commentaire") ||
                    line.Contains("Méthodes") || line.Contains("Méthode") ||
                    line.Contains("Elément") || line.Contains("Type") ||
                    line.Contains("conc") || line.Contains("Mode") ||
                    line.Contains("Précision") || line.Contains("Facteur") || line.Contains("Hauteur"))
                {
                    continue;
                }

                // Replace commas in the concentration value with dots
                var modifiedLine = ReplaceThirdCommaWithDot(line);

                var columns = modifiedLine.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                // Ensure that the line has the expected number of columns
                if (columns.Length == 7)
                {
                    try
                    {
                        var date = DateTime.ParseExact(columns[0].Trim('"'), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        var time = TimeSpan.Parse(columns[1].Trim('"'));
                        var element = columns[2].Trim('"');
                        var sampleID = columns[3].Trim('"');
                        var concentration = double.Parse(columns[4].Trim('"'), CultureInfo.InvariantCulture);
                        var unit = columns[6].Trim('"');

                        dataEntries.Add(new DataEntry2
                        {
                            Date = date,
                            Time = time,
                            Element = element,
                            SampleID = sampleID,
                            Concentration = concentration,
                            Unit = unit
                        });
                    }
                    catch (FormatException ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Error parsing line: {line}. Error: {ex.Message}");
                    }
                }
                else
                {
                    hasErrors = true;
                    _service.LogToFile($"Skipping line with unexpected number of columns: {line}");
                }
            }

            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return dataEntries;
        }

        private string ReplaceThirdCommaWithDot(string line)
        {
            int commaCount = 0;
            int index = 0;

            // Find the index of the third comma
            while (index < line.Length && commaCount < 5)
            {
                if (line[index] == ',')
                {
                    commaCount++;
                }
                index++;
            }

            // Replace the third comma with a dot
            if (commaCount == 5)
            {
                line = line.Substring(0, index - 1) + '.' + line.Substring(index);
            }

            return line;
        }

        public List<DataEntry3> ParseDataFile3(string filePath)
        {
            var dataEntries = new List<DataEntry3>();
            bool hasErrors = false;

            if (!File.Exists(filePath))
            {
                _service.LogToFile($"Error: File '{filePath}' not found.");
                return dataEntries;
            }

            var lines = File.ReadAllLines(filePath);
            foreach (var line in lines)
            {
                // Ignore empty lines and lines that don't contain data
                if (string.IsNullOrWhiteSpace(line) || !line.Contains(",") || line.Contains("Analyste") || line.Contains("Date") || line.Contains("Feuille") || line.Contains("Commentaire") || line.Contains("Méthode") || line.Contains("Fe"))
                {
                    continue;
                }

                var columns = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                // Ensure that the line has the expected number of columns
                if (columns.Length == 7)
                {
                    try
                    {
                        var date = DateTime.ParseExact(columns[0].Trim('"'), "dd/MM/yyyy", CultureInfo.InvariantCulture);
                        var time = TimeSpan.Parse(columns[1].Trim('"'));
                        var element = columns[2].Trim('"');
                        var sampleID = columns[3].Trim('"');
                        var concentration = double.Parse(columns[4].Trim('"'), CultureInfo.InvariantCulture);
                        var unit = columns[6].Trim('"');

                        dataEntries.Add(new DataEntry3
                        {
                            Date = date,
                            Time = time,
                            Element = element,
                            SampleID = sampleID,
                            Concentration = concentration,
                            Unit = unit
                        });
                    }
                    catch (FormatException ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Error parsing line: {line}. Error: {ex.Message}");
                    }
                }
                else
                {
                    hasErrors = true;
                    _service.LogToFile($"Skipping line with unexpected number of columns: {line}");
                }
            }
            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return dataEntries;
        }

        public List<DataEntry4> ParseDataFile4(string filePath)
        {
            var dataEntries = new List<DataEntry4>();
            bool hasErrors = false;

            if (!File.Exists(filePath))
            {
                _service.LogToFile($"Error: File '{filePath}' not found.");
                return dataEntries;
            }

            var lines = File.ReadAllLines(filePath);
            bool isFirstLine = true;

            foreach (var line in lines)
            {
                if (isFirstLine || line.StartsWith("\"Date de l'analyse\""))
                {
                    isFirstLine = false;
                    continue; // Skip header line
                }

                if (string.IsNullOrWhiteSpace(line) || !line.Contains(","))
                {
                    continue;
                }

                var columns = line.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries);

                if (columns.Length == 4)
                {
                    try
                    {
                        var date = DateTime.Parse(columns[0].Trim('"'), CultureInfo.InvariantCulture);
                        var name = columns[1].Trim('"');
                        var averageCarbon = double.Parse(columns[2].Trim('"').TrimEnd(' ', '%'), CultureInfo.InvariantCulture);
                        var averageSulfur = double.Parse(columns[3].Trim('"').TrimEnd(' ', '%'), CultureInfo.InvariantCulture);

                        dataEntries.Add(new DataEntry4
                        {
                            Date = date,
                            Name = name,
                            AverageCarbon = averageCarbon,
                            AverageSulfur = averageSulfur
                        });
                    }
                    catch (FormatException ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Error parsing line: {line}. Error: {ex.Message}");
                    }
                }
                else
                {
                    hasErrors = true;
                    _service.LogToFile($"Erreur de saisie dans cette ligne surement a cause d'une virgule: {line}");
                }
            }
            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return dataEntries;
        }

        public List<DataEntry5> ParseDataFile5(string filePath)
        {
            List<DataEntry5> results = new List<DataEntry5>();
            bool hasErrors = false;

            using (StreamReader reader = new StreamReader(filePath))
            {
                string line;
                while ((line = reader.ReadLine()) != null)
                {
                    // Ignorer les lignes vides ou les lignes non pertinentes
                    if (string.IsNullOrWhiteSpace(line) || line.StartsWith("ICP REPORT") || line.StartsWith(",GEO-MO-126") || line.StartsWith(",,,,,,,,"))
                        continue;

                    // Ignorer les virgules initiales
                    if (line.StartsWith(","))
                    {
                        line = line.TrimStart(',');
                    }

                    // Split the line by commas
                    string[] parts = line.Split(',');

                    // Afficher les lignes rejetées
                    if (parts.Length != 8)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Erreur de format de ligne (nombre incorrect de colonnes): {line}");
                        continue;
                    }

                    // Afficher les lignes rejetées
                    if (!parts[0].StartsWith("R-"))
                    {
                        _service.LogToFile($"Ligne rejetée: {line}");
                        continue;
                    }

                    try
                    {
                        // Parse the data into an AnalysisResult object
                        var result = new DataEntry5
                        {
                            SampleId = parts[0].Trim(),
                            Mn = ParseNullableDouble(parts[2]),
                            Fe = ParseNullableDouble(parts[3]),
                            Mg = ParseNullableDouble(parts[4]),
                            Nb = ParseNullableDouble(parts[5]),
                            Ca = ParseNullableDouble(parts[6])
                        };

                        results.Add(result);
                    }
                    catch (FormatException ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Erreur de parsing de la ligne: {line} :: Exception: {ex.Message}");
                    }
                }
            }
            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return results;
        }

        public double? ParseNullableDouble(string value)
        {
            if (double.TryParse(value, NumberStyles.Any, CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }
            return null;
        }

        public List<DataEntry6> ParseDataFile6(string filePath)
        {
            var results = new List<DataEntry6>();
            bool hasErrors = false;

            using (TextFieldParser parser = new TextFieldParser(filePath))
            {
                parser.TextFieldType = FieldType.Delimited;
                parser.SetDelimiters(",");

                // Ignorer la première ligne (en-tête)
                if (!parser.EndOfData)
                    parser.ReadLine();

                int lineNumber = 2; // Commence à compter à partir de la première ligne de données

                while (!parser.EndOfData)
                {
                    var line = parser.ReadLine();
                    try
                    {
                        var data = ParseCsvLineModified(line);
                        if (data != null)
                        {
                            results.Add(data);
                        }
                    }
                    catch (Exception ex)
                    {
                        hasErrors = true;
                        _service.LogToFile($"Erreur de parsing à la ligne {lineNumber}: {ex.Message}");
                    }
                    lineNumber++;
                }
            }
            if (hasErrors)
            {
                MoveFileToErrorDirectory(filePath);
            }

            return results;
        }


        private DataEntry6 ParseCsvLineModified(string line)
        {
            try
            {
                string trimmedLine = line.Trim('"'); // Supprimer les guillemets au début et à la fin de la ligne
                string cleanedLine = trimmedLine.Replace("\"\"", "\""); // Enlever les guillemets doubles dans la ligne

                // Vérifier la longueur minimale avant de continuer
                if (cleanedLine.Length < 2)
                {
                    throw new Exception($"La ligne est trop courte pour être valide : {line} , erreur de saisie VEUILLEZ BIEN SUIVRE LA SYNTAXE DE SAISIE");
                }

                int lastCommaIndex = cleanedLine.LastIndexOf(','); // Trouver l'index de la dernière virgule

                // Vérifier si lastCommaIndex est valide avant de continuer
                if (lastCommaIndex < 0 || lastCommaIndex >= cleanedLine.Length - 1)
                {
                    throw new Exception($"Impossible de trouver une virgule valide dans la ligne : {line} , erreur de saisie VEUILLEZ BIEN SUIVRE LA SYNTAXE DE SAISIE");
                }

                // Remplacer la dernière virgule par un point
                string modifiedLine = cleanedLine.Substring(0, lastCommaIndex) + '.' + cleanedLine.Substring(lastCommaIndex + 1);

                // Diviser la ligne modifiée en ses valeurs
                string[] values = modifiedLine.Split(',');

                if (values.Length != 6)
                {
                    throw new Exception($"Nombre de champs incorrect, erreur de saisie VEUILLEZ BIEN SUIVRE LA SYNTAXE DE SAISIE");
                }

                // Traitement de ""Val,ue"" avec guillemets doubles et point décimal
                string sixthValue = values[5].Trim('"').Replace("\"\"", "\"");
                decimal numericValue = decimal.Parse(sixthValue, NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture);

                var analysisResult = new DataEntry6
                {
                    SampleID = values[0].Trim(),
                    ResultName = values[1].Trim(),
                    Date = DateTime.ParseExact(values[2].Trim(), "dd/MM/yyyy", CultureInfo.InvariantCulture),
                    Time = TimeSpan.Parse(values[3].Trim()),
                    Element = values[4].Trim(),
                    Value = numericValue
                };

                return analysisResult;
            }
            catch (Exception ex)
            {
                throw new Exception($"Erreur lors de la lecture de la ligne : {line}. Détails : {ex.Message}");
            }
        }



        public DataFileType DetectFileType(string filePath)
        {
            // Logic to detect file type based on file content
            string[] lines = File.ReadAllLines(filePath);
            if (lines.Length >= 2 && lines[1].StartsWith("REMINEX - CENTRE DE RECHERCHE"))
            {
                return DataFileType.Type1;
            }
            else if (lines.Length >= 1 && lines[0].Contains("Analyste") && lines[0].Contains("FADIMA"))
            {
                return DataFileType.Type2;
            }
            else if (lines.Length >= 1 && lines[0].Contains("Analyste"))
            {
                return DataFileType.Type3;
            }
            else if (lines.Length >= 1 && lines[0].Contains("\"Date de l'analyse\",\"Nom\",\"Carbone moyenne\",\"Soufre moyenne\""))
            {
                return DataFileType.Type4;
            }
            else if (lines.Length >= 1 && lines[0].Contains("ICP REPORT,"))
            {
                return DataFileType.Type5;
            }
            else if (lines.Length >= 1 && lines[0].Contains("Sample ID,Result Name,Date,Time,Elem,Reported Conc (Calib)"))
            {
                return DataFileType.Type6;
            }
            else
            {
                return DataFileType.Unknown;
            }
        }

        public List<object> ParseDataFile(string filePath)
        {
            var fileType = DetectFileType(filePath);
            switch (fileType)
            {
                case DataFileType.Type1:
                    return ParseDataFile1(filePath).Cast<object>().ToList();
                case DataFileType.Type2:
                    return ParseDataFile2(filePath).Cast<object>().ToList();
                case DataFileType.Type3:
                    return ParseDataFile3(filePath).Cast<object>().ToList();
                case DataFileType.Type4:
                    return ParseDataFile4(filePath).Cast<object>().ToList();
                case DataFileType.Type5:
                    return ParseDataFile5(filePath).Cast<object>().ToList();
                case DataFileType.Type6:
                    return ParseDataFile6(filePath).Cast<object>().ToList();
                default:
                    Console.WriteLine("Error: Unknown file type.");
                    return new List<object>();
            }
        }
    }

}

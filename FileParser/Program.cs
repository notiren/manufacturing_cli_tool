using CsvHelper;
using CsvHelper.Configuration;
using SharpCompress.Archives;
using SharpCompress.Common;
using System.Globalization;
using System.Text.Json;
using System.Text.RegularExpressions;
using static FileParserv1.NetworkTask;


namespace FileParserv1
{
    class Program
    {
        static void Main(string[] args)
        {
            try
            {
                Console.WriteLine("Select the process to run:");
                Console.WriteLine("1. Process Adapter Adac Data");
                Console.WriteLine("2. Process Poe Network Data");
                Console.Write("Enter your choice (1 or 2): ");

                var choice = Console.ReadLine();

                Console.WriteLine("Drag and drop the file here and press Enter:");

                string rawPath = Console.ReadLine()?.Trim('"');
                string filePath = CleanDragDropFilePath(rawPath);

                if (string.IsNullOrWhiteSpace(filePath) || !File.Exists(filePath))
                {
                    Console.WriteLine("Invalid file path. Please restart and enter a valid path.");
                    return;
                }

                if (filePath.EndsWith("txt") && choice != "1")
                {
                    Console.WriteLine("Wrong choice, for Adapter Adac data must be a .txt file");
                    return;
                }

                switch (choice)
                {
                    case "1":
                        ProcessAdacData(filePath);
                        break;
                    case "2":
                        ProcessPoeNetowrkData(filePath);
                        break;
                    default:
                        Console.WriteLine("Invalid choice. Please restart and enter 1 or 2.");
                        return;
                }
                Console.WriteLine("-----------");
                Console.WriteLine("Successfully processed the file inside folder: 'extracted'\n");
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
        }

        static string CleanDragDropFilePath(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
                return null;

            input = input.Trim();

            if (input.StartsWith("& "))
                input = input.Substring(2).Trim();

            // Remove surrounding quotes (handles multiple layers)
            while (input.Length > 1 &&
                   ((input.StartsWith("\"") && input.EndsWith("\"")) ||
                    (input.StartsWith("'") && input.EndsWith("'"))))
            {
                input = input.Substring(1, input.Length - 2).Trim();
            }

            input = input.Replace("&&", "&");

            return input;
        }

        public static void ConvertToJsonPoE(List<string> serialNumbers
            , List<Dictionary<string
            , List<Dictionary<string, string>>>> allData
            , List<string> headers
            , List<Dictionary<string, string>> allSummarysBySn
            , string filePath
            , List<List<string>> allTaskDetails
            , string outputFolder = null)
        {
            var list = new List<PowerOverEthernetTask>();

            if (serialNumbers.Count != allSummarysBySn.Count)
            {
                Console.WriteLine("SerialNumbers do not match with the data");
                return;
            }

            for (int i = 0; i < serialNumbers.Count; i++)
            {
                var sn = serialNumbers[i];
                var data = allData[i];
                var summaryData = allSummarysBySn[i];
                var poeTask = new PowerOverEthernetTask()
                {
                    SerialNumber = sn,
                    SummaryData = summaryData,
                    HasPassed = summaryData.ElementAt(5).Value == "Pass",
                };

                var taskDetails = allTaskDetails[i];

                var index = 0;
                foreach (var item in data)
                {
                    var values = item.Value.SelectMany(x => x.Values).ToList();

                    var status = taskDetails[index];

                    index++;
                    var networkTask = new NetworkTask
                    {
                        Name = item.Key,
                        Status = status
                    };
                    var taskSections = new List<TaskSection>();
                    for (int k = 0; k < headers.Count; k++)
                    {
                        var name = headers[k];
                        var isDataset = decimal.TryParse(values[k], out var res);

                        taskSections.Add(new TaskSection
                        {
                            Name = name,
                            Value = values[k],
                            IsDataSet = isDataset,
                            NetworkChartType = SetChartTypeBasedOnName(name, isDataset).ToString()
                        });
                    }
                    networkTask.TaskSections = taskSections;
                    poeTask.NetworkTasks.Add(networkTask);
                }
                list.Add(poeTask);
            }

            var serialized = JsonSerializer.Serialize(list);

            var fileName = Path.GetFileNameWithoutExtension(filePath);

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                outputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "extracted");
            }

            Directory.CreateDirectory(outputFolder);

            string outputFilePath = Path.Combine(outputFolder, $"{fileName}.txt");

            File.WriteAllText(outputFilePath, serialized);

        }

        private static NetworkChartType SetChartTypeBasedOnName(string name, bool isDataset)
        {
            if (isDataset)
            {
                var txPackets = new List<string> { "TxPacket", "RxPacket" };
                if (txPackets.Any(pattern => Regex.IsMatch(name, $@"^{Regex.Escape(pattern)}", RegexOptions.IgnoreCase)))
                {
                    return NetworkChartType.TxPacketsVSRxPackets;
                }

                var frames = new List<string> { "64Byte_1", "65-127_1", "128-255Byte_1", "256-511Byte_1", "512-1023Byte_1" };

                if (frames.Any(pattern => Regex.IsMatch(name, $@"^{Regex.Escape(pattern)}", RegexOptions.IgnoreCase)))
                {
                    return NetworkChartType.FrameSizeDis;
                }

                var latency = new List<string> { "Latency CT(us)" };

                if (latency.Any(x => Regex.IsMatch(name, $@"^{Regex.Escape(x)}", RegexOptions.IgnoreCase)))
                {
                    return NetworkChartType.LatencyPerTest;
                }


                var errorPatterns = new List<string> { "IP Checksum Error", "RxSNError", "RxIPCsError" };
                if (errorPatterns.Any(pattern => Regex.IsMatch(name, $@"^{Regex.Escape(pattern)}", RegexOptions.IgnoreCase)))
                {
                    return NetworkChartType.ErrorCounts;
                }


                var txLineRates = new List<string> { "Tx Line Rate" };
                if (txLineRates.Any(pattern => Regex.IsMatch(name, $@"^{Regex.Escape(pattern)}", RegexOptions.IgnoreCase)))
                {
                    return NetworkChartType.TxLineRateVsTestName;
                }

            }
            return NetworkChartType.None;
        }

        public static void WriteToCsv(List<string> serialNumbers
            , List<string> timeStamps
            , Dictionary<string
            , List<string>> data
            , List<string> headers
            , Dictionary<string, string> testStatuses
            , string filePath
            , string outputFolder = null)
        {
            var config = new CsvConfiguration(cultureInfo: CultureInfo.InvariantCulture)
            {
                Delimiter = ";",
                HasHeaderRecord = false
            };

            var fileName = Path.GetFileNameWithoutExtension(filePath);

            if (string.IsNullOrWhiteSpace(outputFolder))
            {
                outputFolder = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "extracted");
            }

            Directory.CreateDirectory(outputFolder);

            string outputFilePath = Path.Combine(outputFolder, $"{fileName}.csv");

            using var writer = new StreamWriter(outputFilePath);
            using var csv = new CsvWriter(writer, config);

            // Skip to row 3 (as row 1 and 2 are empty)
            for (int i = 0; i < 3; i++) // Start counting from 0, so skip 2 rows for row 3
            {
                csv.NextRecord();
            }

            csv.WriteField("SerialNumber");
            csv.WriteField("IsPass");
            csv.WriteField("StartDate");

            foreach (var header in headers)
            {
                csv.WriteField(header);
            }
            csv.NextRecord();

            // Skip to row 4 for writing serial numbers (no more need for row skips since we're already there)
            // Write serial numbers starting from column 1 (row 4 onward)

            for (int i = 0; i < serialNumbers.Count; i++)
            {
                var sn = serialNumbers[i];
                var dictData = data[sn];
                var testStatus = testStatuses[sn];
                var timeStamp = timeStamps[i];

                csv.WriteField(sn);

                for (int j = 0; j < 2; j++)
                {
                    if (j == 0)
                    {
                        csv.WriteField(testStatus);
                    }
                    if (j == 1)
                    {
                        csv.WriteField(timeStamp);
                    }
                }

                foreach (var d in dictData)
                {
                    csv.WriteField(d);
                }
                csv.NextRecord();
            }

            // Now, write dictionary data horizontally starting from row 6, column 4
            for (int i = 0; i < 2; i++) // Skip 2 more rows to get to row 6
            {
                csv.NextRecord();
            }
        }


        static string ExtractValue(string line, string key)
        {
            int startIndex = line.IndexOf(key) + key.Length;
            if (startIndex >= key.Length)
            {
                return line.Substring(startIndex).Trim();
            }
            return string.Empty;
        }


        static void ProcessAdacData(string filePath)
        {
            var files = ExtractArchiveFilesToMemory(filePath);

            string[] lines = Array.Empty<string>();

            if (files.Count > 0 && files[0].FileData != null)
            {
                lines = files[0].FileData;
            }
            else
            {
                Console.WriteLine("No file found or file is empty.");
            }

            // Dictionary to store headers and their values
            var seq1Values = new Dictionary<string, string>();
            var generalInfo = new Dictionary<string, string>();

            var isFirstSequence = true;
            bool isSecSeq = false, isThirdSeq = false, isFourthSeq = false, isFifthSeq = false, isSixSeq = false, isSevenSeq = false, isEightSeq = false,
                isNineSeq = false, isTenSeq = false, isElevenSeq = false;

            bool firstSeqPassed = false;
            bool secSeqPassed = false;

            var serialNumbers = new List<string>();
            var modelNames = new List<string>();
            var timeStamps = new List<string>();

            var data = new Dictionary<string, List<string>>();
            var testStatus = new Dictionary<string, string>();

            for (int k = 0; k < lines.Length; k++)
            {
                string cleanLine = lines[k].Trim();

                if (cleanLine.Contains("Model") && lines[k + 3].Contains("YYYY_MM_DD"))
                {
                    // Extract and store the values
                    generalInfo["Model Name"] = ExtractValue(cleanLine, "Model Name:");
                    generalInfo["Customer"] = ExtractValue(cleanLine, "Customer:");
                    generalInfo["Serial No"] = ExtractValue(cleanLine, "Serial No:");
                    generalInfo["Order No"] = ExtractValue(lines[k + 1], "Order No.:");
                    generalInfo["Lot No"] = ExtractValue(lines[k + 1], "Lot No.:");
                    generalInfo["Total Load No"] = ExtractValue(lines[k + 1], "Total Load No.:");
                    generalInfo["Environment"] = ExtractValue(lines[k + 2], "Environment:");
                    generalInfo["Inspector"] = ExtractValue(lines[k + 2], "Inspector:");
                    generalInfo["YYYY_MM_DD"] = ExtractValue(lines[k + 3], "YYYY_MM_DD:");
                    generalInfo["Begin Time"] = ExtractValue(lines[k + 3], "Begin Time:");
                    generalInfo["End Time"] = ExtractValue(lines[k + 3], "End Time:");

                    serialNumbers.Add(generalInfo["Serial No"]);
                    modelNames.Add(generalInfo["Model Name"]);
                    timeStamps.Add(generalInfo["YYYY_MM_DD"]);

                }
            }
            var testPassed = false;
            var sn = string.Empty;
            for (int j = 0; j < lines.Length; j++)
            {
                string cleanLine = lines[j].Trim();

                if (cleanLine.Contains("Model") && lines[j + 3].Contains("YYYY_MM_DD"))
                {
                    var serialNumber = ExtractValue(cleanLine, "Serial No:");
                    sn = serialNumber;
                    data.Add(serialNumber, new List<string>());
                    testStatus.Add(serialNumber, "");
                    seq1Values.Clear();
                    testPassed = false;

                }

                if (cleanLine.StartsWith("SEQ.1:") && !cleanLine.StartsWith("SEQ.11:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFirstSequence = true;
                    isElevenSeq = false;
                }
                else if (cleanLine.Contains("SEQ.2:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFirstSequence = false;
                    isSecSeq = true;
                }

                else if (cleanLine.Contains("SEQ.3:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSecSeq = false;
                    isThirdSeq = true;
                }
                else if (cleanLine.Contains("SEQ.4:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isThirdSeq = false;
                    isFourthSeq = true;
                }
                else if (cleanLine.Contains("SEQ.5:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isFourthSeq = false;
                    isFifthSeq = true;
                }

                else if (cleanLine.Contains("SEQ.6:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSixSeq = true;
                    isFifthSeq = false;
                }
                else if (cleanLine.Contains("SEQ.7:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSevenSeq = true;
                    isSixSeq = false;
                }

                else if (cleanLine.Contains("SEQ.8:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isSevenSeq = false;
                    isEightSeq = true;
                }

                else if (cleanLine.Contains("SEQ.9:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isEightSeq = false;
                    isNineSeq = true;
                }
                else if (cleanLine.Contains("SEQ.10:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isNineSeq = false;
                    isTenSeq = true;
                }
                else if (cleanLine.StartsWith("SEQ.11:") && !cleanLine.StartsWith("SEQ.1:"))
                {
                    testPassed = cleanLine.Contains("PASS");
                    isTenSeq = false;
                    isElevenSeq = true;
                    isFirstSequence = false;
                }

                if (isFirstSequence)
                {
                    // Extract key-value pairs using regex or string split for '=' sign
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values["Seq1-" + key] = value;

                            var limits = new Dictionary<string, List<string>>
                            {
                                { "Seq1-Ton Read", new List<string> { "3000", "1000" } }
                            };

                            if (limits.TryGetValue("Seq1-" + key, out var val))
                            {
                                if (int.TryParse(val[0], out var upperLimit) && int.TryParse(val[1], out var lowerLimit))
                                {
                                    // Check if the target value is within the limits
                                    if (decimal.Parse(value) >= lowerLimit && decimal.Parse(value) <= upperLimit)
                                    {
                                        firstSeqPassed = true;
                                    }
                                }
                            }

                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ld    TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ld    Ton") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq1-" + keys[i]] = values[i];
                                var value = values[i];
                                var key = keys[i];

                                var limits = new Dictionary<string, List<string>>
                                {
                                    { "Seq1-Ton Read", new List<string> { "3000", "1000" } }
                                };

                                if (limits.TryGetValue("Seq1-" + key, out var val))
                                {
                                    if (int.TryParse(val[0], out var upperLimit) && int.TryParse(val[1], out var lowerLimit))
                                    {
                                        // Check if the target value is within the limits
                                        if (decimal.Parse(value) >= lowerLimit && decimal.Parse(value) <= upperLimit)
                                        {
                                            firstSeqPassed = true;
                                        }
                                    }
                                }

                            }
                        }
                        j++;
                    }

                    //TBD tomorrow.
                    if (cleanLine.Contains("Ref Ton from LOAD:") && lines[j + 5].Contains("Tdls"))
                    {
                        var keys = Regex.Split(lines[j + 1], @"\s{2,}").ToList();


                    }
                }
                if (isSecSeq || isThirdSeq || isFourthSeq)
                {
                    var seqValue = isSecSeq ? "Seq2-" : isThirdSeq ? "Seq3-" : "Seq4-";
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"{seqValue}" + key] = value;
                        }
                    }
                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("BITS-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"(?<!SLEW)\s{1,}(?!Rate)").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if ((cleanLine.Contains("Vpp Max") || cleanLine.Contains("Vdc Max")) && lines[j + 1].Contains("1."))
                    {
                        // here
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                                var key = keys[i];
                                var value = values[i];

                                var limits = new Dictionary<string, List<string>>
                                {
                                    { "Seq2-Vpp-3 RD", new List<string> { "0.07", "0.005" } },
                                    { "Seq2-Vpp-2 RD", new List<string> { "0.07", "0.005" } },
                                    { "Seq2-Vpp-1 RD", new List<string> { "0.1", "0.01" } }
                                };

                                if (limits.TryGetValue("Seq1-" + key, out var val))
                                {
                                    if (int.TryParse(val[0], out var upperLimit) && int.TryParse(val[1], out var lowerLimit))
                                    {
                                        // Check if the target value is within the limits
                                        if (decimal.Parse(value) >= lowerLimit && decimal.Parse(value) <= upperLimit)
                                        {
                                            firstSeqPassed = false;
                                        }
                                    }
                                }

                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("dV(+) Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vn Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();

                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isFifthSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq5-" + key] = value;
                        }
                    }
                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Period-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("I/R-1") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vs Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq5-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isSixSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq6-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq6-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isSevenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq7-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq7-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vdisable Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq7-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }

                if (isEightSeq || isNineSeq)
                {

                    var seqValue = isEightSeq ? "Seq8-" : "Seq9-";
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"{seqValue}" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("RISE") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Vdc Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values[$"{seqValue}" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isTenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq10-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        keys = keys.Select((item, index) => item.Contains("TRIGG") ? $"TRIGG{index + 1}" : item).ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Ton Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                    if (cleanLine.Contains("Tons Source") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq10-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }
                }
                if (isElevenSeq)
                {
                    if (cleanLine.Contains('='))
                    {
                        // Example: Vin Port       =        1      Vin Type        =       AC
                        string[] parts = Regex.Split(cleanLine, @"\s+=\s+|\s{2,}");

                        for (int i = 0; i < parts.Length - 1; i += 2)
                        {
                            string key = parts[i];
                            string value = parts[i + 1];

                            // Store in the dictionary or process further
                            seq1Values[$"Seq11-" + key] = value;
                        }
                    }

                    if (cleanLine.Contains("Load Name") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("TRIG") && lines[j + 2].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        keys = keys.Select((item, index) => item.Contains("TRIGG") ? $"TRIGG{index + 1}" : item).ToList();
                        var values = Regex.Split(lines[j + 2], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);

                        values = values.Where(x => !string.IsNullOrEmpty(x)).ToList();

                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("Thd Max") && lines[j + 1].Contains("1."))
                    {
                        var keys = Regex.Split(cleanLine, @"\s{2,}").ToList();
                        var values = Regex.Split(lines[j + 1], @"\s{2,}").Where(x => !string.IsNullOrEmpty(x)).ToList();
                        values.RemoveAt(0);
                        keys.RemoveAt(0);
                        if (keys.Count > 0 && values.Count > 0)
                        {
                            for (int i = 0; i < values.Count; i++)
                            {
                                seq1Values["Seq11-" + keys[i]] = values[i];
                            }
                        }
                        j++;
                    }

                    if (cleanLine.Contains("-----------------------------------------------------------------------"))
                    {
                        data[sn] = seq1Values.Values.ToList();
                        testStatus[sn] = testPassed ? "PASS" : "FALSE";
                    }
                }

            }

            var headers = seq1Values.Keys.ToList();
            WriteToCsv(serialNumbers, timeStamps, data, headers, testStatus, filePath);
        }


        static void ProcessPoeNetowrkData(string filePath)
        {
            // Update the path to point to the correct directory for the file.

            var files = ExtractArchiveFilesToMemory(filePath);

            var serialNumbers = new List<string>();
            var serialNumbersFromFileText = new List<string>();
            var allData = new List<Dictionary<string, List<Dictionary<string, string>>>>();
            var headers = new List<string>();
            var summaryAllData = new List<Dictionary<string, string>>();

            var allTaskDetails = new List<List<string>>();

            foreach (var (FileName, FileData) in files)
            {
                if (FileData?.Length > 1 && FileName.EndsWith(".log", StringComparison.OrdinalIgnoreCase))
                {
                    string snValue = "Not Found";
                    var summaryData = new List<string>();
                    var lines = FileData;

                    var sections = SplitIntoMultipleLines(lines);

                    string pattern = @"_report_(\d{2})"; // Match `_report_` followed by 2 digits
                    var match = Regex.Match(FileName, pattern);
                    if (match.Success)
                    {
                        var sn = match.Groups[1].Value;
                        serialNumbersFromFileText.Add(sn);
                    }


                    foreach (var line in lines)
                    {
                        // Check for SN1# line
                        if (line.Contains("SN1#:"))
                        {
                            int snIndex = line.IndexOf("SN1#:") + 5; // Skip "SN1#: " part
                            snValue = line[snIndex..].Trim();
                        }
                    }
                    var t = GetSummaryData(FileData);
                    summaryAllData.Add(t);

                    serialNumbers.Add(snValue);
                    var taskDetails = GetTaskDetails(lines);


                    allTaskDetails.Add(taskDetails);

                    var listBySection = new Dictionary<string, List<Dictionary<string, string>>>();
                    foreach (var sect in sections)
                    {
                        var arr = sect.Value.ToArray();
                        var packetSection = GetPacketSection(arr);
                        var learningSection = GetLearningSection(arr);
                        var processDetails = ExtractProcessDetails(arr);
                        var processTimeSummary = GetProcessTimeSummary(arr);
                        var finalResult = ExtractFinalResult1(arr);
                        var streamCounterResults = GetStreamCounterResults(arr);

                        listBySection.Add(sect.Key, new List<Dictionary<string, string>>
                        {
                             packetSection, learningSection, processDetails, processTimeSummary, finalResult, streamCounterResults
                        });

                        if (headers.Count == 0)
                        {
                            headers = listBySection.SelectMany(x => x.Value).SelectMany(x => x.Keys).ToList();

                        }
                    }
                    allData.Add(listBySection);

                }

            }
            ConvertToJsonPoE(serialNumbers, allData, headers, summaryAllData, filePath, allTaskDetails);
        }

        static Dictionary<string, string> GetSummaryData(string[] lines)
        {

            string snValue = "Not Found";
            var summaryData = new Dictionary<string, string>();
            bool inSummarySection = false;

            var sections = SplitIntoMultipleLines(lines);

            foreach (var line in lines)
            {
                // Check for SN1# line
                if (line.Contains("SN1#:"))
                {
                    int snIndex = line.IndexOf("SN1#:") + 5; // Skip "SN1#: " part
                    snValue = line[snIndex..].Trim();
                }

                if (line.Contains("===<< SUMMARY >>"))
                {
                    inSummarySection = true;
                    continue;
                }

                if (inSummarySection && line.StartsWith("----------------------------------------------------------------"))
                {
                    inSummarySection = false;
                    continue;
                }

                if (inSummarySection && !string.IsNullOrWhiteSpace(line))
                {
                    var separatorIndex = line.IndexOf(':');
                    if (separatorIndex > -1)
                    {
                        var key = line[..separatorIndex].Trim();
                        var value = line[(separatorIndex + 1)..].Trim();
                        summaryData[key] = value; // Add to dictionary
                    }
                }

            }
            return summaryData;

        }
        static string[] SplitHeaderLine(string headerLine)
        {
            return new[] { "Index", "Task Name", "Start Time", "End Time", "Elapsed Time", "Result" };
        }

        static Dictionary<string, string> GetProcessTimeSummary(string[] lines)
        {
            var processTimeSummary = new Dictionary<string, string>();
            bool processing = false;

            foreach (var line in lines)
            {
                // Start processing after "Process Time Summary:"
                if (!processing)
                {
                    if (line.Trim() == "Process Time Summary:")
                    {
                        processing = true;
                    }
                    continue;
                }

                // Skip empty lines or separators
                if (string.IsNullOrWhiteSpace(line) || line.Contains("---"))
                    continue;

                if (char.IsDigit(line.TrimStart()[0]))
                {
                    // Use Regex to capture the process name and time used
                    var match = Regex.Match(line, @"^\s*(\d+)\s+([a-zA-Z\s]+)\s+\d{2}:\d{2}:\d{2}\s+\d{2}:\d{2}:\d{2}\s+(\d+)\s*sec");

                    if (match.Success)
                    {
                        // Extract the process name and time used
                        string processItem = match.Groups[1].Value.Trim() + "-" + match.Groups[2].Value.Trim();
                        string timeUsed = match.Groups[3].Value.Trim() + " sec";

                        // Add the process item and time used to the dictionary
                        processTimeSummary[processItem] = timeUsed;
                    }
                }
            }

            return processTimeSummary;
        }
        static Dictionary<string, string> ExtractProcessDetails(string[] lines)
        {
            var processDetails = new Dictionary<string, string>();
            string currentProcessItem = null;

            bool processing = false;

            foreach (var line in lines)
            {
                if (!processing)
                {
                    if (line.Trim() == "Process Detail:")
                    {
                        processing = true;
                    }
                    continue;
                }

                // Skip empty lines or separators
                if (string.IsNullOrWhiteSpace(line) || line.Contains("---"))
                    continue;

                // Check if the line is a process item
                if (char.IsDigit(line.TrimStart()[0]))
                {
                    // Extract the Process Item Name (after the number)
                    currentProcessItem = line.Substring(line.IndexOf(" ") + 1).Trim();
                }
                else if (line.Trim().StartsWith("Test Time"))
                {
                    var testTime = line.Split('=').Last().Trim();

                    if (currentProcessItem != null)
                    {
                        processDetails[currentProcessItem] = testTime;
                        currentProcessItem = null;
                    }
                }
            }
            return processDetails;

        }
        static Dictionary<string, List<string>> SplitIntoMultipleLines(string[] lines)
        {
            string taskPattern = @"Task Name\s+:\s+(.*)";

            var taskSections = new List<List<string>>();
            var tt = new Dictionary<string, List<string>>();

            string currentTaskName = string.Empty;
            var currentTaskLines = new List<string>();

            foreach (var line in lines)
            {
                var taskMatch = Regex.Match(line, taskPattern);
                if (taskMatch.Success)
                {
                    // If a new Task Name is found and we already have lines for the previous task, store it
                    if (!string.IsNullOrEmpty(currentTaskName))
                    {
                        taskSections.Add(currentTaskLines);
                        //tt[currentTaskName] = currentTaskLines;
                        tt.Add(currentTaskName, currentTaskLines);
                    }

                    // Start a new task
                    currentTaskName = taskMatch.Groups[1].Value.Trim();
                    currentTaskLines = new List<string>(); // Reset the lines for the new task
                }

                currentTaskLines.Add(line);
            }

            if (currentTaskLines.Count > 0)
            {
                taskSections.Add(currentTaskLines);
                if (!tt.ContainsKey(currentTaskName))
                {
                    tt.Add(currentTaskName, currentTaskLines);
                }
            }
            return tt;
        }

        static Dictionary<string, string> GetLearningSection(string[] lines)
        {
            var keysToExtract = new List<string>
            {
                "Learning Count",
                "Learning Delay",
                "Learning Gap",
                "Learning Timeout",
                "Allowable Tolerance Loss(Per Port)",
                "Allowable Tolerance Excess(Per Port)",
                "Minimum Collision"
            };

            // Dictionary to store key-value pairs
            var extractedValues = new Dictionary<string, string>();

            foreach (var line in lines)
            {
                foreach (var key in keysToExtract)
                {
                    if (line.StartsWith(key, StringComparison.OrdinalIgnoreCase))
                    {
                        // Extract the value by splitting on the separator ":"
                        var value = line.Split(new[] { ':' }, 2)
                                        .LastOrDefault()
                                        ?.Trim();
                        if (value != null)
                        {
                            extractedValues[key] = value;
                        }
                    }
                }
            }

            return extractedValues;
        }

        static Dictionary<string, string> GetPacketSection(string[] lines)
        {
            // Define the keys we are looking for
            var keysToExtract = new List<string>
            {
                "Frame Count",
                "Frame Gap",
                "Burst Count",
                "Collision Release Gap",
                "Tx Timeout",
                "Wait for Read Counter"
            };

            // Dictionary to store key-value pairs
            var extractedValues = new Dictionary<string, string>();

            foreach (var line in lines)
            {
                foreach (var key in keysToExtract)
                {
                    if (line.StartsWith(key, StringComparison.OrdinalIgnoreCase))
                    {
                        // Extract the value by splitting on the separator ":"
                        var value = line.Split(new[] { ':' }, 2)
                                        .LastOrDefault()
                                        ?.Trim();
                        if (value != null)
                        {
                            extractedValues[key] = value;
                        }
                    }
                }
            }

            return extractedValues;
        }

        static Dictionary<string, string> ExtractFinalResult1(string[] lines)
        {
            var result = new Dictionary<string, string>();

            // Step 1: Extract port identifiers and associated data
            var ports = new List<string>();
            var portValues = new List<List<string>>();
            bool isPortSection = false;
            var portPattern = @"^\d+\(\d+,\d+,\d+\)$"; // Pattern for valid port (e.g., 1(0,4,1))

            var portHeaders = new[] { "TxPacket", "RxPacket", "TxByte", "RxByte", "X-TAG", "Unicast", "Multicast", "Broadcast", "UnderSize", "OverSize",
                "Pause","Fragment Err"};

            var vlanHeaders = new[]
            {
                "VLAN","IP Checksum Error","Latency CT(us)","Latency SF(us)","Tx Line Rate (Mbps)","IPv4","Packet Loss Rate"
            };


            bool isVlanSection = false;
            foreach (var line in lines)
            {
                if (line.StartsWith("Port", StringComparison.OrdinalIgnoreCase))
                {
                    isPortSection = true;
                    continue;
                }

                if (isPortSection)
                {
                    if (line.StartsWith("====")) // Stop when separator line is reached
                    {
                        isPortSection = false;
                        break;
                    }

                    var split = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (split.Length > 0 && Regex.IsMatch(split[0], portPattern)) // Validate port pattern
                    {
                        ports.Add(split[0]); // Add exact port identifier
                        portValues.Add(split.Skip(1).ToList()); // Add values for the port
                    }
                }
            }

            //// Step 2: Assign port values to the dictionary
            //var portHeaders = lines.FirstOrDefault(line => line.StartsWith("Port"))?.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries).Skip(1).ToArray();
            //if (portHeaders != null)
            //{
            //for (int i = 0; i < ports.Count; i++)
            //{
            //    for (int j = 0; j < portHeaders.Length; j++)
            //    {
            //        result[$"{ports[i]}"] = portValues[i][j];
            //    }
            //}
            //}

            // Step 3: Extract byte data
            var byteHeaders = new[] { "64Byte", "65-127Byte", "128-255Byte", "256-511Byte", "512-1023Byte" };
            var byteValues = new List<List<string>>();
            bool isByteSection = false;

            foreach (var line in lines)
            {
                if (line.Trim().StartsWith("64byte", StringComparison.OrdinalIgnoreCase))
                {
                    isByteSection = true;
                    continue;
                }

                if (isByteSection)
                {
                    if (line.StartsWith("====")) // Stop when separator line is reached
                    {
                        isByteSection = false;
                        break;
                    }

                    var split = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (!split.Any(x => x.StartsWith("-----------+")) && !split.Any(x => x.StartsWith("=============")))
                    {
                        byteValues.Add(split.ToList());
                    }
                }
            }

            var vlanValues = new List<List<string>>();

            foreach (var line in lines)
            {
                if (line.Trim().StartsWith("VLAN       IP Checksum Error", StringComparison.OrdinalIgnoreCase))
                {
                    isVlanSection = true;
                    continue;
                }

                if (isVlanSection)
                {
                    if (line.StartsWith("====")) // Stop when separator line is reached
                    {
                        break;
                    }

                    var split = line.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                    if (!split.Any(x => x.StartsWith("-----------+")) && !split.Any(x => x.StartsWith("=============")))
                    {
                        vlanValues.Add(split.ToList());
                    }
                }
            }

            // Step 3: Assign byte values to the dictionary
            for (int i = 0; i < ports.Count; i++)
            {
                for (int j = 0; j < portHeaders.Length; j++)
                {
                    if (i < portValues.Count && j < portValues[i].Count)
                    {
                        result[$"{portHeaders[j]}_{ports[i]}"] = portValues[i][j];
                    }
                }
            }

            // Step 4: Assign byte values to the dictionary
            for (int i = 0; i < ports.Count; i++)
            {
                for (int j = 0; j < byteHeaders.Length; j++)
                {
                    if (i < byteValues.Count && j < byteValues[i].Count)
                    {
                        result[$"{byteHeaders[j]}_{ports[i]}"] = byteValues[i][j];
                    }
                }
            }

            // Step 5: Assign VLAN values to the dict.
            for (int i = 0; i < ports.Count; i++)
            {
                for (int j = 0; j < vlanHeaders.Length; j++)
                {
                    if (i < vlanValues.Count && j < vlanValues[i].Count)
                    {
                        result[$"{vlanHeaders[j]}_{ports[i]}"] = vlanValues[i][j];
                    }
                }
            }
            return result;
        }

        static Dictionary<string, string> GetStreamCounterResults(string[] lines)
        {
            var result = new Dictionary<string, string>();

            // Parse the input
            //var lines = input.Split(new[] { Environment.NewLine }, StringSplitOptions.RemoveEmptyEntries);

            var headers = new List<string>
            {
                "TxPackets",
                "RxPackets",
                "TxBytes",
                "RxBytes",
                "RxLostPacket",
                "RxSNError",
                "RxIPCsError"
            };

            var headerLineIndex = Array.FindIndex(lines, line => line.Trim().StartsWith("SPort"));
            if (headerLineIndex == -1)
            {
                Console.WriteLine("Headers not found");
                return result;
            }

            var dataRegex = new Regex(@"\((\d+, \d+, \d+)\)|\S+"); // Match (x, y, z) or standalone values

            for (int i = headerLineIndex + 2; i < lines.Length; i++)
            {
                var line = lines[i].Trim();
                if (string.IsNullOrEmpty(line) || line.StartsWith("-----")) break; // Stop at separator or empty line

                var matches = dataRegex.Matches(line).Cast<Match>().Select(m => m.Value).ToList();
                if (matches.Count < 2 + headers.Count) continue; // Skip invalid rows

                var sPort = matches[0]; // Extract SPort
                var dPort = matches[1]; // Extract DPort
                var values = matches.Skip(2).ToArray(); // Remaining values

                // Assign headers and values to the dictionary
                for (int j = 0; j < headers.Count; j++)
                {
                    var key = $"{headers[j]}({sPort})({dPort})";
                    result[key] = values[j];
                }
            }


            return result;
        }

        static List<string> ExtractHeaders(string rawHeaders, string headerDividers)
        {
            var headers = new List<string>();
            var dividerIndices = new List<int>();

            // Identify the start positions of each header using the dividers line
            for (int i = 0; i < headerDividers.Length; i++)
            {
                if (headerDividers[i] == '+')
                {
                    dividerIndices.Add(i);
                }
            }

            // Extract each header based on the boundaries
            for (int i = 0; i < dividerIndices.Count - 1; i++)
            {
                var start = dividerIndices[i] + 1;
                var end = dividerIndices[i + 1];
                var header = rawHeaders.Substring(start, end - start).Trim();
                headers.Add(header);
            }

            return headers;
        }
        static Dictionary<string, string> ExtractFinalResult(string[] inputData)
        {
            // Dictionary to hold results
            var result = new Dictionary<string, string>();
            string pattern = @"(\d+)\(\d+,\d+,\d+\)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)\s+(\d+)";


            // Iterate through the input data
            for (int i = 2; i < inputData.Length; i++) // Skip the first two header lines
            {
                var line = inputData[i];

                var match = Regex.Match(line, pattern);
                if (match.Success)
                {
                    // Extract Port number and the values
                    string port = match.Groups[1].Value;
                    for (int j = 1; j <= 12; j++)
                    {
                        string header = GetHeaderByIndex(j - 1, port);
                        string value = match.Groups[j].Value;
                        result[header] = value;
                    }
                }
            }
            return result;
        }

        static string GetHeaderByIndex(int index, string port)
        {
            string[] headers = { "TxPacket", "RxPacket", "TxByte", "RxByte", "X-TAG", "Unicast", "Multicast", "Broadcast", "UnderSize", "OverSize", "Pause", "Fragment Err" };
            if (index >= 0 && index <= headers.Length)
            {
                return $"{headers[index]}{port}";
            }
            return string.Empty;
        }

        static List<string> GetTaskDetails(string[] lines)
        {
            // Find the start of the task section
            var taskSectionStartIndex = Array.FindIndex(lines, line => line.Contains("Index    Task Name"));

            if (taskSectionStartIndex < 0)
            {
                Console.WriteLine("Task details section not found.");
                return new List<string>();
            }

            // Extract headers (preserve original header structure)
            var headerLine = lines[taskSectionStartIndex];
            var headers = SplitHeaderLine(headerLine);

            // Extract task details
            var taskDetails = lines
                .Skip(taskSectionStartIndex + 2) // Skip headers and separator lines
                .TakeWhile(line => line.Trim().Length > 0 && !line.StartsWith("=")) // Until empty or separator
                .Select(line => SplitTaskLine(line, headers)) // Split the line using header positions
                .ToDictionary(
                    details => $"Index{details[0]}", // Key: "Index1", "Index2", etc.
                    details => details.Skip(1).ToArray() // Value: Task details as array (excluding index)
                );

            var results1 = new List<string>();

            // take only the Result from Array.
            foreach (var item in taskDetails)
            {
                results1.Add(item.Value[4]);
            }

            // Add headers as the first item in the dictionary
            var result = new Dictionary<string, string[]>
            {
                { "Headers", headers }
            };


            // Add task details to the dictionary
            foreach (var kvp in taskDetails)
            {
                result[kvp.Key] = kvp.Value;
            }

            return results1;
        }
        // Split task line based on header column widths
        static string[] SplitTaskLine(string taskLine, string[] headers)
        {
            return new[]
            {
                taskLine.Substring(0, 8).Trim(),               // Index
                taskLine.Substring(8, 36).Trim(),             // Task Name
                taskLine.Substring(44, 16).Trim(),            // Start Time
                taskLine.Substring(61, 16).Trim(),            // End Time
                taskLine.Substring(78, 16).Trim(),            // Elapsed Time
                taskLine.Substring(95).Trim()                 // Result
            };
        }

        private static List<(string FileName, string[]? FileData)> ExtractArchiveFilesToMemory(string archivePath)
        {
            var extractedFiles = new List<(string FileName, string[]? FileData)>();


            using var archive = ArchiveFactory.Open(archivePath);
            if (archive != null && archive.Entries.Count() > 0)
            {
                foreach (var entry in archive.Entries.Where(entry => !entry.IsDirectory))
                {
                    using var memoryStream = new MemoryStream();
                    entry.WriteTo(memoryStream);
                    memoryStream.Seek(0, SeekOrigin.Begin);

                    using var reader = new StreamReader(memoryStream);

                    var lines = reader.ReadToEnd().Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);

                    extractedFiles.Add((entry.Key, lines));
                }
            }


            return extractedFiles;
        }
    }
}

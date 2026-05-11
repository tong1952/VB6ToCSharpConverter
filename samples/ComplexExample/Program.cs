using System;
using System.Collections.Generic;
using System.IO;

namespace VB6ToCSharpConverter.Samples.ComplexExample
{
    internal static class Program
    {
        private const string InputFile = "data.txt";
        private const string OutputFile = "output.csv";

        private static int Main(string[] args)
        {
            try
            {
                LoadData(out string[] names, out int[] scores);
                double avgScore = CalculateAverage(scores);
                SaveReport(names, scores, avgScore);

                for (int i = 0; i < names.Length; i++)
                {
                    Console.WriteLine($"Student: {names[i]} Score: {scores[i]}");
                }

                return 0;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Main error: {ex.HResult} - {ex.Message}");
                return 1;
            }
        }

        private static void LoadData(out string[] names, out int[] scores)
        {
            var nameList = new List<string>();
            var scoreList = new List<int>();

            try
            {
                foreach (string line in File.ReadLines(InputFile))
                {
                    string[] parts = line.Trim().Split(',');
                    if (parts.Length >= 2)
                    {
                        string studentName = parts[0].Trim();
                        if (int.TryParse(parts[1].Trim(), out int score))
                        {
                            nameList.Add(studentName);
                            scoreList.Add(score);
                        }
                        else
                        {
                            Console.WriteLine($"Warning: invalid score for '{studentName}', skipping line.");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"LoadData error: {ex.HResult} - {ex.Message}");
            }

            names = nameList.ToArray();
            scores = scoreList.ToArray();
        }

        private static double CalculateAverage(int[] values)
        {
            try
            {
                if (values.Length == 0)
                {
                    return 0;
                }

                long total = 0;
                foreach (int value in values)
                {
                    total += value;
                }

                return (double)total / values.Length;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"CalculateAverage error: {ex.HResult} - {ex.Message}");
                return 0;
            }
        }

        private static void SaveReport(string[] names, int[] scores, double avgScore)
        {
            try
            {
                var reportLines = new string[names.Length + 1];
                reportLines[0] = "Name,Score,Category";

                for (int i = 0; i < names.Length; i++)
                {
                    reportLines[i + 1] = $"{names[i]},{scores[i]},{GetCategory(scores[i])}";
                }

                File.WriteAllLines(OutputFile, reportLines);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"SaveReport error: {ex.HResult} - {ex.Message}");
            }
        }

        private static string GetCategory(int score)
        {
            return score switch
            {
                >= 90 => "Excellent",
                >= 75 => "Good",
                >= 60 => "Average",
                _ => "Needs Improvement",
            };
        }
    }
}

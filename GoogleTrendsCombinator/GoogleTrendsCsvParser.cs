using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.IO;

namespace GoogleTrendsCombinator
{
    public class GoogleTrendsCsvParser
    {
        private readonly List<string> _files;
        private readonly string _csv;
        private int _fileCount = 0;

        private static readonly string InterestOverTimeSection = "Interest over time";
        //private string _separator;
        //private string _searchTerm;

        /// <summary>
        /// Constructor
        /// </summary>
        /// <param name="csv"></param>
        public GoogleTrendsCsvParser(string csv)
        {
            this._csv = csv;
        }

        public GoogleTrendsCsvParser(List<string> files)
        {
            this._files = files;

            StringBuilder holder = new StringBuilder();
            foreach (var file in _files)
            {
                string csvRead = File.ReadAllText(file);
                holder.Append(csvRead);

                _fileCount++;
            }

            this._csv = holder.ToString();
        }

        public int FileCount { get { return _fileCount; } }

        public string GetSearchTerm()
        {
            int start = _csv.IndexOf(": ");
            int length = _csv.IndexOf("\n") - (start +2);
            return _csv.Substring(start + 2, length);
        }

        public string GetTopMostDate()
        {
            string item = GetAllSectionsAsLines(InterestOverTimeSection).FirstOrDefault();

            if (item != null)
                return item.Substring(0, 10);
            else
                return "9999";
        }

        public string GetSectionAsString(string section, int sectionNumber)
        {
            string result = null;

            MatchCollection matches = Regex.Matches(_csv, "^" + section + ".*$", RegexOptions.Multiline);

            if (matches.Count > 0)
            {
                int start = matches[sectionNumber - 1].Index;
                result = _csv.Substring(start, _csv.Length - start);

                int endStart = result.NthIndexOf("\n", 2) + 1;
                int end = result.IndexOf("\n\n");

                result = result.Substring(endStart, end).ToString();
            }
            else
            {
                return String.Empty;
            }

            return result;
        }

        public IEnumerable<string> GetSectionAsLines(string section, int sectionNumber)
        {
            string items = GetSectionAsString(section, sectionNumber);
            foreach (var item in items.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None))
            {
                yield return item;
            }
        }

        public IEnumerable<string> GetAllSections(string section)
        {
            string result = null;
            MatchCollection matches = Regex.Matches(_csv, "^" + section + ".*$", RegexOptions.Multiline);

            foreach (Match match in matches)
            {
                int start = match.Index;
                result = _csv.Substring(start, _csv.Length - start);

                int endStart = result.NthIndexOf("\n", 2) + 1;
                int end = result.IndexOf("\n\n\n");

                result = result.Substring(endStart, end - endStart).ToString();

                yield return result;
            }
        }

        public IEnumerable<string> GetAllSectionsAsLines(string section)
        {
            foreach (var sect in GetAllSections(section))
            {
                string[] lines = sect.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                foreach(var line in lines)
                {
                    yield return line;
                }
            }
        }

        public IEnumerable<string> GetAllSectionsGrouped(string section)
        {
            int group = 0;

            foreach (var sect in GetAllSections(section))
            {
                group++; //increment group number for each section

                string[] lines = sect.Split(new string[] { "\r\n", "\n" }, StringSplitOptions.None);

                foreach (var line in lines)
                {
                    yield return group.ToString() + "Þ" + line;
                }
            }
        }
    }
}

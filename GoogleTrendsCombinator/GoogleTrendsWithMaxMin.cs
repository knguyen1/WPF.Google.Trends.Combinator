using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GoogleTrendsCombinator
{
    class GoogleTrendsWithMaxMin
    {
        public int Group { get; set; }
        public DateTime Date { get; set; }
        public DateTime WeekStart { get; set; }
        public DateTime WeekEnd { get; set; }
        public int DailyIndex { get; set; }
        public int WeeklyIndex { get; set; }
        public int MaxDailyIndex { get; set; }
        public int MinDailyIndex { get; set; }
        public int MaxWeeklyIndex { get; set; }
        public int MinWeeklyIndex { get; set; }
    }
}

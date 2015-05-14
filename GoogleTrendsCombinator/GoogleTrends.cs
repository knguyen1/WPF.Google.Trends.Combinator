using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace GoogleTrendsCombinator
{
    class GoogleTrends
    {
        public DateTime Date { get; set; }
        public DateTime WeekStart { get; set; }
        public DateTime WeekEnd { get; set; }
        public int DailyIndex { get; set; }
        public int WeeklyIndex { get; set; }
    }
}

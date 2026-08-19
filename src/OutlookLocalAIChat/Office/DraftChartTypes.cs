using System;

namespace OutlookLocalAIChat.Office
{
    // Maps the small model-facing chart vocabulary to XlChartType
    // codes shared by Excel and PowerPoint charts. Unknown names
    // fall back to a clustered column chart instead of failing the
    // draft.
    public static class DraftChartTypes
    {
        public const int ColumnClustered = 51;
        public const int BarClustered = 57;
        public const int Line = 4;
        public const int Pie = 5;
        public const int Area = 1;
        public const int Scatter = -4169;

        public static int Resolve(string name)
        {
            var kind = (name ?? string.Empty)
                .Trim()
                .ToLowerInvariant();
            switch (kind)
            {
                case "bar":
                    return BarClustered;
                case "line":
                    return Line;
                case "pie":
                    return Pie;
                case "area":
                    return Area;
                case "scatter":
                    return Scatter;
                default:
                    return ColumnClustered;
            }
        }
    }
}

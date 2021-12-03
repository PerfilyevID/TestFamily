namespace TestFamily.Common
{
    public static class UnitExtentions
    {
        public static double ToFeets(this double value) => value / 304.8;
        public static double ToFeets(this int value) => value / 304.8;
    }
}
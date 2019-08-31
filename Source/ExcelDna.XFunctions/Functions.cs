using ExcelDna.Integration;

namespace ExcelDna.XFunctions
{
    public static class Functions
    {
        [ExcelFunction(Description = "The XLOOKUP function searches a range or an array, and returns an item corresponding to the first match it finds.\r\nIf a match doesn't exist, then XLOOKUP can return the closest (approximate) match.")]

        public static object XLookup_(
            [ExcelArgument(Description="The lookup value (What you're looking for)")] object lookup_value,
            [ExcelArgument(Description="The array or range to search (Where to find it)", AllowReference=true)] object lookup_array,
            [ExcelArgument(Description="The array or range ot return (What to return)", AllowReference = true)] object return_array, 
            [ExcelArgument(
                Name="[match_mode]",
                Description="the match type (optional)\r\n 0 - Exact match. If none found, return #N/A (default)\r\n -1 - Exact match, else return the next smaller item\r\n 1 - Exact match, else return the next larger item\r\n 2 - A wildcard match - ? means any character and * means any run of characters"
            )] object match_mode,
            [ExcelArgument(
                Name = "[search_mode]",
                Description = "the search mode to use (optional)\r\n 1 - Search first-to-last (default)\r\n -1 - Search last-to-first\r\n 2 - Binary search (sorted ascending order)\r\n -2 - Binary search (sorted descending order)"
            )] object search_mode)
        {
            return ExcelError.ExcelErrorNA;
        }

        [ExcelFunction(Description = "The XMATCH function returns the relative position of an item in an array or range of cells. ")]

        public static object XMatch_(
            [ExcelArgument(Description = "The lookup value (What you're looking for)")] object lookup_value,
            [ExcelArgument(Description = "The array or range to search (Where to find it)", AllowReference = true)] object lookup_array,
            [ExcelArgument(
                Name="[match_mode]",
                Description="the match type (optional)\r\n 0 - Exact match. If none found, return #N/A (default)\r\n -1 - Exact match, else return the next smaller item\r\n 1 - Exact match, else return the next larger item\r\n 2 - A wildcard match - ? means any character and * means any run of characters"
            )] object match_mode,
            [ExcelArgument(
                Name = "[search_mode]",
                Description = "the search mode to use (optional)\r\n 1 - Search first-to-last (default)\r\n -1 - Search last-to-first\r\n 2 - Binary search (sorted ascending order)\r\n -2 - Binary search (sorted descending order)"
            )] object search_mode)
        {
            return ExcelError.ExcelErrorNA;
        }

    }
}

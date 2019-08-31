using ExcelDna.Integration;

namespace ExcelDna.XFunctions
{
    // TODO: Various outstanding XMATCH parameter cases
    //       Split implementation from the argument-heavy definitions
    //       Clean up parameter validation and int / double .Equals messiness
    //       Decide what to do with names and registering on an Excel which has native XLOOKUP support
    //       Release packaging etc.

    public static class Functions
    {
        [ExcelFunction(Description = "The XLOOKUP function searches a range or an array, and returns an item corresponding to the first match it finds.\r\nIf a match doesn't exist, then XLOOKUP can return the closest (approximate) match.")]

        public static object XLOOKUP(
            [ExcelArgument(Description="The lookup value (What you're looking for)")] object lookup_value,
            [ExcelArgument(Description="The array or range to search (Where to find it)")] object lookup_array,
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
            // Get the match index from XMATCH
            object matchResult = XMATCH(lookup_value, lookup_array, match_mode, search_mode);
            if (matchResult is ExcelError)
                return matchResult;

            if (!(matchResult is double))
            {
                // Unexpected !? 
                // (Maybe get a result array with multiple values ????)
                return ExcelError.ExcelErrorValue;
            }

            // Now we have a 0-based matchOffset
            int matchOffset = (int)(double)matchResult - 1;

            // Gather some info about lookup_array - need this to shape end result
            bool lookupIsRow;
            if (lookup_array is object[,] lookup_arr)
            {
                // We expect one of these to be 1
                int lookupRows = lookup_arr.GetLength(0);
                int lookupCols = lookup_arr.GetLength(1);
                if (lookupRows == 1)
                {
                    lookupIsRow = true;
                }
                else if (lookupCols == 1)
                {
                    lookupIsRow = false;
                }
                else
                {
                    // Unexpected lookup_array shape?
                    // We should already have returned an error from XMATCH
                    return ExcelError.ExcelErrorValue;
                }
            }
            else
            {
                // scalar - treat as row
                lookupIsRow = true;
            }


            // Return a result if the input sizes and match worked out right
            if (return_array is ExcelReference returnRef)
            {
                // We want to return an ExcelReference from the same sheet
                // Shaped by the lookup_array:
                // * If lookup_array was a row, then return the matched column from return_ref
                // * If lookup_array was a column, then return the matched row from return_ref

                if (lookupIsRow)
                {
                    // Consider it a row - return the column
                    int returnCol = returnRef.ColumnFirst + matchOffset;
                    if (returnCol <= returnRef.ColumnLast)
                    {
                        return new ExcelReference(returnRef.RowFirst, returnRef.RowLast, returnCol, returnCol, returnRef.SheetId);
                    }
                }
                else
                {
                    // Consider it a column - return the row
                    int returnRow = returnRef.RowFirst + matchOffset;
                    if (returnRow <= returnRef.RowLast)
                    {
                        return new ExcelReference(returnRow, returnRow, returnRef.ColumnFirst, returnRef.ColumnLast, returnRef.SheetId);
                    }
                }
            }
            else if (return_array is object[,] returnVals)
            {
                int returnRows = returnVals.GetLength(0);
                int returnCols = returnVals.GetLength(1);

                // return_array is an array of values - we either return a sub-array or a single value
                if (lookupIsRow)
                {
                    // Consider lookup to have been a row - return the column from returnVals
                    int matchCol = matchOffset;
                    if (matchCol < returnCols)
                    {
                        // Make result have the same number of rows as returnVals, and one column
                        var result = new object[returnRows, 1];
                        for (int i = 0; i < returnRows; i++)
                        {
                            result[i, 0] = returnVals[i, matchCol];
                        }
                        return result;
                    }
                }
                else
                {
                    // Consider lookup to have been a column - return the rows from returnVals
                    int matchRow = matchOffset;
                    if (matchRow < returnRows)
                    {
                        // Make result have the same number of cols as returnVals, and one row
                        var result = new object[1, returnCols];
                        for (int i = 0; i < returnCols; i++)
                        {
                            result[0, i] = returnVals[matchRow, i];
                        }
                        return result;
                    }
                }
            }
            else
            {
                // return_array is a single value
                // if the matchPosition is 1 this is OK
                if (matchOffset == 1)
                {
                    return return_array;
                }
            }

            // Any other case means mismatched input sizes
             return ExcelError.ExcelErrorValue;
        }

        // If lookup_value is a scalar (i.e. not object[,]), returns either a double (!) with the (integer) index result, or #N/A
        // Returns #VALUE for a parameter error
        [ExcelFunction(Description = "The XMATCH function returns the relative position of an item in an array or range of cells. ")]
        public static object XMATCH(
            [ExcelArgument(Description = "The lookup value (What you're looking for)")] object lookup_value,
            [ExcelArgument(Description = "The array or range to search (Where to find it)")] object lookup_array,
            [ExcelArgument(
                Name="[match_mode]",
                Description="the match type (optional)\r\n 0 - Exact match. If none found, return #N/A (default)\r\n -1 - Exact match, else return the next smaller item\r\n 1 - Exact match, else return the next larger item\r\n 2 - A wildcard match - ? means any character and * means any run of characters"
            )] object match_mode,
            [ExcelArgument(
                Name = "[search_mode]",
                Description = "the search mode to use (optional)\r\n 1 - Search first-to-last (default)\r\n -1 - Search last-to-first\r\n 2 - Binary search (sorted ascending order)\r\n -2 - Binary search (sorted descending order)"
            )] object search_mode)
        {
            if (lookup_value is object[,])
            {
                // TODO: Does XMATCH support lookup_value that is an array?
                // Do we then loop etc? / return a double[,]?
                return ExcelError.ExcelErrorValue;
            }

            // NOTE: lookup_array is _not_ marked AllowReference=true, so it can only be scalar or object[,]

            // If lookup_array is a scalar, create a 1x1 array to hold it
            if (!(lookup_array is object[,]))
            {
                lookup_array = new object[,] { { lookup_array } };
            }

            // Now we know lookup_array is an object[,]
            var arr = lookup_array as object[,];
            if (arr.GetLength(0) > 1 && arr.GetLength(1) > 1)
                return ExcelError.ExcelErrorValue;

            // Sort out match_mode to correspond with built-in MATCH function

            if (match_mode is ExcelMissing)
                match_mode = 0.0;

            // match_mode must be missing or a number
            if (!(match_mode is double))
            {
                return ExcelError.ExcelErrorValue;
            }

            if (match_mode.Equals(0.0) ||
                match_mode.Equals(1.0) ||
                match_mode.Equals(-1.0))
            {
                // We're not in wildcard mode, escape any wildcards so they don't bother MATCH
                if (lookup_value is string str)
                {
                    str = str.Replace("?", "~?");
                    str = str.Replace("*", "~*");
                    str = str.Replace("~", "~~");
                    lookup_value = str;
                }
            }
            else if (match_mode.Equals(2.0))
            {
                // We are in wildcard mode - keep the wildcards, and set to exact mode (so we can call built-in match)
                match_mode = 0.0;
            }
            else
            {
                // Invalid match_mode
                return ExcelError.ExcelErrorValue;
            }

            // Now deal with search_mode
            if (search_mode is ExcelMissing)
                search_mode = 1.0;

            // search_mode must be missing or a number
            if (!(search_mode is double))
            {
                return ExcelError.ExcelErrorValue;
            }

            // Three cases are fine:
            // match_mode == 0 (exact_match) and search_mode == 1 (first-to-last) ==> match_type = 0 in MATCH
            // match_mode == -1 (exact or next smaller) and search_mode == 2 (sorted ascending) ==> match_type = 1 in MATCH ?????
            // match_mode == 1 (exact or next larger) and search_mode == 2 (sorted descending) ==> match_type = -1 in MATCH ?????

            // We'll use this for the MATCH call
            double match_type;

            if (match_mode.Equals(0.0) && search_mode.Equals(1.0))
            {
                match_type = 0.0;
            }
            else if (match_mode.Equals(-1.0) && search_mode.Equals(2.0))
            {
                match_type = 1.0; // ????? seems strange that XMATCH param -1 corresponds to MATCH param 1
            }
            else if (match_mode.Equals(1.0) && search_mode.Equals(-2.0))
            {
                match_type = -1.0; // ????? seems strange that XMATCH param 1 corresponds to MATCH param -1
            }
            else
            {
                // TODO: Not supported yet - need to sort etc.
                return ExcelError.ExcelErrorValue;

                // Easiest next one is match_mode == 0 && search_mode == -1 (last-to-first)
                // For this we just reverse the lookup_array and incert the result index
                // Next we have match_mode == 0 && (search_mode == 2 || search_mode == -2) 
                // ??? - does the binary search matter here for the exact match?
            }

            // We're OK for MATCH
            return XlCall.Excel(XlCall.xlfMatch, lookup_value, lookup_array, match_type);
        }
    }
}

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using static ExcelDna.Integration.XlCall;

namespace ExcelDna.XFunctions
{
    // TODO: Split implementation from the argument-heavy definitions
    //       Clean up parameter validation and int / double .Equals messiness
    //       Decide what to do with names and registering on an Excel which has native XLOOKUP support
    //       Release packaging etc.
    //       Marking lookup_array also AllowReference=true can improve performance when passed to MATCH (but we need some care otherwise)
    //       (Then we also need to improve our SortedMatch)
    //       Sorting and caching the lookup_array can improve performance
    //       Testing for larger inputs & performance

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
                // TODO: Test along various branches when doing our own comparison
                if (lookup_value is string str)
                {
                    str = str.Replace("~", "~~");
                    str = str.Replace("?", "~?");
                    str = str.Replace("*", "~*");
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
            // Sort out default value
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

            if (match_mode.Equals(0.0) && search_mode.Equals(1.0))
            {
                double match_type = 0.0;
                return Excel(xlfMatch, lookup_value, lookup_array, match_type);
            }
            else if (match_mode.Equals(-1.0) && search_mode.Equals(2.0))
            {
                double match_type = 1.0; // ????? seems strange that XMATCH param -1 corresponds to MATCH param 1
                return Excel(xlfMatch, lookup_value, lookup_array, match_type);
            }
            else if (match_mode.Equals(1.0) && search_mode.Equals(-2.0))
            {
                double match_type = -1.0; // ????? seems strange that XMATCH param 1 corresponds to MATCH param -1
                return Excel(xlfMatch, lookup_value, lookup_array, match_type);
            }
            else
            {
                // Other cases use our ExcelCompare to fix, or our UnsortedMatch
                if (match_mode.Equals(0.0))    // exact match
                {
                    if (search_mode.Equals(-1.0))
                    {
                        return UnsortedMatch(lookup_value, lookup_array, 0, reverse_lookup: true);
                    }
                    else if (search_mode.Equals(2.0))
                    {
                        // Means lookup_array is sorted in ascending order
                        // Do built-in match and check result for equals
                        double match_type = 1.0;
                        var match = Excel(xlfMatch, lookup_value, lookup_array, match_type);
                        if (match is double matchPos)
                        {
                            // Check for equals
                            var matchValue = GetLookupValue(lookup_array, (int)matchPos);
                            if (ExcelCompare(lookup_value, matchValue) == 0)
                                return match;
                            else
                                return ExcelError.ExcelErrorNA;
                        }
                        else
                        {
                            // Some error
                            return match;
                        }
                    }
                    else if (search_mode.Equals(2.0))
                    {
                        // Means lookup_array is sorted in descending order
                        // Do built-in match and check result for equals
                        double match_type = -1.0;
                        var match = Excel(xlfMatch, lookup_value, lookup_array, match_type);
                        if (match is double matchPos)
                        {
                            // Check for equals
                            var matchValue = GetLookupValue(lookup_array, (int)matchPos);
                            if (ExcelCompare(lookup_value, matchValue) == 0)
                                return match;
                            else
                                return ExcelError.ExcelErrorNA;
                        }
                        else
                        {
                            // Some error
                            return match;
                        }
                    }
                }
                else if (match_mode.Equals(-1.0)) // we want best match with match <= lookup
                {
                    var match_type = 1; // Reversed meaning in MATCH and UnsortedMatch
                    if (search_mode.Equals(1.0)) // first-to-last
                    {
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: false);
                    }
                    else if (search_mode.Equals(-1.0)) // last-to-first
                    {
                        // TODO: Check what this case does for equal values - return first hit or last?
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: true);
                    }
                    else if (search_mode.Equals(2.0))
                    {
                        // Already done this case with MATCH - bug if we get here
                        return ExcelError.ExcelErrorValue;
                    }
                    else if (search_mode.Equals(-2.0))
                    {
                        // We have descending sorted data, and want best match at or just after lookup
                        // We'll do the search forward
                        // TODO: Check what this case does for equal values - return first hit or last?
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: false);
                    }

                }
                else if (match_mode.Equals(1.0)) // we want best match with match >= lookup
                {
                    var match_type = -1; // Reversed meaning in MATCH and UnsortedMatch
                    if (search_mode.Equals(1.0)) // first-to-last
                    {
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: false);
                    }
                    else if (search_mode.Equals(-1.0)) // last-to-first
                    {
                        // TODO: Check what this case does for equal values - return first hit or last?
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: true);
                    }
                    else if (search_mode.Equals(2.0))
                    {
                        // We have ascending sorted data, and want best match at or just after lookup
                        // We'll do the search backward
                        // TODO: Check what this case does for equal values - return first hit or last?
                        return UnsortedMatch(lookup_value, lookup_array, match_type, reverse_lookup: true);
                    }
                    else if (search_mode.Equals(-2.0))
                    {
                        // Already done this case with MATCH - bug if we get here
                        return ExcelError.ExcelErrorValue;
                    }
                }
            }

            // TODO: Not supported yet ????
            return ExcelError.ExcelErrorValue;
        }

        // This comparison function must agree exactly with how Excel compares two values
        // I'm avoiding to implment it myself for now - who know what the rules are for strings, errors etc.
        // Instead I base it on a call to MATCH
        // If a < b it returns -1
        // if a == b it returns 0
        // if a > b it returns 1
#if DEBUG
        public  // allows checking on a sheet
#endif
            static int ExcelCompare(object a, object b)
        {
            // Is a <= b ? (MATCH with match_type -1 will return 1.0 if so, else #N/A error)
            object a_leq_b_res = Excel(xlfMatch, a, new object[] { b }, -1.0);
            object b_leq_a_res = Excel(xlfMatch, b, new object[] { a }, -1.0);

            bool a_leq_b = a_leq_b_res is double;
            bool b_leq_a = b_leq_a_res is double;

            if (a_leq_b)
            {
                if (b_leq_a)
                    return 0;
                else
                    return -1;
            }
            else
            {
                return 1;
            }
        }


        // Implements something like MATCH with the same match_type interpretation, but working on unsorted inputs
        // MATCH(lookup_value, lookup_array, [match_type])

        // lookup_value: Required.The value that you want to match in lookup_array.
        // lookup_value argument can be a value (number, text, or logical value) or a cell reference to a number, text, or logical value.
        // lookup_array Required.The range of cells being searched. Either a row or a column, but not a non-trivial 2D array
        // match_type The number -1, 0, or 1. The match_type argument specifies how Excel matches lookup_value with values in lookup_array.The default value for this argument is 1.
        // 1 - finds the largest value that is less than or equal to lookup_value. lookup_array does not need to be sorted.
        // 0 - finds the first value that is exactly equal to lookup_value. NO WILDCARDS.
        // -1 - finds the smallest value that is greater than or equal to lookup_value. lookup_array does not need to be sorted.
#if DEBUG
        public  // allows checking on a sheet
#endif
        static object UnsortedMatch(object lookup_value, object lookup_array, int match_type, bool reverse_lookup = false)
        {
            int currentIndex = 0; // 1-based index
            int bestIndex = 0;    // 1-based index
            object bestValue = null;

            foreach (var currentValue in GetLookupValues(lookup_array))
            {
                currentIndex++;

                int cmp = ExcelCompare(currentValue, lookup_value);
                if (bestValue == null)
                {
                    if (match_type == 0)
                    {
                        if (cmp == 0) // currentValue == lookup_value
                        {
                            // Exact match and we found one - return immediately if not reverse_lookup
                            bestIndex = currentIndex;
                            bestValue = currentValue;
                            if (!reverse_lookup)
                                return (double)currentIndex;
                        }
                    }
                    else if (match_type == 1) // We want best <= lookup
                    {

                        if (cmp == 0 || cmp == -1) // currentValue <= lookup_value
                        {
                            // Take it
                            bestIndex = currentIndex;
                            bestValue = currentValue;
                        }
                    }
                    else if (match_type == -1) // We want best >= lookup
                    {
                        if (cmp == 0 || cmp == 1) // currentValue >= lookup_value
                        {
                            // Take it
                            bestIndex = currentIndex;
                            bestValue = currentValue;
                        }
                    }
                    else
                    {
                        return ExcelError.ExcelErrorValue;
                    }
                }
                else // We have a value already, see if currentValue improves it
                {
                    int cmpBest = ExcelCompare(currentValue, bestValue);

                    if (match_type == 0)
                    {
                        if (cmp == 0) // currentValue == lookup_value
                        {
                            // Exact match and we found one - this is the 'best' so far, since we must be doing reverse_lookup
                            bestIndex = currentIndex;
                            bestValue = currentValue;
                        }
                    }
                    else if (match_type == 1) // We want best <= lookup
                    {
                        if (cmp == 0 || cmp == -1) // currentValue <= lookup_value
                        {
                            if (cmpBest == -1) // current < best
                            {
                                // No good for us
                            }
                            else if (cmpBest == 0) // current == best
                            {
                                // If we are doing reverse_lookup, it's better since it came later
                                if (reverse_lookup)
                                {
                                    bestIndex = currentIndex;
                                    bestValue = currentValue;
                                }
                            }
                            else if (cmpBest == 1) // current > best
                            {
                                // Better than best
                                bestIndex = currentIndex;
                                bestValue = currentValue;
                            }
                        }
                    }
                    else if (match_type == -1) // We want best >= lookup
                    {
                        if (cmp == 0 || cmp == 1) // currentValue >= lookup_value
                        {
                            if (cmpBest == 1) // current > best
                            {
                                // No good for us
                            }
                            else if (cmpBest == 0) // current == best
                            {
                                // If we are doing reverse_lookup, it's better since it came later
                                if (reverse_lookup)
                                {
                                    bestIndex = currentIndex;
                                    bestValue = currentValue;
                                }
                            }
                            else if (cmpBest == -1) // current < best
                            {
                                // Better than best
                                bestIndex = currentIndex;
                                bestValue = currentValue;
                            }
                        }
                    }
                }
            }

            if (bestIndex == 0)
                return ExcelError.ExcelErrorNA;

            // Return 1-based index as double
            return (double)bestIndex;

        }

        // This would change if we allowed lookup_array to be an ExcelReference - we'd want to build single-cell ExcelReferences
        static IEnumerable<object> GetLookupValues(object lookup_array)
        {
            var arr = lookup_array as object[,];
            int rows = arr.GetLength(0);
            int cols = arr.GetLength(1);
            // If it's a single row, we enumerate along the columns
            if (rows == 1)
            {
                for (int i = 0; i < cols; i++)
                {
                    yield return arr[0, i];
                }
            }
            else
            {
                // Else we assume it's a single column, and go down the rows
                for (int i = 0; i < rows; i++)
                {
                    yield return arr[i, 0];
                }
            }
        }

        // TODO: Might deal with ExcelReference later
        static object GetLookupValue(object lookup_array, int oneBasedPosition)
        {
            var arr = lookup_array as object[,];
            int rows = arr.GetLength(0);
            int cols = arr.GetLength(1);
            // If it's a single row, we return from the right column
            if (rows == 1)
            {
                return arr[0, oneBasedPosition - 1];
            }
            else
            {
                // Else we assume it's a single column, and return from the right rows
                return arr[oneBasedPosition - 1, 0];
            }
        }

    }
}

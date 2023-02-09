using System;

namespace PML.Excel.Control
{
    public static class StringHelper
    {
        public static string Before(this string value, string inputString)
        {
            if (string.IsNullOrEmpty(inputString))
                return value;
            int length = value.IndexOf(inputString, StringComparison.Ordinal);
            return length <= 0 ? value : value.Substring(0, length);
        }

        public static string After(this string value, string inputString)
        {
            if (string.IsNullOrEmpty(inputString))
                return value;
            int num = value.IndexOf(inputString, StringComparison.Ordinal);
            return num == -1 || num == value.Length - 1 ? value : value.Substring(num + inputString.Length);
        }
        public static bool StartWithOr(this string value, params string[] input)
        {
            foreach (string inputString in input)
            {
                if (value.StartsWith(inputString))
                    return true;
            }
            return false;
        }
        public static bool IsContain(this string value, params string[] input)
        {
            foreach (string inputString in input)
            {
                if (value.Contains(inputString))
                    return true;
            }
            return false;
        }
        public static double ToDouble(this string value)
        {
            try
            {
                if (value.Contains("."))
                    return Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture);
                else
                    return Convert.ToDouble(value);
            }
            catch { return 0; }
        }
        public static double ToDouble(this int value)
        {
            try
            {
                return Convert.ToDouble(value);
            }
            catch { return 0; }
        }
        public static int ToInt(this string value)
        {
            try { return Convert.ToInt32(value.ToDouble()); }
            catch { return 0; }
        }
    }
    public static class DoubleHelper
    {
        public static double Round(this double value, int digit)
        {
            return Math.Round(value, digit);
        }
        public static double ToInt(this double value)
        {
            try { return Convert.ToInt32(value); }
            catch { return 0; }
        }
    }


}

using System;

namespace Reckon_Connector
{
    public class UtilityHelper
    {
        public static int? IntParseToNull(string stringValue)
        {
            int id;
            return int.TryParse(stringValue, out id) ? (int?)id : null;
        }

        public static int IntParseToDefaultValue(string stringValue, int defaultValue)
        {
            int id;
            return int.TryParse(stringValue, out id) ? id : defaultValue;
        }

        public static double? DblParseToNull(string stringValue)
        {
            double id;
            return double.TryParse(stringValue, out id) ? (double?)id : null;
        }

        public static double DblParseToDefaultValue(string stringValue, double defaultValue)
        {
            double id;
            return double.TryParse(stringValue, out id) ? id : defaultValue;
        }

        public static decimal? DecParseToNull(string stringValue)
        {
            decimal id;
            return decimal.TryParse(stringValue, out id) ? (decimal?)id : null;
        }

        public static DateTime? DatParseToNull(string stringValue)
        {
            DateTime id;
            return DateTime.TryParse(stringValue, out id) ? (DateTime?)id : null;
        }

        public static string DatParseToShortDate(string stringValue)
        {
            DateTime id;
            return DateTime.TryParse(stringValue, out id) ? id.ToShortDateString() : "";
        }

        public static Guid? GuidParseToNull(string stringValue)
        {
            Guid id;
            return Guid.TryParse(stringValue, out id) ? (Guid?)id : null;
        }

        public static string ConvertNullStringToBlank(object value)
        {
            return (value != null) ? value.ToString().Trim() : "";
        }

        public static string GetCurrencyFormat(object value)
        {
            return String.Format("{0:C}", value);
        }

        public static double? ParseCurrencyFormat(string value)
        {
            double id;
            return double.TryParse(value, System.Globalization.NumberStyles.Currency, new System.Globalization.CultureInfo("en-AU"), out id) ? (double?)id : null;
        }

        public static string TimeAgo(DateTime? passedDate)
        {
            const int SECOND = 1;
            const int MINUTE = 60 * SECOND;
            const int HOUR = 60 * MINUTE;
            const int DAY = 24 * HOUR;
            const int MONTH = 30 * DAY;

            DateTime newDate;
            if (passedDate == null)
            {
                return "Never";
            }
            else
            {
                newDate = passedDate.Value;
            }

            var ts = new TimeSpan(DateTime.UtcNow.Ticks - newDate.Ticks);
            double delta = Math.Abs(ts.TotalSeconds);

            if (delta < 1 * MINUTE)
            {
                return ts.Seconds == 1 ? "One second ago" : ts.Seconds + " seconds ago";
            }
            if (delta < 2 * MINUTE)
            {
                return "A minute ago";
            }
            if (delta < 45 * MINUTE)
            {
                return ts.Minutes + " minutes ago";
            }
            if (delta < 90 * MINUTE)
            {
                return "An hour ago";
            }
            if (delta < 24 * HOUR)
            {
                return ts.Hours + " hours ago";
            }
            if (delta < 48 * HOUR)
            {
                return "Yesterday";
            }
            if (delta < 30 * DAY)
            {
                return ts.Days + " days ago";
            }
            if (delta < 12 * MONTH)
            {
                int months = Convert.ToInt32(Math.Floor((double)ts.Days / 30));
                return months <= 1 ? "One month ago" : months + " months ago";
            }
            else
            {
                int years = Convert.ToInt32(Math.Floor((double)ts.Days / 365));
                return years <= 1 ? "One year ago" : years + " years ago";
            }
        }

        public static string Base64Encode(string plainText)
        {
            var plainTextBytes = System.Text.Encoding.UTF8.GetBytes(plainText);
            return System.Convert.ToBase64String(plainTextBytes);
        }

        public static string Base64Decode(string base64EncodedData)
        {
            var base64EncodedBytes = System.Convert.FromBase64String(base64EncodedData);
            return System.Text.Encoding.UTF8.GetString(base64EncodedBytes);
        }

        public static bool IsDecimal(string theValue)
        {
            bool returnVal = false;
            try
            {
                Convert.ToDouble(theValue, System.Globalization.CultureInfo.CurrentCulture);
                returnVal = true;
            }
            catch
            {
                returnVal = false;
            }
            finally
            {
            }

            return returnVal;

        }
    }

}
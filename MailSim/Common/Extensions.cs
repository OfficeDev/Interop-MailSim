using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailSim.Common
{
    public static class Extensions
    {
        public static bool ContainsCaseInsensitive(this string input, string match)
        {
            return input.IndexOf(match, StringComparison.OrdinalIgnoreCase) >= 0;
        }

        public static bool EqualsCaseInsensitive(this string input, string match)
        {
            return input.Equals(match, StringComparison.OrdinalIgnoreCase);
        }
    }
}

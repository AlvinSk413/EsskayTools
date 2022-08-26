using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ABMLengthVSLiveLength
{
    class Util1
    {
        public Util1()
        {

        }
        public static string GetFeetAndInches(double millimeters, int Fractions)
        {
            bool negative = false;
            int value = Math.Sign(millimeters);
            if (value == -1)
            {
                negative = true;
                millimeters = Math.Abs(millimeters);
            }
            string feetAndInches = string.Empty;

            double inches = millimeters / 25.4;

            int a = (int)Math.Truncate(inches / 12);
            inches -= a * 12;

            int b = (int)Math.Truncate(inches);
            inches -= b;

            int c = (int)Math.Round(inches * Fractions);
            int d = Fractions;

            b = b + c / d;
            c = c % d;
            a = a + b / 12;
            b = b % 12;

            while (c > 1 && (c % 2) == 0 && d > 1 && (d % 2) == 0)
            {
                c /= 2;
                d /= 2;
            }

            if (a == 0 && b == 0 && c == 0)
            {
                feetAndInches = "0\"";
            }
            else
            {
                if (millimeters < 0)
                {
                    feetAndInches += '-';
                }

                if (a > 0)
                {
                    feetAndInches += a + "'";
                    if (b >= 0)
                    {
                        feetAndInches += "-";
                    }
                }

                if (b >= 0 || (a > 0 && c > 0))
                {
                    feetAndInches += b.ToString(CultureInfo.InvariantCulture) + '"';
                }

                if (c > 0)
                {
                    feetAndInches += c.ToString(CultureInfo.InvariantCulture);
                    feetAndInches += '/';
                    feetAndInches += d.ToString(CultureInfo.InvariantCulture);
                }
            }
            if (negative == true)
            {
                feetAndInches = "-" + feetAndInches;
            }
            return feetAndInches;
        }
    }
}

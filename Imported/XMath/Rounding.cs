using System;
using System.Collections.Generic;
using System.Text;

namespace CSBoost
{
    public partial class XMath
    {

        public static double fmod(double x, double y)
        {
            double i = Math.Floor(x / y);
            return x - (y * i);
        }

        public static double mceiling(double x, double m)
        {
            double r = fmod(x, m);
            if(r <= XMath.epsilon) return x;
            return x + m - r;
        }

        public static double mfloor(double x, double m)
        {
            double r = fmod(x, m);
            if (r <= XMath.epsilon) return x;
            return x - r;
        }

        public static double mround(double x, double m)
        {
            double r = fmod(x, m);
            if (r <= XMath.epsilon) return x;
            if (r >= m / 2) return x + m - r;
            return x - r;
        }

        public static double even(double x)
        {
            return mround(x, 2);
        }

        public static double odd(double x)
        {
            return mround(x + 1, 2) - 1;
        }

        public static double quotient(double n, double d)
        {
            return Math.Floor(n / d);
        }

        public static double roundup(double x, int n)
        {
            double r = Math.Pow(10, -n);
            return mceiling(x, r);
        }

        public static double rounddown(double x, int n)
        {
            double r = Math.Pow(10, -n);
            return mfloor(x, r);
        }
    }
}

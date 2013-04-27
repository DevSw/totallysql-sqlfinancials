using System;
using System.Collections.Generic;

using System.Text;

namespace CSBoost
{
    public static partial class XMath
    {
        public const int max_root_iterations = 300;
        public const int max_series_iterations = 1000000;

        public abstract class series<T>
        {
            public abstract T next();
        }

        public static double sum_series(series<double> s, double factor, ref int max_terms, double init_value)
        {
            int counter = max_terms;

            double result = init_value;
            double next_term;
            do
            {
                next_term = s.next();
                result += next_term;
            }
            while ((Math.Abs(factor * result) < Math.Abs(next_term)) && --counter > 0);

            // set max_terms to the actual number of terms of the series evaluated:
            max_terms = max_terms - counter;

            return result;
        }

        public static double continued_fraction_a(series<pair<double>> g, double factor, ref int max_terms)
        {
            double tiny = XMath.min_value;

            pair<double> v = g.next();

            double f, C, D, delta, a0;
            f = v.v2;
            a0 = v.v1;
            if (f == 0) f = tiny;
            C = f;
            D = 0;

            int counter = max_terms;

            do
            {
                v = g.next();
                D = v.v2 + v.v1 * D;
                if (D == 0) D = tiny;
                C = v.v2 + v.v1 / C;
                if (C == 0) C = tiny;
                D = 1.0 / D;
                delta = C * D;
                f *= delta;
            } while ((Math.Abs(delta - 1.0) > factor) && --counter > 0);

            max_terms = max_terms - counter;

            return a0 / f;
        }

        public static double continued_fraction_b(series<pair<double>> g, double factor, ref int max_terms)
        {

            double tiny = XMath.min_value;

            pair<double> v = g.next();

            double f, C, D, delta;
            f = v.v2;
            if (f == 0) f = tiny;
            C = f;
            D = 0;

            int counter = max_terms;

            do
            {
                v = g.next();
                D = v.v2 + v.v1 * D;
                if (D == 0) D = tiny;
                C = v.v2 + v.v1 / C;
                if (C == 0) C = tiny;
                D = 1 / D;
                delta = C * D;
                f = f * delta;
            } while ((Math.Abs(delta - 1) > factor) && --counter > 0);

            max_terms = max_terms - counter;

            return f;
        }

        public static double evaluate_even_polynomial(double[] poly, double z, int count)
        {
            return evaluate_polynomial(poly, z * z, count);
        }

        public static double evaluate_even_polynomial(double[] poly, double z)
        {
            return evaluate_polynomial(poly, z * z);
        }

        public static double evaluate_odd_polynomial(double[] poly, double z, int count)
        {
            double[] p2 = new double[poly.Length - 1];
            poly.CopyTo(p2, 1);
            return poly[0] + z * evaluate_polynomial(p2, z * z, count - 1);
        }

        public static double evaluate_odd_polynomial(double[] poly, double z)
        {
            double[] p2 = new double[poly.Length - 1];
            poly.CopyTo(p2, 1);
            return poly[0] + z * evaluate_polynomial(p2, z * z);
        }

        public static double evaluate_polynomial(double[] poly, double z, int count)
        {
            double sum = poly[count - 1];
            for (int i = count - 2; i >= 0; --i)
            {
                sum *= z;
                sum += poly[i];
            }
            return sum;
        }

        public static double evaluate_polynomial(double[] poly, double z)
        {
            int count = poly.Length;
            double sum = poly[count - 1];
            for (int i = count - 2; i >= 0; --i)
            {
                sum *= z;
                sum += poly[i];
            }
            return sum;
        }

        public static double evaluate_rational(double[] num, double[] denom, double z_)
        {
            int count = num.Length;
            double z = z_;
            double s1, s2;
            if (z <= 1)
            {
                s1 = num[count - 1];
                s2 = (double)(denom[count - 1]);
                for (int i = count - 2; i >= 0; --i)
                {
                    s1 *= z;
                    s2 *= z;
                    s1 += num[i];
                    s2 += denom[i];
                }
            }
            else
            {
                z = 1 / z;
                s1 = num[0];
                s2 = (double)(denom[0]);
                for (uint i = 1; i < count; ++i)
                {
                    s1 *= z;
                    s2 *= z;
                    s1 += num[i];
                    s2 += denom[i];
                }
            }
            return s1 / s2;
        }

    }
}

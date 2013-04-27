using System;
using System.Collections.Generic;

using System.Text;

namespace CSBoost
{
    public static partial class XMath
    {
        public static double log1p(double x)
        {
            if (x <= -1) throw new Exception(string.Format("log1p(x) requires x > -1, but got x = {0:G}.", x));
            double u = 1 + x;
            if (u == 1.0)
                return x;
            else
                return Math.Log(u) * (x / (u - 1.0));
        }

        class log1p_series : series<double>
        {
            int k;
            double m_mult;
            double m_prod;

            public log1p_series(double x)
            {
                k = 0;
                m_mult = -x;
                m_prod = -1.0;
            }

            public override double next()
            {
                m_prod *= m_mult;
                return m_prod / ++k;
            }

            int count()
            {
                return k;
            }
        }

        public static double log1pmx(double x)
        {
            if (x < -1)
                throw new Exception(string.Format("log1pmx(x) requires x > -1, but got x = {0:G}.", x));
            if (x == -1)
                throw new OverflowException();

            double a = Math.Abs(x);
            if (a > 0.95) return Math.Log(1 + x) - x;
            // Note that without numeric_limits specialisation support, 
            // epsilon just returns zero, and our "optimisation" will always fail:
            if (a < XMath.epsilon) return -x * x / 2;
            log1p_series s = new log1p_series(x);
            s.next();
            int max_iter = max_series_iterations;
            double result = sum_series(s, XMath.epsilon, ref max_iter, 0.0);
            return result;
        }
    }
}

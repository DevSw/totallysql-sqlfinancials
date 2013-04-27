using System;
using System.Collections.Generic;

using System.Text;

namespace CSBoost
{
    public static partial class XMath
    {
        public struct pair<T>
        {
            public pair(T _v1, T _v2)
            {
                v1 = _v1;
                v2 = _v2;
            }
            public T v1;
            public T v2;
        }

        public static double frexp(double value, out int exp)
        {
            // Translate the double into sign, exponent and mantissa. 
            long bits = BitConverter.DoubleToInt64Bits(value);

            // Note that the shift is sign-extended, hence the test against -1 not 1 
            bool negative = (bits < 0);
            int exponent = (int)((bits >> 52) & 0x7ffL);
            long mantissa = bits & 0xfffffffffffffL;

            // Subnormal numbers; exponent is effectively one higher, 
            // but there's no extra normalisation bit in the mantissa 
            if (exponent == 0)
            {
                exponent++;
            }
            // Normal numbers; leave exponent as it is but add extra 
            // bit to the front of the mantissa 
            else
            {
                mantissa = mantissa | (1L << 52);
            }

            // Bias the exponent. It's actually biased by 1023, but we're 
            // treating the mantissa as m.0 rather than 0.m, so we need 
            // to subtract another 52 from it. 
            exponent -= 1075;

            if (mantissa != 0)
            {

                /* Normalize */
                while ((mantissa & 1) == 0)
                {    /*  i.e., Mantissa is even */
                    mantissa >>= 1;
                    exponent++;
                }
            }

            exp = exponent;
            return mantissa;
        }

        public static void swap(ref double a, ref double b)
        {
            double d = a;
            a = b;
            b = d;
        }

        public static double min(double a, double b)
        {
            return a > b ? b : a;
        }

        public static double max(double a, double b)
        {
            return a > b ? a : b;
        }

        public static double ldexp(double x, int n)
        {
            return x * Math.Pow(2, n);
        }

        public abstract class iterand<T1, T2>
        {
            public abstract T1 next(T2 x);
        }

        public static int sign(double x)
        {
            return x < 0 ? -1 : (x == 0 ? 0 : 1);
        }

        public struct Tuple<T1, T2>
        {
            public Tuple(T1 a, T2 b)
            {
                v1 = a;
                v2 = b;
            }
            public T1 v1;
            public T2 v2;
        }

        public struct Tuple<T1, T2, T3>
        {
            public Tuple(T1 a, T2 b, T3 c)
            {
                v1 = a;
                v2 = b;
                v3 = c;
            }
            public T1 v1;
            public T2 v2;
            public T3 v3;
        }

        private static void handle_zero_derivative(iterand<Tuple<double, double>, double> f,
                                                    ref double last_f0,
                                                    double f0,
                                                    ref double delta,
                                                    ref double result,
                                                    ref double guess,
                                                    double min,
                                                    double max)
        {
            if (last_f0 == 0)
            {
                // this must be the first iteration, pretend that we had a
                // previous one at either min or max:
                if (result == min)
                {
                    guess = max;
                }
                else
                {
                    guess = min;
                }
                last_f0 = f.next(guess).v1;
                delta = guess - result;
            }
            if (sign(last_f0) * sign(f0) < 0)
            {
                // we've crossed over so move in opposite direction to last step:
                if (delta < 0)
                {
                    delta = (result - min) / 2;
                }
                else
                {
                    delta = (result - max) / 2;
                }
            }
            else
            {
                // move in same direction as last step:
                if (delta < 0)
                {
                    delta = (result - max) / 2;
                }
                else
                {
                    delta = (result - min) / 2;
                }
            }
        }

        private static void handle_zero_derivative(iterand<Tuple<double, double, double>, double> f,
                                            ref double last_f0,
                                            double f0,
                                            ref double delta,
                                            ref double result,
                                            ref double guess,
                                            double min,
                                            double max)
        {
            if (last_f0 == 0)
            {
                // this must be the first iteration, pretend that we had a
                // previous one at either min or max:
                if (result == min)
                {
                    guess = max;
                }
                else
                {
                    guess = min;
                }
                last_f0 = f.next(guess).v1;
                delta = guess - result;
            }
            if (sign(last_f0) * sign(f0) < 0)
            {
                // we've crossed over so move in opposite direction to last step:
                if (delta < 0)
                {
                    delta = (result - min) / 2;
                }
                else
                {
                    delta = (result - max) / 2;
                }
            }
            else
            {
                // move in same direction as last step:
                if (delta < 0)
                {
                    delta = (result - max) / 2;
                }
                else
                {
                    delta = (result - min) / 2;
                }
            }
        }

        public static double newton_raphson_iterate(iterand<Tuple<double, double>, double> f, double guess, double min, double max, int digits, int max_iter)
        {
            double f0 = 0.0, f1, last_f0 = 0.0;
            double result = guess;

            double factor = ldexp(1.0, 1 - digits);
            double delta = 1;
            double delta1 = double.MaxValue;
            double delta2 = double.MaxValue;

            int count = max_iter;

            do
            {
                last_f0 = f0;
                delta2 = delta1;
                delta1 = delta;
                Tuple<double, double> p = f.next(result);
                f0 = p.v1;
                f1 = p.v2;
                if (0 == f0) break;
                if (f1 == 0)
                {
                    // Oops zero derivative!!!
                    handle_zero_derivative(f, ref last_f0, f0, ref delta, ref result, ref guess, min, max);
                }
                else
                {
                    delta = f0 / f1;
                }
                if (Math.Abs(delta * 2) > Math.Abs(delta2))
                {
                    // last two steps haven't converged, try bisection:
                    delta = (delta > 0) ? (result - min) / 2 : (result - max) / 2;
                }
                guess = result;
                result -= delta;
                if (result <= min)
                {
                    delta = 0.5 * (guess - min);
                    result = guess - delta;
                    if ((result == min) || (result == max))
                        break;
                }
                else if (result >= max)
                {
                    delta = 0.5 * (guess - max);
                    result = guess - delta;
                    if ((result == min) || (result == max))
                        break;
                }
                // update brackets:
                if (delta > 0)
                    max = guess;
                else
                    min = guess;
            } while (--count > 0 && (Math.Abs(result * factor) < Math.Abs(delta)));

            max_iter -= count;

            return result;
        }

        public static double halley_iterate(iterand<Tuple<double, double, double>, double> f, double guess, double min, double max, int digits, int max_iter)
        {
            double f0 = 0, f1, f2;
            double result = guess;

            double factor = ldexp(1.0, 1 - digits);
            double delta = XMath.max(10000000.0 * guess, 10000000.0);  // arbitarily large delta
            double last_f0 = 0;
            double delta1 = delta;
            double delta2 = delta;

            bool out_of_bounds_sentry = false;

            int count = max_iter;

            do
            {
                last_f0 = f0;
                delta2 = delta1;
                delta1 = delta;
                Tuple<double, double, double> p = f.next(result);
                f0 = p.v1;
                f1 = p.v2;
                f2 = p.v3;

                if (0 == f0) break;
                if ((f1 == 0) && (f2 == 0))
                {
                    // Oops zero derivative!!!
                    handle_zero_derivative(f, ref last_f0, f0, ref delta, ref result, ref guess, min, max);
                }
                else
                {
                    if (f2 != 0)
                    {
                        double denom = 2 * f0;
                        double num = 2 * f1 - f0 * (f2 / f1);

                        if ((Math.Abs(num) < 1) && (Math.Abs(denom) >= Math.Abs(num) * double.MaxValue))
                        {
                            // possible overflow, use Newton step:
                            delta = f0 / f1;
                        }
                        else
                            delta = denom / num;
                        if (delta * f1 / f0 < 0)
                        {
                            // probably cancellation error, try a Newton step instead:
                            delta = f0 / f1;
                        }
                    }
                    else
                        delta = f0 / f1;
                }
                double convergence = Math.Abs(delta / delta2);
                if ((convergence > 0.8) && (convergence < 2))
                {
                    // last two steps haven't converged, try bisection:
                    delta = (delta > 0) ? (result - min) / 2 : (result - max) / 2;
                    if (Math.Abs(delta) > result)
                        delta = sign(delta) * result; // protect against huge jumps!
                    // reset delta2 so that this branch will *not* be taken on the
                    // next iteration:
                    delta2 = delta * 3;
                }
                guess = result;
                result -= delta;

                // check for out of bounds step:
                if (result < min)
                {
                    double diff = ((Math.Abs(min) < 1) && (Math.Abs(result) > 1) && (double.MaxValue / Math.Abs(result) < Math.Abs(min))) ? 1e3 : result / min;
                    if (Math.Abs(diff) < 1)
                        diff = 1 / diff;
                    if (!out_of_bounds_sentry && (diff > 0) && (diff < 3))
                    {
                        // Only a small out of bounds step, lets assume that the result
                        // is probably approximately at min:
                        delta = 0.99 * (guess - min);
                        result = guess - delta;
                        out_of_bounds_sentry = true; // only take this branch once!
                    }
                    else
                    {
                        delta = (guess - min) / 2;
                        result = guess - delta;
                        if ((result == min) || (result == max))
                            break;
                    }
                }
                else if (result > max)
                {
                    double diff = ((Math.Abs(max) < 1) && (Math.Abs(result) > 1) && (double.MaxValue / Math.Abs(result) < Math.Abs(max))) ? 1e3 : result / max;
                    if (Math.Abs(diff) < 1)
                        diff = 1 / diff;
                    if (!out_of_bounds_sentry && (diff > 0) && (diff < 3))
                    {
                        // Only a small out of bounds step, lets assume that the result
                        // is probably approximately at min:
                        delta = 0.99 * (guess - max);
                        result = guess - delta;
                        out_of_bounds_sentry = true; // only take this branch once!
                    }
                    else
                    {
                        delta = (guess - max) / 2;
                        result = guess - delta;
                        if ((result == min) || (result == max))
                            break;
                    }
                }
                // update brackets:
                if (delta > 0)
                    max = guess;
                else
                    min = guess;
            } while (--count > 0 && (Math.Abs(result * factor) < Math.Abs(delta)));

            max_iter -= count;

            return result;
        }
    }
}

using System;
using System.Collections.Generic;

using System.Text;

namespace CSBoost
{
    public static partial class XMath
    {
        public class eps_tolerance
        {
            double eps;

            public eps_tolerance(int bits)
            {
                eps = max(ldexp(1.0, 1 - bits), 2.0 * XMath.epsilon);
            }

            public bool tolerant(double a, double b)
            {
                return (Math.Abs(a - b) / min(Math.Abs(a), Math.Abs(b))) <= eps;
            }
        }

        internal static pair<double> bracket_and_solve_root(iterand<double, double> f, double guess, double factor, bool rising, eps_tolerance tol, ref int max_iter)
        {
            //
            // Set up inital brackets:
            //
            double a = guess;
            double b = a;
            double fa = f.next(a);
            double fb = fa;
            //
            // Set up invocation count:
            //
            int count = max_iter - 1;

            if ((fa < 0) == (guess < 0 ? !rising : rising))
            {
                //
                // Zero is to the right of b, so walk upwards
                // until we find it:
                //
                while (sign(fb) == sign(fa))
                {
                    if (count == 0) throw new Exception(string.Format("Unable to bracket root, last nearest value was {0:G}", b));
                    //
                    // Heuristic: every 20 iterations we double the growth factor in case the
                    // initial guess was *really* bad !
                    //
                    if ((max_iter - count) % 20 == 0)
                        factor *= 2;
                    //
                    // Now go ahead and move our guess by "factor":
                    //
                    a = b;
                    fa = fb;
                    b *= factor;
                    fb = f.next(b);
                    --count;
                }
            }
            else
            {
                //
                // Zero is to the left of a, so walk downwards
                // until we find it:
                //
                while (sign(fb) == sign(fa))
                {
                    if (Math.Abs(a) < XMath.min_value)
                    {
                        // Escape route just in case the answer is zero!
                        max_iter -= count;
                        max_iter += 1;
                        return a > 0 ? new pair<double>(0.0, a) : new pair<double>(a, 0.0);
                    }
                    if (count == 0) throw new Exception(string.Format("Unable to bracket root, last nearest value was {0:G}", a));
                    //
                    // Heuristic: every 20 iterations we double the growth factor in case the
                    // initial guess was *really* bad !
                    //
                    if ((max_iter - count) % 20 == 0)
                        factor *= 2;
                    //
                    // Now go ahead and move are guess by "factor":
                    //
                    b = a;
                    fb = fa;
                    a /= factor;
                    fa = f.next(a);
                    --count;
                }
            }
            max_iter -= count;
            max_iter += 1;
            pair<double> r = toms748_solve(f, (a < 0 ? b : a), (a < 0 ? a : b), (a < 0 ? fb : fa), (a < 0 ? fa : fb), tol, ref count);
            max_iter += count;
            return r;
        }

        public static pair<double> toms748_solve(iterand<double, double> f, double ax, double bx, double fax, double fbx, eps_tolerance tol, ref int max_iter)
        {
            int count = max_iter;
            double a, b, fa, fb, c, u, fu, a0, b0, d = 0.0, fd, e, fe;
            double mu = 0.5;

            // initialise a, b and fa, fb:
            a = ax;
            b = bx;
            if (a >= b) throw new Exception(string.Format("Parameters a and b out of order: a={0:G}, b={1:G}", a, b));
            fa = fax;
            fb = fbx;

            if (tol.tolerant(a, b) || (fa == 0) || (fb == 0))
            {
                max_iter = 0;
                if (fa == 0)
                    b = a;
                else if (fb == 0)
                    a = b;
                return new pair<double>(a, b);
            }

            if (sign(fa) * sign(fb) > 0)
                throw new Exception(string.Format("Parameters a and b do not bracket the root: a={0:G}, b={1:G}", a, b));
            // dummy value for fd, e and fe:
            fe = e = fd = 1e5;

            if (fa != 0)
            {
                //
                // On the first step we take a secant step:
                //
                c = secant_interpolate(a, b, fa, fb);
                bracket(f, ref a, ref b, c, ref fa, ref fb, ref d, ref fd);
                --count;

                if (count > 0 && (fa != 0) && !tol.tolerant(a, b))
                {
                    //
                    // On the second step we take a quadratic interpolation:
                    //
                    c = quadratic_interpolate(a, b, d, fa, fb, fd, 2);
                    e = d;
                    fe = fd;
                    bracket(f, ref a, ref b, c, ref fa, ref fb, ref d, ref fd);
                    --count;
                }
            }

            while (count > 0 && (fa != 0) && !tol.tolerant(a, b))
            {
                // save our brackets:
                a0 = a;
                b0 = b;
                //
                // Starting with the third step taken
                // we can use either quadratic or cubic interpolation.
                // Cubic interpolation requires that all four function values
                // fa, fb, fd, and fe are distinct, should that not be the case
                // then variable prof will get set to true, and we'll end up
                // taking a quadratic step instead.
                //
                double min_diff = XMath.min_value * 32;
                bool prof = (Math.Abs(fa - fb) < min_diff) || (Math.Abs(fa - fd) < min_diff) || (Math.Abs(fa - fe) < min_diff) || (Math.Abs(fb - fd) < min_diff) || (Math.Abs(fb - fe) < min_diff) || (Math.Abs(fd - fe) < min_diff);
                if (prof) c = quadratic_interpolate(a, b, d, fa, fb, fd, 2);
                else c = cubic_interpolate(a, b, d, e, fa, fb, fd, fe);
                //
                // re-bracket, and check for termination:
                //
                e = d;
                fe = fd;
                bracket(f, ref a, ref b, c, ref fa, ref fb, ref d, ref fd);
                if ((0 == --count) || (fa == 0) || tol.tolerant(a, b)) break;
                //
                // Now another interpolated step:
                //
                prof = (Math.Abs(fa - fb) < min_diff) || (Math.Abs(fa - fd) < min_diff) || (Math.Abs(fa - fe) < min_diff) || (Math.Abs(fb - fd) < min_diff) || (Math.Abs(fb - fe) < min_diff) || (Math.Abs(fd - fe) < min_diff);
                if (prof) c = quadratic_interpolate(a, b, d, fa, fb, fd, 3);
                else c = cubic_interpolate(a, b, d, e, fa, fb, fd, fe);
                //
                // Bracket again, and check termination condition, update e:
                //
                bracket(f, ref a, ref b, c, ref fa, ref fb, ref d, ref fd);
                if ((0 == --count) || (fa == 0) || tol.tolerant(a, b)) break;
                //
                // Now we take a double-length secant step:
                //
                if (Math.Abs(fa) < Math.Abs(fb))
                {
                    u = a;
                    fu = fa;
                }
                else
                {
                    u = b;
                    fu = fb;
                }
                c = u - 2 * (fu / (fb - fa)) * (b - a);
                if (Math.Abs(c - u) > (b - a) / 2)
                {
                    c = a + (b - a) / 2;
                }
                //
                // Bracket again, and check termination condition:
                //
                e = d;
                fe = fd;
                bracket(f, ref a, ref b, c, ref fa, ref fb, ref d, ref fd);
                if ((0 == --count) || (fa == 0) || tol.tolerant(a, b)) break;
                //
                // And finally... check to see if an additional bisection step is 
                // to be taken, we do this if we're not converging fast enough:
                //
                if ((b - a) < mu * (b0 - a0)) continue;
                //
                // bracket again on a bisection:
                //
                e = d;
                fe = fd;
                bracket(f, ref a, ref b, a + (b - a) / 2, ref fa, ref fb, ref d, ref fd);
                --count;
            } // while loop

            max_iter -= count;
            if (fa == 0)
            {
                b = a;
            }
            else if (fb == 0)
            {
                a = b;
            }
            return new pair<double>(a, b);
        }

        static void bracket(iterand<double, double> f, ref double a, ref double b, double c, ref double fa, ref double fb, ref double d, ref double fd)
        {
            //
            // Given a point c inside the existing enclosing interval
            // [a, b] sets a = c if f(c) == 0, otherwise finds the new 
            // enclosing interval: either [a, c] or [c, b] and sets
            // d and fd to the point that has just been removed from
            // the interval.  In other words d is the third best guess
            // to the root.
            //
            double tol = XMath.epsilon * 2;
            //
            // If the interval [a,b] is very small, or if c is too close 
            // to one end of the interval then we need to adjust the
            // location of c accordingly:
            //
            if ((b - a) < 2 * tol * a)
            {
                c = a + (b - a) / 2;
            }
            else if (c <= a + Math.Abs(a) * tol)
            {
                c = a + Math.Abs(a) * tol;
            }
            else if (c >= b - Math.Abs(b) * tol)
            {
                c = b - Math.Abs(a) * tol;
            }
            //
            // OK, lets invoke f(c):
            //
            double fc = f.next(c);
            //
            // if we have a zero then we have an exact solution to the root:
            //
            if (fc == 0)
            {
                a = c;
                fa = 0;
                d = 0;
                fd = 0;
                return;
            }
            //
            // Non-zero fc, update the interval:
            //
            if (sign(fa) * sign(fc) < 0)
            {
                d = b;
                fd = fb;
                b = c;
                fb = fc;
            }
            else
            {
                d = a;
                fd = fa;
                a = c;
                fa = fc;
            }
        }

        static double secant_interpolate(double a, double b, double fa, double fb)
        {
            //
            // Performs standard secant interpolation of [a,b] given
            // function evaluations f(a) and f(b).  Performs a bisection
            // if secant interpolation would leave us very close to either
            // a or b.  Rationale: we only call this function when at least
            // one other form of interpolation has already failed, so we know
            // that the function is unlikely to be smooth with a root very
            // close to a or b.
            //
            double tol = XMath.epsilon * 5;
            double c = a - (fa / (fb - fa)) * (b - a);
            if ((c <= a + Math.Abs(a) * tol) || (c >= b - Math.Abs(b) * tol))
                return (a + b) / 2;
            return c;
        }

        static double quadratic_interpolate(double a, double b, double d, double fa, double fb, double fd, int count)
        {
            //
            // Performs quadratic interpolation to determine the next point,
            // takes count Newton steps to find the location of the
            // quadratic polynomial.
            //
            // Point d must lie outside of the interval [a,b], it is the third
            // best approximation to the root, after a and b.
            //
            // Note: this does not guarentee to find a root
            // inside [a, b], so we fall back to a secant step should
            // the result be out of range.
            //
            // Start by obtaining the coefficients of the quadratic polynomial:
            //
            double B = safe_div(fb - fa, b - a, double.MaxValue);
            double A = safe_div(fd - fb, d - b, double.MaxValue);
            A = safe_div(A - B, d - a, 0);

            if (a == 0)
            {
                // failure to determine coefficients, try a secant step:
                return secant_interpolate(a, b, fa, fb);
            }
            //
            // Determine the starting point of the Newton steps:
            //
            double c;
            if (sign(A) * sign(fa) > 0)
            {
                c = a;
            }
            else
            {
                c = b;
            }
            //
            // Take the Newton steps:
            //
            for (int i = 1; i <= count; ++i)
            {
                //c -= safe_div(B * c, (B + A * (2 * c - a - b)), 1 + c - a);
                c -= safe_div(fa + (B + A * (c - b) * (c - a)), B + A * (2 * c - a - b), 1 + c - a);
            }
            if ((c <= a) || (c >= b))
            {
                // Oops, failure, try a secant step:
                c = secant_interpolate(a, b, fa, fb);
            }
            return c;
        }

        static double cubic_interpolate(double a, double b, double d, double e, double fa, double fb, double fd, double fe)
        {
            //
            // Uses inverse cubic interpolation of f(x) at points 
            // [a,b,d,e] to obtain an approximate root of f(x).
            // Points d and e lie outside the interval [a,b]
            // and are the third and forth best approximations
            // to the root that we have found so far.
            //
            // Note: this does not guarentee to find a root
            // inside [a, b], so we fall back to quadratic
            // interpolation in case of an erroneous result.
            double q11 = (d - e) * fd / (fe - fd);
            double q21 = (b - d) * fb / (fd - fb);
            double q31 = (a - b) * fa / (fb - fa);
            double d21 = (b - d) * fd / (fd - fb);
            double d31 = (a - b) * fb / (fb - fa);
            double q22 = (d21 - q11) * fb / (fe - fb);
            double q32 = (d31 - q21) * fa / (fd - fa);
            double d32 = (d31 - q21) * fd / (fd - fa);
            double q33 = (d32 - q22) * fa / (fe - fa);
            double c = q31 + q32 + q33 + a;

            if ((c <= a) || (c >= b))
            {
                // Out of bounds step, fall back to quadratic interpolation:
                c = quadratic_interpolate(a, b, d, fa, fb, fd, 3);
            }
            return c;
        }

        static double safe_div(double num, double denom, double r)
        {
            //
            // return num / denom without overflow,
            // return r if overflow would occur.
            //

            if (Math.Abs(denom) < 1)
            {
                if (Math.Abs(denom * double.MaxValue) <= Math.Abs(num))
                    return r;
            }
            return num / denom;
        }

        public static pair<double> bisect(iterand<double, double> f, double min, double max, eps_tolerance tol, ref int max_iter)
        {
            double fmin = f.next(min);
            double fmax = f.next(max);
            if (fmin == 0) return new pair<double>(min, min);
            if (fmax == 0) return new pair<double>(max, max);

            //
            // Error checking:
            //
            if (min >= max)
            {
                throw new Exception("bisect : arguments in wrong order");
            }
            if (fmin * fmax >= 0)
            {
                throw new Exception("bisect: no change of sign: either there is no root to find, or there are multiple roots in the interval");
            }

            //
            // Three function invocations so far:
            //
            int count = max_iter;
            if (count < 3) count = 0;
            else count -= 3;

            while (count != 0 && (!tol.tolerant(min, max)))
            {
                double mid = (min + max) / 2;
                double fmid = f.next(mid);
                if ((mid == max) || (mid == min))
                    break;
                if (fmid == 0)
                {
                    min = max = mid;
                    break;
                }
                else if (sign(fmid) * sign(fmin) < 0)
                {
                    max = mid;
                    fmax = fmid;
                }
                else
                {
                    min = mid;
                    fmin = fmid;
                }
                --count;
            }

            max_iter -= count;

            return new pair<double>(min, max);
        }
}
}

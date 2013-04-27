using System;
using System.Collections.Generic;
using System.Text;

namespace NETFinancials
{
    public static class XLFinancials
    {
        static readonly CSBoost.XMath.eps_tolerance ep = new CSBoost.XMath.eps_tolerance(53);

        #region generic date functions

        public const int LeapDay = 60;       //Feb 29th is the 60th day of the year

        public static bool IsEndOfFeb(DateTime d)
        {
            return d.Month == 2 && IsEndOfMonth(d);
        }

        public static bool IsEndOfMonth(DateTime d)
        {
            return (d.AddDays(1)).Day == 1;
        }

        public static DateTime EndOfMonth(DateTime d)
        {
            DateTime d2 = d;
            d2 = d2.AddDays(1-d2.Day);
            d2 = d2.AddMonths(1);
            d2 = d2.AddDays(-1);
            return d2;
        }

        public static DateTime StartOfMonth(DateTime d)
        {
            DateTime d2 = d;
            d2.AddDays(-d2.Day);
            return d2;
        }

        private static bool LessThanOrEqualToOneYearApart(DateTime d1, DateTime d2)
        {
            if (d1.Year == d2.Year) return true;
            if (d2.Year == d1.Year + 1 && (d1.Month > d2.Month || (d1.Month == d2.Month && d1.Day >= d2.Day))) return true;
            return false;
        }

        private static bool isin(int i, params int[] p)
        {
            foreach (int j in p) if (i == j) return true;
            return false;
        }

        private static void check_frequency(int frequency, string fun)
        {
            if (!isin(frequency, 1, 2, 4, 6, 12)) throw new Exception(string.Format(fun + ": Frequency must be 1 (annual), 2 (bi-annual), 4 (quarterly), 6 (bi-monthly) or 12 (monthly): (got {0:G}).", frequency));
        }

        #endregion

        #region depreciation

        public static double DB(double cost, double salvage, double life, double period, double my1)
        {
            //Check parameters
            if (life < 0) throw new System.ArgumentException(string.Format("DB: Asset life cannot be negative (got {0:G}).", life));
            if (cost < 0) throw new System.ArgumentException(string.Format("DB: Asset cost cannot be negative (got {0:G}).", cost));
            if (my1 < 1 || my1 > 12) throw new System.ArgumentException(string.Format("DB: Months in Year 1 must be a number between 1 and 12 (got {0:G}).", my1));
            if (salvage < 0.0 || salvage > cost) throw new System.ArgumentException(string.Format("DB: Salvage value must be between 0 and original cost (got 0 <= {0:G} <= {1:G}).", salvage, cost));
            if (period < 0 || period > life + 1) throw new System.ArgumentException(string.Format("DB: Period must be between 0 and original life of asset (got 1 <= {0:G} <= {1:G}).", period, life));

            //Deal with edge cases - for periods outside the life return 0 rather than an error
            if (period < 1 || (my1 == 12 && period > life) || (my1 < 12 && period > life + 1)) return 0.0;

            // Now do main body of calculation
            double rate = Math.Round(1.0 - Math.Pow(salvage / cost, 1.0 / (double)life), 3);        // Fixed rate applied each year
            double fpd = cost * rate * (double)my1 / 12.0;                                          // First period depreciation
            double accd = fpd;                                                                      // Cumulative depreciation
            double retval = fpd;                                                                    // Return value
            double bal = 0.0;
            for (int i = 2; i <= period; i++)                                                       // Periods 2 et seq
            {
                bal = cost - accd;
                retval = bal * rate;
                accd += retval;
            }
            if (my1 < 12 && period == life + 1)                  // Final period depreciation
                retval = (cost - (accd - retval)) * rate * (12.0 - my1) / 12.0;
            return retval;
        }

        public static double DDB(double cost, double salvage, double life, double period, double factor)
        {
            //Check parameters
            if (life < 0) throw new System.ArgumentException(string.Format("DDB: Asset life cannot be negative (got {0:G}).", life));
            if (cost < 0) throw new System.ArgumentException(string.Format("DDB: Asset cost cannot be negative (got {0:G}).", cost));
            if (salvage < 0.0 || salvage > cost) throw new System.ArgumentException(string.Format("DDB: Salvage value must be between 0 and original cost (got 0 <= {0:G} <= {1:G}).", salvage, cost));
            if (period < 0 || period > life) throw new System.ArgumentException(string.Format("DDB: Period must be between 0 and original life of asset (got 1 <= {0:G} <= {1:G}).", period, life));
            if (factor < 0) throw new System.ArgumentException(string.Format("DDB: Factor must be greater than or equal to zero (got {0:G}).", factor));

            //Deal with edge cases - for periods outside the life return 0 rather than an error
            if (period < 1 || period > life) return 0.0;

            // Now do main body of calculation
            double rate = factor / life;
            double retval = 0.0;
            double accd = 0.0;
            double a = 0.0, b = 0.0;
            for (int i = 1; i <= period; i++)
            {
                a = (cost - accd) * rate;
                b = cost - salvage - accd;
                retval = a < b ? a : b;
                accd += retval;
            }
            return retval;
        }

        public static double SLN(double cost, double salvage, double life)
        {
            if (life < 0) throw new System.ArgumentException(string.Format("SLN: Asset life cannot be negative (got {0:G}).", life));
            if (cost < 0) throw new System.ArgumentException(string.Format("SLN: Asset cost cannot be negative (got {0:G}).", cost));
            if (salvage < 0.0 || salvage > cost) throw new System.ArgumentException(string.Format("SLN: Salvage value must be between 0 and original cost (got 0 <= {0:G} <= {1:G}).", salvage, cost));

            // Now do main body of calculation
            double net = cost - salvage;
            return net / life;
        }

        public static double SYD(double cost, double salvage, double life, double period)
        {
            //Check parameters
            if (life < 0) throw new System.ArgumentException(string.Format("SYD: Asset life cannot be negative (got {0:G}).", life));
            if (cost < 0) throw new System.ArgumentException(string.Format("SYD: Asset cost cannot be negative (got {0:G}).", cost));
            if (salvage < 0.0 || salvage > cost) throw new System.ArgumentException(string.Format("SYD: Salvage value must be between 0 and original cost (got 0 <= {0:G} <= {1:G}).", salvage, cost));
            if (period < 0 || period > life) throw new System.ArgumentException(string.Format("SYD: Period must be between 0 and original life of asset (got 1 <= {0:G} <= {1:G}).", period, life));

            //Deal with edge cases - for periods outside the life return 0 rather than an error
            if (period < 1 || period > life) return 0.0;

            // Now do main body of calculation
            double retval = (cost - salvage) * (life - period + 1) * 2.0 / (life * (life + 1));
            return retval;
        }

        public static double VDB(double cost, double salvage, double life, double start_period, double end_period, double factor, bool no_switch)
        {
            //Check parameters
            if (life < 0) throw new System.ArgumentException(string.Format("VDB: Asset life cannot be negative (got {0:G}).", life));
            if (cost < 0) throw new System.ArgumentException(string.Format("VDB: Asset cost cannot be negative (got {0:G}).", cost));
            if (salvage < 0.0 || salvage > cost) throw new System.ArgumentException(string.Format("VDB: Salvage value must be between 0 and original cost (got 0 <= {0:G} <= {1:G}).", salvage, cost));
            if (start_period > end_period) throw new System.ArgumentException(string.Format("VDB: Start period must be before or equal to end period (got {0:G} -> {1:G}).", start_period, end_period));
            if (start_period < 0) throw new System.ArgumentException(string.Format("VDB: Start period must be greater than or equalt to zero (got {0:G}).", start_period));
            if (end_period > life) throw new System.ArgumentException(string.Format("VDB: End period must be less than or equal to life of asset (got {0:G} <= {1:G}).", end_period, life));
            if (factor < 0) throw new System.ArgumentException(string.Format("VDB: Factor must be greater than or equal to zero (got {0:G}).", factor));

            //Deal with edge cases - for periods outside the life return 0 rather than an error
            if (end_period < 1 || start_period > life) return 0.0;

            // Now do main body of calculation
            double start = start_period < 0 ? 0 : start_period;
            double end = end_period > life ? life : end_period;

            double rate = factor / life;
            double pval = 0.0;
            double accd = 0.0;
            double a = 0.0, b = 0.0, c = 0.0;
            double retval = 0.0;
            for (int i = 1; i <= end; i++)
            {
                a = (cost - accd) * rate;
                b = cost - salvage - accd;
                pval = a < b ? a : b;

                if (!no_switch)                 //Calculate straight line for comparison if no_switch == false
                {
                    c = b / (life - i + 1);
                    pval = pval < c ? c : pval;
                }

                accd += pval;
                if (i > start && i <= end) retval += pval;
            }
            return retval;


        }

        public static double AMORLINC(double cost, DateTime purchased, DateTime first_period, double salvage, double period, double rate, ExcelDCDBasis basis)
        {
            double result;
            if (basis == ExcelDCDBasis.Actual_360) throw new ArgumentException("AMORLINC: Basis 2 (Actual / 360) is not supported for this function.");
            if (purchased > first_period) throw new ArgumentException(string.Format("AMORLINC: Purchase date must be <= first_period end date (got {0:G} -> {1:G}).", purchased, first_period));
            if (salvage > cost) throw new ArgumentException(string.Format("AMORLINC: Initial cost must be >= salvage value (got {0:G} -> {1:G}).", cost, salvage));
            if (cost < 0) throw new ArgumentException(string.Format("AMORLINC: Initial cost must be >= 0 (got {0:G}).", cost));
            if (period < 0) throw new ArgumentException(string.Format("AMORLINC: Period must be >= 0 (got {0:G}).", period));
            if (rate < 0) throw new ArgumentException(string.Format("AMORLINC: Rate must be >= 0 (got {0:G}).", rate));
            if (cost > salvage && rate == 0) throw new ArgumentException(string.Format("AMORLINC: Rate must be > 0 when initial cost > salvage value (got {0:G}).", rate));
            if (cost == salvage && cost > 0 && rate > 0) throw new ArgumentException(string.Format("AMORLINC: Initial cost must be > salvage value when rate > 0 (got {0:G} -> {1:G} @ {2:G}).", cost, salvage, rate));

            if (cost == salvage || rate == 0) return 0;
            double DY;
            DateTime p = purchased, f = first_period;
            if (basis == ExcelDCDBasis.Actual_Actual || basis == ExcelDCDBasis.Actual_365)
            {
                if (IsEndOfFeb(p) && p.Day == 29) p = p.AddDays(-1);
                if (IsEndOfFeb(f) && f.Day == 29) f = f.AddDays(-1);
            }
            if (basis == ExcelDCDBasis.Actual_Actual) DY = DateTime.IsLeapYear(purchased.Year) ? 366 : 365;
            else DY = DAYSINYEAR(p, p, basis);
            double days = DAYCOUNT(p, f, basis);
            if (days == 0) days = DY;
            double d0 = Math.Min(rate * cost * days / DY, cost);
            result = d0;
            double value = cost - d0;
            double d = cost * rate;
            for (int i = 1; i <= period; i++)
            {
                if (value <= salvage) return 0;
                if (value - d <= salvage) d = value - salvage;
                value -= d;
                result = d;
            }
            if (result > cost - salvage) result = cost - salvage;
            return result;
        }

        public static double AMORDEGRC(double cost, DateTime purchased, DateTime first_period, double salvage, double period, double rate, ExcelDCDBasis basis)
        {
            double result;
            if (basis == ExcelDCDBasis.Actual_360) throw new ArgumentException("AMORDEGRC: Basis 2 (Actual / 360) is not supported for this function.");
            if (purchased > first_period) throw new ArgumentException(string.Format("AMORDEGRC: Purchase date must be <= first_period end date (got {0:G} -> {1:G}).", purchased, first_period));
            if (salvage > cost) throw new ArgumentException(string.Format("AMORDEGRC: Initial cost must be >= salvage value (got {0:G} -> {1:G}).", cost, salvage));
            if (cost < 0) throw new ArgumentException(string.Format("AMORDEGRC: Initial cost must be >= 0 (got {0:G}).", cost));
            if (period < 0) throw new ArgumentException(string.Format("AMORDEGRC: Period must be >= 0 (got {0:G}).", period));
            if (rate < 0) throw new ArgumentException(string.Format("AMORDEGRC: Rate must be >= 0 (got {0:G}).", rate));
            if (cost > salvage && rate == 0) throw new ArgumentException(string.Format("AMORDEGRC: Rate must be > 0 when initial cost > salvage value (got {0:G}).", rate));
            if (cost == salvage && cost > 0 && rate > 0) throw new ArgumentException(string.Format("AMORDEGRC: Initial cost must be > salvage value when rate > 0 (got {0:G} -> {1:G} @ {2:G}).", cost, salvage, rate));

            if (cost == salvage || rate == 0) return 0;
            double life = Math.Ceiling(1.0 / rate);
            double r = rate;
            if (life >= 3 && life <= 4) r *= 1.5;
            else if (life >= 5 && life <= 6) r *= 2;
            else if (life > 6) r *= 2.5;
            double DY;
            DateTime p = purchased, f = first_period;
            if (basis == ExcelDCDBasis.Actual_Actual || basis == ExcelDCDBasis.Actual_365)
            {
                if (IsEndOfFeb(p) && p.Day == 29) p = p.AddDays(-1);
                if (IsEndOfFeb(f) && f.Day == 29) f = f.AddDays(-1);
            }
            if (basis == ExcelDCDBasis.Actual_Actual) DY = DateTime.IsLeapYear(purchased.Year) ? 366 : 365;
            else DY = DAYSINYEAR(p, p, basis);
            double days = DAYCOUNT(p, f, basis);
            if (days == 0)
            {
                days = DY;
                life--;
            }
            double d0 = Math.Min(r * cost * days / DY, cost);
            result = Math.Round(Math.Round(d0, 13, MidpointRounding.AwayFromZero), MidpointRounding.AwayFromZero);
            double value = cost - result;
            double d = cost * r;
            for (int i = 1; i <= period; i++)
            {
                if (life - i == 1) d = value;
                else
                {
                    if (value <= salvage) return 0;
                    if (life - i == 2) d = value * 0.5;
                    else d = value * r;
//                    if (value - d <= salvage) d = value - salvage;
                }
                value -= d;
                result = Math.Round(Math.Round(d, 13, MidpointRounding.AwayFromZero), MidpointRounding.AwayFromZero);
            }
            if (result > cost - salvage) result = cost - salvage;
            return result;
        }

        #endregion

        #region payment

        public static double IPMT(double rate, double per, double nper, double pv, double fv, bool pay_in_advance)
        {

            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("IPMT: The number of periods must be greater than zero (got {0:G}).", nper));
            if (per <= 0 || per > nper) throw new System.ArgumentException(string.Format("IPMT: The period specified must be between 1 and the number of periods (got {0:G} / {1:G}).", per, nper));
            if (fv == 0 && pv == 0) throw new System.ArgumentException(string.Format("IPMT: At least one of present value, future value must be non-zero (got 0 & 0)."));
            if (rate < 0 && Math.Floor(nper) != nper) throw new System.ArgumentException(string.Format("IPMT: If the rate is negative the number of periods must be an integer (got {0:G}).", nper));
            if (rate < 0 && Math.Floor(per) != per) throw new System.ArgumentException(string.Format("IPMT: If the rate is negative the period specified must be an integer (got {0:G}).", per));

            double fpmt = PMT(rate, nper, pv, fv, pay_in_advance);
            if (pay_in_advance && per == 1) return 0;            //First payment is 100% repayment if payment is in advance
            double interest = 0.0;
            double repayment = 0.0;
            double bp = pay_in_advance ? fv / (1 + rate) : fv;      //Gross down balloon payment for pay_in_advance because it's paid 1 period later than final payment
            double bpp = bp * rate;                                 //Fixed monthly interest on balloon payment portion
            double mp = pay_in_advance ? pv + fpmt : pv;            //If pay in advance reduce main portion of payment by 1 months' payment
            mp += bp;                                               //Amount to pay off is net of balloon payment (which is negative)

            double cap = pay_in_advance ? per - 1 : per;
            for (int i = 1; i <= cap; i++)
            {
                interest = mp * rate;
                repayment = fpmt + interest - bpp;
                mp += repayment;
            }
            return bpp - interest;
        }

        public static double ISPMT(double rate, double per, double nper, double pv)
        {
            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("ISPMT: The number of periods must be greater than zero (got {0:G}).", nper));
            if (per < 0 || per > nper) throw new System.ArgumentException(string.Format("ISPMT: The period specified must be between 0 and the number of periods (got {0:G} / {1:G}).", per, nper));

            double repayment = pv / nper;
            double principle = pv - (repayment * per);
            double interest = principle * rate;
            return -interest;
        }

        public static double PMT(double rate, double nper, double pv, double fv, bool pay_in_advance)
        {
            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("PMT: The number of periods must be greater than zero (got {0:G}).", nper));
            if (fv == 0 && pv == 0) throw new System.ArgumentException(string.Format("PMT: At least one of present value, future value must be non-zero (got 0 & 0)."));
            if (rate < 0 && Math.Floor(nper) != nper) throw new System.ArgumentException(string.Format("PMT: If the rate is negative the number of periods must be an integer (got {0:G}).", nper));

            double bp = pay_in_advance ? fv / (1 + rate) : fv;
            double bpp = bp * rate;
            double n = pay_in_advance ? nper - 1 : nper;
            double F = 1 - Math.Pow(1 + rate, -n);
            double DN = pay_in_advance ? F + rate : F;
            double retval = -(rate * pv + bpp * (1 - F)) / DN;
            return retval;
        }

        public static double PPMT(double rate, double per, double nper, double pv, double fv, bool pay_in_advance)
        {
            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("PPMT: The number of periods must be greater than zero (got {0:G}).", nper));
            if (per <= 0 || per > nper) throw new System.ArgumentException(string.Format("PPMT: The period specified must be between 1 and the number of periods (got {0:G} / {1:G}).", per, nper));
            if (fv == 0 && pv == 0) throw new System.ArgumentException(string.Format("PPMT: At least one of present value, future value must be non-zero (got 0 & 0)."));
            if (rate < 0 && Math.Floor(nper) != nper) throw new System.ArgumentException(string.Format("PPMT: If the rate is negative the number of periods must be an integer (got {0:G}).", nper));
            if (rate < 0 && Math.Floor(per) != per) throw new System.ArgumentException(string.Format("PPMT: If the rate is negative the period specified must be an integer (got {0:G}).", per));

            double fpmt = PMT(rate, nper, pv, fv, pay_in_advance);
            if (pay_in_advance && per == 1) return fpmt;            //First payment is 100% repayment if payment is in advance
            double interest = 0.0;
            double repayment = 0.0;
            double bp = pay_in_advance ? fv / (1 + rate) : fv;      //Gross down balloon payment for pay_in_advance because it's paid 1 period later than final payment
            double bpp = bp * rate;                                 //Fixed monthly interest on balloon payment portion
            double mp = pay_in_advance ? pv + fpmt : pv;            //If pay in advance reduce main portion of payment by 1 months' payment
            mp += bp;                                               //Amount to pay off is net of balloon payment (which is negative)

            double cap = pay_in_advance ? per - 1 : per;
            for (int i = 1; i <= cap; i++)
            {
                interest = mp * rate;
                repayment = fpmt + interest - bpp;
                mp += repayment;
            }
            return repayment;
        }

        public static double NPER(double rate, double pmt, double pv, double fv, bool pay_in_advance)
        {
            double t1, t2, retval;
            double bp = fv;                             //Balloon payment at end = future value (usually negative)
            double mp = pv;                             //Main portion of payment - will subtract adjusted ballon payment later

            if (pay_in_advance)
            {
                bp /= 1.0 + rate;                       //Reduce the balloon payment because it it paid 1 period later than the final installment
                mp += pmt;                              //Reduce the outstanding balance by 1 months payment because it is paid in advance (no accrued interest)
            }
            mp += bp;                                   //Main portion to pay off is net of adjusted balloon payment (remember bp is usually negative)

            double bpp = bp * rate;                     //Per-period fixed amount to account for interest on balloon payment
            double mpp = pmt - bpp;                     //Net payment applied to declining part of loan

            t2 = Math.Log(rate + 1.0);
            t1 = -Math.Log(1.0 - rate * mp / -mpp);
            retval = t1 / t2;
            if (pay_in_advance) retval += 1;
            return retval;
        }

        public static double CUMIPMT(double rate, double nper, double pv, double start_period, double end_period, bool pay_in_advance)
        {
            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("CUMIPMT: The number of periods must be greater than zero (got {0:G}).", nper));
            if (start_period > end_period) throw new System.ArgumentException(string.Format("CUMIPMT: Start period must be less than or equal to end period (got {0:G} -> {1:G}).", start_period, end_period));
            if (start_period <= 0) throw new System.ArgumentException(string.Format("CUMIPMT: Start period must be greater than zero (got {0:G}).", start_period));
            if (end_period > nper) throw new System.ArgumentException(string.Format("CUMIPMT: End period must be less than or equal to number of periods (got {0:G} <= {1:G}).", end_period, nper));
            if (rate <= 0) throw new System.ArgumentException(string.Format("CUMIPMT: Rate cannot be negative (got {0:G}).", rate));
            if (pv <= 0) throw new System.ArgumentException(string.Format("CUMIPMT: Present value cannot be negative (got {0:G}).", pv));

            double result;
            double pmt = PMT(rate, nper, pv, 0, pay_in_advance);
            if(pay_in_advance)
                result = -pmt * (end_period - start_period + 1) - (start_period > 1 ? FV(rate, start_period - 2, -pmt, -(pv + pmt), false) : pv) + Math.Max(0, FV(rate, end_period - 1, -pmt, -(pv + pmt), false));
            else
                result = -pmt * (end_period - start_period + 1) - FV(rate, start_period - 1, -pmt, -pv, false) + Math.Max(0, FV(rate, end_period, -pmt, -pv, false));
            return -result;
        }

        public static double CUMPRINC(double rate, double nper, double pv, double start_period, double end_period, bool pay_in_advance)
        {
            //Check parameters
            if (nper <= 0) throw new System.ArgumentException(string.Format("CUMPRINC: The number of periods must be greater than zero (got {0:G}).", nper));
            if (start_period > end_period) throw new System.ArgumentException(string.Format("CUMPRINC: Start period must be less than or equal to end period (got {0:G} -> {1:G}).", start_period, end_period));
            if (start_period <= 0) throw new System.ArgumentException(string.Format("CUMPRINC: Start period must be greater than zero (got {0:G}).", start_period));
            if (end_period > nper) throw new System.ArgumentException(string.Format("CUMPRINC: End period must be less than or equal to number of periods (got {0:G} <= {1:G}).", end_period, nper));
            if (rate <= 0) throw new System.ArgumentException(string.Format("CUMPRINC: Rate cannot be negative (got {0:G}).", rate));
            if (pv <= 0) throw new System.ArgumentException(string.Format("CUMPRINC: Present value cannot be negative (got {0:G}).", pv));

            double result;
            double pmt = PMT(rate, nper, pv, 0, pay_in_advance);
            if (pay_in_advance)
                result = (start_period > 1 ? FV(rate, start_period - 2, -pmt, -(pv + pmt), false) : pv) - Math.Max(0, FV(rate, end_period - 1, -pmt, -(pv + pmt), false));
            else
                result = FV(rate, start_period - 1, -pmt, -pv, false) - Math.Max(0, FV(rate, end_period, -pmt, -pv, false));
            return -result;
        }

        public static double EFFECT(double nominal_rate, double npery)
        {
            if (nominal_rate <= 0) throw new System.ArgumentException(string.Format("EFFECT: Nominal rate must be greater than zero (got {0:G}).", nominal_rate));
            if (npery <= 0) throw new System.ArgumentException(string.Format("EFFECT: Number of periods must be 1 or more (got {0:G}).", npery));

            double result = Math.Pow(nominal_rate / npery + 1, npery) - 1;
            return result;
        }

        public static double NOMINAL(double effect_rate, double npery)
        {
            if (effect_rate <= 0) throw new System.ArgumentException(string.Format("NOMINAL: Effective rate must be greater than zero (got {0:G}).", effect_rate));
            if (npery <= 0) throw new System.ArgumentException(string.Format("NOMINAL: Number of periods must be 1 or more (got {0:G}).", npery));

            double result = (Math.Pow(effect_rate + 1, 1.0 / npery) - 1) * npery;
            return result;
        }

        #endregion 

        #region future value

        public static double FV(double rate, double nper, double pmt, double pv, bool pay_in_advance)
        {
            if (pmt == 0 && pv == 0) throw new System.ArgumentException(string.Format("FV: At least one of present value, payment must be non-zero (got 0 & 0)."));
            if (rate < 0 && Math.Floor(nper) != nper) throw new System.ArgumentException(string.Format("FV: If the rate is negative the number of periods must be an integer (got {0:G}).", nper));
            if (rate == -1 && nper <= 0) throw new System.ArgumentException(string.Format("FV: Rate cannot be -100% when number of periods is <= 0."));

            double n = pay_in_advance ? nper + 1 : nper;
            double G = Math.Pow(1 + rate, nper);
            double H = Math.Pow(1 + rate, n);
            double retval = pmt * (H - 1) / rate + G * pv;
            if (pay_in_advance) retval -= pmt;
            return -retval;
        }

        public static double PV(double rate, double nper, double pmt, double fv, bool pay_in_advance)
        {
            if (pmt == 0 && fv == 0) throw new System.ArgumentException(string.Format("PV: At least one of future value, payment must be non-zero (got 0 & 0)."));
            if (rate < 0 && Math.Floor(nper) != nper) throw new System.ArgumentException(string.Format("PV: If the rate is negative the number of periods must be an integer (got {0:G}).", nper));
            if (rate == -1) throw new System.ArgumentException(string.Format("PV: Rate cannot be -100%."));

            double n = pay_in_advance ? nper - 1 : nper;
            double F = 1 - Math.Pow(1 + rate, -n);
            double bp = pay_in_advance ? fv / (1 + rate) : fv;
            double t1 = pmt * F / rate;
            double t2 = (1 - F) * bp;
            double t3 = pay_in_advance ? pmt : 0;
            double retval = t1 + t2 + t3;
            return -retval;
        }

        private class rate_iterand : CSBoost.XMath.iterand<double, double>
        {
            private int np;
            private double pmt, pv, fv;
            private bool pia;

            public rate_iterand(int nper, double _pmt, double _pv, double _fv, bool pay_in_advance)
            {
                np = nper;
                pmt = _pmt;
                pv = _pv;
                fv = _fv;
                pia = pay_in_advance;
            }

            public override double next(double rate)
            {
                return fv - FV(rate, np, pmt, pv, pia);
            }
        }

        public static double RATE(int nper, double pmt, double pv, double fv, bool pay_in_advance, double guess)
        {
            if (pmt == 0 && pv == 0) throw new Exception(string.Format("RATE: Either payment or present value (or both) must be non-zero (got {0:G} and {1:G}).", pmt, pv));
            int max_iter = 100;
            rate_iterand ri = new rate_iterand(nper, pmt, pv, fv, pay_in_advance);
            double ax = guess - 0.01, bx = guess + 0.01;
            if (ax <= -1) ax = -1 + CSBoost.XMath.epsilon;
            double fax = ri.next(ax), fbx = ri.next(bx);
            bool holda = false, holdb = false;
            double olda, oldb;
            while (Math.Sign(fax) == Math.Sign(fbx))
            {
                if (max_iter-- <= 0) throw new Exception("RATE: No solution found - try a different initial guess");

                olda = ax;
                oldb = bx;
                if (!holda)
                {
                    ax -= 1.6 * (bx - ax);
                    fax = ri.next(ax);
                    if (double.IsNaN(fax) || double.IsInfinity(fax))
                    {
                        ax = olda;
                        holda = true;
                    }
                }
                if (!holdb)
                {
                    bx += 1.6 * (bx - olda);    //in case ax is already changed
                    fbx = ri.next(bx);
                    if (double.IsNaN(fbx) || double.IsInfinity(fbx))
                    {
                        bx = oldb;
                        holdb = true;
                    }
                }
            }
            max_iter = 100;
            CSBoost.XMath.pair<double> respair = CSBoost.XMath.toms748_solve(ri, ax, bx, fax, fbx, ep, ref max_iter);
            if (max_iter == 100 && Math.Abs(respair.v1 - respair.v2) > 2 * CSBoost.XMath.epsilon) throw new Exception("RATE: No solution found - try a different initial guess");
            return (respair.v1 + respair.v2) / 2.0;
        }

        //Aggregates
        public static double FVSCHEDULE(double principal, List<double> rates)
        {
            double p = principal;
            foreach (double d in rates) p *= 1.0 + d;
            return p;
        }

        public static double IRR(List<double> flows)
        {
            if (flows.Count == 0) throw new Exception("IRR: No cashflows supplied");
            //Bug in SQL2005 - lambda expressions not allowed
            //int a = flows.FindAll(p => p < 0).Count;
            //int b = flows.FindAll(p => p > 0).Count;
            int a = 0, b = 0;
            foreach (double p in flows) if (p < 0) a++; else b++;
            if (a == 0 || b == 0) throw new Exception("IRR: list of cashflows must contain at least one positive and one negative cashflow.");
            int max_iter = 1000;
            double guess = 0.1;
            irr_iterand xi = new irr_iterand(flows);
            double ax = guess - 0.1, bx = guess + 0.1;
            double fax = xi.next(ax), fbx = xi.next(bx);
            while (Math.Sign(fax) == Math.Sign(fbx))
            {
                if (Math.Sign(fax) > 0)
                {
                    ax -= 0.1;
                    fax = xi.next(ax);
                }
                else
                {
                    bx += 0.1;
                    fbx = xi.next(bx);
                }
                if (double.IsInfinity(fax) || double.IsNaN(fax) || double.IsInfinity(fbx) || double.IsNaN(fbx)) throw new Exception("IRR: Can't find suitable initial conditions");
                if (--max_iter <= 0) throw new Exception("IRR: Can't find suitable initial conditions");
            }
            max_iter = 100;
            CSBoost.XMath.pair<double> respair = CSBoost.XMath.toms748_solve(xi, ax, bx, fax, fbx, ep, ref max_iter);
            if (max_iter == 100 && Math.Abs(respair.v1 - respair.v2) > 2 * CSBoost.XMath.epsilon) throw new Exception("IRR: No solution found after 100 iterations");
            return (respair.v1 + respair.v2) / 2.0;
        }

        public static double MIRR(List<double> flows, double finance_rate, double reinvest_rate)
        {
            if(finance_rate == -1) throw new Exception("MIRR: Finance rate cannot be -100%.");
            if(reinvest_rate == -1) throw new Exception("MIRR: Reinvest rate cannot be -100%.");
            if(flows.Count < 2) throw new Exception(string.Format("MIRR: Insufficient cash flows supplied (must be > 1 - got {0:G}).", flows.Count));
            //Bug in SQL2005 - lambda expressions not allowed
            //int a = flows.FindAll(p => p < 0).Count;
            //int b = flows.FindAll(p => p > 0).Count;
            int a = 0, b = 0;
            foreach (double p in flows) if (p < 0) a++; else b++;
            if (a == 0 || b == 0) throw new Exception("MIRR: list of cashflows must contain at least one positive and one negative cashflow.");
            double fv = 0.0;
            double pv = 0.0;
            int n = flows.Count;
            double d;
            for (int i = 0; i < n; i++)
            {
                d = flows[i];
                if (d < 0) pv += d / Math.Pow(1 + finance_rate, i);
                else fv += d * Math.Pow(1 + reinvest_rate, n - i - 1);
            }
            double retval = Math.Pow(-fv / pv, 1.0 / (n - 1)) - 1;
            return retval;
        }

        public static double NPV(List<double> flows, double rate)
        {
            if (rate == -1) throw new Exception("NPV: Rate cannot be -100%.");
            double val = 0.0;
            for (int i = 1; i <= flows.Count; i++)
            {
                val += flows[i-1] / Math.Pow(1 + rate, i);
            }
            return val;
        }

        public static double XIRR(List<cashflow> flows)
        {
            if (flows.Count == 0) throw new Exception("XIRR: No cashflows supplied");
            //Bug in SQL2005 - lambda expressions not allowed
            //int a = flows.FindAll(p => Math.Sign(p.payment) < 0).Count;
            //int b = flows.FindAll(p => Math.Sign(p.payment) > 0).Count;
            int a = 0, b = 0;
            foreach (cashflow p in flows) if (p.payment < 0) a++; else b++;
            if (a == 0 || b == 0) throw new Exception("XIRR: list of cashflows must contain at least one positive and one negative cashflow.");
            int max_iter = 1000;
            double guess = 0.1;
            xirr_iterand xi = new xirr_iterand(flows);
            double ax = guess - 0.1, bx = guess + 0.1;
            double fax = xi.next(ax), fbx = xi.next(bx);
            while (Math.Sign(fax) == Math.Sign(fbx))
            {
                if (Math.Sign(fax) < 0)
                {
                    ax -= 0.1;
                    fax = xi.next(ax);
                }
                else
                {
                    bx += 0.1;
                    fbx = xi.next(bx);
                }
                if (double.IsInfinity(fax) || double.IsNaN(fax) || double.IsInfinity(fbx) || double.IsNaN(fbx)) throw new Exception("XIRR: Can't find suitable initial conditions");
                if (--max_iter <= 0) throw new Exception("XIRR: Can't find suitable initial conditions");
            }
            max_iter = 100;
            CSBoost.XMath.pair<double> respair = CSBoost.XMath.toms748_solve(xi, ax, bx, fax, fbx, ep, ref max_iter);
            if (max_iter == 100 && Math.Abs(respair.v1 - respair.v2) > 2 * CSBoost.XMath.epsilon) throw new Exception("XIRR: No solution found after 100 iterations");
            return (respair.v1 + respair.v2) / 2.0;
        }

        public static double XNPV(List<cashflow> flows, double rate)
        {
            if (rate == -1) throw new Exception("XNPV: Rate cannot be -100%.");
            cashflow P0 = flows[0];
            double result = 0;
            foreach (cashflow cf in flows)
            {
                result += cf.payment / Math.Pow(1 + rate, cf.date.Subtract(P0.date).Days / 365.0);
            }
            return result;
        }

        //Internal helper functions
        public struct cashflow : IComparable
        {
            public cashflow(double p, DateTime d)
            {
                payment = p;
                date = d;
            }

            public double payment;
            public DateTime date;

            public int CompareTo(object obj)
            {
                cashflow f = (cashflow)obj;
                return date.CompareTo(f.date);
            }
        }

        public class irr_iterand : CSBoost.XMath.iterand<double, double>
        {
            private List<double> clist;

            public irr_iterand(List<double> theList)
            {
                clist = theList;
            }

            public override double next(double rate)
            {
                return NPV(clist, rate);
            }
        }

        public class xirr_iterand : CSBoost.XMath.iterand<double, double>
        {
            private List<cashflow> clist;

            public xirr_iterand(List<cashflow> theList)
            {
                clist = theList;
            }

            public override double next(double rate)
            {
                return XNPV(clist, rate);
            }
        }

        #endregion

        #region securities

        public enum ExcelDCDBasis {US30_360 = 0, Actual_Actual = 1, Actual_360 = 2,  Actual_365 = 3,  EU30_360 = 4};

        public static double US360_daycount(DateTime date1, DateTime date2)
        {
            if (date1 == date2) return 0;

            int d1 = date1.Day, xd1 = d1, m1 = date1.Month, y1 = date1.Year;
            int d2 = date2.Day, m2 = date2.Month, y2 = date2.Year;

            if (IsEndOfFeb(date2)) d2 = 30;
            if (d2 == 31) d2 = 30;
            if (d1 == 31) d1 = 30;
            if (IsEndOfFeb(date1)) d1 = 30;
            int N = (d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1);
            return N;
        }

        public static double DAYCOUNT(DateTime date1, DateTime date2, ExcelDCDBasis basis)
        {
            if (date1 == date2) return 0;

            int d1 = date1.Day, xd1 = d1, m1 = date1.Month, y1 = date1.Year;
            int d2 = date2.Day, m2 = date2.Month, y2 = date2.Year;

            // Calculate no. interest-bearing days
            int N;
            switch (basis)
            {
                case ExcelDCDBasis.US30_360:
                    //new version
                    if (d2 == 31 && (d1 >= 30)) d2 = 30;
                    if (d1 == 31) d1 = 30;
                    if (IsEndOfFeb(date2) && (IsEndOfFeb(date1))) d2 = 30;
                    if (IsEndOfFeb(date1)) d1 = 30;
                    N = (d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1);
                    break;
                case ExcelDCDBasis.Actual_Actual:
                case ExcelDCDBasis.Actual_360:
                case ExcelDCDBasis.Actual_365:
                    N = (date2.Subtract(date1)).Days;
                    break;
                case ExcelDCDBasis.EU30_360:
                    if (d1 == 31) d1 = 30;
                    if (d2 == 31) d2 = 30;
                    N = (d2 - d1) + 30 * (m2 - m1) + 360 * (y2 - y1);
                    break;
                default:
                    throw new System.ArgumentException("Basis must be 0,1,2,3 or 4");
            }
            return N;
        }

        public static double DAYSINYEAR(DateTime date1, DateTime date2, ExcelDCDBasis basis)
        {
            double N = 0.0;
            switch (basis)
            {
                case ExcelDCDBasis.US30_360:
                case ExcelDCDBasis.EU30_360:
                case ExcelDCDBasis.Actual_360:
                    N = 360; 
                    break;
                case ExcelDCDBasis.Actual_365:
                    N = 365; 
                    break;
                case ExcelDCDBasis.Actual_Actual:
                    if (LessThanOrEqualToOneYearApart(date1, date2))
                    {
                        if (date1.Year == date2.Year && DateTime.IsLeapYear(date1.Year)) N = 366;
                        else if (DateTime.IsLeapYear(date1.Year) && date1.Month <= 2) N = 366;
                        else if (DateTime.IsLeapYear(date2.Year) && date2.Month > 2) N = 366;
                        else if (DateTime.IsLeapYear(date2.Year) && IsEndOfFeb(date2)) N = 366;
                        else N = 365;
                    }
                    else
                    {
                        int days = 0;
                        int years = 0;
                        for (int y = date1.Year; y <= date2.Year; y++)
                        {
                            years++;
                            days += (int)(new DateTime(y + 1, 1, 1).Subtract(new DateTime(y, 1, 1)).TotalDays);
                        }
                        N = (double)days / (double)years;
                    }
                    //Original version
                    //for (i = date1.Year; i <= date2.Year; i++)
                    //{
                    //    if (DateTime.IsLeapYear(i))
                    //    {
                    //        if (i == date1.Year && date1.DayOfYear >= LeapDay) N += 365;
                    //        else if (i == date2.Year && date2.DayOfYear < LeapDay) N += 365;
                    //        else N += 366;
                    //    }
                    //    else N += 365;
                    //}
                    //N /= (i - date1.Year) + 1.0;
                    break;
                default:
                    throw new System.ArgumentException("Basis must be 0,1,2,3 or 4");
            }
            return N;
        }

        struct period:System.IComparable
        {
            public DateTime start;
            public DateTime end;
            public double days;

            public period(DateTime st, DateTime nd, ExcelDCDBasis bs)
            {
                start = st;
                end = nd;
//                days = (int)DAYCOUNT(start, end, bs);
                switch (bs)
                {
                    case ExcelDCDBasis.Actual_Actual:
                        days = end.Subtract(start).Days;
                        break;
                    case ExcelDCDBasis.Actual_365:
                        //if (end.Subtract(start).Days >= 365) days = 365;
                        //else days = end.Subtract(start).Days;
                        if (end.Subtract(start).Days >= 360) days = 365;
                        else if (end.Subtract(start).Days >= 170) days = 182.5;
                        else if (end.Subtract(start).Days >= 80) days = 91.25;
                        else days = 0;
                        break;
                    case ExcelDCDBasis.Actual_360:
                        //if (end.Subtract(start).Days >= 365) days = 360;
                        //else days = end.Subtract(start).Days;
                        //break;
                    case ExcelDCDBasis.EU30_360:
                    case ExcelDCDBasis.US30_360:
                        if (end.Subtract(start).Days >= 360) days = 360;
                        else if (end.Subtract(start).Days >= 170) days = 180;
                        else if (end.Subtract(start).Days >= 80) days = 90;
                        else days = 0;
                        break;
                    default:
                        throw new System.ArgumentException("Basis must be 0,1,2,3 or 4");
                }
            }

            public int CompareTo(Object o)
            {
                period p = (period)o;
                if (p.start < this.start) return 1;
                if (p.start == this.start) return 0;
                return -1;
            }
        }

        private static double accrint_factor(DateTime pcd, DateTime ncd, DateTime issue, byte frequency, ExcelDCDBasis basis, bool calc_method)
        {
            DateTime d1 = issue > pcd ? issue : pcd;
            double days, coupdays;
            if (basis == ExcelDCDBasis.US30_360 && issue <= pcd) days = US360_daycount(d1, ncd);
            else days = DAYCOUNT(d1, ncd, basis);
            switch (basis)
            {
                case ExcelDCDBasis.US30_360: coupdays = US360_daycount(pcd, ncd); break;
                case ExcelDCDBasis.Actual_365: coupdays = 365.0 / frequency; break;
                case ExcelDCDBasis.Actual_360: coupdays = 360.0 / frequency; break;
                default: coupdays = DAYCOUNT(pcd, ncd, basis); break;
            }
            if (issue <= pcd) return calc_method ? 1 : 0;
            else return days / coupdays;
        }

        public static double ACCRINT(DateTime issue, DateTime first_interest, DateTime settlement, double rate, double par, byte frequency, ExcelDCDBasis basis, bool calc_method)
        {
            if (settlement <= issue) throw new ArgumentException(string.Format("ACCRINT: Settlement date must be later than issue date (got {0:G} and {1:G} respectively).", settlement, issue));
            if (rate <= 0) throw new ArgumentException(string.Format("ACCRINT: Rate must be greater than zero (got {0:G}).", rate));
            if (par <= 0) throw new ArgumentException(string.Format("ACCRINT: Par value must be greater than zero (got {0:G}).", par));
            check_frequency(frequency, "ACCRINT");

            DateTime d1, d2;
            int m = 12 / frequency;
            DateTime pcd, ncd;
            if (settlement > first_interest && calc_method)
            {
                pcd = first_interest;
                ncd = pcd;
                while (ncd < settlement)
                {
                    pcd = ncd;
                    ncd = pcd.AddMonths(m);
                    if (IsEndOfMonth(first_interest)) ncd = EndOfMonth(ncd);
                }
            }
            else
            {
                pcd = first_interest.AddMonths(-m);
                if (IsEndOfMonth(first_interest)) pcd = EndOfMonth(pcd);
            }
            d1 = issue > pcd ? issue : pcd;
            double days = DAYCOUNT(d1, settlement, basis);
            double coupdays;
            if(pcd >= first_interest && basis == ExcelDCDBasis.Actual_Actual)
            {
                DateTime ppcd = first_interest.AddMonths(-m);
                if (IsEndOfMonth(first_interest) ) ppcd = EndOfMonth(ppcd);
                coupdays = first_interest.Subtract(ppcd).Days;
            }
            else coupdays = COUPDAYS(pcd, first_interest, frequency, basis);
            double a = days / coupdays;
            d1 = pcd; 
            while (d1 > issue)
            {
                d2 = d1;
                d1 = d1.AddMonths(-m);
                if (IsEndOfMonth(first_interest)) d1 = EndOfMonth(d1);
                a += accrint_factor(d1, d2, issue, frequency, basis, calc_method);
            }
            double result = a * par * rate / frequency;
            return result;
        }

        public static double ACCRINTM(DateTime issue, DateTime settlement, double rate, double par, ExcelDCDBasis basis)
        {
            if (settlement <= issue) throw new ArgumentException(string.Format("ACCRINTM: Settlement date must be later than issue date (got {0:G} and {1:G} respectively).", settlement, issue));
            if (rate <= 0) throw new ArgumentException(string.Format("ACCRINTM: Rate must be greater than zero (got {0:G}).", rate));
            if (par <= 0) throw new ArgumentException(string.Format("ACCRINTM: Par value must be greater than zero (got {0:G}).", par));

            double A = DAYCOUNT(issue, settlement, basis);
            double D = DAYSINYEAR(issue, settlement, basis);
            double retval = par * rate * A / D;
            return retval;
        }

        public static double COUPDAYS(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
//            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPDAYS: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            check_frequency(frequency, "COUPDAYS");

            switch (basis)
            {
                case ExcelDCDBasis.US30_360:
                case ExcelDCDBasis.EU30_360:
                case ExcelDCDBasis.Actual_360:
                    return 360 / frequency;
                case ExcelDCDBasis.Actual_365:
                    return 365.0 / (double)frequency;
                case ExcelDCDBasis.Actual_Actual:

                    int m = 12 / frequency;
                    DateTime d1, d2;

                    d1 = maturity;

                    do
                    {
                        d2 = d1;
                        d1 = d2.AddMonths(-m);
                        if (IsEndOfMonth(d2) && settlement >= d1) d1 = EndOfMonth(d1);
                    } while (settlement < d1);

                    return d2.Subtract(d1).Days;
            }
            return 0;
        }

        public static double COUPDAYSBS(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
//            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPDAYSBS: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            check_frequency(frequency, "COUPDAYSBS");

            DateTime d1;
            d1 = COUPPCD(settlement, maturity, frequency, basis);
            return DAYCOUNT(d1, settlement, basis);
        }

        public static double COUPDAYSNC(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
//            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPDAYSNC: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            check_frequency(frequency, "COUPDAYSNC");

            switch (basis)
            {
                case ExcelDCDBasis.US30_360 :
                    return COUPDAYS(settlement, maturity, frequency, basis) - COUPDAYSBS(settlement, maturity, frequency, basis);
                case ExcelDCDBasis.EU30_360 :
                    return DAYCOUNT(settlement, COUPNCD(settlement, maturity, frequency, basis), basis);
                case ExcelDCDBasis.Actual_Actual:
                case ExcelDCDBasis.Actual_360 :
                case ExcelDCDBasis.Actual_365 :
                    return COUPNCD(settlement, maturity, frequency, basis).Subtract(settlement).Days;
                default:
                    throw new System.ArgumentException("Basis must be 0,1,2,3 or 4");
            }
        }

        public static DateTime COUPNCD(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
//            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPNCD: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            check_frequency(frequency, "COUPNCD");

            int m = 12 / frequency;
            DateTime d1, d2;

            d1 = maturity;
            do
            {
                d2 = d1;
                d1 = d2.AddMonths(-m);
                if (IsEndOfMonth(maturity)) d1 = EndOfMonth(d1);
                while (maturity.Day > d1.Day && !IsEndOfMonth(d1)) d1 = d1.AddDays(1);
            } while (d1 > settlement);

            return d2;
        }

        public static int COUPNUM(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
//            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPNUM: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            check_frequency(frequency, "COUPNUM");

            int m = 12 / frequency;
            DateTime d1, d2;
            int i = 0;

            d1 = maturity;
            do
            {
                d2 = d1;
                d1 = d2.AddMonths(-m);
                if (IsEndOfMonth(maturity)) d1 = EndOfMonth(d1);
                i++;
            } while (d1 > settlement);

            return i;
        }

        public static DateTime COUPPCD(DateTime settlement, DateTime maturity, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("COUPPCD: Settlement must be before maturity (got {0:G} and {1:G} respectively", settlement, maturity));
            check_frequency(frequency, "COUPPCD");

            int m = 12 / frequency;
            DateTime d1, d2;

            d1 = maturity;
            do
            {
                d2 = d1;
                d1 = d2.AddMonths(-m);
                if (IsEndOfMonth(maturity)) d1 = EndOfMonth(d1);
                while (maturity.Day > d1.Day && !IsEndOfMonth(d1)) d1 = d1.AddDays(1);
            } while (d1 > settlement);

            return d1;
        }

        public static double DISC(DateTime settlement, DateTime maturity, double price, double redemption, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("DISC: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (price <= 0) throw new System.ArgumentException(string.Format("DISC: Price must be greater than zero (got {0:G}).", price));
            if (redemption <= 0) throw new System.ArgumentException(string.Format("DISC: Redemption must be greater than zero (got {0:G}).", redemption));

            double B = DAYSINYEAR(settlement, maturity, basis);
            double DSM = DAYCOUNT(settlement, maturity, basis);
            double result = (redemption - price) / redemption;
            result *= B / DSM;
            return result;
        }

        public static double PRICE(DateTime settlement, DateTime maturity, double rate, double yld, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("PRICE: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (rate < 0) throw new System.ArgumentException(string.Format("PRICE: Rate must be >= zero (got {0:G}).", rate));
            if (yld < 0) throw new System.ArgumentException(string.Format("PRICE: Yield must be >= zero (got {0:G}).", yld));
            if (redemption < 0) throw new System.ArgumentException(string.Format("PRICE: Redemption must be >= zero (got {0:G}).", redemption));
            check_frequency(frequency, "PRICE");
            return price(settlement, maturity, rate, yld, redemption, frequency, basis);
        }

        private static double price(DateTime settlement, DateTime maturity, double rate, double yld, double redemption, byte frequency, ExcelDCDBasis basis)
        {

            DateTime ncd = COUPNCD(settlement, maturity, frequency, basis);
            DateTime pcd = COUPPCD(settlement, maturity, frequency, basis);
            double E = basis == ExcelDCDBasis.Actual_Actual ? DAYCOUNT(pcd, ncd, basis) : COUPDAYS(settlement, maturity, frequency, basis);
            double N = COUPNUM(settlement, maturity, frequency, basis);
            double A = COUPDAYSBS(settlement, maturity, frequency, basis);
            double DSC = E - A;
            double price;
            if (N == 1.0)
            {
                price = (redemption + 100 * rate / frequency) / (1.0 + (DSC / E) * yld / frequency);
            }
            else
            {
                price = redemption / Math.Pow(1.0 + yld / frequency, N - 1.0 + DSC / E);
                for (int k = 1; k <= N; k++)
                {
                    price += (100.0 * rate / frequency) / Math.Pow(1.0 + yld / frequency, (double)k - 1.0 + DSC / E);
                }
            }
            price -= (100 * rate / frequency) * (A / E);
            return price;
        }

        public static double PRICEDISC(DateTime settlement, DateTime maturity, double discount, double redemption, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("PRICEDISC: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (discount <= 0) throw new System.ArgumentException(string.Format("PRICEDISC: Discount must be greater than zero (got {0:G}).", discount));
            if (redemption <= 0) throw new System.ArgumentException(string.Format("PRICEDISC: Redemption must be greater than zero (got {0:G}).", redemption));

            double B = DAYSINYEAR(settlement, maturity, basis);
            double DSM = DAYCOUNT(settlement, maturity, basis);
            double result = redemption - discount * redemption * DSM / B;
            return result;
        }

        public static double PRICEMAT(DateTime settlement, DateTime maturity, DateTime issue, double rate, double yld, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("PRICEMAT: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (rate <= 0) throw new System.ArgumentException(string.Format("PRICEMAT: Rate must be greater than zero (got {0:G}).", rate));
            if (yld <= 0) throw new System.ArgumentException(string.Format("PRICEMAT: Yield must be greater than zero (got {0:G}).", yld));

            double B = DAYSINYEAR(issue, settlement, basis);
            double DIM = DAYCOUNT(issue, maturity, basis);
            double A = DAYCOUNT(issue, settlement, basis);
            double DSM = DIM - A;
            double result = (100 + DIM / B * rate * 100) / (1 + DSM / B * yld) - (A / B * rate * 100);
            return result;
        }

        public static double ODDFPRICE(DateTime settlement, DateTime maturity, DateTime issue, DateTime first_interest, double rate, double yld, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity.Month != first_interest.Month || (maturity.Day != first_interest.Day && !(IsEndOfFeb(maturity) && IsEndOfFeb(first_interest)))) throw new System.ArgumentException(string.Format("ODDFPRICE: Maturity and first interest must have the same month and day (got {0:G} and {1:G} respectively).", maturity, first_interest));
            if (maturity < first_interest) throw new System.ArgumentException(string.Format("ODDFPRICE: Maturity must be after first interest payment (got {0:G} and {1:G} respectively).", maturity, first_interest));
            if (first_interest < settlement) throw new System.ArgumentException(string.Format("ODDFPRICE: First interest must be after settlement (got {0:G} and {1:G} respectively).", first_interest, settlement));
            if (settlement < issue) throw new System.ArgumentException(string.Format("ODDFPRICE: Settlement must be after issue (got {0:G} and {1:G} respectively).", settlement, issue));
            if (rate < 0) throw new System.ArgumentException(string.Format("ODDFPRICE: Rate must be >= 0 (got {0:G}).", rate));
            if (yld < 0) throw new System.ArgumentException(string.Format("ODDFPRICE: Yield must be >=0 (got {0:G}).", yld));
            if (redemption < 0) throw new System.ArgumentException(string.Format("ODDFPRICE: Redemption must be >=0 (got {0:G}).", redemption));
            check_frequency(frequency, "ODDFPRICE");
            return oddfprice(settlement, maturity, issue, first_interest, rate, yld, redemption, frequency, basis);
        }

        private static double oddfprice(DateTime settlement, DateTime maturity, DateTime issue, DateTime first_interest, double rate, double yld, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            double price;
            double E = COUPDAYS(settlement, first_interest, frequency, basis);
            double DFC = DAYCOUNT(issue, first_interest, basis);
            double x = 1 + yld / frequency;
            double y = 100 * rate / frequency;
            if (DFC < E)
            {
                double N = COUPNUM(settlement, maturity, frequency, basis);
                double DSC = Math.Max(DAYCOUNT(settlement, first_interest, basis), 0);
                double A = DAYCOUNT(issue, settlement, basis);
                double z = DSC / E;
                double p1 = x;
                double p3 = Math.Pow(p1, N - 1 + z);
                double t1 = redemption / p3;
                double t2 = y * DFC / E / Math.Pow(p1, z);
                double t3 = 0;
                for (int k = 2; k <= N; k++) t3 += y / Math.Pow(p1, k - 1 + z);
                double t4 = y * A / E;
                price = t1 + t2 + t3 - t4;
            }
            else
            {
                double N = COUPNUM(first_interest, maturity, frequency, basis);
                double NC = COUPNUM(issue, first_interest, frequency, basis);
                DateTime ncd = COUPNCD(settlement, first_interest, frequency, basis);
                DateTime pcd = COUPPCD(settlement, first_interest, frequency, basis);
                double DSC;
                if (basis == ExcelDCDBasis.Actual_360 || basis == ExcelDCDBasis.Actual_365)
                    DSC = DAYCOUNT(settlement, ncd, basis);
                else
                    DSC = E - DAYCOUNT(pcd, settlement, basis);
                double Nq = cnum(first_interest, settlement, frequency, basis);
                double t1 = redemption / Math.Pow(x, N + Nq + DSC / E);
                double t2a = 0, t4a = 0;

                ncd = first_interest;
                pcd = ncd.AddMonths(-12 / frequency);

                for (int i = (int)NC; i >= 1; i--)
                {
                    double NLi = basis == ExcelDCDBasis.Actual_Actual ? DAYCOUNT(pcd, ncd, basis) : E;
                    double DCi = i > 1 ? 1 : DAYCOUNT(issue, ncd, basis) / NLi;
                    DateTime d1 = issue > pcd ? issue : pcd;
                    DateTime d2 = settlement < ncd ? settlement : ncd;
                    t2a += DCi;
                    t4a += Math.Max(DAYCOUNT(d1, d2, basis) / NLi, 0);
                    ncd = pcd;
                    pcd = ncd.AddMonths(-12 / frequency);
                    if (ncd < issue) break;
                }
                double t2 = y * t2a / Math.Pow(x, Nq + DSC / E);
                double t4 = y * t4a;
                double t3 = 0;
                for (int i = 1; i <= N; i++) t3 += y / Math.Pow(x, i + Nq + DSC / E);
                price = t1 + t2 + t3 - t4;
            }
            return price;
        }

        public static double ODDLPRICE(DateTime settlement, DateTime maturity, DateTime last_interest, double rate, double yld, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("ODDLPRICE: Maturity must be after settlement (got {0:G} and {1:G} respectively).", maturity, settlement));
            if (settlement < last_interest) throw new System.ArgumentException(string.Format("ODDLPRICE: Settlement must be after last interest payment (got {0:G} and {1:G} respectively).", last_interest, settlement));
            if (rate < 0) throw new System.ArgumentException(string.Format("ODDLPRICE: Rate must be >= 0 (got {0:G}).", rate));
            if (yld < 0) throw new System.ArgumentException(string.Format("ODDLPRICE: Yield must be >= 0 (got {0:G}).", yld));
            if (redemption < 0) throw new System.ArgumentException(string.Format("ODDLPRICE: Redemption must be >= 0 (got {0:G}).", redemption));
            check_frequency(frequency, "ODDLPRICE");

            double NC = COUPNUM(last_interest, maturity, frequency, basis);
            DateTime d1 = last_interest, d2;
            double t1a = 0, t2a = 0, t3a = 0;
            for (int i = 1; i <= NC; i++)
            {
                d2 = d1.AddMonths(12 / frequency);
                double Nl, DCi;
                if (basis == ExcelDCDBasis.US30_360)
                {
                    Nl = US360_daycount(d1, d2);
                    DCi = i < NC ? Nl : US360_daycount(d1, maturity);
                }
                else
                {
                    Nl = DAYCOUNT(d1, d2, basis);
                    DCi = i < NC ? Nl : DAYCOUNT(d1, maturity, basis);
                }
                double Ai = d2 < settlement ? DCi : d1 < settlement ? Math.Max(0, DAYCOUNT(d1, settlement, basis)) : 0;
                DateTime d3 = settlement > d1 ? settlement : d1;
                DateTime d4 = maturity < d2 ? maturity : d2;
                double DSC = Math.Max(0, DAYCOUNT(d3, d4, basis));
                d1 = d2;
                t1a += DCi / Nl;
                t2a += DSC / Nl;
                t3a += Ai / Nl;
            }
            double x = 100 * rate / frequency;
            double t1 = t1a * x + redemption;
            double t2 = t2a * yld / frequency + 1;
            double t3 = t3a * x;
            double result = t1 / t2 - t3;
            return result;
        }

        public static double DURATION(DateTime settlement, DateTime maturity, double rate, double yld, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("DURATION: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (rate < 0) throw new System.ArgumentException(string.Format("DURATION: Rate must be >= 0 (got {0:G}).", rate));
            if (yld < 0) throw new System.ArgumentException(string.Format("DURATION: Yield must be >= 0 (got {0:G}).", yld));
            check_frequency(frequency, "DURATION");

            double FV = 100.0;
            DateTime ncd = COUPNCD(settlement, maturity, frequency, basis);
            DateTime pcd = COUPNCD(settlement, maturity, frequency, basis);
            //double E = basis == ExcelDCDBasis.Actual_Actual ? DAYCOUNT(pcd, ncd, basis) : COUPDAYS(settlement, maturity, frequency, basis);
            double E = COUPDAYS(settlement, maturity, frequency, basis);
            double A = COUPDAYSBS(settlement, maturity, frequency, basis); ;
            double DSC = E - A;

            double m = COUPNUM(settlement, maturity, frequency, basis);
            double a = DSC / E;
            double x = 0.0, y = 0.0;
            for (int i = 1; i <= m; i++)
            {
                double t1 = i - 1 + a;
                double t2 = (FV * rate / frequency) / Math.Pow(1 + yld / frequency, t1);
                x += t2 * t1;
                y += t2;
            }
            double t3 = (a + m - 1) * FV / Math.Pow(1 + yld / frequency, a + m - 1) + x;
            double t4 = FV / Math.Pow(1 + yld / frequency, a + m - 1) + y;
            double DUR = (t3 / t4) / frequency;
            
            return DUR;
        }

        public static double MDURATION(DateTime settlement, DateTime maturity, double rate, double yld, byte frequency, ExcelDCDBasis basis)
        {
            return DURATION(settlement, maturity, rate, yld, frequency, basis) / (1 + yld / frequency);
        }

        private class yielditerand : CSBoost.XMath.iterand<double, double>
        {
            DateTime set, mat;
            double rat, pr, red;
            byte freq;
            ExcelDCDBasis bas;

            public yielditerand(DateTime settlement, DateTime maturity, double rate, double price, double redemption, byte frequency, ExcelDCDBasis basis)
            {
                set = settlement;
                mat = maturity;
                rat = rate;
                pr = price;
                red = redemption;
                freq = frequency;
                bas = basis;
            }

            public override double next(double yield)
            {
                return pr - price(set, mat, rat, yield, red, freq, bas);
            }
        }

        public static double YIELD(DateTime settlement, DateTime maturity, double rate, double pr, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("YIELD: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (rate < 0) throw new System.ArgumentException(string.Format("YIELD: Rate must be >= zero (got {0:G}).", rate));
            if (pr < 0) throw new System.ArgumentException(string.Format("YIELD: Price must be >= zero (got {0:G}).", pr));
            if (redemption < 0) throw new System.ArgumentException(string.Format("YIELD: Redemption must be >= zero (got {0:G}).", redemption));
            check_frequency(frequency, "YIELD");

            int n = COUPNUM(settlement, maturity, frequency, basis);
            double result;
            double A = COUPDAYSBS(settlement, maturity, frequency, basis);
            double E = COUPDAYS(settlement, maturity, frequency, basis);
            double DSR = DAYCOUNT(settlement, maturity, basis);
            double t1 = pr / 100 + (A / E) * (rate / frequency);
            result = (((redemption / 100 + rate / frequency) - t1) / t1) * frequency * E / DSR;
            if (n < 1) return result;
            int max_iter = 100;
            yielditerand yi = new yielditerand(settlement, maturity, rate, pr, redemption, frequency, basis);
            double ax = result - 0.1, bx = result + 0.1;
            double fax = yi.next(ax), fbx = yi.next(bx);
            while (Math.Sign(fax) == Math.Sign(fbx))
            {
                if (Math.Sign(fax) > 0)
                {
                    ax -= 0.1;
                    fax = yi.next(ax);
                }
                else
                {
                    bx += 0.1;
                    fbx = yi.next(bx);
                }
            }
            CSBoost.XMath.pair<double> respair = CSBoost.XMath.toms748_solve(yi, ax, bx, fax, fbx, ep, ref max_iter);
            if (max_iter == 100 && Math.Abs(respair.v1 - respair.v2) > 2 * CSBoost.XMath.epsilon) throw new Exception();
            GC.Collect();
            return (respair.v1 + respair.v2) / 2.0;
        }

        public static double YIELDDISC(DateTime settlement, DateTime maturity, double pr, double redemption, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("YIELDDISC: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (pr <= 0) throw new System.ArgumentException(string.Format("YIELDDISC: Price must be greater than zero (got {0:G}).", pr));
            if (redemption <= 0) throw new System.ArgumentException(string.Format("YIELDDISC: Redemption must be greater than zero (got {0:G}).", redemption));
 

            double result = ((redemption - pr) / pr) * DAYSINYEAR(settlement, maturity, basis) / DAYCOUNT(settlement, maturity, basis);
            return result;
        }

        public static double YIELDMAT(DateTime settlement, DateTime maturity, DateTime issue, double rate, double pr, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("YIELDMAT: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (pr <= 0) throw new System.ArgumentException(string.Format("YIELDMAT: Price must be greater than zero (got {0:G}).", pr));
            if (rate <= 0) throw new System.ArgumentException(string.Format("YIELDMAT: Rate must be greater than zero (got {0:G}).", rate));

            double dcim = DAYCOUNT(issue, maturity, basis);
            double dcis = DAYCOUNT(issue, settlement, basis);
            double dyis = DAYSINYEAR(issue, settlement, basis);
            double dcsm = dcim - dcis;
            return (dcim / dyis * rate + 1 - pr / 100 - dcis / dyis * rate) / (pr / 100 + dcis / dyis * rate) * (dyis / dcsm);
        }

        private class oddfyielditerand : CSBoost.XMath.iterand<double, double>
        {
            DateTime set, mat, iss, fint;
            double rat, pr, red;
            byte freq;
            ExcelDCDBasis bas;

            public oddfyielditerand(DateTime settlement, DateTime maturity, DateTime issue, DateTime first_interest, double rate, double price, double redemption, byte frequency, ExcelDCDBasis basis)
            {
                set = settlement;
                mat = maturity;
                rat = rate;
                pr = price;
                red = redemption;
                freq = frequency;
                bas = basis;
                iss = issue;
                fint = first_interest;
            }

            public override double next(double yield)
            {
                return pr - oddfprice(set, mat, iss, fint, rat, yield, red, freq, bas);
            }
        }

        public static double ODDFYIELD(DateTime settlement, DateTime maturity, DateTime issue, DateTime first_interest, double rate, double pr, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity.Month != first_interest.Month || (maturity.Day != first_interest.Day && !(IsEndOfFeb(maturity) && IsEndOfFeb(first_interest)))) throw new System.ArgumentException(string.Format("ODDFYIELD: Maturity and first interest must have the same month and day (got {0:G} and {1:G} respectively).", maturity, first_interest));
            if (maturity < first_interest) throw new System.ArgumentException(string.Format("ODDFYIELD: Maturity must be after first interest payment (got {0:G} and {1:G} respectively).", maturity, first_interest));
            if (first_interest < settlement) throw new System.ArgumentException(string.Format("ODDFYIELD: First interest must be after settlement (got {0:G} and {1:G} respectively).", first_interest, settlement));
            if (settlement < issue) throw new System.ArgumentException(string.Format("ODDFYIELD: Settlement must be after issue (got {0:G} and {1:G} respectively).", settlement, issue));
            if (rate < 0) throw new System.ArgumentException(string.Format("ODDFYIELD: Rate must be greater than zero (got {0:G}).", rate));
            if (pr < 0) throw new System.ArgumentException(string.Format("ODDFYIELD: Price must be greater than zero (got {0:G}).", pr));
            if (redemption < 0) throw new System.ArgumentException(string.Format("ODDFYIELD: Redemption must be greater than zero (got {0:G}).", redemption));
            check_frequency(frequency, "ODDFYIELD");

            int max_iter = 100;
            oddfyielditerand yi = new oddfyielditerand(settlement, maturity, issue, first_interest, rate, pr, redemption, frequency, basis);
            double ax = 0.4, bx = 0.6;
            double fax = yi.next(ax), fbx = yi.next(bx);
            while (Math.Sign(fax) == Math.Sign(fbx))
            {
                if (Math.Sign(fax) > 0)
                {
                    ax -= 0.1;
                    fax = yi.next(ax);
                }
                else
                {
                    bx += 0.1;
                    fbx = yi.next(bx);
                }
            }
            CSBoost.XMath.pair<double> respair = CSBoost.XMath.toms748_solve(yi, ax, bx, fax, fbx, ep, ref max_iter);
            if (max_iter == 100 && Math.Abs(respair.v1 - respair.v2) > 2 * CSBoost.XMath.epsilon) throw new Exception();
            return (respair.v1 + respair.v2) / 2.0;
        }

        public static double ODDLYIELD(DateTime settlement, DateTime maturity, DateTime last_interest, double rate, double pr, double redemption, byte frequency, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("ODDLYIELD: Maturity must be after settlement (got {0:G} and {1:G} respectively).", maturity, settlement));
            if (settlement < last_interest) throw new System.ArgumentException(string.Format("ODDLYIELD: Settlement must be after last interest payment (got {0:G} and {1:G} respectively).", last_interest, settlement));
            if (rate < 0) throw new System.ArgumentException(string.Format("ODDLYIELD: Rate must be greater than zero (got {0:G}).", rate));
            if (pr < 0) throw new System.ArgumentException(string.Format("ODDLYIELD: Price must be greater than zero (got {0:G}).", pr));
            if (redemption < 0) throw new System.ArgumentException(string.Format("ODDLYIELD: Redemption must be greater than zero (got {0:G}).", redemption));
            check_frequency(frequency, "ODDLYIELD");

            double NC = COUPNUM(last_interest, maturity, frequency, basis);
            DateTime d1 = last_interest, d2;
            double t1a = 0, t2a = 0, t3a = 0;
            for (int i = 1; i <= NC; i++)
            {
                d2 = d1.AddMonths(12 / frequency);
                double Nl, DCi;
                if (basis == ExcelDCDBasis.US30_360)
                {
                    Nl = US360_daycount(d1, d2);
                    DCi = i < NC ? Nl : US360_daycount(d1, maturity);
                }
                else
                {
                    Nl = DAYCOUNT(d1, d2, basis);
                    DCi = i < NC ? Nl : DAYCOUNT(d1, maturity, basis);
                }
                double Ai = d2 < settlement ? DCi : d1 < settlement ? Math.Max(0, DAYCOUNT(d1, settlement, basis)) : 0;
                DateTime d3 = settlement > d1 ? settlement : d1;
                DateTime d4 = maturity < d2 ? maturity : d2;
                double DSC = Math.Max(0, DAYCOUNT(d3, d4, basis));
                d1 = d2;
                t1a += DCi / Nl;
                t2a += Ai / Nl;
                t3a += DSC / Nl;
            }
            double x = 100 * rate / frequency;
            double t1 = t1a * x + redemption;
            double t2 = t2a * x + pr;
            double t3 = frequency / t3a;
            double result = t3 * (t1 - t2) / t2;
            return result;
        }

        public static double INTRATE(DateTime settlement, DateTime maturity, double investment, double redemption, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("INTRATE: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (investment <= 0) throw new System.ArgumentException(string.Format("INTRATE: Investment must be greater than zero (got {0:G}).", investment));
            if (redemption <= 0) throw new System.ArgumentException(string.Format("INTRATE: Redemption must be greater than zero (got {0:G}).", redemption));

            double B = DAYSINYEAR(settlement, maturity, basis);
            double DIM = DAYCOUNT(settlement, maturity, basis);
            double result = ((redemption - investment) / investment) * (B / DIM);
            return result;
        }

        public static double RECEIVED(DateTime settlement, DateTime maturity, double investment, double discount, ExcelDCDBasis basis)
        {
            if (maturity < settlement) throw new System.ArgumentException(string.Format("RECEIVED: Settlement must be before maturity (got {0:G} and {1:G} respectively).", settlement, maturity));
            if (investment <= 0) throw new System.ArgumentException(string.Format("RECEIVED: Investment must be greater than zero (got {0:G}).", investment));
            if (discount <= 0) throw new System.ArgumentException(string.Format("RECEIVED: Discount must be greater than zero (got {0:G}).", discount));

            double B = DAYSINYEAR(settlement, maturity, basis);
            double DIM = DAYCOUNT(settlement, maturity, basis);
            if (discount * DIM / B == 1) throw new System.ArgumentException(string.Format("RECEIVED: Result cannot be evaluated for with the supplied arguments (discount * (maturity - settlement) = days in year).", settlement, maturity));
            double result = investment / (1 - discount * DIM / B);
            return result;
        }

        private static double cnum(DateTime mat, DateTime settl, int freq, ExcelDCDBasis basis)
        {
            bool endOfMonthTemp = IsEndOfMonth(mat);
            bool endOfMonth = !endOfMonthTemp && mat.Month != 2 && mat.Day > 28 && mat.Day < DateTime.DaysInMonth(mat.Year, mat.Month) ? IsEndOfMonth(settl) : endOfMonthTemp;
            DateTime startDate = endOfMonth ? EndOfMonth(settl) : settl;
            int coupons = settl < startDate ? 1 : 0;
            DateTime endDate = startDate.AddMonths(12 / freq);
            if (endOfMonth) endDate = EndOfMonth(endDate);
            while (endDate < mat)
            {
                startDate = endDate;
                endDate = startDate.AddMonths(12 / freq);
                if (endOfMonth) endDate = EndOfMonth(endDate);
                coupons++;
            }
            return coupons;
        }

        public static double TBILLEQ(DateTime settlement, DateTime maturity, double discount)
        {
            if (maturity < settlement) throw new ArgumentException(string.Format("TBILLEQ: Settlement must be before maturity (got {0:G} -> {1:G}).", settlement, maturity));
            if (maturity > settlement.AddYears(1)) throw new ArgumentException(string.Format("TBILLEQ: Settlement to maturity must be <= 1 year (got {0:G} -> {1:G}).", settlement, maturity));
            if (discount <= 0) throw new ArgumentException(string.Format("TBILLEQ: Discount rate must be > 0 (got {0:G}%).", Math.Round(100*discount, 2)));

            double DSM = DAYCOUNT(settlement, maturity, ExcelDCDBasis.Actual_360);
            double result;
            if (DSM > 182)
            {
                double price = (100 - discount * 100 * DSM / 360) / 100;
                double days = DSM == 366 ? 366 : 365;
                double t1 = DSM/days;
                double t2 = Math.Sqrt(Math.Pow(t1, 2) - (2 * t1 - 1) * (1 - 1 / price));
                result = 2 * (t2 - t1) / (2 * t1 - 1);
            }
            else
                result = (365 * discount) / (360 - (discount * DSM));
            return result;
        }

        public static double TBILLPRICE(DateTime settlement, DateTime maturity, double discount)
        {
            if (maturity < settlement) throw new ArgumentException(string.Format("TBILLPRICE: Settlement must be before maturity (got {0:G} -> {1:G}).", settlement, maturity));
            if (maturity > settlement.AddYears(1)) throw new ArgumentException(string.Format("TBILLPRICE: Settlement to maturity must be <= 1 year (got {0:G} -> {1:G}).", settlement, maturity));
            if (discount <= 0) throw new ArgumentException(string.Format("TBILLPRICE: Discount rate must be > 0 (got {0:G}%).", Math.Round(100 * discount, 2)));

            double DSM = DAYCOUNT(settlement, maturity, ExcelDCDBasis.Actual_Actual);
            double result = 100 * (1 - discount * DSM / 360);
            return result;
        }

        public static double TBILLYIELD(DateTime settlement, DateTime maturity, double pr)
        {
            if (maturity < settlement) throw new ArgumentException(string.Format("TBILLYIELD: Settlement must be before maturity (got {0:G} -> {1:G}).", settlement, maturity));
            if (maturity > settlement.AddYears(1)) throw new ArgumentException(string.Format("TBILLYIELD: Settlement to maturity must be <= 1 year (got {0:G} -> {1:G}).", settlement, maturity));
            if (pr <= 0) throw new ArgumentException(string.Format("TBILLYIELD: Price must be > 0 (got {0:G}%).", Math.Round(100 * pr, 2)));

            double DSM = DAYCOUNT(settlement, maturity, ExcelDCDBasis.Actual_Actual);
            double result = (360 / DSM) * (100 - pr) / pr;
            return result;
        }

        public static double YEARFRAC(DateTime start, DateTime end, ExcelDCDBasis basis)
        {
            return DAYCOUNT(start, end, basis) / DAYSINYEAR(start, end, basis);
        }

        #endregion

        #region misc

        public static double DOLLARDE(double fractional_dollar, int fraction)
        {
            if(fraction <= 0) throw new ArgumentException(string.Format("DOLLARDE: fraction argument must be > 0 (got {0:G}).", fraction));
            double intpart = fractional_dollar > 0 ? Math.Floor(fractional_dollar) : Math.Ceiling(fractional_dollar);
            double remainder = fractional_dollar - intpart;
            double digits = Math.Pow(10, Math.Ceiling(Math.Log10(fraction)));
            double result = intpart + remainder * digits / fraction;
            return result;
        }

        public static double DOLLARFR(double decimal_dollar, int fraction)
        {
            if (fraction <= 0) throw new ArgumentException(string.Format("DOLLARFR: fraction argument must be > 0 (got {0:G}).", fraction));
            double intpart = decimal_dollar > 0 ? Math.Floor(decimal_dollar) : Math.Ceiling(decimal_dollar);
            double remainder = decimal_dollar - intpart;
            double digits = Math.Abs(Math.Pow(10, Math.Ceiling(Math.Log10(fraction))));
            double result = intpart + remainder * fraction / digits;
            return result;
        }

        #endregion

    }
}


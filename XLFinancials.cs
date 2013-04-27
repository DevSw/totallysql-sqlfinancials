using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Collections.Generic;

public partial class SQLFinancials
{

    #region KeyValidation
    private static readonly List<bool> license = new List<bool>();
    internal static void ValidateKey()
    {
#if TRIAL
        if (license.Count > 0) return;
        using (SqlConnection conn = new SqlConnection("context connection=true"))
        {
            SqlCommand command = new SqlCommand("SELECT SERVERPROPERTY('edition')", conn);
            conn.Open();
            object o = command.ExecuteScalar();
            if (o != null && o.GetType().ToString() == "System.String")
            {
                string s = (string)o;
                if (s.StartsWith("Developer Edition"))
                {
                    license.Add(true);
                    return;
                }
            }

            command.CommandText = "SELECT af.content FROM sys.assemblies a INNER JOIN sys.assembly_files af ON a.assembly_id = af.assembly_id WHERE a.name like 'SQLFinancials%' AND af.name = 'license.key'";
            o = command.ExecuteScalar();
            if (o == null || o.GetType().ToString() != "System.Byte[]") throw new Exception("Error: No valid trial license key found for SQLFinancials. Please visit www.totallysql.com to purchase.");
            byte[] key = (byte[])o;
            if (key.Length != 8) throw new Exception("Error: No valid trial license key found for SQLFinancials. Please visit www.totallysql.com to purchase.");
            ushort a, ticks = 0, prodcust = 0, cust2 = 0;
            for (int j = 3; j >= 0; j--)
            {
                a = BitConverter.ToUInt16(key, j * 2);
                for (int i = 0; i < 4; i++)
                {
                    a = (ushort)(a >> 1);
                    cust2 = (ushort)((cust2 << 1) + (a & 1));
                    a = (ushort)(a >> 1);
                    prodcust = (ushort)((prodcust << 1) + (a & 1));
                    a = (ushort)(a >> 1);
                    ticks = (ushort)((ticks << 1) + (a & 1));
                    a = (ushort)(a >> 1);
                }
            }
            ulong ticktotal = (ulong)new DateTime(2010, 1, 1).Ticks / 864000000000 + (ulong)ticks;
            DateTime d = new DateTime((long)ticktotal * 864000000000);
            byte product = (byte)(prodcust & 0xFF);
            uint customer = cust2;
            customer = (uint)((customer << 8) + (prodcust >> 8));
            if (product != 0xE4) throw new Exception("Error: License key is not valid for this product (SQLFinancials). Please visit www.totallysql.com to purchase.");
            if (d.CompareTo(new DateTime(2010, 1, 6)) == 0)
                throw new Exception("Error - you need to activate the trial license first: EXEC dbo.ActivateSQLFinancialsTrial ");
            if (d <= DateTime.UtcNow) throw new Exception("Error: trial license for SQLFinancials has expired. Please visit www.totallysql.com to purchase.");
        }
        license.Add(true);
#endif
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDateTime SQLFinancialsExpires()
    {
#if TRIAL
        object o;

        using (SqlConnection conn = new SqlConnection("context connection=true"))
        {
            SqlCommand command = new SqlCommand("SELECT SERVERPROPERTY('edition')", conn);
            conn.Open();
            o = command.ExecuteScalar();
            if (o != null && o.GetType().ToString() == "System.String")
            {
                string s = (string)o;
                if (s.StartsWith("Developer Edition")) return SqlDateTime.MaxValue;
            }

            command.CommandText = "SELECT af.content FROM sys.assemblies a INNER JOIN sys.assembly_files af ON a.assembly_id = af.assembly_id WHERE a.name like 'SQLFinancials%' AND af.name = 'license.key'";
            o = command.ExecuteScalar();
        }

        if (o == null || o.GetType().ToString() != "System.Byte[]") throw new Exception("Error: No valid trial license key found for SQLFinancials. Please visit www.totallysql.com to purchase.");
        byte[] key = (byte[])o;
        if (key.Length != 8) throw new Exception("Error: No valid trial license key found for SQLFinancials. Please visit www.totallysql.com to purchase.");
        ushort a, ticks = 0, prodcust = 0, cust2 = 0;
        for (int j = 3; j >= 0; j--)
        {
            a = BitConverter.ToUInt16(key, j * 2);
            for (int i = 0; i < 4; i++)
            {
                a = (ushort)(a >> 1);
                cust2 = (ushort)((cust2 << 1) + (a & 1));
                a = (ushort)(a >> 1);
                prodcust = (ushort)((prodcust << 1) + (a & 1));
                a = (ushort)(a >> 1);
                ticks = (ushort)((ticks << 1) + (a & 1));
                a = (ushort)(a >> 1);
            }
        }
        ulong ticktotal = (ulong)new DateTime(2010, 1, 1).Ticks / 864000000000 + (ulong)ticks;
        DateTime d = new DateTime((long)ticktotal * 864000000000);
        byte product = (byte)(prodcust & 0xFF);
        uint customer = cust2;
        customer = (uint)((customer << 8) + (prodcust >> 8));
        if (product != 0xE4) throw new Exception("Error: License key is not valid for this product (SQLFinancials). Please visit www.totallysql.com to purchase.");
        return new SqlDateTime(d);
#else
        return SqlDateTime.MaxValue;
#endif
    }

#if TRIAL
    [SqlProcedure]
    public static void ActivateSQLFinancialsTrial()
    {
        DateTime expiry = SQLFinancialsExpires().Value;
        if (expiry.CompareTo(new DateTime(2010, 1, 6)) == 0)
        {
            string key = new_key();
            using (SqlConnection conn = new SqlConnection("context connection=true"))
            {
                SqlCommand command = new SqlCommand("alter assembly [SQLFinancials2005XTrial] drop file 'license.key'", conn);
                conn.Open();
                command.ExecuteNonQuery();
                command.CommandText = "ALTER ASSEMBLY [SQLFinancials2005XTrial] ADD FILE FROM " + key + " AS 'license.key'";
                command.ExecuteNonQuery();
                conn.Close();
            }
            SqlContext.Pipe.Send("SQLFinancials trial license successfully activated - expires: " + DateTime.Today.AddDays(31).ToShortDateString());
        }
        else
            SqlContext.Pipe.Send("Error: trial activation is not available on this installation. Please remove and install a fresh copy if you wish to extend the trial, or visit www.totallysql.com to purchase a perpetual license.");
    }

    private static readonly Random rand = new Random();

    private static string new_key()
    {
        DateTime Expiry = DateTime.Today.AddDays(31);
        ulong tickbase = (ulong)new DateTime(2010, 1, 1).Ticks / 864000000000;
        ushort ticks = (ushort)((ulong)Expiry.Ticks / 864000000000 - tickbase);
        ushort seed = (ushort)rand.Next();
        uint customer = (ushort)rand.Next();
        byte product = 0xE4;
        ushort prodcust = (ushort)((customer << 8) + product);
        ushort cust2 = (ushort)(customer >> 8);
        byte[] key = new byte[8];
        ushort a = 0;
        for (int j = 0; j < 4; j++)
        {
            a = 0;
            for (int i = 0; i < 4; i++)
            {
                a = (ushort)((a << 1) + (ticks & 1)); ticks = (ushort)(ticks >> 1);
                a = (ushort)((a << 1) + (prodcust & 1)); prodcust = (ushort)(prodcust >> 1);
                a = (ushort)((a << 1) + (cust2 & 1)); cust2 = (ushort)(cust2 >> 1);
                a = (ushort)((a << 1) + (seed & 1)); seed = (ushort)(seed >> 1);
            }
            BitConverter.GetBytes(a).CopyTo(key, j * 2);
        }
        string s = "0x" + BitConverter.ToString(key).Replace("-", "");
        return s;
    }
#endif


    #endregion

    #region depreciation

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DB(SqlDouble cost, SqlDouble salvage, SqlInt32 life, SqlInt32 period, SqlInt32 my1)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DB(cost.Value, salvage.Value, life.Value, period.Value, my1.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DDB(SqlDouble cost, SqlDouble salvage, SqlInt32 life, SqlInt32 period, SqlDouble factor)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DDB(cost.Value, salvage.Value, life.Value, period.Value, factor.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble SLN(SqlDouble cost, SqlDouble salvage, SqlInt32 life)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.SLN(cost.Value, salvage.Value, life.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble SYD(SqlDouble cost, SqlDouble salvage, SqlInt32 life, SqlInt32 period)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.SYD(cost.Value, salvage.Value, life.Value, period.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble VDB(SqlDouble cost, SqlDouble salvage, SqlInt32 life, SqlInt32 start_period, SqlInt32 end_period, SqlDouble factor, SqlBoolean noswitch)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.VDB(cost.Value, salvage.Value, life.Value, start_period.Value, end_period.Value, factor.Value, noswitch.Value);
    }

    #endregion

    #region payment

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble IPMT(SqlDouble rate, SqlInt32 per, SqlInt32 nper, SqlDouble pv, SqlDouble fv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.IPMT(rate.Value, per.Value, nper.Value, pv.Value, fv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ISPMT(SqlDouble rate, SqlInt32 per, SqlInt32 nper, SqlDouble pv)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ISPMT(rate.Value, per.Value, nper.Value, pv.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PMT(SqlDouble rate, SqlInt32 nper, SqlDouble pv, SqlDouble fv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PMT(rate.Value, nper.Value, pv.Value, fv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PPMT(SqlDouble rate, SqlInt32 per, SqlInt32 nper, SqlDouble pv, SqlDouble fv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PPMT(rate.Value, per.Value, nper.Value, pv.Value, fv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble NPER(SqlDouble rate, SqlDouble pmt, SqlDouble pv, SqlDouble fv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.NPER(rate.Value, pmt.Value, pv.Value, fv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble CUMIPMT(SqlDouble rate, SqlInt32 nper, SqlDouble pv, SqlInt32 start_period, SqlInt32 end_period, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.CUMIPMT(rate.Value, nper.Value, pv.Value, start_period.Value, end_period.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble CUMPRINC(SqlDouble rate, SqlInt32 nper, SqlDouble pv, SqlInt32 start_period, SqlInt32 end_period, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.CUMPRINC(rate.Value, nper.Value, pv.Value, start_period.Value, end_period.Value, payinadvance.Value);
    }

    #endregion

    #region future value - need to figure out how to process lists

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble FV(SqlDouble rate, SqlInt32 nper, SqlDouble pmt, SqlDouble pv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.FV(rate.Value, nper.Value, pmt.Value, pv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PV(SqlDouble rate, SqlInt32 nper, SqlDouble pmt, SqlDouble fv, SqlBoolean payinadvance)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PV(rate.Value, nper.Value, pmt.Value, fv.Value, payinadvance.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble RATE(SqlInt32 nper, SqlDouble pmt, SqlDouble pv, SqlDouble fv, SqlBoolean payinadvance, SqlDouble guess)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.RATE(nper.Value, pmt.Value, pv.Value, fv.Value, payinadvance.Value, guess.Value);
    }

    //public static SqlDouble FVSCHEDULE(SqlDouble principal, LIST OF VALUES)
    //public static SqlDouble IRR(LIST OF VALUES, SqlDouble guess)
    //public static SqlDouble MIRR(LIST OF VALUES, SqlDouble finance_rate, SqlDouble reinvest_rate)
    //public static SqlDouble NPV(LIST OF VALUES, SqlDouble rate)

    #endregion

    #region securities

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DAYCOUNT(SqlDateTime d1, SqlDateTime d2, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DAYCOUNT(d1.Value, d2.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DAYSINYEAR(SqlDateTime d1, SqlDateTime d2, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DAYSINYEAR(d1.Value, d2.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ACCRINT(SqlDateTime issue, SqlDateTime first_interest, SqlDateTime settlement, SqlDouble rate, SqlDouble par, SqlByte frequency, SqlByte basis, SqlBoolean calc_method)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ACCRINT(issue.Value, first_interest.Value, settlement.Value, rate.Value, par.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value, calc_method.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ACCRINTM(SqlDateTime issue, SqlDateTime settlement, SqlDouble rate, SqlDouble par, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ACCRINTM(issue.Value, settlement.Value, rate.Value, par.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble COUPDAYS(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPDAYS(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble COUPDAYSBS(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPDAYSBS(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble COUPDAYSNC(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPDAYSNC(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDateTime COUPNCD(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPNCD(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDateTime COUPPCD(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPPCD(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlInt32 COUPNUM(SqlDateTime settlement, SqlDateTime maturity, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.COUPNUM(settlement.Value, maturity.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DISC(SqlDateTime settlement, SqlDateTime maturity, SqlDouble price, SqlDouble redemption, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DISC(settlement.Value, maturity.Value, price.Value, redemption.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PRICE(SqlDateTime settlement, SqlDateTime maturity, SqlDouble rate, SqlDouble yld, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PRICE(settlement.Value, maturity.Value, rate.Value, yld.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PRICEDISC(SqlDateTime settlement, SqlDateTime maturity, SqlDouble discount, SqlDouble redemption, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PRICEDISC(settlement.Value, maturity.Value, discount.Value, redemption.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble PRICEMAT(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime issue, SqlDouble rate, SqlDouble yld, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.PRICEMAT(settlement.Value, maturity.Value, issue.Value, rate.Value, yld.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ODDFPRICE(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime issue, SqlDateTime first_interest, SqlDouble rate, SqlDouble yld, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ODDFPRICE(settlement.Value, maturity.Value, issue.Value, first_interest.Value, rate.Value, yld.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ODDLPRICE(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime last_interest, SqlDouble rate, SqlDouble yld, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ODDLPRICE(settlement.Value, maturity.Value, last_interest.Value, rate.Value, yld.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DURATION(SqlDateTime settlement, SqlDateTime maturity, SqlDouble coupon, SqlDouble yld, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DURATION(settlement.Value, maturity.Value, coupon.Value, yld.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble MDURATION(SqlDateTime settlement, SqlDateTime maturity, SqlDouble coupon, SqlDouble yld, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.MDURATION(settlement.Value, maturity.Value, coupon.Value, yld.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble YIELD(SqlDateTime settlement, SqlDateTime maturity, SqlDouble rate, SqlDouble price, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.YIELD(settlement.Value, maturity.Value, rate.Value, price.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble YIELDDISC(SqlDateTime settlement, SqlDateTime maturity, SqlDouble price, SqlDouble redemption, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.YIELDDISC(settlement.Value, maturity.Value, price.Value, redemption.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble YIELDMAT(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime issue, SqlDouble rate, SqlDouble price, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.YIELDMAT(settlement.Value, maturity.Value, issue.Value, rate.Value, price.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ODDFYIELD(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime issue, SqlDateTime first_interest, SqlDouble rate, SqlDouble price, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ODDFYIELD(settlement.Value, maturity.Value, issue.Value, first_interest.Value, rate.Value, price.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble ODDLYIELD(SqlDateTime settlement, SqlDateTime maturity, SqlDateTime last_interest, SqlDouble rate, SqlDouble price, SqlDouble redemption, SqlByte frequency, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.ODDLYIELD(settlement.Value, maturity.Value, last_interest.Value, rate.Value, price.Value, redemption.Value, frequency.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble INTRATE(SqlDateTime settlement, SqlDateTime maturity, SqlDouble investment, SqlDouble redemption, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.INTRATE(settlement.Value, maturity.Value, investment.Value, redemption.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble RECEIVED(SqlDateTime settlement, SqlDateTime maturity, SqlDouble investment, SqlDouble discount, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.RECEIVED(settlement.Value, maturity.Value, investment.Value, discount.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble EFFECT(SqlDouble rate, SqlInt32 nper)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.EFFECT(rate.Value, nper.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble NOMINAL(SqlDouble rate, SqlInt32 nper)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.NOMINAL(rate.Value, nper.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble TBILLEQ(SqlDateTime settlement, SqlDateTime maturity, SqlDouble discount)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.TBILLEQ(settlement.Value, maturity.Value, discount.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble TBILLPRICE(SqlDateTime settlement, SqlDateTime maturity, SqlDouble discount)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.TBILLPRICE(settlement.Value, maturity.Value, discount.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble TBILLYIELD(SqlDateTime settlement, SqlDateTime maturity, SqlDouble price)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.TBILLYIELD(settlement.Value, maturity.Value, price.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble YEARFRAC(SqlDateTime start, SqlDateTime end, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.YEARFRAC(start.Value, end.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble AMORLINC(SqlDouble cost, SqlDateTime purchased, SqlDateTime first_period, SqlDouble salvage, SqlInt32 period, SqlDouble rate, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.AMORLINC(cost.Value, purchased.Value, first_period.Value, salvage.Value, period.Value, rate.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble AMORDEGRC(SqlDouble cost, SqlDateTime purchased, SqlDateTime first_period, SqlDouble salvage, SqlInt32 period, SqlDouble rate, SqlByte basis)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.AMORDEGRC(cost.Value, purchased.Value, first_period.Value, salvage.Value, period.Value, rate.Value, (NETFinancials.XLFinancials.ExcelDCDBasis)basis.Value);
    }


    #endregion

    #region miscellaneous

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DOLLARDE(SqlDouble fractional_dollar, SqlInt32 fraction)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DOLLARDE(fractional_dollar.Value, fraction.Value);
    }

    [SqlFunction(SystemDataAccess = SystemDataAccessKind.Read)]
    public static SqlDouble DOLLARFR(SqlDouble decimal_dollar, SqlInt32 fraction)
    {
        ValidateKey();
        return NETFinancials.XLFinancials.DOLLARFR(decimal_dollar.Value, fraction.Value);
    }

    #endregion

};


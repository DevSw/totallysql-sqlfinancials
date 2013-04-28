
using System;
using System.Data;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using Microsoft.SqlServer.Server;
using System.Collections.Generic;
using NETFinancials;
#if EXTENDED
    using SQLCore;
#endif

public struct flow : IComparable
{
    public flow(double v, int o)
    {
        value = v;
        order = (ushort)o;
    }
    public double value;
    public ushort order;

    #region IComparable Members

    public int CompareTo(object obj)
    {
        flow f = (flow)obj;
        return order.CompareTo(f.order);
    }

    #endregion
}

#if EXTENDED
public partial class SqlFinancials
{
    internal const bool IsExtended = true;
}
#else
public partial class SqlFinancials
{
    internal const bool IsExtended = false;
}
#endif

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize=8000)]
public struct FVSCHEDULE: IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<flow>> theLists = new ThreadSafeDictionary<Guid, List<flow>>();
#endif

    double p;
    List<flow> theList;

    public void Init()
    {
        p = double.NaN; 
        theList = new List<flow>();
    }

#if SQL2005
    public void Accumulate(tuple_3 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 798) throw new Exception("FVSCHEDULE: Too many entries - the SQL2005 version of this function has a maximum of 799 rate entries");
        SqlDouble principal = Tuples.GetDouble(Value.v1);
        SqlInt32 number = Tuples.GetInt(Value.v2);
        SqlDouble rate = Tuples.GetDouble(Value.v3);
#else
    public void Accumulate(SqlDouble principal, SqlInt32 number, SqlDouble rate)
    {
#endif
        if (principal.IsNull || number.IsNull || rate.IsNull) throw new Exception("FVSCHEDULE: Null values cannot be passed to this function");
        if (double.IsNaN(p)) p = principal.Value;
        else if (p != principal.Value)
            throw new Exception(string.Format("FVSCHEDULE: Principal must be constant (got {0:G} and {1:G}).", p, principal.Value));
        theList.Add(new flow(rate.Value, number.Value));
    }

    public void Merge(FVSCHEDULE Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 798) throw new Exception("FVSCHEDULE: Too many entries - the SQL2005 version of this function has a maximum of 799 rate entries");
        if (double.IsNaN(p)) p = Group.p;
        else if (!double.IsNaN(Group.p) && p != Group.p)
            throw new Exception(string.Format("FVSCHEDULE: Principal must be constant (got {0:G} and {1:G}).", p, Group.p));
        theList.AddRange(Group.theList);
    }

    public SqlDouble Terminate()
    {
        theList.Sort();
        List<double> rates = new List<double>();
        foreach (flow f in theList) rates.Add(f.value);
        try
        {
            return new SqlDouble(XLFinancials.FVSCHEDULE(p, rates));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }


    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
        p = r.ReadDouble();
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<flow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            flow f = new flow();
            f.value = r.ReadDouble();
            f.order = r.ReadUInt16();
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
        w.Write(p);
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (flow f in theList)
        {
            w.Write(f.value);
            w.Write(f.order);
        }
#endif
    }

    #endregion
}

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = 8000)]
public struct NPV : IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<flow>> theLists = new ThreadSafeDictionary<Guid, List<flow>>();
#endif

    double p;
    List<flow> theList;

    public void Init()
    {
        p = double.NaN;
        theList = new List<flow>();
    }

#if SQL2005
    public void Accumulate(tuple_3 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 798) throw new Exception("NPV: Too many entries - the SQL2005 version of this function has a maximum of 799 cashflow entries");
        SqlDouble rate = Tuples.GetDouble(Value.v1);
        SqlInt32 number = Tuples.GetInt(Value.v2);
        SqlDouble payment = Tuples.GetDouble(Value.v3);
#else
    public void Accumulate(SqlDouble rate, SqlInt32 number, SqlDouble payment)
    {
#endif
        if (rate.IsNull || number.IsNull || payment.IsNull) throw new Exception("NPV: Null values cannot be passed to this function");
        if (double.IsNaN(p)) p = rate.Value;
        else if (p != rate.Value)
            throw new Exception(string.Format("NPV: Rate must be constant (got {0:G} and {1:G}).", p, rate.Value));
        theList.Add(new flow(payment.Value, number.Value));
    }

    public void Merge(NPV Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 798) throw new Exception("NPV: Too many entries - the SQL2005 version of this function has a maximum of 799 cashflow entries");
        if (double.IsNaN(p)) p = Group.p;
        else if (!double.IsNaN(Group.p) && p != Group.p)
            throw new Exception(string.Format("NPV: Rate must be constant (got {0:G} and {1:G}).", p, Group.p));
        theList.AddRange(Group.theList);
    }

    public SqlDouble Terminate()
    {
        theList.Sort();
        List<double> flows = new List<double>();
        foreach (flow f in theList) flows.Add(f.value);
        try
        {
            return new SqlDouble(XLFinancials.NPV(flows, p));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }


    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
        p = r.ReadDouble();
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<flow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            flow f = new flow();
            f.value = r.ReadDouble();
            f.order = r.ReadUInt16();
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
        w.Write(p);
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (flow f in theList)
        {
            w.Write(f.value);
            w.Write(f.order);
        }
#endif
    }

    #endregion
}

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = 8000)]
public struct IRR : IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<flow>> theLists = new ThreadSafeDictionary<Guid, List<flow>>();
#endif
    List<flow> theList;

    public void Init()
    {
        theList = new List<flow>();
    }

#if SQL2005
    public void Accumulate(tuple_2 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 798) throw new Exception("IRR: Too many entries - the SQL2005 version of this function has a maximum of 799 cashflow entries");
        SqlInt32 number = Tuples.GetInt(Value.v1);
        SqlDouble payment = Tuples.GetDouble(Value.v2);
#else
    public void Accumulate(SqlInt32 number, SqlDouble payment)
    {
#endif
        if (number.IsNull || payment.IsNull) throw new Exception("IRR: Null values cannot be passed to this function");
        theList.Add(new flow(payment.Value, number.Value));
    }


    public void Merge(IRR Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 798) throw new Exception("IRR: Too many entries - the SQL2005 version of this function has a maximum of 799 cashflow entries");
        theList.AddRange(Group.theList);
    }

    public SqlDouble Terminate()
    {
        theList.Sort();
        List<double> flows = new List<double>();
        foreach (flow f in theList) flows.Add(f.value);
        try
        {
            return new SqlDouble(XLFinancials.IRR(flows));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }


    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<flow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            flow f = new flow();
            f.value = r.ReadDouble();
            f.order = r.ReadUInt16();
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (flow f in theList)
        {
            w.Write(f.value);
            w.Write(f.order);
        }
#endif
    }

    #endregion
}

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = 8000)]
public struct XNPV : IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<XLFinancials.cashflow>> theLists = new ThreadSafeDictionary<Guid, List<XLFinancials.cashflow>>();
#endif

    double p;
    List<XLFinancials.cashflow> theList;

    public void Init()
    {
        p = double.NaN;
        theList = new List<XLFinancials.cashflow>();
    }

#if SQL2005
    public void Accumulate(tuple_3 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 498) throw new Exception("XNPV: Too many entries - the SQL2005 version of this function has a maximum of 499 cashflow entries");
        SqlDouble rate = Tuples.GetDouble(Value.v1);
        SqlDateTime date = Tuples.GetDateTime(Value.v2);
        SqlDouble payment = Tuples.GetDouble(Value.v3);
#else
    public void Accumulate(SqlDouble rate, SqlDateTime date, SqlDouble payment)
    {
#endif
        if (rate.IsNull || date.IsNull || payment.IsNull) throw new Exception("XNPV: Null values cannot be passed to this function");
        if (double.IsNaN(p)) p = rate.Value;
        else if (p != rate.Value)
            throw new Exception(string.Format("XNPV: Rate must be constant (got {0:G} and {1:G}).", p, rate.Value));
        theList.Add(new XLFinancials.cashflow(payment.Value, date.Value));
    }

    public void Merge(XNPV Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 498) throw new Exception("XNPV: Too many entries - the SQL2005 version of this function has a maximum of 499 cashflow entries");
        if (double.IsNaN(p)) p = Group.p;
        else if (!double.IsNaN(Group.p) && p != Group.p)
            throw new Exception(string.Format("XNPV: Rate must be constant (got {0:G} and {1:G}).", p, Group.p));
        theList.AddRange(Group.theList);
    }

    public SqlDouble Terminate()
    {
        theList.Sort();
        try
        {
            return new SqlDouble(XLFinancials.XNPV(theList, p));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }


    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
        p = r.ReadDouble();
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<XLFinancials.cashflow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            XLFinancials.cashflow f = new XLFinancials.cashflow();
            f.payment = r.ReadDouble();
            f.date = new DateTime(r.ReadInt64());
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
        w.Write(p);
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (XLFinancials.cashflow f in theList)
        {
            w.Write(f.payment);
            w.Write(f.date.Ticks);
        }
#endif
    }

    #endregion
}

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = 8000)]
public struct XIRR : IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<XLFinancials.cashflow>> theLists = new ThreadSafeDictionary<Guid, List<XLFinancials.cashflow>>();
#endif

    List<XLFinancials.cashflow> theList;

    public void Init()
    {
        theList = new List<XLFinancials.cashflow>();
    }

#if SQL2005
    public void Accumulate(tuple_2 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 498) throw new Exception("XIRR: Too many entries - the SQL2005 version of this function has a maximum of 499 cashflow entries");
        SqlDateTime date = Tuples.GetDateTime(Value.v1);
        SqlDouble payment = Tuples.GetDouble(Value.v2);
#else
    public void Accumulate(SqlDateTime date, SqlDouble payment)
    {
#endif
        if (date.IsNull || payment.IsNull) throw new Exception("XIRR: Null values cannot be passed to this function");
        theList.Add(new XLFinancials.cashflow(payment.Value, date.Value));
    }


    public void Merge(XIRR Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 498) throw new Exception("XIRR: Too many entries - the SQL2005 version of this function has a maximum of 499 cashflow entries");
        theList.AddRange(Group.theList);
    }
   
    public SqlDouble Terminate()
    {
        theList.Sort();
        try
        {
            return new SqlDouble(XLFinancials.XIRR(theList));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }


    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<XLFinancials.cashflow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            XLFinancials.cashflow f = new XLFinancials.cashflow();
            f.payment = r.ReadDouble();
            f.date = new DateTime(r.ReadInt64());
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (XLFinancials.cashflow f in theList)
        {
            w.Write(f.payment);
            w.Write(f.date.Ticks);
        }
#endif
    }


    #endregion
}

[Serializable]
[Microsoft.SqlServer.Server.SqlUserDefinedAggregate(Format.UserDefined, MaxByteSize = 8000)]
public struct MIRR : IBinarySerialize
{
#if EXTENDED
    readonly static ThreadSafeDictionary<Guid, List<flow>> theLists = new ThreadSafeDictionary<Guid, List<flow>>();
#endif

    double frate, rrate;
    List<flow> theList;

    public void Init()
    {
        theList = new List<flow>();
        rrate = frate = double.NaN;
    }

#if SQL2005
    public void Accumulate(tuple_4 Value)
    {
        if (Value.IsNull) return;
        if (!SqlFinancials.IsExtended && theList.Count > 797) throw new Exception("MIRR: Too many entries - the SQL2005 version of this function has a maximum of 798 cashflow entries");
        SqlInt32 number = Tuples.GetInt(Value.v1);
        SqlDouble payment = Tuples.GetDouble(Value.v2);
        SqlDouble finance_rate = Tuples.GetDouble(Value.v3);
        SqlDouble reinvest_rate = Tuples.GetDouble(Value.v4);
#else
    public void Accumulate(SqlInt32 number, SqlDouble payment, SqlDouble finance_rate, SqlDouble reinvest_rate)
    {
#endif
        if (number.IsNull || payment.IsNull || finance_rate.IsNull || reinvest_rate.IsNull) throw new Exception("MIRR: Null values cannot be passed to this function");
        if (double.IsNaN(frate)) frate = finance_rate.Value;
        else if (frate != finance_rate.Value)
            throw new Exception(string.Format("MIRR: Finance Rate must be constant (got {0:G} and {1:G}).", frate, finance_rate.Value));
        if (double.IsNaN(rrate)) rrate = reinvest_rate.Value;
        else if (rrate != reinvest_rate.Value)
            throw new Exception(string.Format("MIRR: Reinvest Rate must be constant (got {0:G} and {1:G}).", rrate, reinvest_rate.Value));
        theList.Add(new flow(payment.Value, number.Value));
    }


    public void Merge(MIRR Group)
    {
        if (!SqlFinancials.IsExtended && theList.Count + Group.theList.Count > 797) throw new Exception("MIRR: Too many entries - the SQL2005 version of this function has a maximum of 798 cashflow entries");
        theList.AddRange(Group.theList);
        if (double.IsNaN(frate)) frate = Group.frate;
        else if (!double.IsNaN(Group.frate) && frate != Group.frate)
            throw new Exception(string.Format("MIRR: Finance Rate must be constant (got {0:G} and {1:G}).", frate, Group.frate));
        if (double.IsNaN(rrate)) rrate = Group.rrate;
        else if (!double.IsNaN(Group.rrate) && rrate != Group.rrate)
            throw new Exception(string.Format("MIRR: Reinvest Rate must be constant (got {0:G} and {1:G}).", rrate, Group.rrate));
    }

    public SqlDouble Terminate()
    {
        theList.Sort();
        List<double> flows = new List<double>();
        foreach (flow f in theList) flows.Add(f.value);
        try
        {
            return new SqlDouble(XLFinancials.MIRR(flows, frate, rrate));
        }
        catch
        {
            return SqlDouble.Null;
        }
    }    
    
    #region IBinarySerialize Members

    public void Read(System.IO.BinaryReader r)
    {
        frate = r.ReadDouble();
        rrate = r.ReadDouble();
#if EXTENDED
        Guid g = new Guid(r.ReadBytes(16));
        try
        {
            this.theList = theLists[g];
        }
        finally
        {
            theLists.Remove(g);
        }
#else
        theList = new List<flow>();
        int n = r.ReadUInt16();
        for (int i = 0; i < n; i++)
        {
            flow f = new flow();
            f.value = r.ReadDouble();
            f.order = r.ReadUInt16();
            theList.Add(f);
        }
#endif
    }

    public void Write(System.IO.BinaryWriter w)
    {
        w.Write(frate);
        w.Write(rrate);
#if EXTENDED
        Guid g = Guid.NewGuid();
        try
        {
            theLists.Add(g, this.theList);
            w.Write(g.ToByteArray());
        }
        catch (Exception e)
        {
            if (theLists.ContainsKey(g)) theLists.Remove(g);
            throw new Exception(e.ToString());
        }
#else
        w.Write((ushort)theList.Count);
        foreach (flow f in theList)
        {
            w.Write(f.value);
            w.Write(f.order);
        }
#endif
    }

    #endregion
}



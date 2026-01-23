using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using ExcelAddIn1;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class GB
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<string, Range> A;

		public static Func<Range, string> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal Range A(string A)
		{
			return GB.B(A);
		}

		[SpecialName]
		internal string B(Range A)
		{
			return GB.A(A);
		}
	}

	private readonly Dictionary<string, List<string>> m_A;

	public GB()
	{
		this.m_A = new Dictionary<string, List<string>>();
	}

	internal List<Range> A(Range A)
	{
		try
		{
			List<Range> list = B(A) ?? C(A);
			object result;
			if (list.Count != 0)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				result = list;
			}
			else
			{
				result = null;
			}
			return (List<Range>)result;
		}
		finally
		{
			List<Range> list = null;
		}
	}

	internal int A(Range A)
	{
		List<string> value = null;
		if (!this.m_A.TryGetValue(GB.A(A), out value))
		{
			while (true)
			{
				switch (4)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return 0;
				}
			}
		}
		return value.Count;
	}

	private List<Range> B(Range A)
	{
		List<string> value = null;
		if (!this.m_A.TryGetValue(GB.A(A), out value))
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return null;
				}
			}
		}
		List<string> source = value;
		Func<string, Range> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (string a) => B(a));
		}
		else
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			selector = _Closure_0024__.A;
		}
		return source.Select(selector).ToList();
	}

	private List<Range> C(Range A)
	{
		try
		{
			List<Range> list = RangeHelpers.D(A);
			if (list == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					break;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
				}
				list = new List<Range>();
			}
			List<Range> list2 = list;
			this.A(A, list2);
			return list2;
		}
		finally
		{
			List<Range> list2 = null;
		}
	}

	private void A(Range A, List<Range> B)
	{
		Dictionary<string, List<string>> a = this.m_A;
		string key = GB.A(A);
		Func<Range, string> selector;
		if (_Closure_0024__.A == null)
		{
			selector = (_Closure_0024__.A = [SpecialName] (Range a2) => GB.A(a2));
		}
		else
		{
			while (true)
			{
				switch (7)
				{
				case 0:
					continue;
				}
				break;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			selector = _Closure_0024__.A;
		}
		a[key] = B.Select(selector).ToList();
	}

	internal void A()
	{
		this.m_A.Clear();
	}

	private static string A(Range A)
	{
		return A.get_Address(RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlReferenceStyle.xlA1, (object)true, RuntimeHelpers.GetObjectValue(Missing.Value));
	}

	private static Range B(string A)
	{
		return ((_Application)MH.A.Application).get_Range((object)A, RuntimeHelpers.GetObjectValue(Missing.Value));
	}
}

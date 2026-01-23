using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class ZB : XB
{
	[Serializable]
	[CompilerGenerated]
	internal new sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Worksheet, Range> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal Range A(Worksheet A)
		{
			return RangeHelpers.A(A);
		}
	}

	[CompilerGenerated]
	internal sealed class YB
	{
		public Analysis A;

		public Action<Observations, Range, Worksheet> A;

		public Action<Observations, Worksheet> A;

		public ZB A;

		[SpecialName]
		internal void A(Worksheet A)
		{
			Observations observations = new Observations();
			try
			{
				ZB zB = this.A;
				Analysis b = this.A;
				Observations c = observations;
				string d = VH.A(2658);
				Func<Worksheet, Range> e;
				if (_Closure_0024__.A == null)
				{
					e = (_Closure_0024__.A = _Closure_0024__.A.A);
				}
				else
				{
					while (true)
					{
						switch (6)
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
					e = _Closure_0024__.A;
				}
				zB.B(A, b, c, d, e, this.A);
			}
			finally
			{
				Action<Observations, Worksheet> action = this.A;
				if (action == null)
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
				}
				else
				{
					action(observations, A);
				}
				observations = null;
			}
		}
	}

	internal ZB(Analysis A, string B, Action<Observations, Range, Worksheet> C, Action<Observations, Worksheet> D, Func<bool> E = null)
		: base(B, null, E)
	{
		ZB A2 = this;
		base.WsAction = [SpecialName] (Worksheet worksheet) =>
		{
			Observations observations = new Observations();
			try
			{
				ZB zB = A2;
				Analysis b = A;
				Observations c = observations;
				string d = VH.A(2658);
				Func<Worksheet, Range> e;
				if (_Closure_0024__.A == null)
				{
					e = (_Closure_0024__.A = _Closure_0024__.A.A);
				}
				else
				{
					while (true)
					{
						switch (6)
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
					e = _Closure_0024__.A;
				}
				zB.B(worksheet, b, c, d, e, C);
			}
			finally
			{
				Action<Observations, Worksheet> action = D;
				if (action == null)
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
				}
				else
				{
					action(observations, worksheet);
				}
				observations = null;
			}
		};
	}
}

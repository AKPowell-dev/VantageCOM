using System;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;

namespace A;

internal sealed class BC : XB
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
			return A.UsedRange;
		}
	}

	[CompilerGenerated]
	internal sealed class AC
	{
		public Analysis A;

		public Action<Observations, Range, Worksheet> A;

		public Action<Observations> A;

		public BC A;

		[SpecialName]
		internal void A(Worksheet A)
		{
			Observations observations = new Observations();
			try
			{
				BC bC = this.A;
				Analysis b = this.A;
				Observations c = observations;
				string d = VH.A(2512);
				Func<Worksheet, Range> e;
				if (_Closure_0024__.A == null)
				{
					e = (_Closure_0024__.A = _Closure_0024__.A.A);
				}
				else
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
					e = _Closure_0024__.A;
				}
				bC.B(A, b, c, d, e, this.A);
			}
			finally
			{
				Action<Observations> action = this.A;
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
					action(observations);
				}
				observations = null;
			}
		}
	}

	internal BC(Analysis A, string B, Action<Observations, Range, Worksheet> C, Action<Observations> D, Func<bool> E = null)
		: base(B, null, E)
	{
		BC A2 = this;
		base.WsAction = [SpecialName] (Worksheet a) =>
		{
			Observations observations = new Observations();
			try
			{
				BC bC = A2;
				Analysis b = A;
				Observations c = observations;
				string d = VH.A(2512);
				Func<Worksheet, Range> e;
				if (_Closure_0024__.A == null)
				{
					e = (_Closure_0024__.A = _Closure_0024__.A.A);
				}
				else
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
					e = _Closure_0024__.A;
				}
				bC.B(a, b, c, d, e, C);
			}
			finally
			{
				Action<Observations> action = D;
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
					action(observations);
				}
				observations = null;
			}
		};
	}
}

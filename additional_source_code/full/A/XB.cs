using System;
using System.Collections;
using System.Runtime.CompilerServices;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Audit.Check.Observations.Raw;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal class XB : RB
{
	[CompilerGenerated]
	internal sealed class VB
	{
		public Action<Observations, Worksheet> A;

		public Action<Observations> A;

		[SpecialName]
		internal void A(Worksheet A)
		{
			Observations observations = new Observations();
			try
			{
				this.A(observations, A);
			}
			finally
			{
				Action<Observations> action = this.A;
				if (action == null)
				{
					while (true)
					{
						switch (1)
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
				}
				else
				{
					action(observations);
				}
				observations = null;
			}
		}
	}

	[CompilerGenerated]
	internal sealed class WB
	{
		public Analysis A;

		[SpecialName]
		internal bool A(Worksheet A)
		{
			return !this.A.ItemCancelled();
		}
	}

	[CompilerGenerated]
	private new Action<Worksheet> m_A;

	internal Action<Worksheet> WsAction
	{
		[CompilerGenerated]
		get
		{
			return this.m_A;
		}
		[CompilerGenerated]
		set
		{
			this.m_A = value;
		}
	}

	internal XB(string A, Action<Worksheet> B, Func<bool> C = null)
	{
		base.CheckDesc = A;
		WsAction = B;
		base.CondFunc = C;
	}

	internal XB(string A, Action<Observations, Worksheet> B, Action<Observations> C, Func<bool> D = null)
	{
		base.CheckDesc = A;
		WsAction = [SpecialName] (Worksheet arg) =>
		{
			Observations observations = new Observations();
			try
			{
				B(observations, arg);
			}
			finally
			{
				Action<Observations> action = C;
				if (action == null)
				{
					while (true)
					{
						switch (1)
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
				}
				else
				{
					action(observations);
				}
				observations = null;
			}
		};
		base.CondFunc = D;
	}

	internal override void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null)
	{
		OB.A(D, WsAction, [SpecialName] (long numItems) =>
		{
			base.AssociatedAction.NumItems = numItems;
		}, [SpecialName] (Worksheet worksheet) => !A.ItemCancelled());
	}

	protected void B(Worksheet A, Analysis B, Observations C, string D, Func<Worksheet, Range> E, Action<Observations, Range, Worksheet> F)
	{
		Range range = E(A);
		if (range == null)
		{
			return;
		}
		int a = B.A;
		try
		{
			B.ActionStarted(D, Conversions.ToLong(range.CountLarge));
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = range.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Range arg = (Range)enumerator.Current;
					if (B.ItemCancelled())
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
								return;
							}
						}
					}
					F(C, arg, A);
				}
				while (true)
				{
					switch (4)
					{
					case 0:
						break;
					default:
						return;
					}
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						(enumerator as IDisposable).Dispose();
						break;
					}
				}
			}
		}
		finally
		{
			B.A(a);
			Range arg = null;
			range = null;
		}
	}

	[SpecialName]
	[CompilerGenerated]
	private void B(long A)
	{
		base.AssociatedAction.NumItems = A;
	}
}

using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using ExcelAddIn1.Audit.Check;
using ExcelAddIn1.Audit.Check.Helpers;
using MacabacusMacros;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.CompilerServices;

namespace A;

internal abstract class RB
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<RB, bool> A;

		public static Func<RB, bool> B;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(RB A)
		{
			if (A.B())
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return !(A is SB);
					}
				}
			}
			return false;
		}

		[SpecialName]
		internal bool B(RB A)
		{
			return A is SB;
		}
	}

	[CompilerGenerated]
	private string m_A;

	[CompilerGenerated]
	private Func<bool> m_A;

	[CompilerGenerated]
	private ActionItem m_A;

	internal string CheckDesc
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

	internal Func<bool> CondFunc
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

	internal ActionItem AssociatedAction
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

	internal abstract void A(Analysis A, Application B = null, Workbook C = null, Sheets D = null);

	internal static void B(Analysis A, List<RB> B, Application C = null, Workbook D = null, Sheets E = null)
	{
		List<RB> list = B.Where([SpecialName] (RB rB) =>
		{
			if (rB.B())
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						return !(rB is SB);
					}
				}
			}
			return false;
		}).ToList();
		A.B(list);
		checked
		{
			try
			{
				A.B = list.Count;
				int num = 0;
				using List<RB>.Enumerator enumerator = list.GetEnumerator();
				while (enumerator.MoveNext())
				{
					RB current = enumerator.Current;
					A.B();
					if (A.A(A: true))
					{
						while (true)
						{
							switch (7)
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
					ActionItem associatedAction = current.AssociatedAction;
					if (associatedAction != null && associatedAction.A(CB.F, CB.G))
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
						num++;
						continue;
					}
					int a = A.A(current, 1L);
					current.AssociatedAction = A.A();
					try
					{
						num = (A.C = num + 1);
						current.A(A, C, D, E);
					}
					catch (ThreadAbortException ex)
					{
						ProjectData.SetProjectError(ex);
						ThreadAbortException ex2 = ex;
						ProjectData.ClearProjectError();
					}
					catch (Exception ex3)
					{
						ProjectData.SetProjectError(ex3);
						Exception ex4 = ex3;
						if (!(ex4 is DB))
						{
							while (true)
							{
								switch (4)
								{
								case 0:
									continue;
								}
								break;
							}
							clsReporting.LogException(ex4);
						}
						current.AssociatedAction.Exception = ex4;
						ProjectData.ClearProjectError();
					}
					finally
					{
						A.A(a);
					}
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
				IEnumerator<SB> enumerator2 = default(IEnumerator<SB>);
				try
				{
					Func<RB, bool> predicate;
					if (_Closure_0024__.B == null)
					{
						predicate = (_Closure_0024__.B = [SpecialName] (RB rB) => rB is SB);
					}
					else
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
						predicate = _Closure_0024__.B;
					}
					enumerator2 = B.Where(predicate).Cast<SB>().GetEnumerator();
					while (enumerator2.MoveNext())
					{
						SB current2 = enumerator2.Current;
						if (!current2.B())
						{
							continue;
						}
						while (true)
						{
							switch (4)
							{
							case 0:
								continue;
							}
							break;
						}
						current2.B(A);
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							goto end_IL_01e1;
						}
						continue;
						end_IL_01e1:
						break;
					}
				}
				finally
				{
					enumerator2?.Dispose();
				}
			}
		}
	}

	private bool B()
	{
		Func<bool> condFunc = CondFunc;
		bool? obj;
		if (condFunc == null)
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
			obj = null;
		}
		else
		{
			obj = condFunc();
		}
		return !object.Equals(obj, false);
	}
}

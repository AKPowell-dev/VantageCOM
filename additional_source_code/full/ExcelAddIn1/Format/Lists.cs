using System;
using System.Collections;
using System.Runtime.CompilerServices;
using A;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Format;

public sealed class Lists
{
	[CompilerGenerated]
	private static int m_A;

	internal static int CycleIndex
	{
		[CompilerGenerated]
		get
		{
			return Lists.m_A;
		}
		[CompilerGenerated]
		set
		{
			Lists.m_A = value;
		}
	}

	public static void Cycle()
	{
		checked
		{
			if (Licensing.AllowRestrictedMode())
			{
				CycleIndex++;
				switch (CycleIndex)
				{
				case 1:
					Bullets();
					break;
				case 2:
					Dashes();
					break;
				case 3:
					Numbers();
					break;
				case 4:
					LettersUpper();
					break;
				case 5:
					LettersLower();
					break;
				case 6:
					RomanUpper();
					break;
				case 7:
					RomanLower();
					break;
				default:
					None();
					break;
				}
				if (CycleIndex == 1)
				{
					Base.LogActivity(VH.A(150010));
				}
			}
		}
	}

	public static void None()
	{
		A("");
		CycleIndex = 0;
	}

	public static void Bullets()
	{
		A(Conversions.ToString(Strings.Chr(149)) + VH.A(41385));
	}

	public static void Dashes()
	{
		A(VH.A(150031));
	}

	private static void A(string A)
	{
		int try0000_dispatch = -1;
		int num3 = default(int);
		int num = default(int);
		int num2 = default(int);
		Range range = default(Range);
		bool flag = default(bool);
		Application application = default(Application);
		IEnumerator enumerator = default(IEnumerator);
		Range range2 = default(Range);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0007;
				case 333:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0007;
						case 3:
							goto IL_0012;
						case 4:
							goto IL_001e;
						case 5:
							goto IL_002e;
						case 6:
							goto IL_0038;
						case 7:
							goto IL_0054;
						case 8:
							goto IL_0072;
						case 9:
							goto IL_008f;
						case 10:
							goto IL_00a7;
						case 11:
							goto IL_00bf;
						case 12:
							goto IL_00ca;
						case 13:
							goto IL_00cd;
						case 14:
							goto IL_00de;
						case 15:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 16:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_00de:
					num2 = 14;
					JH.A(range, VH.A(82757));
					break;
					IL_0007:
					num2 = 2;
					range = JH.A((Range)null);
					goto IL_0012;
					IL_0012:
					num2 = 3;
					flag = JH.A(range);
					goto IL_001e;
					IL_001e:
					num2 = 4;
					application = MH.A.Application;
					goto IL_002e;
					IL_002e:
					num2 = 5;
					application.ScreenUpdating = false;
					goto IL_0038;
					IL_0038:
					num2 = 6;
					enumerator = range.GetEnumerator();
					goto IL_0092;
					IL_0092:
					if (enumerator.MoveNext())
					{
						range2 = (Range)enumerator.Current;
						goto IL_0054;
					}
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_00a7;
					IL_0054:
					num2 = 7;
					if (Base.IsString(range2))
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
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						goto IL_0072;
					}
					goto IL_008f;
					IL_00a7:
					num2 = 10;
					if (enumerator is IDisposable)
					{
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00bf;
					IL_008f:
					num2 = 9;
					goto IL_0092;
					IL_00bf:
					num2 = 11;
					application.ScreenUpdating = true;
					goto IL_00ca;
					IL_00ca:
					application = null;
					goto IL_00cd;
					IL_00cd:
					num2 = 13;
					if (!flag)
					{
						break;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_00de;
					IL_0072:
					num2 = 8;
					range2.NumberFormat = A + VH.A(48146);
					goto IL_008f;
					end_IL_0000_2:
					break;
				}
				num2 = 15;
				range = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 333;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void Numbers()
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Range range2 = default(Range);
		int num = default(int);
		int num3 = default(int);
		Range range = default(Range);
		int num5 = default(int);
		bool flag = default(bool);
		Application application = default(Application);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					range2 = JH.A((Range)null);
					goto IL_0009;
				case 372:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_0009;
						case 3:
							goto IL_0015;
						case 4:
							goto IL_001a;
						case 5:
							goto IL_0021;
						case 6:
							goto IL_0033;
						case 7:
							goto IL_003d;
						case 8:
							goto IL_005b;
						case 9:
							goto IL_0079;
						case 10:
							goto IL_00a7;
						case 11:
							goto IL_00b0;
						case 12:
							goto IL_00c8;
						case 13:
							goto IL_00ea;
						case 14:
							goto IL_00f5;
						case 15:
							goto IL_00f8;
						case 16:
							goto IL_00ff;
						case 17:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 18:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_0079:
					num2 = 9;
					range.NumberFormat = VH.A(48247) + Conversions.ToString(num5) + VH.A(150036);
					goto IL_00a7;
					IL_0009:
					num2 = 2;
					flag = JH.A(range2);
					goto IL_0015;
					IL_0015:
					num2 = 3;
					num5 = 1;
					goto IL_001a;
					IL_001a:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0021;
					IL_0021:
					num2 = 5;
					application = MH.A.Application;
					goto IL_0033;
					IL_0033:
					num2 = 6;
					application.ScreenUpdating = false;
					goto IL_003d;
					IL_003d:
					num2 = 7;
					enumerator = range2.GetEnumerator();
					goto IL_00b3;
					IL_00b3:
					if (enumerator.MoveNext())
					{
						range = (Range)enumerator.Current;
						goto IL_005b;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_00c8;
					IL_005b:
					num2 = 8;
					if (Base.IsString(range))
					{
						while (true)
						{
							switch (2)
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
						goto IL_0079;
					}
					goto IL_00b0;
					IL_00c8:
					num2 = 12;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_00ea;
					IL_00a7:
					num2 = 10;
					num5 = checked(num5 + 1);
					goto IL_00b0;
					IL_00b0:
					num2 = 11;
					goto IL_00b3;
					IL_00ea:
					num2 = 13;
					application.ScreenUpdating = true;
					goto IL_00f5;
					IL_00f5:
					application = null;
					goto IL_00f8;
					IL_00f8:
					num2 = 15;
					if (!flag)
					{
						break;
					}
					goto IL_00ff;
					IL_00ff:
					num2 = 16;
					JH.A(range2, VH.A(82757));
					break;
					end_IL_0000_2:
					break;
				}
				num2 = 17;
				range2 = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 372;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (5)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}

	public static void LettersUpper()
	{
		B(VH.A(148729));
	}

	public static void LettersLower()
	{
		B(VH.A(148740));
	}

	public static void RomanUpper()
	{
		B(VH.A(150045));
	}

	public static void RomanLower()
	{
		B(VH.A(150060));
	}

	private static void B(string A)
	{
		int try0000_dispatch = -1;
		int num2 = default(int);
		Range range2 = default(Range);
		int num = default(int);
		int num3 = default(int);
		Range range = default(Range);
		string[] array = default(string[]);
		int num5 = default(int);
		bool flag = default(bool);
		Application application = default(Application);
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			try
			{
				/*Note: ILSpy has introduced the following switch to emulate a goto from catch-block to try-block*/;
				switch (try0000_dispatch)
				{
				default:
					num2 = 1;
					range2 = JH.A((Range)null);
					goto IL_000b;
				case 2141:
					{
						num = num2;
						switch (num3)
						{
						case 1:
							break;
						default:
							goto end_IL_0000;
						}
						int num4 = num + 1;
						num = 0;
						switch (num4)
						{
						case 1:
							break;
						case 2:
							goto IL_000b;
						case 3:
							goto IL_0017;
						case 4:
							goto IL_001c;
						case 5:
							goto IL_0023;
						case 6:
							goto IL_0035;
						case 7:
							goto IL_003f;
						case 9:
							goto IL_00b4;
						case 11:
							goto IL_024a;
						case 13:
							goto IL_03e4;
						case 15:
							goto IL_0578;
						case 8:
						case 10:
						case 12:
						case 14:
						case 16:
						case 17:
							goto IL_0705;
						case 18:
							goto IL_0722;
						case 19:
							goto IL_073a;
						case 20:
							goto IL_0766;
						case 21:
							goto IL_076f;
						case 22:
							goto IL_0787;
						case 23:
							goto IL_07a9;
						case 24:
							goto IL_07b4;
						case 25:
							goto IL_07b7;
						case 26:
							goto IL_07be;
						case 27:
							goto end_IL_0000_2;
						default:
							goto end_IL_0000;
						case 28:
							goto end_IL_0000_3;
						}
						goto default;
					}
					IL_073a:
					num2 = 19;
					range.NumberFormat = VH.A(39830) + array[num5] + VH.A(150553);
					goto IL_0766;
					IL_000b:
					num2 = 2;
					flag = JH.A(range2);
					goto IL_0017;
					IL_0017:
					num2 = 3;
					num5 = 0;
					goto IL_001c;
					IL_001c:
					ProjectData.ClearProjectError();
					num3 = 1;
					goto IL_0023;
					IL_0023:
					num2 = 5;
					application = MH.A.Application;
					goto IL_0035;
					IL_0035:
					num2 = 6;
					application.ScreenUpdating = false;
					goto IL_003f;
					IL_003f:
					num2 = 7;
					if (Operators.CompareString(A, VH.A(150045), TextCompare: false) == 0)
					{
						goto IL_00b4;
					}
					while (true)
					{
						switch (2)
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
					if (Operators.CompareString(A, VH.A(150060), TextCompare: false) == 0)
					{
						goto IL_024a;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					if (Operators.CompareString(A, VH.A(148729), TextCompare: false) == 0)
					{
						goto IL_03e4;
					}
					goto IL_0578;
					IL_0766:
					num2 = 20;
					num5 = checked(num5 + 1);
					goto IL_076f;
					IL_07a9:
					num2 = 23;
					application.ScreenUpdating = true;
					goto IL_07b4;
					IL_07b7:
					num2 = 25;
					if (!flag)
					{
						break;
					}
					goto IL_07be;
					IL_07b4:
					application = null;
					goto IL_07b7;
					IL_07be:
					num2 = 26;
					JH.A(range2, VH.A(82757));
					break;
					IL_076f:
					num2 = 21;
					goto IL_0772;
					IL_0578:
					num2 = 15;
					array = new string[26]
					{
						VH.A(150484),
						VH.A(150487),
						VH.A(150490),
						VH.A(150493),
						VH.A(150496),
						VH.A(150499),
						VH.A(150502),
						VH.A(150505),
						VH.A(150251),
						VH.A(150508),
						VH.A(150511),
						VH.A(150514),
						VH.A(150517),
						VH.A(150520),
						VH.A(150523),
						VH.A(150526),
						VH.A(150529),
						VH.A(150532),
						VH.A(150535),
						VH.A(150538),
						VH.A(150541),
						VH.A(150271),
						VH.A(150544),
						VH.A(140671),
						VH.A(150547),
						VH.A(150550)
					};
					goto IL_0705;
					IL_03e4:
					num2 = 13;
					array = new string[26]
					{
						VH.A(57237),
						VH.A(77555),
						VH.A(57572),
						VH.A(150424),
						VH.A(150427),
						VH.A(150430),
						VH.A(150433),
						VH.A(150436),
						VH.A(150075),
						VH.A(150439),
						VH.A(150442),
						VH.A(150445),
						VH.A(150448),
						VH.A(150451),
						VH.A(150454),
						VH.A(150457),
						VH.A(150460),
						VH.A(150463),
						VH.A(150466),
						VH.A(150469),
						VH.A(150472),
						VH.A(150095),
						VH.A(150475),
						VH.A(150124),
						VH.A(150478),
						VH.A(150481)
					};
					goto IL_0705;
					IL_024a:
					num2 = 11;
					array = new string[26]
					{
						VH.A(150251),
						VH.A(150254),
						VH.A(150259),
						VH.A(150266),
						VH.A(150271),
						VH.A(150274),
						VH.A(150279),
						VH.A(150286),
						VH.A(150295),
						VH.A(140671),
						VH.A(150300),
						VH.A(150305),
						VH.A(150312),
						VH.A(150321),
						VH.A(150328),
						VH.A(150333),
						VH.A(150340),
						VH.A(150349),
						VH.A(150360),
						VH.A(150367),
						VH.A(150372),
						VH.A(150379),
						VH.A(150388),
						VH.A(150399),
						VH.A(150408),
						VH.A(150415)
					};
					goto IL_0705;
					IL_00b4:
					num2 = 9;
					array = new string[26]
					{
						VH.A(150075),
						VH.A(150078),
						VH.A(150083),
						VH.A(150090),
						VH.A(150095),
						VH.A(150098),
						VH.A(150103),
						VH.A(150110),
						VH.A(150119),
						VH.A(150124),
						VH.A(150127),
						VH.A(150132),
						VH.A(150139),
						VH.A(150148),
						VH.A(150155),
						VH.A(150160),
						VH.A(150167),
						VH.A(150176),
						VH.A(150187),
						VH.A(150194),
						VH.A(150199),
						VH.A(150206),
						VH.A(150215),
						VH.A(150226),
						VH.A(150235),
						VH.A(150242)
					};
					goto IL_0705;
					IL_0705:
					num2 = 17;
					enumerator = range2.GetEnumerator();
					goto IL_0772;
					IL_0772:
					if (enumerator.MoveNext())
					{
						range = (Range)enumerator.Current;
						goto IL_0722;
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						break;
					}
					goto IL_0787;
					IL_0722:
					num2 = 18;
					if (Base.IsString(range))
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
						goto IL_073a;
					}
					goto IL_076f;
					IL_0787:
					num2 = 22;
					if (enumerator is IDisposable)
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
						(enumerator as IDisposable).Dispose();
					}
					goto IL_07a9;
					end_IL_0000_2:
					break;
				}
				num2 = 27;
				range2 = null;
				break;
				end_IL_0000:;
			}
			catch (object obj) when (obj is Exception && num3 != 0 && num == 0)
			{
				ProjectData.SetProjectError((Exception)obj);
				try0000_dispatch = 2141;
				continue;
			}
			throw ProjectData.CreateProjectError(-2146828237);
			continue;
			end_IL_0000_3:
			break;
		}
		if (num == 0)
		{
			return;
		}
		while (true)
		{
			switch (6)
			{
			case 0:
				continue;
			}
			ProjectData.ClearProjectError();
			return;
		}
	}
}

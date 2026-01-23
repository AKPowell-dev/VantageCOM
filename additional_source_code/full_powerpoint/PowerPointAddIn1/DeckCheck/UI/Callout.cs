using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using MacabacusMacros.Proofing.UI;

namespace PowerPointAddIn1.DeckCheck.UI;

public sealed class Callout
{
	public static readonly int POINTER_X_OFFSET = 25;

	[CompilerGenerated]
	private static List<Rect> m_A;

	[CompilerGenerated]
	private static wpfCallout m_A;

	[CompilerGenerated]
	private static wpfMarchingAnts m_A;

	[CompilerGenerated]
	private static bool m_A;

	internal static List<Rect> DashBoxes
	{
		[CompilerGenerated]
		get
		{
			return Callout.m_A;
		}
		[CompilerGenerated]
		set
		{
			Callout.m_A = value;
		}
	}

	internal static wpfCallout Dialog
	{
		[CompilerGenerated]
		get
		{
			return Callout.m_A;
		}
		[CompilerGenerated]
		set
		{
			Callout.m_A = value;
		}
	}

	internal static wpfMarchingAnts MarchingAnts
	{
		[CompilerGenerated]
		get
		{
			return Callout.m_A;
		}
		[CompilerGenerated]
		set
		{
			Callout.m_A = value;
		}
	}

	internal static bool DoNotClose
	{
		[CompilerGenerated]
		get
		{
			return Callout.m_A;
		}
		[CompilerGenerated]
		set
		{
			Callout.m_A = value;
		}
	}

	internal static void A()
	{
		if (MarchingAnts == null)
		{
			return;
		}
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			MarchingAnts.CloseByCode = true;
			((Window)(object)MarchingAnts).Close();
			MarchingAnts = null;
			return;
		}
	}

	internal static void A(wpfCallout A, double B, double C)
	{
		wpfCallout wpfCallout2 = A;
		wpfCallout2.Left = B - (double)POINTER_X_OFFSET;
		wpfCallout2.Top = C - wpfCallout2.gridMain.ActualHeight - wpfCallout2.gridMain.Margin.Top;
		if (MarchingAnts != null)
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
			((Window)(object)MarchingAnts).Top = wpfCallout2.Top + wpfCallout2.ActualHeight;
			((Window)(object)MarchingAnts).Left = B - wpfCallout2.XOffset;
		}
		wpfCallout2 = null;
	}
}

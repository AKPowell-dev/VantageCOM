using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using A;
using MacabacusMacros;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointAddIn1.DeckCheck;

public sealed class Main
{
	[CompilerGenerated]
	private static Analysis m_A;

	public static Analysis Analysis
	{
		[CompilerGenerated]
		get
		{
			return Main.m_A;
		}
		[CompilerGenerated]
		set
		{
			Main.m_A = value;
		}
	}

	public static void EnsureSlidePaneActive()
	{
		DocumentWindow activeWindow = NG.A.Application.ActiveWindow;
		if (activeWindow.ActivePane.ViewType != PpViewType.ppViewSlide)
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = activeWindow.Panes.GetEnumerator();
				while (enumerator.MoveNext())
				{
					Pane pane = (Pane)enumerator.Current;
					if (pane.ViewType != PpViewType.ppViewSlide)
					{
						continue;
					}
					while (true)
					{
						switch (1)
						{
						case 0:
							continue;
						}
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						pane.Activate();
						break;
					}
					break;
				}
			}
			finally
			{
				if (enumerator is IDisposable)
				{
					while (true)
					{
						switch (2)
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
		activeWindow = null;
	}

	internal static void A(Exception A, int? B, Chart C, int[] D = null)
	{
		if (D != null)
		{
			if (!clsCharts.A(C, D))
			{
				goto IL_0087;
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
		}
		if (A is NotImplementedException)
		{
			return;
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
		if (B.HasValue)
		{
			COMException obj = A as COMException;
			int? obj2;
			if (obj == null)
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
				obj2 = null;
			}
			else
			{
				obj2 = obj.ErrorCode;
			}
			if (object.Equals(obj2, B.Value))
			{
				return;
			}
		}
		goto IL_0087;
		IL_0087:
		clsReporting.LogException((Exception)new TimeZoneNotFoundException(string.Format(AH.A(49065), C.ChartType, A.Message), A));
	}
}

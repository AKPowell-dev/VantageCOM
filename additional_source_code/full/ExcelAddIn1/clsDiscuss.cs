using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class clsDiscuss
{
	public enum Storage
	{
		Local = 1,
		Remote,
		Embedded,
		WebLink
	}

	private static readonly string m_A = VH.A(197950);

	private static readonly string m_B = VH.A(197979);

	public static readonly string HIDDEN_SHEET_NAME = VH.A(198002);

	private static bool m_A;

	private static readonly string C = VH.A(198035);

	private static bool A
	{
		get
		{
			return clsDiscuss.m_A;
		}
		set
		{
			clsDiscuss.m_A = value;
		}
	}

	public static void DiscussPaneToggle(bool blnPressed)
	{
		try
		{
			if (blnPressed)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						A();
						clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)15, VH.A(197778));
						return;
					}
				}
			}
			B();
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}

	private static void A()
	{
		//IL_0045: Unknown result type (might be due to invalid IL or missing references)
		//IL_004b: Expected O, but got Unknown
		ctpDiscuss control = new ctpDiscuss();
		clsDiscuss.A = true;
		try
		{
			CustomTaskPane customTaskPane = MH.A.CustomTaskPanes.Add(control, clsDiscuss.m_A, MH.A.Application.ActiveWindow);
			try
			{
				customTaskPane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionRight;
				customTaskPane.DockPositionRestrict = MsoCTPDockPositionRestrict.msoCTPDockPositionRestrictNoChange;
				clsDisplay val = new clsDisplay();
				customTaskPane.Width = checked((int)Math.Round(410.0 * val.X));
				val = null;
				customTaskPane.Visible = true;
				customTaskPane.VisibleChanged += A;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				clsReporting.LogException(ex2);
				ProjectData.ClearProjectError();
			}
			customTaskPane = null;
		}
		catch (Exception ex3)
		{
			ProjectData.SetProjectError(ex3);
			Exception ex4 = ex3;
			clsReporting.LogException(ex4);
			ProjectData.ClearProjectError();
		}
	}

	private static void B()
	{
		clsPanes.A(clsDiscuss.m_A);
		clsDiscuss.A = false;
		KH.A.InvalidateControl(clsDiscuss.m_B);
	}

	private static void A(object A, EventArgs B)
	{
		if ((A as CustomTaskPane).Visible)
		{
			return;
		}
		while (true)
		{
			switch (2)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (KH.A)
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
				clsDiscuss.B();
				return;
			}
		}
	}

	public static bool IsDiscussPaneOpen()
	{
		return clsDiscuss.A;
	}

	public static void VerifySources()
	{
		List<string> source = new List<string>();
		List<string> source2 = new List<string>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = MH.A.Application.ActiveWorkbook.CustomXMLParts.GetEnumerator();
			while (enumerator.MoveNext())
			{
				CustomXMLPart customXMLPart = (CustomXMLPart)enumerator.Current;
				if (customXMLPart.XML.Contains(C))
				{
					XPathQuery(customXMLPart);
				}
			}
			while (true)
			{
				switch (6)
				{
				case 0:
					continue;
				}
				if (1 == 0)
				{
					/*OpCode not supported: LdMemberToken*/;
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
					switch (6)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		if (!source.Any() && !source2.Any())
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
			MessageBox.Show(VH.A(197813), VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Asterisk);
		}
		source = null;
		source2 = null;
	}

	public static string XPathQuery(CustomXMLPart part)
	{
		string text = part.NamespaceManager.LookupPrefix(part.NamespaceURI);
		return VH.A(197945) + text + VH.A(2826);
	}
}

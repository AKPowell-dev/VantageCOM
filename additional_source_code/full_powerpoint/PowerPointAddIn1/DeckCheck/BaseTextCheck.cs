using System;
using System.Collections;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Text.RegularExpressions;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.DeckCheck;

public abstract class BaseTextCheck : BaseCheck
{
	[CompilerGenerated]
	private Regex m_A;

	[CompilerGenerated]
	private string m_A;

	internal Regex RegexObj
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

	internal string Fix
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

	public abstract void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp, TextRange2 para, string strText);

	public override void Check(Slide sld, Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
	}

	internal List<TextRange2> A(TextRange2 A, string B, Regex C, int D)
	{
		List<TextRange2> list = new List<TextRange2>();
		try
		{
			IEnumerator enumerator = default(IEnumerator);
			try
			{
				enumerator = C.Matches(B).GetEnumerator();
				while (enumerator.MoveNext())
				{
					Group obj = ((Match)enumerator.Current).Groups[D];
					list.Add(A.get_Characters(checked(obj.Index + 1), obj.Length));
					obj = null;
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
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return list;
	}
}

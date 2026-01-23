using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Proofing;

public sealed class Language
{
	public static string LanguagesMenu()
	{
		StringBuilder stringBuilder = new StringBuilder(XC.A(36369));
		List<int> list = new List<int>();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = InputLanguage.InstalledInputLanguages.GetEnumerator();
			while (enumerator.MoveNext())
			{
				CultureInfo culture = ((InputLanguage)enumerator.Current).Culture;
				if (!list.Contains(culture.LCID))
				{
					while (true)
					{
						switch (3)
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
					stringBuilder.Append(XC.A(36507) + culture.LCID + XC.A(36548) + culture.DisplayName + XC.A(36567) + culture.LCID + XC.A(36628));
					list.Add(culture.LCID);
				}
				culture = null;
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
		list = null;
		stringBuilder.Append(XC.A(37850));
		return stringBuilder.ToString();
	}

	public static void SetProofingLanguage(IRibbonControl control)
	{
		Conversions.ToInteger(control.Tag);
		_ = PC.A.Application.ActiveDocument;
		PC.A.Application.UndoRecord.StartCustomRecord(XC.A(37865));
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)12, XC.A(37900));
	}
}

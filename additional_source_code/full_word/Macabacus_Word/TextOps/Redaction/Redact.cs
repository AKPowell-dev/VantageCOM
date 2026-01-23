using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros.Auth;
using MacabacusMacros.UI;
using Macabacus_Word.TextOps.Redaction.Process;
using Macabacus_Word.TextOps.Redaction.Values;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.TextOps.Redaction;

public sealed class Redact
{
	[Serializable]
	[CompilerGenerated]
	internal sealed class _Closure_0024__
	{
		public static readonly _Closure_0024__ A;

		public static Func<Paragraph, bool> A;

		static _Closure_0024__()
		{
			_Closure_0024__.A = new _Closure_0024__();
		}

		[SpecialName]
		internal bool A(Paragraph A)
		{
			return !string.IsNullOrWhiteSpace(A.Range.Text);
		}
	}

	public static void RedactSelection()
	{
		if (!A())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			wpfRedactSelection wpfRedactSelection = new wpfRedactSelection();
			IEnumerable<Paragraph> source = PC.A.Application.Selection.Paragraphs.Cast<Paragraph>();
			Func<Paragraph, bool> predicate;
			if (_Closure_0024__.A == null)
			{
				predicate = (_Closure_0024__.A = [SpecialName] (Paragraph A) => !string.IsNullOrWhiteSpace(A.Range.Text));
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
				predicate = _Closure_0024__.A;
			}
			if (source.Count(predicate) > 19)
			{
				wpfRedactSelection.Show();
			}
			else
			{
				wpfRedactSelection.RedactNoProgressBar();
			}
			wpfRedactSelection = null;
			return;
		}
	}

	public static void FindAndRedact()
	{
		if (!A())
		{
			return;
		}
		SelectionValue selectionValue = new SelectionValue(PC.A.Application);
		wpfFindAndRedact wpfFindAndRedact = new wpfFindAndRedact(selectionValue);
		string text = "";
		try
		{
			RedactUtilities.ExpandAtInsertionPoint(selectionValue.WdApp.Selection);
			if (selectionValue.SelType == WdSelectionType.wdSelectionNormal)
			{
				IEnumerable<Range> wordList = RedactUtilities.GetWordList(selectionValue.WdApp.Selection.Range, trimList: true);
				if (wordList.Count() == 1)
				{
					text = wordList.ElementAtOrDefault(0).Text;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			text = "";
			ProjectData.ClearProjectError();
		}
		wpfFindAndRedact wpfFindAndRedact2 = wpfFindAndRedact;
		wpfFindAndRedact2.txtFind.Text = text;
		wpfFindAndRedact2.txtFind.Focus();
		wpfFindAndRedact2.txtFind.SelectAll();
		wpfFindAndRedact2.ShowDialog();
		_ = null;
		wpfFindAndRedact = null;
		selectionValue = null;
	}

	private static bool A()
	{
		if (B())
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
			if (E())
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
				if (!C())
				{
					if (!D())
					{
						return true;
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
				}
			}
		}
		return false;
	}

	private static bool B()
	{
		if (PC.A.Application.Documents.Count > 0)
		{
			while (true)
			{
				switch (6)
				{
				case 0:
					break;
				default:
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					return true;
				}
			}
		}
		Forms.WarningMessage(XC.A(19665));
		return false;
	}

	private static bool C()
	{
		bool result;
		try
		{
			if (PC.A.Application.ActiveDocument.ProtectionType != WdProtectionType.wdNoProtection)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(XC.A(19786));
					result = true;
					break;
				}
			}
			else
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool D()
	{
		bool result;
		try
		{
			if (PC.A.Application.ActiveDocument.ReadOnly)
			{
				while (true)
				{
					switch (3)
					{
					case 0:
						continue;
					}
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					Forms.WarningMessage(XC.A(19901));
					result = true;
					break;
				}
			}
			else
			{
				result = false;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			result = false;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool E()
	{
		return Access.AllowWordOperation((PlanType)5, (Restriction)1, false);
	}
}

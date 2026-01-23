using System;
using System.CodeDom.Compiler;
using System.Diagnostics;
using Microsoft.Office.Tools;
using PowerPointAddIn1;

namespace A;

[GeneratedCode("Microsoft.VisualStudio.Tools.Office.ProgrammingModel.dll", "15.0.0.0")]
[DebuggerNonUserCode]
internal sealed class NG
{
	private static ThisAddIn m_A;

	private static Factory m_A;

	private static LG m_A;

	internal static ThisAddIn A
	{
		get
		{
			return NG.m_A;
		}
		set
		{
			if (NG.m_A == null)
			{
				NG.m_A = value;
				return;
			}
			throw new NotSupportedException();
		}
	}

	internal static Factory A
	{
		get
		{
			return NG.m_A;
		}
		set
		{
			if (NG.m_A == null)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						if (1 == 0)
						{
							/*OpCode not supported: LdMemberToken*/;
						}
						NG.m_A = value;
						return;
					}
				}
			}
			throw new NotSupportedException();
		}
	}

	internal static LG A
	{
		get
		{
			if (NG.m_A == null)
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
				NG.m_A = new LG(NG.m_A.GetRibbonFactory());
			}
			return NG.m_A;
		}
	}

	private NG()
	{
	}
}

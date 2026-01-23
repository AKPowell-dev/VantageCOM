using System;
using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Presentation;

namespace PowerPointAddIn1.Library2.Insert;

public sealed class Common
{
	internal static Microsoft.Office.Interop.PowerPoint.Presentation A(string A, Application B, [Optional][DefaultParameterValue(null)] ref List<string> C)
	{
		Microsoft.Office.Interop.PowerPoint.Presentation presentation = null;
		try
		{
			presentation = B.Presentations[Path.GetFileName(A)];
			if (Operators.CompareString(presentation.FullName, A, TextCompare: false) != 0)
			{
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
					presentation = null;
					break;
				}
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		if (presentation == null)
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
			try
			{
				presentation = Helpers.OpenQuietly(B, A);
				if (presentation != null && C != null)
				{
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						C.Add(A);
						break;
					}
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
		return presentation;
	}
}

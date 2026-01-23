using System;
using System.IO;
using A;
using ExcelAddIn1.ExcelApp;
using MacabacusMacros;
using MacabacusMacros.Libraries;
using MacabacusMacros.UI;
using Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1.Library2;

public sealed class Models
{
	public static string BuildFilesMenu()
	{
		return Ribbon.BuildFilesMenu(Base.LIB_MODELS_FOLDER_NAME, VH.A(86378));
	}

	public static void OpenFile(string strPath)
	{
		Application application = MH.A.Application;
		try
		{
			string text = Path.Combine(Interaction.Environ(VH.A(86342)), VH.A(86439) + Path.GetExtension(strPath));
			File.Copy(strPath, text, overwrite: true);
			application.Workbooks.Add(text);
			File.Delete(text);
			clsReporting.LogActivity((ActivityApp)1, (ActivityCategory)6, VH.A(86448));
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			if (!EditMode.IsEditMode(application))
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
				Forms.ErrorMessage(ex2.Message);
				clsReporting.LogException(ex2);
			}
			ProjectData.ClearProjectError();
		}
		application = null;
	}
}

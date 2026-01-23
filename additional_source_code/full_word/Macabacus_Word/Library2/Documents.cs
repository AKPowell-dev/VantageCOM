using System;
using System.Reflection;
using System.Runtime.CompilerServices;
using A;
using MacabacusMacros;
using MacabacusMacros.Auth;
using MacabacusMacros.Libraries;
using MacabacusMacros.UI;
using Macabacus_Word.DocBuilder;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic.CompilerServices;

namespace Macabacus_Word.Library2;

public sealed class Documents
{
	public static string BuildFilesMenu()
	{
		return Ribbon.BuildFilesMenu(Base.LIB_DOCUMENTS_FOLDER_NAME, XC.A(9537));
	}

	public static void OpenFile(string strPath)
	{
		//IL_0067: Unknown result type (might be due to invalid IL or missing references)
		//IL_006c: Unknown result type (might be due to invalid IL or missing references)
		//IL_006f: Invalid comparison between Unknown and I4
		try
		{
			Microsoft.Office.Interop.Word.Documents documents = PC.A.Application.Documents;
			object Template = strPath;
			object NewTemplate = RuntimeHelpers.GetObjectValue(Missing.Value);
			object DocumentType = RuntimeHelpers.GetObjectValue(Missing.Value);
			object Visible = RuntimeHelpers.GetObjectValue(Missing.Value);
			Document document = documents.Add(ref Template, ref NewTemplate, ref DocumentType, ref Visible);
			strPath = Conversions.ToString(Template);
			Document doc = document;
			if (Access.IsEnterprisePlanOrTrialMode() && (int)Base.UserProfile.LicenseType == 2)
			{
				Base.InspectDocument(doc, blnManual: false);
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			Forms.ErrorMessage(ex2.Message);
			ProjectData.ClearProjectError();
		}
		finally
		{
			Document doc = null;
		}
		clsReporting.LogActivity((ActivityApp)3, (ActivityCategory)6, XC.A(9596));
	}
}

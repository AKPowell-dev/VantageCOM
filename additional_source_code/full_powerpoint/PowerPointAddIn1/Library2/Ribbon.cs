using System;
using A;
using Microsoft.VisualBasic.CompilerServices;
using PowerPointAddIn1.Library2.UI;

namespace PowerPointAddIn1.Library2;

public sealed class Ribbon
{
	public static bool PictureOverride()
	{
		bool result = false;
		if (!Pane.IsVisible() && A(AH.A(67974)))
		{
			try
			{
				Pane.Show(blnSlides: false, blnShapes: false, blnImages: true, blnCharts: false, blnText: false, blnDecks: false, blnVideos: false);
				result = true;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		return result;
	}

	public static bool ChartOverride()
	{
		bool result = false;
		if (!Pane.IsVisible())
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			if (A(AH.A(67991)))
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
					Pane.Show(blnSlides: false, blnShapes: false, blnImages: false, blnCharts: true, blnText: false, blnDecks: false, blnVideos: false);
					result = true;
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
			}
		}
		return result;
	}

	private static bool A(string A)
	{
		bool result;
		try
		{
			result = Conversions.ToBoolean(KG.A.SettingsXml.DocumentElement.SelectSingleNode(AH.A(68004) + A).InnerText);
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
}

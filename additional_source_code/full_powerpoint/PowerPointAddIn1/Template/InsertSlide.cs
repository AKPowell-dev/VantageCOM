using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Windows;
using A;
using MacabacusMacros.UI;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Template;

public sealed class InsertSlide
{
	[CompilerGenerated]
	private static Dictionary<string, ObservableCollection<LayoutThumb>> A;

	public static Dictionary<string, ObservableCollection<LayoutThumb>> LayoutThumbnails
	{
		[CompilerGenerated]
		get
		{
			return A;
		}
		[CompilerGenerated]
		set
		{
			A = value;
		}
	} = null;

	public static void ShowDialog()
	{
		if (!Licensing.AllowTemplateOperation())
		{
			return;
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
			if (IG.A(NG.A.Application.Presentations) > 0)
			{
				while (true)
				{
					switch (6)
					{
					case 0:
						break;
					default:
					{
						bool flag = false;
						try
						{
							IEnumerable<wpfInsertSlide> source = Application.Current.Windows.OfType<wpfInsertSlide>();
							if (source.Any())
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
								source.ElementAt(0).Activate();
								flag = true;
							}
							source = null;
						}
						catch (Exception ex)
						{
							ProjectData.SetProjectError(ex);
							Exception ex2 = ex;
							ProjectData.ClearProjectError();
						}
						if (!flag)
						{
							while (true)
							{
								switch (7)
								{
								case 0:
									break;
								default:
									if (LayoutThumbnails == null)
									{
										while (true)
										{
											switch (1)
											{
											case 0:
												continue;
											}
											break;
										}
										LayoutThumbnails = new Dictionary<string, ObservableCollection<LayoutThumb>>();
									}
									new wpfInsertSlide().Show();
									_ = null;
									return;
								}
							}
						}
						return;
					}
					}
				}
			}
			Forms.WarningMessage(AH.A(59235));
			return;
		}
	}
}

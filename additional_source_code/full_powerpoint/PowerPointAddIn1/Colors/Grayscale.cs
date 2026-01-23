using System;
using System.Collections;
using A;
using MacabacusMacros;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Colors;

public sealed class Grayscale
{
	internal static void A(Microsoft.Office.Interop.PowerPoint.Presentation A = null)
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		if (A == null)
		{
			try
			{
				A = NG.A.Application.ActivePresentation;
			}
			catch (Exception ex)
			{
				ProjectData.SetProjectError(ex);
				Exception ex2 = ex;
				ProjectData.ClearProjectError();
			}
		}
		if (A == null)
		{
			return;
		}
		NG.A.Application.StartNewUndoEntry();
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.Slides.GetEnumerator();
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				Slide slide = (Slide)enumerator.Current;
				{
					enumerator2 = slide.Shapes.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							Grayscale.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current);
						}
						while (true)
						{
							switch (7)
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
						IDisposable disposable = enumerator2 as IDisposable;
						if (disposable != null)
						{
							disposable.Dispose();
						}
					}
				}
			}
			while (true)
			{
				switch (3)
				{
				case 0:
					break;
				default:
					goto end_IL_00c2;
				}
				continue;
				end_IL_00c2:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		A = null;
		clsReporting.LogActivity((ActivityApp)2, (ActivityCategory)1, AH.A(13625));
	}

	internal static void A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		try
		{
			MsoShapeType type = A.Type;
			if (type <= MsoShapeType.msoLine)
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
				if (type <= MsoShapeType.msoChart)
				{
					if (type != MsoShapeType.msoAutoShape)
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
						if (type != MsoShapeType.msoChart)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									break;
								default:
									return;
								}
							}
						}
					}
				}
				else
				{
					if (type == MsoShapeType.msoGroup)
					{
						IEnumerator enumerator = default(IEnumerator);
						try
						{
							enumerator = A.GroupItems.GetEnumerator();
							while (enumerator.MoveNext())
							{
								Grayscale.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current);
							}
							return;
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
					while (true)
					{
						switch (7)
						{
						case 0:
							continue;
						}
						break;
					}
					if (type != MsoShapeType.msoLine)
					{
						return;
					}
				}
			}
			else if (type <= MsoShapeType.msoTextBox)
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
				if (type != MsoShapeType.msoPlaceholder)
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
					if (type != MsoShapeType.msoTextBox)
					{
						while (true)
						{
							switch (5)
							{
							case 0:
								break;
							default:
								return;
							}
						}
					}
				}
			}
			else if (type != MsoShapeType.msoTable)
			{
				while (true)
				{
					switch (7)
					{
					case 0:
						continue;
					}
					break;
				}
				if (type != MsoShapeType.msoSmartArt)
				{
					while (true)
					{
						switch (7)
						{
						case 0:
							break;
						default:
							return;
						}
					}
				}
			}
			A.BlackWhiteMode = MsoBlackWhiteMode.msoBlackWhiteGrayScale;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
	}
}

using System;
using System.Collections;
using System.Runtime.CompilerServices;
using System.Text;
using System.Xml;
using A;
using MacabacusMacros;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Interop.Word;
using Microsoft.VisualBasic;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Resize
{
	public static void StandardSize(int i)
	{
		//IL_0028: Unknown result type (might be due to invalid IL or missing references)
		//IL_002d: Unknown result type (might be due to invalid IL or missing references)
		//IL_0131: Unknown result type (might be due to invalid IL or missing references)
		//IL_0143: Unknown result type (might be due to invalid IL or missing references)
		//IL_00ff: Unknown result type (might be due to invalid IL or missing references)
		//IL_0115: Unknown result type (might be due to invalid IL or missing references)
		//IL_0157: Unknown result type (might be due to invalid IL or missing references)
		//IL_016b: Unknown result type (might be due to invalid IL or missing references)
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
		while (true)
		{
			switch (4)
			{
			case 0:
				continue;
			}
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			bool flag = false;
			Microsoft.Office.Interop.PowerPoint.ShapeRange shapeRange;
			try
			{
				shapeRange = Base.SelectedShapes();
				try
				{
					StandardSize standardSize = clsPublish.GetStandardSize(i);
					NG.A.Application.StartNewUndoEntry();
					try
					{
						Microsoft.Office.Interop.PowerPoint.Shape shape2;
						for (enumerator = shapeRange.GetEnumerator(); enumerator.MoveNext(); shape2 = null)
						{
							Microsoft.Office.Interop.PowerPoint.Shape shape = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							shape2 = shape;
							if (shape2.Type != MsoShapeType.msoPicture)
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
								if (shape2.Type != MsoShapeType.msoLinkedPicture)
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
									if (!Images.IsGraphic(shape))
									{
										shape2.Height = clsPublish.InchesToPoints(standardSize.Height);
										shape2.Width = clsPublish.InchesToPoints(standardSize.Width);
										continue;
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
								}
							}
							if (!flag)
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
								wpfPictureSizePrompt obj = new wpfPictureSizePrompt();
								obj.ShowDialog();
								flag = obj.chkApplyAll.IsChecked.Value;
								_ = null;
							}
							switch (PB.Settings.PictureSizePrompt)
							{
							case 1:
								shape2.Width = clsPublish.InchesToPoints(standardSize.Width);
								break;
							case 2:
								shape2.Height = clsPublish.InchesToPoints(standardSize.Height);
								break;
							case 0:
								shape2.LockAspectRatio = MsoTriState.msoFalse;
								shape2.Height = clsPublish.InchesToPoints(standardSize.Height);
								shape2.Width = clsPublish.InchesToPoints(standardSize.Width);
								break;
							}
						}
						while (true)
						{
							switch (3)
							{
							case 0:
								break;
							default:
								goto end_IL_018b;
							}
							continue;
							end_IL_018b:
							break;
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
				}
				catch (Exception ex)
				{
					ProjectData.SetProjectError(ex);
					Exception ex2 = ex;
					ProjectData.ClearProjectError();
				}
				Base.LogActivity(AH.A(83115));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Base.AlignError();
				ProjectData.ClearProjectError();
			}
			shapeRange = null;
			return;
		}
	}

	internal static void A()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = null;
		try
		{
			application = InstanceManagement.GetExcelInstance(false);
			if (application == null)
			{
				throw new Exception();
			}
			try
			{
				object objectValue = RuntimeHelpers.GetObjectValue(application.Selection);
				float height = Conversions.ToSingle(NewLateBinding.LateGet(objectValue, null, AH.A(83176), new object[0], null, null, null));
				float width = Conversions.ToSingle(NewLateBinding.LateGet(objectValue, null, AH.A(83189), new object[0], null, null, null));
				try
				{
					NG.A.Application.StartNewUndoEntry();
					IEnumerator enumerator = default(IEnumerator);
					try
					{
						enumerator = Base.SelectedShapes().GetEnumerator();
						while (enumerator.MoveNext())
						{
							Microsoft.Office.Interop.PowerPoint.Shape obj = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
							obj.Height = height;
							obj.Width = width;
						}
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
							break;
						}
					}
					finally
					{
						if (enumerator is IDisposable)
						{
							while (true)
							{
								switch (7)
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
				Base.LogActivity(AH.A(83200));
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				Forms.WarningMessage(AH.A(82118));
				ProjectData.ClearProjectError();
			}
		}
		catch (Exception ex5)
		{
			ProjectData.SetProjectError(ex5);
			Exception ex6 = ex5;
			Forms.WarningMessage(AH.A(83263));
			ProjectData.ClearProjectError();
		}
		if (application != null)
		{
			JG.A(application);
			application = null;
		}
	}

	internal static void B()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		IEnumerator enumerator = default(IEnumerator);
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
			try
			{
				Microsoft.Office.Interop.Word.Application application = (Microsoft.Office.Interop.Word.Application)Interaction.GetObject(null, AH.A(82290));
				if (application == null)
				{
					while (true)
					{
						switch (5)
						{
						case 0:
							break;
						default:
							throw new Exception();
						}
					}
				}
				try
				{
					float[] array = clsPublish.SelectedWordShapeSize(application);
					if (!(array[0] > 0f))
					{
						throw new Exception();
					}
					while (true)
					{
						switch (3)
						{
						case 0:
							continue;
						}
						try
						{
							NG.A.Application.StartNewUndoEntry();
							try
							{
								enumerator = Base.SelectedShapes().GetEnumerator();
								while (enumerator.MoveNext())
								{
									Microsoft.Office.Interop.PowerPoint.Shape obj = (Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current;
									obj.Height = array[1];
									obj.Width = array[0];
								}
								while (true)
								{
									switch (2)
									{
									case 0:
										break;
									default:
										goto end_IL_00b0;
									}
									continue;
									end_IL_00b0:
									break;
								}
							}
							finally
							{
								if (enumerator is IDisposable)
								{
									while (true)
									{
										switch (4)
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
						Base.LogActivity(AH.A(83336));
						break;
					}
				}
				catch (Exception ex3)
				{
					ProjectData.SetProjectError(ex3);
					Exception ex4 = ex3;
					Forms.WarningMessage(AH.A(83397));
					ProjectData.ClearProjectError();
				}
				application = null;
				return;
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				Forms.WarningMessage(AH.A(83448));
				ProjectData.ClearProjectError();
				return;
			}
		}
	}

	internal static string A()
	{
		int num = 1;
		StringBuilder stringBuilder = new StringBuilder(AH.A(47526));
		stringBuilder.Append(AH.A(83519));
		stringBuilder.Append(AH.A(83636));
		stringBuilder.Append(AH.A(84102));
		stringBuilder.Append(AH.A(84556));
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = clsPublish.StandardSizeNodes().GetEnumerator();
			while (enumerator.MoveNext())
			{
				string text = ((XmlNode)enumerator.Current).Attributes[AH.A(82505)].Value.Replace(AH.A(82514), AH.A(82517));
				string text2 = num + AH.A(82538) + text;
				if (num < 10)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					text2 = AH.A(82543) + num + AH.A(82538) + text;
				}
				else if (num == 10)
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
					text2 = AH.A(84681) + text;
				}
				else
				{
					text2 = num + AH.A(82538) + text;
				}
				stringBuilder.Append(AH.A(84700) + num + AH.A(47705) + text2 + AH.A(84765) + num + AH.A(82654) + text + AH.A(84826));
				num = checked(num + 1);
			}
			while (true)
			{
				switch (2)
				{
				case 0:
					break;
				default:
					goto end_IL_01f0;
				}
				continue;
				end_IL_01f0:
				break;
			}
		}
		finally
		{
			if (enumerator is IDisposable)
			{
				while (true)
				{
					switch (1)
					{
					case 0:
						continue;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
		stringBuilder.Append(AH.A(49007));
		return stringBuilder.ToString();
	}
}

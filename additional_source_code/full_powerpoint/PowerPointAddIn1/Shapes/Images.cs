using System;
using System.Collections;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows.Forms;
using A;
using MacabacusMacros.UI;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.VisualBasic.CompilerServices;

namespace PowerPointAddIn1.Shapes;

public sealed class Images
{
	internal static readonly int A = 28;

	internal static readonly int B = 29;

	public static bool HasPictureOrGraphic(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result = false;
		try
		{
			Microsoft.Office.Interop.PowerPoint.Shape shape = shp;
			switch (shape.Type)
			{
			case MsoShapeType.msoLinkedPicture:
			case MsoShapeType.msoPicture:
				result = true;
				break;
			case MsoShapeType.msoPlaceholder:
			{
				MsoShapeType containedType = shape.PlaceholderFormat.ContainedType;
				if (containedType != MsoShapeType.msoLinkedPicture)
				{
					if (containedType != MsoShapeType.msoPicture)
					{
						if (shape.PlaceholderFormat.ContainedType == (MsoShapeType)Images.A || shape.PlaceholderFormat.ContainedType == (MsoShapeType)Images.B)
						{
							result = true;
						}
						break;
					}
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
				}
				result = true;
				break;
			}
			default:
				if (shape.Type != (MsoShapeType)Images.A)
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
					if (shape.Type != (MsoShapeType)Images.B)
					{
						break;
					}
				}
				result = true;
				break;
			}
			shape = null;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	private static bool A(Microsoft.Office.Interop.PowerPoint.Shape A)
	{
		bool result = false;
		try
		{
			MsoShapeType type = A.Type;
			if (type == MsoShapeType.msoLinkedPicture)
			{
				goto IL_0032;
			}
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
			if (type == MsoShapeType.msoPicture)
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
				goto IL_0032;
			}
			if (A.Type == (MsoShapeType)Images.A)
			{
				goto IL_0068;
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
			if (A.Type == (MsoShapeType)Images.B)
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
				goto IL_0068;
			}
			goto end_IL_0002;
			IL_0068:
			result = true;
			goto end_IL_0002;
			IL_0032:
			result = true;
			end_IL_0002:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool HasGraphic(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result = false;
		try
		{
			if (shp.Type == (MsoShapeType)Images.A)
			{
				goto IL_0033;
			}
			while (true)
			{
				switch (6)
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
			if (shp.Type == (MsoShapeType)Images.B)
			{
				goto IL_0033;
			}
			if (shp.Type == MsoShapeType.msoPlaceholder)
			{
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					if (shp.PlaceholderFormat.ContainedType != (MsoShapeType)Images.A)
					{
						if (shp.PlaceholderFormat.ContainedType != (MsoShapeType)Images.B)
						{
							break;
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
					result = true;
					break;
				}
			}
			goto end_IL_0002;
			IL_0033:
			result = true;
			end_IL_0002:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool IsGraphic(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result = false;
		try
		{
			if (shp.Type == (MsoShapeType)Images.A)
			{
				goto IL_0031;
			}
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
			if (shp.Type == (MsoShapeType)Images.B)
			{
				goto IL_0031;
			}
			goto end_IL_0002;
			IL_0031:
			result = true;
			end_IL_0002:;
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool HasPicture(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool result = false;
		try
		{
			switch (shp.Type)
			{
			case MsoShapeType.msoLinkedPicture:
			case MsoShapeType.msoPicture:
				result = true;
				break;
			case MsoShapeType.msoPlaceholder:
			{
				MsoShapeType containedType = shp.PlaceholderFormat.ContainedType;
				if (containedType != MsoShapeType.msoLinkedPicture)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (containedType != MsoShapeType.msoPicture)
					{
						break;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				result = true;
				break;
			}
			case MsoShapeType.msoOLEControlObject:
				break;
			}
		}
		catch (Exception ex)
		{
			ProjectData.SetProjectError(ex);
			Exception ex2 = ex;
			ProjectData.ClearProjectError();
		}
		return result;
	}

	public static bool HasPictureOrOLE(Microsoft.Office.Interop.PowerPoint.Shape shp)
	{
		bool flag = HasPicture(shp);
		if (!flag)
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
			if (1 == 0)
			{
				/*OpCode not supported: LdMemberToken*/;
			}
			MsoShapeType type = shp.Type;
			if (type != MsoShapeType.msoEmbeddedOLEObject)
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
				if (type != MsoShapeType.msoLinkedOLEObject)
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
					if (type != MsoShapeType.msoOLEControlObject)
					{
						goto IL_004a;
					}
				}
			}
			flag = true;
		}
		goto IL_004a;
		IL_004a:
		return flag;
	}

	public static void FixScale()
	{
		if (!Licensing.AllowRestrictedMode())
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Application application = NG.A.Application;
		int B = 0;
		bool flag = false;
		try
		{
			Selection selection = application.ActiveWindow.Selection;
			_ = selection.SlideRange;
			if (selection.Type == PpSelectionType.ppSelectionShapes)
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
					flag = true;
					try
					{
						bool C = A();
						application.StartNewUndoEntry();
						if (selection.HasChildShapeRange)
						{
							while (true)
							{
								switch (6)
								{
								case 0:
									continue;
								}
								A(selection.ChildShapeRange, ref B, ref C);
								break;
							}
						}
						else
						{
							A(selection.ShapeRange, ref B, ref C);
						}
					}
					catch (InvalidTimeZoneException ex)
					{
						ProjectData.SetProjectError(ex);
						InvalidTimeZoneException ex2 = ex;
						flag = false;
						ProjectData.ClearProjectError();
					}
					break;
				}
			}
			else
			{
				try
				{
					if (selection.SlideRange.Count == 1)
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
						flag = MessageBox.Show(AH.A(74542), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK;
					}
					else if (selection.SlideRange.Count > 0)
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
						flag = MessageBox.Show(AH.A(74688), AH.A(5874), MessageBoxButtons.OKCancel, MessageBoxIcon.Exclamation) == DialogResult.OK;
					}
					if (flag)
					{
						IEnumerator enumerator = default(IEnumerator);
						while (true)
						{
							switch (3)
							{
							case 0:
								continue;
							}
							try
							{
								bool C = A();
								application.StartNewUndoEntry();
								try
								{
									enumerator = selection.SlideRange.GetEnumerator();
									while (enumerator.MoveNext())
									{
										A(((Slide)enumerator.Current).Shapes.Range(RuntimeHelpers.GetObjectValue(Missing.Value)), ref B, ref C);
									}
									while (true)
									{
										switch (7)
										{
										case 0:
											break;
										default:
											goto end_IL_0197;
										}
										continue;
										end_IL_0197:
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
							catch (InvalidTimeZoneException ex3)
							{
								ProjectData.SetProjectError(ex3);
								InvalidTimeZoneException ex4 = ex3;
								flag = false;
								ProjectData.ClearProjectError();
							}
							break;
						}
					}
				}
				catch (Exception ex5)
				{
					ProjectData.SetProjectError(ex5);
					Exception ex6 = ex5;
					ProjectData.ClearProjectError();
				}
			}
		}
		catch (Exception ex7)
		{
			ProjectData.SetProjectError(ex7);
			Exception ex8 = ex7;
			ProjectData.ClearProjectError();
		}
		if (flag)
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
			if (B > 0)
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
				Forms.SuccessMessage(AH.A(74852) + B + AH.A(74901));
			}
			else
			{
				Forms.InfoMessage(AH.A(74922));
			}
			Base.LogActivity(AH.A(75031));
		}
		application = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.ShapeRange A, ref int B, ref bool C)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = A.GetEnumerator();
			while (enumerator.MoveNext())
			{
				Images.A((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, ref B, ref C);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, ref int B, ref bool C)
	{
		bool flag = false;
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		checked
		{
			if (Images.A(A))
			{
				float height = shape.Height;
				float width = shape.Width;
				MsoTriState lockAspectRatio = shape.LockAspectRatio;
				shape.LockAspectRatio = MsoTriState.msoFalse;
				Images.A(A, 1f, 1f);
				float num = width / shape.Width;
				float num2 = height / shape.Height;
				if (Math.Round(num2, 2) != Math.Round(num, 2))
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					if (C)
					{
						Images.A(A, num, num);
					}
					else
					{
						Images.A(A, num2, num2);
					}
					flag = true;
				}
				else
				{
					Images.A(A, num2, num);
				}
				shape.LockAspectRatio = MsoTriState.msoTrue;
				if (!flag)
				{
					if (lockAspectRatio == MsoTriState.msoTrue)
					{
						goto IL_0126;
					}
					while (true)
					{
						switch (5)
						{
						case 0:
							continue;
						}
						break;
					}
				}
				B++;
			}
			else if (shape.Type == MsoShapeType.msoGroup)
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
				foreach (Microsoft.Office.Interop.PowerPoint.Shape groupItem in A.GroupItems)
				{
					Images.A(groupItem, ref B, ref C);
				}
			}
			goto IL_0126;
		}
		IL_0126:
		shape = null;
	}

	private static void A(Microsoft.Office.Interop.PowerPoint.Shape A, float B, float C)
	{
		A.ScaleHeight(B, MsoTriState.msoTrue);
		A.ScaleWidth(C, MsoTriState.msoTrue);
		_ = null;
	}

	private static bool A()
	{
		wpfFixDistortion wpfFixDistortion2 = new wpfFixDistortion();
		wpfFixDistortion2.ShowDialog();
		if (wpfFixDistortion2.DialogResult.HasValue)
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
			if (wpfFixDistortion2.DialogResult.Value)
			{
				while (true)
				{
					switch (5)
					{
					case 0:
						break;
					default:
						return wpfFixDistortion2.radAdjHeight.IsChecked.Value;
					}
				}
			}
		}
		throw new InvalidTimeZoneException();
	}

	public static void FixDistortion(Slide sld, bool blnFixWidth, bool blnFixHeight)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = sld.CustomLayout.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				B((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, blnFixWidth, blnFixHeight);
			}
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
				return;
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
	}

	public static void FixDistortion(CustomLayouts layouts, bool blnFixWidth, bool blnFixHeight)
	{
		IEnumerator enumerator = layouts.GetEnumerator();
		try
		{
			IEnumerator enumerator2 = default(IEnumerator);
			while (enumerator.MoveNext())
			{
				CustomLayout customLayout = (CustomLayout)enumerator.Current;
				{
					enumerator2 = customLayout.Shapes.GetEnumerator();
					try
					{
						while (enumerator2.MoveNext())
						{
							B((Microsoft.Office.Interop.PowerPoint.Shape)enumerator2.Current, blnFixWidth, blnFixHeight);
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
						IDisposable disposable2 = enumerator2 as IDisposable;
						if (disposable2 != null)
						{
							disposable2.Dispose();
						}
					}
				}
			}
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
		finally
		{
			IDisposable disposable = enumerator as IDisposable;
			if (disposable != null)
			{
				disposable.Dispose();
			}
		}
	}

	public static void FixDistortion(Design des, bool blnFixWidth, bool blnFixHeight)
	{
		IEnumerator enumerator = default(IEnumerator);
		try
		{
			enumerator = des.SlideMaster.Shapes.GetEnumerator();
			while (enumerator.MoveNext())
			{
				B((Microsoft.Office.Interop.PowerPoint.Shape)enumerator.Current, blnFixWidth, blnFixHeight);
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					(enumerator as IDisposable).Dispose();
					break;
				}
			}
		}
	}

	private static void B(Microsoft.Office.Interop.PowerPoint.Shape A, bool B, bool C)
	{
		if (!Images.A(A))
		{
			return;
		}
		Microsoft.Office.Interop.PowerPoint.Shape shape = A;
		if (shape.LockAspectRatio == MsoTriState.msoTrue)
		{
			float height = shape.Height;
			float width = shape.Width;
			shape.ScaleHeight(1f, MsoTriState.msoTrue);
			float num = height / shape.Height;
			float num2 = width / shape.Width;
			if (num != num2)
			{
				if (B && !C)
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
					if (1 == 0)
					{
						/*OpCode not supported: LdMemberToken*/;
					}
					shape.ScaleWidth(num, MsoTriState.msoTrue);
				}
				else if (C && !B)
				{
					shape.ScaleHeight(num2, MsoTriState.msoTrue);
				}
				else
				{
					shape.ScaleWidth(num, MsoTriState.msoTrue);
				}
			}
		}
		shape = null;
	}
}

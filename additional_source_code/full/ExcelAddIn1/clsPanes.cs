using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Forms;
using System.Windows.Interop;
using System.Windows.Media;
using A;
using MacabacusMacros;
using MacabacusMacros.ExcelHelpers;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Tools;
using Microsoft.VisualBasic.CompilerServices;

namespace ExcelAddIn1;

public sealed class clsPanes
{
	public static void WorkbookBeforeClose(Microsoft.Office.Interop.Excel.Workbook Wb, ref bool Cancel)
	{
		if (Wb.Saved)
		{
			return;
		}
		Microsoft.Office.Interop.Excel.Application application = Wb.Application;
		DialogResult dialogResult = System.Windows.Forms.MessageBox.Show(VH.A(173724) + Wb.Name + VH.A(43025), VH.A(43304), MessageBoxButtons.YesNoCancel, MessageBoxIcon.Exclamation);
		if (dialogResult != DialogResult.Cancel)
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
			if (dialogResult != DialogResult.Yes)
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
				if (dialogResult != DialogResult.No)
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
				}
				else
				{
					Wb.Saved = true;
				}
			}
			else if (Wb.Path.Length > 0)
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
				Wb.Save();
			}
			else
			{
				Microsoft.Office.Core.FileDialog fileDialog = ((_Application)application).get_FileDialog(MsoFileDialogType.msoFileDialogSaveAs);
				fileDialog.Title = VH.A(173785);
				fileDialog.Show();
				FileDialogSelectedItems selectedItems = fileDialog.SelectedItems;
				if (selectedItems.Count > 0)
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
					application.DisplayAlerts = false;
					try
					{
						Wb.SaveAs(RuntimeHelpers.GetObjectValue(selectedItems.Cast<object>().ElementAtOrDefault(0)), XlFileFormat.xlOpenXMLWorkbook, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), XlSaveAsAccessMode.xlNoChange, RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value), RuntimeHelpers.GetObjectValue(Missing.Value));
						Wb.Saved = true;
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						System.Windows.Forms.MessageBox.Show(ex2.Message, VH.A(40448), MessageBoxButtons.OK, MessageBoxIcon.Hand);
						Cancel = true;
						ProjectData.ClearProjectError();
					}
					application.DisplayAlerts = true;
				}
				else
				{
					Cancel = true;
				}
				_ = null;
			}
		}
		else
		{
			Cancel = true;
		}
		application = null;
	}

	public static bool IsValidWorkbook(Microsoft.Office.Interop.Excel.Workbook wb)
	{
		return Workbooks.IsValid(wb);
	}

	public static bool PaneAlreadyExists(Microsoft.Office.Interop.Excel.Workbook wb, string strPaneTitle)
	{
		return PaneAlreadyExists(wb.Windows[1], strPaneTitle);
	}

	public static bool PaneAlreadyExists(Microsoft.Office.Interop.Excel.Window win, string strPaneTitle)
	{
		int hwnd = win.Hwnd;
		bool result = false;
		checked
		{
			CustomTaskPane customTaskPane;
			try
			{
				int num = MH.A.CustomTaskPanes.Count - 1;
				while (true)
				{
					if (num >= 0)
					{
						try
						{
							customTaskPane = MH.A.CustomTaskPanes[num];
							try
							{
								if (Operators.CompareString(customTaskPane.Title, strPaneTitle, TextCompare: false) == 0 && customTaskPane.Visible)
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
										if (Operators.ConditionalCompareObjectEqual(NewLateBinding.LateGet(customTaskPane.Window, null, VH.A(173800), new object[0], null, null, null), hwnd, TextCompare: false))
										{
											customTaskPane = null;
											result = true;
											goto end_IL_00e0;
										}
										break;
									}
								}
							}
							catch (Exception ex)
							{
								ProjectData.SetProjectError(ex);
								Exception ex2 = ex;
								MH.A.CustomTaskPanes.Remove(customTaskPane);
								ProjectData.ClearProjectError();
							}
						}
						catch (Exception ex3)
						{
							ProjectData.SetProjectError(ex3);
							Exception ex4 = ex3;
							ProjectData.ClearProjectError();
						}
						customTaskPane = null;
						num += -1;
						continue;
					}
					while (true)
					{
						switch (4)
						{
						case 0:
							break;
						default:
							goto end_IL_00e7;
						}
						continue;
						end_IL_00e7:
						break;
					}
					break;
					continue;
					end_IL_00e0:
					break;
				}
			}
			catch (Exception ex5)
			{
				ProjectData.SetProjectError(ex5);
				Exception ex6 = ex5;
				ProjectData.ClearProjectError();
			}
			customTaskPane = null;
			return result;
		}
	}

	public static void EliminateTaskPaneFlicker(Visual visual)
	{
		((HwndSource)PresentationSource.FromVisual(visual)).CompositionTarget.RenderMode = RenderMode.SoftwareOnly;
	}

	internal static void A(string A)
	{
		checked
		{
			try
			{
				CustomTaskPaneCollection customTaskPanes = MH.A.CustomTaskPanes;
				for (int i = customTaskPanes.Count - 1; i >= 0; i += -1)
				{
					try
					{
						if (Operators.CompareString(customTaskPanes[i].Title, A, TextCompare: false) != 0)
						{
							continue;
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
							customTaskPanes.RemoveAt(i);
							break;
						}
					}
					catch (Exception ex)
					{
						ProjectData.SetProjectError(ex);
						Exception ex2 = ex;
						ProjectData.ClearProjectError();
					}
				}
				while (true)
				{
					switch (2)
					{
					case 0:
						continue;
					}
					customTaskPanes = null;
					return;
				}
			}
			catch (Exception ex3)
			{
				ProjectData.SetProjectError(ex3);
				Exception ex4 = ex3;
				ProjectData.ClearProjectError();
			}
		}
	}

	private static CustomTaskPane A(Dictionary<int, CustomTaskPane> A)
	{
		CustomTaskPane value = null;
		try
		{
			if (A != null)
			{
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
					A.TryGetValue(MH.A.Application.ActiveWindow.Hwnd, out value);
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
		return value;
	}

	internal static bool A(Dictionary<int, CustomTaskPane> A)
	{
		bool result;
		try
		{
			result = clsPanes.A(A).Visible;
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

	internal static void A(CustomTaskPane A, clsDisplay B, int C = 0)
	{
		if (C == 0)
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
			C = K.Settings.TaskPaneWidth;
		}
		try
		{
			A.Width = checked((int)Math.Round((double)C * B.X));
		}
		catch (ArgumentException ex)
		{
			ProjectData.SetProjectError(ex);
			ArgumentException ex2 = ex;
			A.Width = 430;
			ProjectData.ClearProjectError();
		}
	}
}

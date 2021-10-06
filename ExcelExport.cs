using System;
using System.Windows.Controls;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;

namespace TaskWPF
{
	public class ExcelExport: Excel.Window
	{
		public static void ExcelExportDataGrid(DataGrid dataGrid)
		{
			Excel.Application excel = new Excel.Application();
			excel.Visible = true;
			Workbook workbook = excel.Workbooks.Add(System.Reflection.Missing.Value);
			Worksheet sheet1 = (Worksheet)workbook.Sheets[1];

			for (int j = 0; j < dataGrid.Columns.Count; j++)
			{
				Range myRange = (Range)sheet1.Cells[1, j + 1];
				sheet1.Cells[1, j + 1].Font.Bold = true;
				sheet1.Columns[j + 1].ColumnWidth = 15;
				myRange.Value2 = dataGrid.Columns[j].Header;
			}
			for (int i = 0; i < dataGrid.Columns.Count; i++)
			{
				for (int j = 0; j < dataGrid.Items.Count; j++)
				{
					TextBlock b = dataGrid.Columns[i].GetCellContent(dataGrid.Items[j]) as TextBlock;
					Microsoft.Office.Interop.Excel.Range myRange = (Microsoft.Office.Interop.Excel.Range)sheet1.Cells[j + 2, i + 1];
					myRange.Value2 = b.Text;
				}
			}
		}

		#region Realization Interface

		dynamic Excel.Window.Activate()
		{
			throw new NotImplementedException();
		}

		public dynamic ActivateNext()
		{
			throw new NotImplementedException();
		}

		public dynamic ActivatePrevious()
		{
			throw new NotImplementedException();
		}

		public bool Close(object SaveChanges, object Filename, object RouteWorkbook)
		{
			throw new NotImplementedException();
		}

		public dynamic LargeScroll(object Down, object Up, object ToRight, object ToLeft)
		{
			throw new NotImplementedException();
		}

		public Excel.Window NewWindow()
		{
			throw new NotImplementedException();
		}

		public dynamic _PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
		{
			throw new NotImplementedException();
		}

		public dynamic PrintPreview(object EnableChanges)
		{
			throw new NotImplementedException();
		}

		public dynamic ScrollWorkbookTabs(object Sheets, object Position)
		{
			throw new NotImplementedException();
		}

		public dynamic SmallScroll(object Down, object Up, object ToRight, object ToLeft)
		{
			throw new NotImplementedException();
		}

		public int PointsToScreenPixelsX(int Points)
		{
			throw new NotImplementedException();
		}

		public int PointsToScreenPixelsY(int Points)
		{
			throw new NotImplementedException();
		}

		public dynamic RangeFromPoint(int x, int y)
		{
			throw new NotImplementedException();
		}

		public void ScrollIntoView(int Left, int Top, int Width, int Height, object Start)
		{
			throw new NotImplementedException();
		}

		public dynamic PrintOut(object From, object To, object Copies, object Preview, object ActivePrinter, object PrintToFile, object Collate, object PrToFileName)
		{
			throw new NotImplementedException();
		}

		public Excel.Application Application => throw new NotImplementedException();

		public XlCreator Creator => throw new NotImplementedException();

		dynamic Excel.Window.Parent => throw new NotImplementedException();

		public Range ActiveCell => throw new NotImplementedException();

		public Chart ActiveChart => throw new NotImplementedException();

		public Pane ActivePane => throw new NotImplementedException();

		public dynamic ActiveSheet => throw new NotImplementedException();

		public dynamic Caption { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayFormulas { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayGridlines { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayHeadings { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayHorizontalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayOutline { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool _DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayVerticalScrollBar { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayWorkbookTabs { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayZeros { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool EnableResize { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool FreezePanes { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public int GridlineColor { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public XlColorIndex GridlineColorIndex { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public int Index => throw new NotImplementedException();

		public string OnWindow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public Panes Panes => throw new NotImplementedException();

		public Range RangeSelection => throw new NotImplementedException();

		public int ScrollColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public int ScrollRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public Sheets SelectedSheets => throw new NotImplementedException();

		public dynamic Selection => throw new NotImplementedException();

		public bool Split { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public int SplitColumn { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double SplitHorizontal { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public int SplitRow { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double SplitVertical { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double TabRatio { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public XlWindowType Type => throw new NotImplementedException();

		public double UsableHeight => throw new NotImplementedException();

		public double UsableWidth => throw new NotImplementedException();

		public bool Visible { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public Range VisibleRange => throw new NotImplementedException();

		public int WindowNumber => throw new NotImplementedException();

		XlWindowState Excel.Window.WindowState { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public dynamic Zoom { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public XlWindowView View { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayRightToLeft { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public SheetViews SheetViews => throw new NotImplementedException();

		public dynamic ActiveSheetView => throw new NotImplementedException();

		public bool DisplayRuler { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool AutoFilterDateGrouping { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public bool DisplayWhitespace { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

		public int Hwnd => throw new NotImplementedException();

		public double Height { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double Left { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double Top { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		public double Width { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }
		#endregion
	}
}

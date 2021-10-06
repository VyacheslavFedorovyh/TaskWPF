using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Data;
using System.Data.SqlClient;
using System.Configuration;
using System.Globalization;


namespace TaskWPF
{
	public partial class MainWindow : Window
	{
		private static SqlConnection sqlConnection = null;
		private DataTable dtComboBox = null;
		private string commandText = string.Empty;
		private string condition = null;

		public MainWindow()
		{
			InitializeComponent();
			bindYearComboBox();
			bindBrandComboBox();
		}

		private void bindYearComboBox()
		{
			commandText = "SELECT YEAR(Date) as Months FROM SalesVolume GROUP BY YEAR(Date);";
			dtComboBox = SqlConnection(commandText);
			YearComboBox.ItemsSource = dtComboBox.DefaultView;
			YearComboBox.SelectedIndex = YearComboBox.Items.Count - 1;
		}

		private void bindBrandComboBox()
		{
			commandText = "SELECT Brand FROM SalesVolume GROUP BY Brand;";
			dtComboBox = SqlConnection(commandText);
			dtComboBox.DefaultView.AddNew();
			BrandComboBox.ItemsSource = dtComboBox.DefaultView;
		}

		private void myComboBox_SelectionChanged(object sender, SelectionChangedEventArgs e)
		{
			DataRowView yearDataRowView = YearComboBox.SelectedItem as DataRowView;
			DataRowView brandDataRowView = BrandComboBox.SelectedItem as DataRowView;

			if (yearDataRowView != null)
				condition = yearDataRowView.Row["Months"].ToString();

			if (brandDataRowView != null)
			{
				string brand = brandDataRowView.Row["Brand"].ToString();

				if (brand != string.Empty)
					condition = condition + string.Format(" AND Brand = '{0}'", brand);
			}
		}

		private void Button_Click(object sender, RoutedEventArgs e)
		{
			sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Cars"].ConnectionString);

			sqlConnection.Open();
			if (sqlConnection.State == ConnectionState.Open && condition != null)
			{
				SqlCommand cmd = new SqlCommand();

				cmd.CommandText = string.Format(@"SELECT Brand, [1] as January,[2] as February,[3] as March,[4] as April,[5] as May,
														[6] as June,[7] as July,[8] as August,[9] as September,[10] as October,[11] as November,[12] as December 
													FROM
													(SELECT Brand, Price, MONTH(Date) as Months
													FROM SalesVolume
													WHERE YEAR(Date) = {0}) SourceTable
													PIVOT
													(SUM(Price)
													FOR Months
													IN([1],[2],[3],[4],[5],[6],[7],[8],[9],[10],[11],[12])
													) PivotTable; ", condition);

				cmd.Connection = sqlConnection;
				SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
				DataTable dt = new DataTable("SalesVolume");
				dataAdapter.Fill(dt);

				DataGrid.ItemsSource = dt.DefaultView;

				sqlConnection.Close();
			}
		}

		private DataTable SqlConnection(string commandText)
		{
			sqlConnection = new SqlConnection(ConfigurationManager.ConnectionStrings["Cars"].ConnectionString);
			sqlConnection.Open();
			SqlCommand cmd = new SqlCommand();
			cmd.CommandText = commandText;
			cmd.Connection = sqlConnection;
			SqlDataAdapter dataAdapter = new SqlDataAdapter(cmd);
			DataTable dt = new DataTable("SalesVolume");
			dataAdapter.Fill(dt);
			return dt;
		}

		private void ExcelExportButton_Click(object sender, RoutedEventArgs e)
		{
			ExcelExport.ExcelExportDataGrid(DataGrid);
		}
	}

	class ToColorConverter : IValueConverter
	{
		public object Convert(object value, Type targetType, object parameter, CultureInfo culture)
		{
			//return (long)value > 25000000 ? new SolidColorBrush(Colors.Green) : new SolidColorBrush(Colors.White);
			return new SolidColorBrush(Colors.White);
		}

		public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
		{
			throw new NotImplementedException();
		}
	}
}

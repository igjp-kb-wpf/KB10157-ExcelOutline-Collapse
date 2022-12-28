using Infragistics.Documents.Excel;
using KB10157_ExcelOutline_Collapse.Model;
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

namespace KB10157_ExcelOutline_Collapse;

/// <summary>
/// Interaction logic for MainWindow.xaml
/// </summary>
public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private void button1_Click(object sender, RoutedEventArgs e)
    {
        List<Person> people = ((MainWindowViewModel)(DataContext)).People.OrderBy(person => person.Prefecture).ToList();

        // https://jp.infragistics.com/help/wpf/excelengine-creating-a-workbook
        Workbook workbook1 = new Workbook(WorkbookFormat.Excel2007);
        Worksheet worksheet1 = workbook1.Worksheets.Add("Sheet 1");

        worksheet1.Rows[0].Cells[0].Value = "ID";
        worksheet1.Rows[0].Cells[0].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
        worksheet1.Rows[0].Cells[0].CellFormat.Fill = CellFill.CreateSolidFill(System.Drawing.Color.Silver);
        worksheet1.Rows[0].Cells[1].Value = "FamilyName";
        worksheet1.Rows[0].Cells[1].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
        worksheet1.Rows[0].Cells[1].CellFormat.Fill = CellFill.CreateSolidFill(System.Drawing.Color.Silver);
        worksheet1.Rows[0].Cells[2].Value = "GivenName";
        worksheet1.Rows[0].Cells[2].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
        worksheet1.Rows[0].Cells[2].CellFormat.Fill = CellFill.CreateSolidFill(System.Drawing.Color.Silver);
        worksheet1.Rows[0].Cells[3].Value = "Prefecture";
        worksheet1.Rows[0].Cells[3].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
        worksheet1.Rows[0].Cells[3].CellFormat.Fill = CellFill.CreateSolidFill(System.Drawing.Color.Silver);
        worksheet1.Rows[0].Cells[4].Value = "City";
        worksheet1.Rows[0].Cells[4].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
        worksheet1.Rows[0].Cells[4].CellFormat.Fill = CellFill.CreateSolidFill(System.Drawing.Color.Silver);

        for (int peopleIndex = 0, rowIndex = 1; peopleIndex < people.Count; peopleIndex++, rowIndex++)
        {
            worksheet1.Rows[rowIndex].Cells[0].Value = people[peopleIndex].ID;
            worksheet1.Rows[rowIndex].Cells[1].Value = people[peopleIndex].FamilyName;
            worksheet1.Rows[rowIndex].Cells[2].Value = people[peopleIndex].GivenName;
            worksheet1.Rows[rowIndex].Cells[3].Value = people[peopleIndex].Prefecture;
            worksheet1.Rows[rowIndex].Cells[4].Value = people[peopleIndex].City;

            // アウトラインレベルを設定する。
            worksheet1.Rows[rowIndex].OutlineLevel = 1;

            // アウトラインを折りたたむ。
            worksheet1.Rows[rowIndex].Hidden = true;

            if (peopleIndex < people.Count - 1 && people[peopleIndex].Prefecture != people[peopleIndex + 1].Prefecture)
            {
                rowIndex++;
                worksheet1.Rows[rowIndex].Cells[0].Value = people[peopleIndex].Prefecture;
                worksheet1.Rows[rowIndex].Cells[0].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                worksheet1.Rows[rowIndex].Cells[1].Value = people.Where(person => person.Prefecture == people[peopleIndex].Prefecture).Count();
                worksheet1.Rows[rowIndex].Cells[1].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
                worksheet1.Rows[rowIndex].Cells[2].Value = "件";
                worksheet1.Rows[rowIndex].Cells[2].CellFormat.Font.Bold = ExcelDefaultableBoolean.True;
            }
        }

        workbook1.Save("Workbook1.xlsx");

        MessageBox.Show("Done!");
    }
}

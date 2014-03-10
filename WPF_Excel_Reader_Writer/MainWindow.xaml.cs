using System;
using System.Windows;
using System.Windows.Controls;

namespace WPF_Excel_Reader_Writer
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        //The object of the DataAccess class
        DataAccess objDs;
        public MainWindow()
        {
            InitializeComponent();
        }

        //The Employee Object for Edit
        Employee emp = new Employee();
        /// <summary>
        /// On Load get data from the Excel
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            objDs = new DataAccess(); 
            try
            {
                    dgEmp.ItemsSource = objDs.GetDataFormExcelAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// TO Synchronize the Excel Workbook with the Application 
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private   void btnsync_Click(object sender, RoutedEventArgs e)
        {
            try
            {
               dgEmp.ItemsSource =   objDs.GetDataFormExcelAsync().Result;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// Read Data entered in each Cell
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgEmp_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            try
            {
                    FrameworkElement eleEno = dgEmp.Columns[0].GetCellContent(e.Row);
                    if (eleEno.GetType() == typeof(TextBox))
                    {
                        emp.EmpNo = Convert.ToInt32(((TextBox)eleEno).Text);
                    }

                    FrameworkElement eleEname = dgEmp.Columns[1].GetCellContent(e.Row);
                    if (eleEname.GetType() == typeof(TextBox))
                    {
                        emp.EmpName = ((TextBox)eleEname).Text;
                    }

                    FrameworkElement eleSal = dgEmp.Columns[2].GetCellContent(e.Row);
                    if (eleSal.GetType() == typeof(TextBox))
                    {
                        emp.Salary = Convert.ToInt32(((TextBox)eleSal).Text);
                    }

                    FrameworkElement eleDname = dgEmp.Columns[3].GetCellContent(e.Row);
                    if (eleDname.GetType() == typeof(TextBox))
                    {
                        emp.DeptName = ((TextBox)eleDname).Text;
                    }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        /// <summary>
        /// Get the Complete row
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgEmp_RowEditEnding(object sender, DataGridRowEditEndingEventArgs e)
        {
            try
            {
              bool IsSave = objDs.InsertOrUpdateRowInExcelAsync(emp).Result;
              if (IsSave)
              {
                  MessageBox.Show("Record Saved Successfully");
              }
              else
              {
                  MessageBox.Show("Problem Occured");
              }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            
        }
        /// <summary>
        /// Select the Recod for the Update
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void dgEmp_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            emp = dgEmp.SelectedItem as Employee;
        }
    }
}

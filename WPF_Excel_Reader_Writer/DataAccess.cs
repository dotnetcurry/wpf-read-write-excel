using System;
using System.Threading.Tasks;

using System.Data.OleDb;
using System.Collections.ObjectModel;

namespace WPF_Excel_Reader_Writer
{

    public class Employee
    {
        public int EmpNo { get;set; }
        public string EmpName { get; set; }
        public int Salary { get; set; }
        public string DeptName { get; set; }
    }
    public class DataAccess
    {
        OleDbConnection Conn;
        OleDbCommand Cmd;

        public DataAccess()
        {
            Conn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0;Data Source=F:\\FromC\\VS2013\\WPF_45_DEMOS\\Employee.xlsx;Extended Properties=\"Excel 12.0 Xml;HDR=YES;\""); 
        }

        /// <summary>
        /// Method to Get All the Records from Excel
        /// </summary>
        /// <returns></returns>
        public async Task<ObservableCollection<Employee>> GetDataFormExcelAsync()
        {
            ObservableCollection<Employee> Employees = new ObservableCollection<Employee>(); 
            await Conn.OpenAsync();
            Cmd = new OleDbCommand();
            Cmd.Connection = Conn;
            Cmd.CommandText = "Select * from [Sheet1$]";
            var Reader = await Cmd.ExecuteReaderAsync();
            while (Reader.Read())
            {
                Employees.Add(new Employee() { 
                    EmpNo = Convert.ToInt32(Reader["EmpNo"]),
                    EmpName = Reader["EmpName"].ToString(),
                    DeptName = Reader["DeptName"].ToString(),
                    Salary = Convert.ToInt32(Reader["Salary"]) 
                });
            }
            Reader.Close();
            Conn.Close();
            return Employees;
        }

        /// <summary>
        /// Method to Insert Record in the Excel
        /// S1. If the EmpNo =0, then the Operation is Skipped.
        /// S2. If the Employee is already exist, then it is taken for Update
        /// </summary>
        /// <param name="Emp"></param>
        public async Task<bool> InsertOrUpdateRowInExcelAsync(Employee Emp)
        {
            bool IsSave = false;
            //S1
            if (Emp.EmpNo != 0)
            {
                await Conn.OpenAsync();
                Cmd = new OleDbCommand();
                Cmd.Connection = Conn;
                Cmd.Parameters.AddWithValue("@EmpNo", Emp.EmpNo);
                Cmd.Parameters.AddWithValue("@EmpName", Emp.EmpName);
                Cmd.Parameters.AddWithValue("@Salary", Emp.Salary);
                Cmd.Parameters.AddWithValue("@DeptName", Emp.DeptName);
                //S2
                if (!CheckIfRecordExistAsync(Emp).Result)
                {
                    Cmd.CommandText = "Insert into [Sheet1$] values (@EmpNo,@EmpName,@Salary,@DeptName)";
                }
                else
                {
                    if (Emp.EmpName != String.Empty || Emp.DeptName != String.Empty)
                    {
                        Cmd.CommandText = "Update [Sheet1$] set EmpNo=@EmpNo,EmpName=@EmpName,Salary=@Salary,DeptName=@DeptName where EmpNo=@EmpNo";
                    }
                }
                int result = await Cmd.ExecuteNonQueryAsync();
                if (result > 0)
                {
                    IsSave = true;
                }
                Conn.Close();
            }
            return IsSave;

        }

         

        /// <summary>
        /// The method to check if the record is already available 
        /// in the workgroup
        /// </summary>
        /// <param name="emp"></param>
        /// <returns></returns>
        private async Task<bool> CheckIfRecordExistAsync(Employee emp)
        {
            bool IsRecordExist = false;
            Cmd.CommandText = "Select * from [Sheet1$] where EmpNo=@EmpNo";
            var Reader = await Cmd.ExecuteReaderAsync();
            if (Reader.HasRows)
            {
                IsRecordExist = true;
            }
             
            Reader.Close();
            return IsRecordExist;
        }
    }
     
}

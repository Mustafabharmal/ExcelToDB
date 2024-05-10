// Employee.cs
using System.ComponentModel.DataAnnotations;
namespace xltodb.Models
{
    public class Employee
    {
        public int EmployeeId { get; set; }

        [Required]
        public string EmployeeName { get; set; }

        public string ContactNumber { get; set; }

        public string Address { get; set; }

        public string Gender { get; set; }

        public string Education { get; set; }
    }
}
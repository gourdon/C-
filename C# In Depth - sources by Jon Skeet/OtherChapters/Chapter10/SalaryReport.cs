using System;
using System.Collections.Generic;
using System.Linq;

namespace Chapter10
{
    class Employee
    {
        public string Name { get; set; }
        public decimal Salary { get; set; }
    }

    class Department
    {
        public string Name { get; set; }

        List<Employee> employees = new List<Employee>();

        public IList<Employee> Employees
        {
            get { return employees; }
        }
    }

    class Company
    {
        public string Name { get; set; }

        List<Department> departments = new List<Department>();

        public IList<Department> Departments
        {
            get { return departments; }
        }
    }

    class SalaryReport
    {
        static void Main()
        {
            var company = new Company
            {
                Name = "Skeetysoft",
                Departments =
                {
                    new Department 
                    { 
                        Name = "Development", 
                        Employees = 
                        {
                            new Employee { Name = "Tim Trotter", Salary = 75000m },
                            new Employee { Name = "Tara Tutu", Salary = 50000m },
                            new Employee { Name = "Deborah Denton", Salary = 90000m },
                            new Employee { Name = "Darren Dahlia", Salary = 45000m },
                            new Employee { Name = "Mary Malcop", Salary = 150000m }
                        }
                    },
                    new Department 
                    { 
                        Name = "Marketing", 
                        Employees = 
                        {
                            new Employee { Name = "Molly Molton", Salary = 200000m },
                            new Employee { Name = "Mark Mainton", Salary = 120000m }
                        }
                    }
                }
            };

            var query = company.Departments
                               .Select(dept => new { dept.Name, Cost = dept.Employees.Sum(person => person.Salary) })
                               .OrderByDescending(deptWithCost => deptWithCost.Cost);

            foreach (var item in query)
            {
                Console.WriteLine(item);
            }
        }
    }
}

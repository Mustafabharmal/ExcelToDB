using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Entity;
using System.Data.SqlClient;
using System.IO;
using System.Runtime.Remoting.Messaging;
using System.Web;
using System.Web.Mvc;
using ExcelDataReader;
using xltodb.Models;

public class EmployeeController : Controller
{
    private string connectionString = "data source = localhost;initial catalog = SpclDB; user id = sa; password=reallyStrongPwd123";

    // GET: Employee
    public ActionResult Index()
    {
        List<Employee> employees = new List<Employee>();

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand cmd = new SqlCommand("SELECT * FROM Employees", connection))
            {
                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        employees.Add(new Employee
                        {
                            EmployeeId = reader.GetInt32(reader.GetOrdinal("Id")),
                            EmployeeName = reader.GetString(reader.GetOrdinal("EmployeeName")),
                            ContactNumber = reader.GetString(reader.GetOrdinal("ContactNumber")),
                            Address = reader.GetString(reader.GetOrdinal("Address")),
                            Gender = reader.GetString(reader.GetOrdinal("Gender")),
                            Education = reader.GetString(reader.GetOrdinal("Education"))
                        });
                    }
                }
            }
        }

        return View(employees);
    }

    // GET: Employee/Add
    public ActionResult Add()
    {
        return View();
    }

    // POST: Employee/Add
    [HttpPost]
    public ActionResult Add(Employee employee)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand cmd = new SqlCommand("INSERT INTO Employees (EmployeeName, ContactNumber, Address, Gender, Education) VALUES (@EmployeeName, @ContactNumber, @Address, @Gender, @Education)", connection))
            {
                cmd.Parameters.AddWithValue("@EmployeeName", employee.EmployeeName);
                cmd.Parameters.AddWithValue("@ContactNumber", employee.ContactNumber);
                cmd.Parameters.AddWithValue("@Address", employee.Address);
                cmd.Parameters.AddWithValue("@Gender", employee.Gender);
                cmd.Parameters.AddWithValue("@Education", employee.Education);

                cmd.ExecuteNonQuery();
            }
        }

        return RedirectToAction("Index");
    }

    // GET: Employee/Edit/5
    public ActionResult Edit(int id)
    {
        Employee employee = null;

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand cmd = new SqlCommand("SELECT * FROM Employees WHERE Id = @Id", connection))
            {
                cmd.Parameters.AddWithValue("@Id", id);

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    if (reader.Read())
                    {
                        employee = new Employee
                        {
                            EmployeeId = reader.GetInt32(reader.GetOrdinal("Id")),
                            EmployeeName = reader.GetString(reader.GetOrdinal("EmployeeName")),
                            ContactNumber = reader.GetString(reader.GetOrdinal("ContactNumber")),
                            Address = reader.GetString(reader.GetOrdinal("Address")),
                            Gender = reader.GetString(reader.GetOrdinal("Gender")),
                            Education = reader.GetString(reader.GetOrdinal("Education"))
                        };
                    }
                }
            }
        }

        if (employee == null)
        {
            return HttpNotFound();
        }

        return View(employee);
    }

    // POST: Employee/Update
    [HttpPost]
    public ActionResult Update(Employee employee)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand cmd = new SqlCommand("UPDATE Employees SET EmployeeName = @EmployeeName, ContactNumber = @ContactNumber, Address = @Address, Gender = @Gender, Education = @Education WHERE Id = @Id", connection))
            {
                cmd.Parameters.AddWithValue("@Id", employee.EmployeeId);
                cmd.Parameters.AddWithValue("@EmployeeName", employee.EmployeeName);
                cmd.Parameters.AddWithValue("@ContactNumber", employee.ContactNumber);
                cmd.Parameters.AddWithValue("@Address", employee.Address);
                cmd.Parameters.AddWithValue("@Gender", employee.Gender);
                cmd.Parameters.AddWithValue("@Education", employee.Education);

                cmd.ExecuteNonQuery();
            }
        }

        return RedirectToAction("Index");
    }

    // GET: Employee/Delete/5
    public ActionResult Delete(int id)
    {
        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand cmd = new SqlCommand("DELETE FROM Employees WHERE Id = @Id", connection))
            {
                cmd.Parameters.AddWithValue("@Id", id);
                cmd.ExecuteNonQuery();
            }
        }

        return RedirectToAction("Index");
    }

    [HttpPost]
    public ActionResult UploadExcel(HttpPostedFileBase file)
    {
        if (file != null && file.ContentLength > 0)
        {
            try
            {
                using (var stream = file.InputStream)
                {
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        do
                        {
                            while (reader.Read())
                            {
                                Employee employee = new Employee
                                {
                                    EmployeeName = reader[0].ToString(),
                                    ContactNumber = reader[1].ToString(),
                                    Address = reader[2].ToString(),
                                    Gender = reader[3].ToString(),
                                    Education = reader[4].ToString()
                                };

                                using (SqlConnection connection = new SqlConnection(connectionString))
                                {
                                    connection.Open();

                                    using (SqlCommand cmd = new SqlCommand("INSERT INTO Employees (EmployeeName, ContactNumber, Address, Gender, Education) VALUES (@EmployeeName, @ContactNumber, @Address, @Gender, @Education)", connection))
                                    {
                                        cmd.Parameters.AddWithValue("@EmployeeName", employee.EmployeeName);
                                        cmd.Parameters.AddWithValue("@ContactNumber", employee.ContactNumber);
                                        cmd.Parameters.AddWithValue("@Address", employee.Address);
                                        cmd.Parameters.AddWithValue("@Gender", employee.Gender);
                                        cmd.Parameters.AddWithValue("@Education", employee.Education);

                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }
                        } while (reader.NextResult());
                    }
                }

                ViewBag.Message = "Excel file uploaded and data stored successfully.";
            }
            catch (Exception ex)
            {
                ViewBag.Message = "An error occurred: " + ex.Message;
            }
        }
        else
        {
            ViewBag.Message = "Please select a file.";
        }

        return View("Index");
    }
}

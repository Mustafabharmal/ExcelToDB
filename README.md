# ExcelToDB 📊

Welcome to **ExcelToDB**, a web application built with ASP.NET MVC that simplifies importing employee data from Excel into a database.

## 🚀 Features

- **Excel File Upload**: Upload an Excel file with employee information.
- **Data Import**: Read and store data from the uploaded Excel file into a database table.

## 💻 Technologies Used

- **ASP.NET MVC**: Framework for web development.
- **C#**: Server-side programming language.
- **SQL Server**: Database management system.
- **Entity Framework**: Object-Relational Mapping (ORM) tool.
- **HTML/CSS**: Frontend design and styling.
- **Bootstrap**: Frontend framework for responsive UI.

## 🛠️ Getting Started

To run this project locally, follow these steps:

1. **Clone the Repository**

   ```bash
   git clone https://github.com/Mustafabharmal/ExcelToDB.git
   ```

2. **Set up Database**

   - Execute the SQL script (`DatabaseScript.sql`) in SQL Server Management Studio to create the necessary database (`Excel`) and table (`Employees`).

3. **Open Project in Visual Studio**

   - Launch Visual Studio and open the project.
   - Ensure all required dependencies for ASP.NET MVC are installed.

4. **Configure Connection String**

   - Update the database connection string in `Web.config` to match your SQL Server instance.

     ```xml
     <connectionStrings>
       <add name="DefaultConnection" connectionString="Data Source=YOUR_SERVER;Initial Catalog=Excel;Integrated Security=True" providerName="System.Data.SqlClient" />
     </connectionStrings>
     ```

5. **Build and Run**

   - Build the solution and start the application.
   - Access the application in your web browser (`http://localhost:{port}/Employee/Index`).

6. **Upload Excel File**

   - Click on "Choose File" to select an Excel file (.xls or .xlsx) containing employee data.
   - Click "Upload" to import the data into the database.

## 📂 Project Structure

- **Controllers**: Contains MVC controller classes.
  - `EmployeeController.cs`: Manages file upload and data import logic.
- **Models**: Holds model classes.
- **Views**: Stores HTML views rendered by the application.
  - `Index.cshtml`: View for the file upload form.

## 📝 License

This project is licensed under the [MIT License](LICENSE).

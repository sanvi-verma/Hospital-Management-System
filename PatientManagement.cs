using System;
using System.Collections.Generic;
using System.Data;
using MySql.Data.MySqlClient;
using OfficeOpenXml; // For handling Excel exports
using OfficeOpenXml.Style; // For styling Excel cells

public class PatientManagement
{
    private DatabaseHelper dbHelper;

    public PatientManagement()
    {
        dbHelper = new DatabaseHelper();
    }

    // Method to add a new patient
    public void AddNewPatient(string firstName, string lastName, DateTime dob, string gender, string phone, string email, string address)
    {
        using (var connection = dbHelper.GetConnection())
        {
            connection.Open();
            string query = "INSERT INTO patients (first_name, last_name, dob, gender, phone_number, email, address) VALUES (@firstName, @lastName, @dob, @gender, @phone_number, @Email, @Address)";
            
            using (var cmd = new MySqlCommand(query, connection))
            {
                cmd.Parameters.AddWithValue("@firstName", firstName);
                cmd.Parameters.AddWithValue("@lastName", lastName);
                cmd.Parameters.AddWithValue("@dob", dob);
                cmd.Parameters.AddWithValue("@gender", gender);
                cmd.Parameters.AddWithValue("@phone_number", phone);
                cmd.Parameters.AddWithValue("@Email", email);
                cmd.Parameters.AddWithValue("@Address", address);

                cmd.ExecuteNonQuery();
            }
        }
    }

    // Method to retrieve all patients
    public List<Patient> GetPatients()
    {
        List<Patient> patients = new List<Patient>();

        using (var connection = dbHelper.GetConnection())
        {
            connection.Open();
            string query = "SELECT patient_id, CONCAT(first_name, ' ', last_name) AS name, dob FROM patients";

            using (var cmd = new MySqlCommand(query, connection))
            using (var reader = cmd.ExecuteReader())
            {
                while (reader.Read())
                {
                    Patient patient = new Patient
                    {
                        PatientID = reader.GetInt32("patient_id"),
                        Name = reader.GetString("name"),
                        DOB = reader.GetDateTime("dob")
                    };
                    patients.Add(patient);
                }
            }
        }

        return patients;
    }

    // Method to export patient data to Excel
   public void ExportPatientsToExcel(string filePath)
{
    try
    {
        var patients = GetPatients();  // Retrieve patient data from database

        // Set the license context for EPPlus (required for non-commercial use)
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            // Add a new worksheet to the package
            var worksheet = package.Workbook.Worksheets.Add("Patients");

            // Add header row to the worksheet
            worksheet.Cells[1, 1].Value = "Patient ID";
            worksheet.Cells[1, 2].Value = "Name";
            worksheet.Cells[1, 3].Value = "DOB";

            // Style the header row
            using (var range = worksheet.Cells[1, 1, 1, 3])
            {
                range.Style.Font.Bold = true;
                range.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                range.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            }

            // Add patient data to the worksheet
            int row = 2;  // Start from row 2 since row 1 is the header
            foreach (var patient in patients)
            {
                worksheet.Cells[row, 1].Value = patient.PatientID;
                worksheet.Cells[row, 2].Value = patient.Name;
                worksheet.Cells[row, 3].Value = patient.DOB.ToString("yyyy-MM-dd");
                row++;
            }

            // Auto-fit columns for better readability
            worksheet.Cells.AutoFitColumns();

            // Validate the file path, ensure it can be written
            var fileInfo = new FileInfo(filePath);

            // Save the Excel file
            package.SaveAs(fileInfo);
            Console.WriteLine($"Patients data exported to {filePath}");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error exporting patients to Excel: {ex.Message}");
    }
}

}

// Patient class to represent the data model
public class Patient
{
    public int PatientID { get; set; }
    public string? Name { get; set; }
    public DateTime DOB { get; set; }
}

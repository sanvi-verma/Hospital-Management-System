using System;

namespace HospitalManagementSystem
{
    class Program
    {
        static void Main(string[] args)
        {
            PatientManagement patientManagement = new PatientManagement();

            // Prompting user to input patient details
            Console.WriteLine("Enter patient details to add a new patient:");

            Console.Write("First Name: ");
            string firstName = Console.ReadLine();

            Console.Write("Last Name: ");
            string lastName = Console.ReadLine();

            Console.Write("Date of Birth (YYYY-MM-DD): ");
            DateTime dob;
            while (!DateTime.TryParse(Console.ReadLine(), out dob))
            {
                Console.WriteLine("Invalid date format. Please enter the date in YYYY-MM-DD format.");
            }

            Console.Write("Gender (Male/Female/Other): ");
            string gender = Console.ReadLine();

            Console.Write("Phone Number: ");
            string phone = Console.ReadLine();

            Console.Write("Email: ");
            string email = Console.ReadLine();

            Console.Write("Address: ");
            string address = Console.ReadLine();

            // Add the new patient with the input details
            patientManagement.AddNewPatient(firstName, lastName, dob, gender, phone, email, address);
            Console.WriteLine("Patient added successfully.");

            // Fetch and print patient details
            var patients = patientManagement.GetPatients();
            foreach (var patient in patients)
            {
                Console.WriteLine($"ID: {patient.PatientID}, Name: {patient.Name}, DOB: {patient.DOB}");
            }

            // Export patient data to Excel (provide the correct file path)
            string excelFilePath = "Patients.xlsx";
            patientManagement.ExportPatientsToExcel(excelFilePath);
            Console.WriteLine($"Patient data exported to {excelFilePath}");
        }
    }
}

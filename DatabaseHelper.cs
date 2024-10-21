using System;
using MySql.Data.MySqlClient;

public class DatabaseHelper
{
    private string connectionString;

    public DatabaseHelper()
    {
        // Connection string for MySQL (replace with your actual database details)
        connectionString = "Server=localhost;Database=hospital_management;Uid=root;Pwd=root;";
    }

    // Method to get a connection to the database
    public MySqlConnection GetConnection()
    {
        return new MySqlConnection(connectionString);
    }
}

using System;
using System.Data;
using System.Data.OleDb;

class Program
{
    static void Main()
    {
        // Connection string to the Access database
        string connectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + "C:\\Users\\Mahad Ghauri\\Downloads\\Northwind.mdb;User Id=admin;Password=;";

        // SQL query with a parameter placeholder
        string queryString = "SELECT ProductID, UnitPrice, ProductName FROM products " + "WHERE UnitPrice > ? " + "ORDER BY UnitPrice DESC;";

        // Parameter value for the query
        int paramValue = 5;

        // Create a connection to the database
        OleDbConnection connection = new OleDbConnection(connectionString);

        // Create a command to execute the SQL query
        OleDbCommand command = new OleDbCommand(queryString, connection);
        command.Parameters.AddWithValue("@pricePoint", paramValue);

        try
        {
            // Open the connection
            connection.Open();

            // Execute the query and get a reader for the results
            OleDbDataReader reader = command.ExecuteReader();

            // Read and display each row of the result set
            Console.WriteLine("ProductID\tUnitPrice\tProductName");
            Console.WriteLine("---------------------------------------");
            while (reader.Read())
            {
                Console.WriteLine("{0}\t\t{1}\t\t{2}", reader[0], reader[1], reader[2]);
            }

            // Close the reader
            reader.Close();
        }
        catch (Exception ex)
        {
            // Display any errors that occur
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            // Ensure the connection is always closed
            if (connection.State == ConnectionState.Open)
            {
                connection.Close();
            }
        }

        // Pause the console for the user to view results
        Console.ReadLine();
    }
}

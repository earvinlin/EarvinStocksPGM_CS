using MySqlConnector;
using System;

public static class DbHelper
{
    private static string ConnStr = "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";
           
public static MySqlConnection GetConnection()
    {
        using var conn = new MySqlConnection(ConnStr);
        return new MySqlConnection(ConnStr);
    }
}

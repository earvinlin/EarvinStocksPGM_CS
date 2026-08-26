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

    public static void TestConnectDB()
    {
        string connStr =
    "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";

        using var conn = new MySqlConnection(connStr);

        try
        {
            conn.Open();
            Console.WriteLine("連線成功");
            MessageBox.Show("連線成功");
            string sql = "SELECT DATE, START_PRICE, HIGH_PRICE, LOW_PRICE, END_PRICE, VOLUME FROM TAIWAN_DATA_POLARIS WHERE STOCK_NO = @stock_no ORDER BY DATE ";
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@stock_no", "1101");
            MySqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                System.Diagnostics.Debug.WriteLine($"{reader["DATE"]}, " + $"{reader["START_PRICE"]}, " + $"{reader["END_PRICE"]}");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"連線失敗: {ex.Message}");
            MessageBox.Show("連線失敗");
        }
    }
}

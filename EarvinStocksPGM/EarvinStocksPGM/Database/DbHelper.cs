using MySqlConnector;
using System;
using System.Data;

public static class DbHelper
{
    private static string ConnStr = "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";
    
    public class StockData
    {
        public String StockNo { get; set; }
        public long TradeDate { get; set; }
        public decimal StartPrice { get; set; }
        public decimal HighPrice { get; set; }
        public decimal LowPrice { get; set; }
        public decimal EndPrice { get; set; }
        public decimal Volume { get; set; }
    }

    public static MySqlConnection GetConnection()
    {
        using var conn = new MySqlConnection(ConnStr);
        return new MySqlConnection(ConnStr);
    }

    public static StockData[] TestConnectDB()
    {
        string connStr =
    "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";

        using var conn = new MySqlConnection(connStr);

        try
        {
            conn.Open();
//            Console.WriteLine("連線成功");
            MessageBox.Show("連線成功");
            string sql = "SELECT DATE, START_PRICE, HIGH_PRICE, LOW_PRICE, END_PRICE, VOLUME FROM TAIWAN_DATA_POLARIS WHERE STOCK_NO = @stock_no ORDER BY DATE ";
            MySqlCommand cmd = new MySqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("@stock_no", "1101");
            MySqlDataReader reader = cmd.ExecuteReader();
            //while (reader.Read())
            //{
            //    System.Diagnostics.Debug.WriteLine($"{reader["DATE"]}, " + $"{reader["START_PRICE"]}, " + $"{reader["END_PRICE"]}");
            //}
            DataTable dt = new DataTable();
            dt.Load(reader);
            StockData[] sd = new StockData[dt.Rows.Count];
            int i = 0;
            foreach (DataRow row in dt.Rows)
            {
                sd[i] = new StockData();
                sd[i].TradeDate = Convert.ToInt64(row["DATE"]);
                sd[i].StartPrice = Convert.ToDecimal(row["START_PRICE"]);
                sd[i].HighPrice = Convert.ToDecimal(row["HIGH_PRICE"]);
                sd[i].LowPrice = Convert.ToDecimal(row["LOW_PRICE"]);
                sd[i].EndPrice = Convert.ToDecimal(row["END_PRICE"]);
                sd[i].Volume = Convert.ToInt64(row["VOLUME"]);
                
                i++;
            }
            //// print data
            //for (int j = 0; j < sd.Length; j++)
            //{
            //    System.Diagnostics.Debug.WriteLine($"{sd[j].TradeDate}, " + $"{sd[j].StartPrice}, " + $"{sd[j].EndPrice}");
            //}
            System.Diagnostics.Debug.WriteLine("總筆數：" + dt.Rows.Count);

            return sd;

        }
        catch (Exception ex)
        {
            Console.WriteLine($"連線失敗: {ex.Message}");
            MessageBox.Show("連線失敗");
            return null;
        }
    }
}

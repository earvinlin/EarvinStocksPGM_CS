using MySqlConnector;
using System;
using System.Collections.Generic;
using System.Data;
using System.Text;

namespace EarvinStocksPGM.Models
{
    public class StockData
    {
        public String StockNo { get; set; }
        public long TradeDate { get; set; }
        public double StartPrice { get; set; }
        public double HighPrice { get; set; }
        public double LowPrice { get; set; }
        public double EndPrice { get; set; }
        public double Volume { get; set; }
    }

    public static class StockModule
    {
        //public static IndexData[] CalculateAverage()
        //{
        //    return new IndexData[0];
        //}

        public static StockData[] GetStockData(String stockNo)
        {
            string connStr = "Server=localhost;Database=stocksdb;User ID=root;Password=lin32ledi;";

            using var conn = new MySqlConnection(connStr);

            try
            {
                conn.Open();

                string sql = "SELECT DATE, START_PRICE, HIGH_PRICE, LOW_PRICE, END_PRICE, VOLUME FROM TAIWAN_DATA_POLARIS WHERE STOCK_NO = @stock_no ORDER BY DATE ";
                MySqlCommand cmd = new MySqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@stock_no", stockNo);
                MySqlDataReader reader = cmd.ExecuteReader();
                DataTable dt = new DataTable();
                dt.Load(reader);
                StockData[] sd = new StockData[dt.Rows.Count];
                int i = 0;
                foreach (DataRow row in dt.Rows)
                {
                    sd[i] = new StockData();
                    sd[i].TradeDate = Convert.ToInt64(row["DATE"]);
                    sd[i].StartPrice = Convert.ToDouble(row["START_PRICE"]);
                    sd[i].HighPrice = Convert.ToDouble(row["HIGH_PRICE"]);
                    sd[i].LowPrice = Convert.ToDouble(row["LOW_PRICE"]);
                    sd[i].EndPrice = Convert.ToDouble(row["END_PRICE"]);
                    sd[i].Volume = Convert.ToDouble(row["VOLUME"]);

                    i++;
                }
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
}

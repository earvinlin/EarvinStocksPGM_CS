public class StockService
{
	public void QueryStock(string stockNo)
	{
		using var conn =
			DbHelper.GetConnection();

		conn.Open();

		string sql =
			@"SELECT *
              FROM TAIWAN_DATA_POLARIS
              WHERE STOCK_NO = @stock_no";

		using var cmd =
			new MySqlCommand(sql, conn);

		cmd.Parameters.AddWithValue("@stock_no", stockNo);

		using var reader = cmd.ExecuteReader();

		while (reader.Read())
		{
			Console.WriteLine(
				reader["DATE"]);
		}
	}
}
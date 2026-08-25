public class StockData
{
    public DateTime Date { get; set; }
    public decimal EndPrice { get; set; }
}

/*
public List<StockData> QueryStock(string stockNo)
{
    List<StockData> result = new();

    while (reader.Read())
    {
        result.Add(new StockData
        {
            Date = Convert.ToDateTime(reader["DATE"]),
            EndPrice = Convert.ToDecimal(reader["END_PRICE"])
        });
    }

    return result;
}
*/

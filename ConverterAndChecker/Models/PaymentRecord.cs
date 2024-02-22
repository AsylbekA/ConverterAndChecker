namespace ConverterAndChecker.Models;

public class PaymentRecord
{
    public string LastName { get; set; }
    public string FirstName { get; set; }
    public string MiddleName { get; set; }
    public string IIN { get; set; }
    public decimal Amount { get; set; }
    public DateTime Date { get; set; }
}

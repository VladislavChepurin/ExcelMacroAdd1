namespace ExcelMacroAdd.Interfaces
{
    public interface ICurrentVendor
    {      
        string VendorAttribute { get; set; }
      
        string Formula_1 { get; set; }
      
        string Formula_2 { get; set; }
       
        string Formula_3 { get; set; }
              
        int Discont { get; set; }
        
        string Date { get; set; }
    }
}

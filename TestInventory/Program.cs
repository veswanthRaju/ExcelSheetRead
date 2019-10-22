using System;
using InventoyReport;

namespace TestInventory
{
    class Program
    {
        static void Main(string[] args)
        {
            var data = Inventory.ReadExcel(@"C:\Users\walkingtree\Desktop\Sunpharma\InventoryManagement.xlsx", 
                                            "InventoryManagement", 
                                            "Column_Name", "Value");
            Console.WriteLine("Data from the Excel..", data);
        }
    }
}
